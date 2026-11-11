/* Minimal ZIP reader/writer (STORE + DEFLATE) built on the browser's native
   CompressionStream / DecompressionStream. No dependencies. */
(function (global) {
  'use strict';
  const enc = new TextEncoder();
  const dec = new TextDecoder('utf-8');

  const CRC = (() => {
    const t = new Uint32Array(256);
    for (let n = 0; n < 256; n++) {
      let c = n;
      for (let k = 0; k < 8; k++) c = (c & 1) ? (0xEDB88320 ^ (c >>> 1)) : (c >>> 1);
      t[n] = c >>> 0;
    }
    return t;
  })();

  function crc32(b) {
    let c = 0xFFFFFFFF;
    for (let i = 0; i < b.length; i++) c = CRC[(c ^ b[i]) & 255] ^ (c >>> 8);
    return (c ^ 0xFFFFFFFF) >>> 0;
  }

  async function pipe(data, stream) {
    const w = stream.writable.getWriter();
    const p = w.write(data).then(() => w.close());
    const out = await new Response(stream.readable).arrayBuffer();
    await p;
    return new Uint8Array(out);
  }
  const inflate = (d) => pipe(d, new DecompressionStream('deflate-raw'));
  const deflate = (d) => pipe(d, new CompressionStream('deflate-raw'));

  /* ---------- Reader ---------- */
  async function read(input) {
    const buf = input instanceof Uint8Array ? input : new Uint8Array(input);
    const dv = new DataView(buf.buffer, buf.byteOffset, buf.byteLength);
    let eocd = -1;
    for (let i = buf.length - 22; i >= Math.max(0, buf.length - 65558); i--) {
      if (dv.getUint32(i, true) === 0x06054b50) { eocd = i; break; }
    }
    if (eocd < 0) throw new Error('This file is not a valid Office (ZIP) document.');
    const count = dv.getUint16(eocd + 10, true);
    let p = dv.getUint32(eocd + 16, true);
    const entries = new Map();
    for (let i = 0; i < count; i++) {
      if (dv.getUint32(p, true) !== 0x02014b50) break;
      const method = dv.getUint16(p + 10, true);
      const csize = dv.getUint32(p + 20, true);
      const nl = dv.getUint16(p + 28, true);
      const el = dv.getUint16(p + 30, true);
      const cl = dv.getUint16(p + 32, true);
      const off = dv.getUint32(p + 42, true);
      const name = dec.decode(buf.subarray(p + 46, p + 46 + nl));
      entries.set(name, { method, csize, off });
      p += 46 + nl + el + cl;
    }
    const norm = (n) => String(n).replace(/^\/+/, '');
    const cache = new Map();
    return {
      names: () => [...entries.keys()],
      has: (n) => entries.has(norm(n)),
      async bytes(n) {
        n = norm(n);
        const e = entries.get(n);
        if (!e) throw new Error('Missing part in file: ' + n);
        if (cache.has(n)) return cache.get(n);
        const lnl = dv.getUint16(e.off + 26, true);
        const lel = dv.getUint16(e.off + 28, true);
        const start = e.off + 30 + lnl + lel;
        const raw = buf.subarray(start, start + e.csize);
        let out;
        if (e.method === 0) out = raw.slice();
        else if (e.method === 8) out = await inflate(raw);
        else throw new Error('Unsupported ZIP compression method ' + e.method);
        cache.set(n, out);
        return out;
      },
      async text(n) { return dec.decode(await this.bytes(n)); },
      async textOrNull(n) { return this.has(n) ? this.text(n) : null; },
    };
  }

  /* ---------- Writer ---------- */
  function writer() {
    const files = [];
    return {
      add(name, data) {
        if (typeof data === 'string') data = enc.encode(data);
        files.push({ name, data });
      },
      async toBlob(mime) {
        const parts = [];
        const central = [];
        let offset = 0;
        const now = new Date();
        const dtime = (now.getHours() << 11) | (now.getMinutes() << 5) | (now.getSeconds() >> 1);
        const ddate = ((now.getFullYear() - 1980) << 9) | ((now.getMonth() + 1) << 5) | now.getDate();
        for (const f of files) {
          const nameB = enc.encode(f.name);
          const crc = crc32(f.data);
          let method = 0, body = f.data;
          if (f.data.length > 64) {
            const d = await deflate(f.data);
            if (d.length < f.data.length) { method = 8; body = d; }
          }
          const lh = new Uint8Array(30 + nameB.length);
          const l = new DataView(lh.buffer);
          l.setUint32(0, 0x04034b50, true);
          l.setUint16(4, 20, true);
          l.setUint16(6, 0x0800, true);
          l.setUint16(8, method, true);
          l.setUint16(10, dtime, true);
          l.setUint16(12, ddate, true);
          l.setUint32(14, crc, true);
          l.setUint32(18, body.length, true);
          l.setUint32(22, f.data.length, true);
          l.setUint16(26, nameB.length, true);
          l.setUint16(28, 0, true);
          lh.set(nameB, 30);
          const ch = new Uint8Array(46 + nameB.length);
          const c = new DataView(ch.buffer);
          c.setUint32(0, 0x02014b50, true);
          c.setUint16(4, 20, true);
          c.setUint16(6, 20, true);
          c.setUint16(8, 0x0800, true);
          c.setUint16(10, method, true);
          c.setUint16(12, dtime, true);
          c.setUint16(14, ddate, true);
          c.setUint32(16, crc, true);
          c.setUint32(20, body.length, true);
          c.setUint32(24, f.data.length, true);
          c.setUint16(28, nameB.length, true);
          c.setUint32(42, offset, true);
          ch.set(nameB, 46);
          central.push(ch);
          parts.push(lh, body);
          offset += lh.length + body.length;
        }
        let cdSize = 0;
        for (const c of central) { parts.push(c); cdSize += c.length; }
        const end = new Uint8Array(22);
        const e = new DataView(end.buffer);
        e.setUint32(0, 0x06054b50, true);
        e.setUint16(8, files.length, true);
        e.setUint16(10, files.length, true);
        e.setUint32(12, cdSize, true);
        e.setUint32(16, offset, true);
        parts.push(end);
        return new Blob(parts, { type: mime || 'application/zip' });
      },
    };
  }

  global.Zip = { read, writer, crc32, enc, dec };
})(window);
