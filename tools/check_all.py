import glob, zipfile, xml.dom.minidom
bad = 0
for f in sorted(glob.glob('_testout/*.*')):
    if not f.endswith(('.docx', '.xlsx', '.pptx')):
        continue
    z = zipfile.ZipFile(f)
    names = z.namelist()
    problems = []
    for n in names:
        if n.endswith(('.xml', '.rels', '.vml')):
            try:
                xml.dom.minidom.parseString(z.read(n))
            except Exception as e:
                problems.append((n, str(e)[:60]))
    # every relationship target must exist
    for n in names:
        if n.endswith('.rels'):
            base = n.replace('_rels/', '').rsplit('.rels', 1)[0]
            d = base.rsplit('/', 1)[0] if '/' in base else ''
            dom = xml.dom.minidom.parseString(z.read(n))
            for r in dom.getElementsByTagName('Relationship'):
                if r.getAttribute('TargetMode') == 'External':
                    continue
                t = r.getAttribute('Target')
                import posixpath
                p = t.lstrip('/') if t.startswith('/') else posixpath.normpath(posixpath.join(d, t))
                if p not in names:
                    problems.append((n, 'missing target ' + p))
    print(('OK   ' if not problems else 'BAD  ') + f, len(names), problems[:3])
    bad += bool(problems)
print('files with problems:', bad)
