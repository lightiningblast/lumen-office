OFFICE APPS (WRITE, CELLS, SLIDES)
Word processor (Write), spreadsheet (Cells) and presentation editor (Slides).
Opens and saves .docx, .xlsx, .csv and .pptx files. Runs on your own computer.

INSTALL (Windows 10 / 11)
  1. Unzip LumenOffice-<version>.zip anywhere.
  2. Double-click Setup.cmd. No administrator rights are needed.
     - Write, Cells, Slides and Office Apps are real Windows programs (own window, icon and
       taskbar entry). They use the Microsoft Edge WebView2 runtime that comes with Windows 10/11. No Python needed.
     - Installs to %LOCALAPPDATA%\Programs\Office Apps
     - Adds a "Office Apps" folder to the Start menu, and optional desktop shortcuts.
     - Adds "Open with > Write / Cells / Slides" for .docx .xlsx .csv .pptx files.
       Your default apps are not changed. To make the default for a type: right-click a file >
       Open with > Choose another app > pick > tick "Always use this app".
     - Adds Office Apps to Settings > Apps > Installed apps, where you can uninstall it.
  Windows may show "Windows protected your PC" because the setup is not code-signed: click More info > Run anyway.

OPENING AND SAVING
  Files opened from Explorer save straight back to the same file with Ctrl+S. F12 is Save As.

UPDATES FROM GITHUB
  The apps check the GitHub repository named in update.json once a day and show a banner when a newer
  release exists. "Update now" downloads it, checks its SHA-256 checksum and installs it. You can also run
  "Check for updates" from the Start menu. A copy of the replaced files is kept in
  %LOCALAPPDATA%\Office Apps\previous-version.

  Setting it up (once):
    1. Create a repository on github.com, for example yourname/lumen-office (public, so updates can be
       downloaded without a password).
    2. Put "repo": "yourname/lumen-office" in update.json.
    3. Upload all files of this folder to the repository (or use git), including .github/workflows/release.yml.

  Publishing a new version:
    1. Change "version" in version.json, for example to 1.1.0, and upload it.
    2. Create a tag/release named v1.1.0 (on github.com: Releases > Draft a new release > Choose a tag > v1.1.0 > Publish,
       or "git tag v1.1.0 && git push --tags").
    3. The "Release" GitHub Action builds LumenOffice-1.1.0.zip and its .sha256 file and attaches them
       to the release. Installed copies pick it up within a day, or at once via "Check for updates".
       Note: if you create the release on github.com yourself, the Action adds the files to a second release
       with the same tag only when the tag is new; the simplest path is to push the tag and let the Action
       create the release.

UNINSTALL
  Settings > Apps > Installed apps > Office Apps > Uninstall, or Start menu > Office Apps > Uninstall Office Apps.
  Your documents are never touched.
