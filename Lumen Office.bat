@echo off
rem Opens Lumen Office as a standalone app window (Edge or Chrome "app mode").
set "HERE=%~dp0"
set "URL=file:///%HERE:\=/%index.html"
where msedge >nul 2>nul && (start "" msedge --app="%URL%" & exit /b)
where chrome >nul 2>nul && (start "" chrome --app="%URL%" & exit /b)
start "" "%HERE%index.html"
