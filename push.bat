@echo off
echo Pushing code to Google Apps Script...
node .\node_modules\@google\clasp\build\src\index.js push -f
echo.
echo Push completed!
pause