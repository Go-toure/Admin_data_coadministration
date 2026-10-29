@echo off
cd /d "%~dp0"
echo Publish the latest corrected admin data to SharePoint
echo.
python publish_admin_data.py %*
echo.
pause
