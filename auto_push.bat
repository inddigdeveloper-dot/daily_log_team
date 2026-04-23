@echo off
cd /d "C:\daily_commit"
git add .
git commit -m "Automated daily local push: %date%"
git push origin master