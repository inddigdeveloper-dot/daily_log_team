@echo off
cd /d "C:\daily_commit"
"C:\Program Files\Git\bin\git.exe" add task_hk.txt
"C:\Program Files\Git\bin\git.exe" commit -m "Daily update: %date%"
"C:\Program Files\Git\bin\git.exe" push origin master
git add .
git commit -m "Daily update: %date%"
git push origin master
:: Remove the pause command for automation