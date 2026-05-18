Get-ChildItem -Path "Z:\04 上海操作中心" -Recurse -ErrorAction SilentlyContinue | Where-Object { $_.Name -like "*Alphaliner*" -or $_.Name -like "*alphaliner*" } | Select-Object FullName | Out-Host
