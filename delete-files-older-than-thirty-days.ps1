Get-ChildItem -Path "C:\Replace\This\With\Your\Path" -Recurse -File | Where-Object { $_.LastWriteTime -lt (Get-Date).AddDays(-30) } | Remove-Item -Force
