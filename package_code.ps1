$src = "D:\Dokumen\@ POKJA 2026\V19_Scheduler\WPy64-313110"
$dest = "$src\Govem_Clean_Code"

Write-Host "Creating clean directory..."
New-Item -ItemType Directory -Force -Path $dest | Out-Null

Write-Host "Copying Root Files..."
Copy-Item "$src\*.py" $dest
Copy-Item "$src\*.bat" $dest
Copy-Item "$src\*.txt" $dest -ErrorAction SilentlyContinue

Write-Host "Copying Modules..."
Copy-Item "$src\_Govem" $dest -Recurse
Copy-Item "$src\_Launcher" $dest -Recurse

Write-Host "Removing Secrets and Junk..."
# Remove explicit secrets
Get-ChildItem $dest -Include "credentials.json","token.json","secret_supabase.env","*.csv","*.log","*.pyc" -Recurse | Remove-Item -Force -ErrorAction SilentlyContinue

# Remove __pycache__
Get-ChildItem $dest -Include "__pycache__" -Recurse | Remove-Item -Recurse -Force -ErrorAction SilentlyContinue

# Remove History/Session data inside _Govem (optional but good)
Remove-Item "$dest\_Govem\govem_scheduler.log" -ErrorAction SilentlyContinue
Remove-Item "$dest\_Govem\history.json" -ErrorAction SilentlyContinue
# Keep config.ini? Yes.

Write-Host "Zipping..."
$zipPath = "$src\Govem_Source_Code.zip"
Compress-Archive -Path "$dest\*" -DestinationPath $zipPath -Force

Write-Host "Cleaning up temp folder..."
Remove-Item $dest -Recurse -Force

Write-Host "Done! Zip created at: $zipPath"
