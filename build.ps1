$exclude = @("venv", "crisiptusaopaulotabas.zip")
$files = Get-ChildItem -Path . -Exclude $exclude
Compress-Archive -Path $files -DestinationPath "crisiptusaopaulotabas.zip" -Force