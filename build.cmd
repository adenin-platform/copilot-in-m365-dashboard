dotnet publish ./adenin.PnPTool -p:PublishProfile=FolderProfile
powershell Compress-Archive .\adenin.PnPTool\bin\Release\net8.0\publish\win-x64 .\build.zip -Update