dotnet publish ./adenin.PnpTool -p:PublishProfile=FolderProfile
powershell Compress-Archive .\adenin.PnpTool\bin\Release\net8.0\publish\win-x64 .\build.zip -Update
