Invoke-PS2EXE '.\Preview Cricut Project.ps1' '.\Preview Cricut Project.exe' -noConsole -iconFile .\cricut.ico -title 'Preview Cricut Project' -noOutput -noError
signtool sign /a /fd SHA256 /tr 'http://timestamp.digicert.com' /td SHA256 '.\Preview Cricut Project.exe'
