Set objShell = CreateObject("Wscript.Shell")
objShell.Run "powershell.exe -ExecutionPolicy Bypass -File ""C:\Users\ext.luis.campos\Documents\content\dist\extraer_reporte.ps1""", 0, False
