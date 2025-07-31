Set objShell = CreateObject("Wscript.Shell")
Set objFSO = CreateObject("Scripting.FileSystemObject")

' Obtener la ruta del script actual (.vbs)
strScriptPath = objFSO.GetParentFolderName(WScript.ScriptFullName)

' Construir la ruta del script PowerShell (asumiendo que está en la misma carpeta)
strPSScriptPath = objFSO.BuildPath(strScriptPath, "extraer_reporte.ps1")

' Construir la ruta del ejecutable datos.exe
strExePath = objFSO.BuildPath(strScriptPath, "datos.exe")

' Ejecutar PowerShell con la ruta relativa (oculto)
objShell.Run "powershell.exe -ExecutionPolicy Bypass -File """ & strPSScriptPath & """", 0, False

' Ejecutar datos.exe (oculto)
objShell.Run """" & strExePath & """", 0, False
