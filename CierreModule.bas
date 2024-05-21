Attribute VB_Name = "CierreModule"

Public Function EscribeLog(cTexto As String)
'Escribe en el archivo de Log de Administracion
Dim ADMIN_LOG As String
ADMIN_LOG = WindowsDirectory
ADMIN_LOG = ADMIN_LOG & "\ADMLOG.SOL"
Open ADMIN_LOG For Append As #1
Print #1, Format(Time & Date, "GENERAL DATE") & Chr(9) & _
    "Usuario : " & nUserNumber & Chr(9) & cTexto
Close #1
End Function

