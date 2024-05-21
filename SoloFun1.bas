Attribute VB_Name = "SoloFun1"
Public Const GW_HWNDPREV = 3
Declare Function OpenIcon Lib "user32" (ByVal hwnd As Long) As Long
Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Declare Function GetWindow Lib "user32" (ByVal hwnd As Long, ByVal wCmd As Long) As Long
Declare Function SetForegroundWindow Lib "user32" (ByVal hwnd As Long) As Long
Public iISC As Single
Public iISCTransaccion As Single
Public rsISC As New ADODB.Recordset
Public Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
Public Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpstring As Any, ByVal lpFileName As String) As Long

Public Enum GetMaxOrMin
   Getmax = 0
   GetMin = 1
End Enum

'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
'NUEVO MESSAGEBOX AGOSTO 2010
Public ButtonNumber As Integer

Public Enum msgResult
'     vbtimedout = -1
'     vbCancel = 0
'     vbOK = 1
'     vbRetry = 2
     vbYes = 6
     vbNo = 7
'     vbHelp = 5
'     vbAbort = 6
'     vbIgnore = 7
 End Enum
 Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" _
   (ByVal hwnd As Long, ByVal lpOperation As String, _
   ByVal lpFile As String, ByVal lpParameters As String, _
   ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
 '~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
 '~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
Public Function GetContened(nID As Integer) As String
Dim rsContened As ADODB.Recordset

Set rsContened = New ADODB.Recordset
rsContened.Open "SELECT DESCRIP FROM CONTENED WHERE CONTENEDOR = " & nID, msConn, adOpenStatic, adLockOptimistic
If rsContened.EOF Then
    GetContened = "Sin Envase"
Else
    GetContened = rsContened!DESCRIP
End If
rsContened.Close
Set rsContened = Nothing

End Function
Public Function GetSecuritySetting(nUser As Integer, cPantallaID As String) As String
'REGRESA EL NIVEL DE SEGURIDAD DE LA PANTALLA QUE SE ESTA USANDO
Dim rsPantalla As ADODB.Recordset
Dim rsSeguridad As ADODB.Recordset
Dim cSQL As String
Dim nIDPantalla As Integer

On Error GoTo ErrAdm:

Set rsPantalla = New ADODB.Recordset
cSQL = "SELECT ID FROM PANTALLAS WHERE ID_NAME = '" & cPantallaID & "'"
rsPantalla.Open cSQL, msConn, adOpenStatic, adLockOptimistic
If Not rsPantalla.EOF Then
    nIDPantalla = rsPantalla!ID
Else
    EscribeLog "PANTALLA NO ESTA DEFINIDA EN EL SISTEMA: " & cPantallaID
    'ShowMsg "SISTEMA DE SEGURIDAD" & vbCrLf & "PANTALLA NO ESTA DEFINIDA EN EL SISTEMA", vbRed
    'GIVE FULL RIGHTS
    GetSecuritySetting = "CEMV"
    rsPantalla.Close
    Set rsPantalla = Nothing
    Exit Function
End If

cSQL = "SELECT NIVEL FROM SEGURIDAD WHERE NUMERO = " & nUser & " AND ID = " & nIDPantalla
Set rsSeguridad = New ADODB.Recordset
rsSeguridad.Open cSQL, msConn, adOpenStatic, adLockOptimistic

If Not rsSeguridad.EOF Then
    GetSecuritySetting = rsSeguridad!NIVEL
Else
    'POR DEFAULT DARLE TODOS LOS DERECHOS
    EscribeLog "USUARIO y PANTALLA NO ESTAN DEFINIDOS (User) " & nUser & " (Pantalla) " & nIDPantalla
    GetSecuritySetting = "CEMV"
End If

If rsPantalla.State = adStateOpen Then rsPantalla.Close
    Set rsPantalla = Nothing
rsSeguridad.Close
Set rsSeguridad = Nothing
On Error GoTo 0
Exit Function

ErrAdm:
    ShowMsg Err.Number & " (Seguridad.ShowSecurityList) - " & vbCrLf & Err.Description, vbRed
    GetSecuritySetting = "CEMV"
End Function
Public Function GetFromArray(vArray, Optional ByVal GetMaxOrMin As GetMaxOrMin = Getmax) As String
Dim i As Integer
Dim nReturn As Long
Dim nTemp As Long
Dim cOption As String
'*************************************
'GETS REQUESTED VALUE FROM AN ARRAY
'FORMAT:
'nMax = GetFromArray(cArreglo, GetMax)
'nMin = GetFromArray(cArreglo, GetMin)
'21/MAY/2009
'*************************************

cOption = IIf(GetMaxOrMin = Getmax, "MAX", "MIN")

nReturn = vArray(0)
For i = 0 To UBound(vArray, 1)
    nTemp = vArray(i)
    If cOption = "MAX" Then
        If nTemp > nReturn Then nReturn = nTemp
    End If
    If cOption = "MIN" Then
        If nTemp < nReturn Then nReturn = nTemp
    End If
Next
GetFromArray = "'" & nReturn & "'"
End Function
Public Function GetNumber_FromArray(vArray, Optional ByVal GetMaxOrMin As GetMaxOrMin = Getmax) As Long
Dim i As Integer
Dim nReturn As Long
Dim nTemp As Long
Dim cOption As String
'*************************************
'GETS REQUESTED VALUE FROM AN ARRAY
'FORMAT:
'nMax = GetFromArray(cArreglo, GetMax)
'nMin = GetFromArray(cArreglo, GetMin)
'27/OCT/2010
'*************************************

cOption = IIf(GetMaxOrMin = Getmax, "MAX", "MIN")

nReturn = vArray(0)
For i = 0 To UBound(vArray, 1)
    nTemp = vArray(i)
    If cOption = "MAX" Then
        If nTemp > nReturn Then nReturn = nTemp
    End If
    If cOption = "MIN" Then
        If nTemp < nReturn Then nReturn = nTemp
    End If
Next
GetNumber_FromArray = nReturn
End Function

Public Function ExportToExcelOrCSVList(SGSolo As SGGrid)
'INFO: ACTUALIZADO EL 10MAY2012
Dim cSelect As String
cSelect = InputBox("(1) Para Excel, (2) Para Lista", "Exportar datos a una carpeta de su selección", "1")

On Error GoTo ErrAdm:

Select Case cSelect
    Case "1"        'EXCEL
    
        'Screen.MousePointer = vbHourglass
        'SGSolo.ExportData App.Path & "\Listado.xls", sgFormatExcel, sgExportOverwrite + sgExportFieldNames
        'Screen.MousePointer = vbDefault
        'ShowMsg "Se exporto datos a " & App.Path & "\Listado.xls", , , vbOKOnly
        
        MainMant.dlgDialog.CancelError = True
        With MainMant.dlgDialog
            .Filter = "(*.XLS)|*.xls"
            .DialogTitle = "Guardar Listado en formato EXCEL"
            .InitDir = App.Path
            .FileName = "Listado.xls"
            .ShowSave
        End With
        
        If MainMant.dlgDialog.FileName <> "" Then
            'If Right(UCase(dlgDialog.FileName), Len("\SOLO.MDB")) = "\SOLO.MDB" Then
                MainMant.dlgDialog.InitDir = Dir(MainMant.dlgDialog.FileName)
                SGSolo.ExportData MainMant.dlgDialog.FileName, sgFormatExcel, sgExportFieldNames + sgExportOverwrite
                'DD_PEDDETALLE.ExportData App.Path & "\Listado.xls", sgFormatExcel, sgExportFieldNames + sgExportOverwrite
                'INFO: 16FEB2013 / OPCION DE launch the file
                If ShowMsg("Datos Exportados: " & MainMant.dlgDialog.FileName & vbCrLf & vbCrLf & _
                                      "¿ Abrir Archivo ?", vbYellow, vbBlue, vbYesNo) = vbYes Then
                    ShellExecute MainMant.hwnd, "", MainMant.dlgDialog.FileName, "", "", vbNormal
                End If
        End If
        On Error GoTo 0

    Case "2"        'CSV
        'Screen.MousePointer = vbHourglass
        'SGSolo.ExportData App.Path & "\Listado.csv", sgFormatCharSeparatedValue, sgExportOverwrite
        'Screen.MousePointer = vbDefault
        'ShowMsg "Se exporto datos a " & App.Path & "\Listado.csv", , , vbOKOnly
        MainMant.dlgDialog.CancelError = True
        With MainMant.dlgDialog
            .Filter = "(*.CSV)|*.csv"
            .DialogTitle = "Guardar Listado en formato CSV de Texto"
            .InitDir = App.Path
            .FileName = "Listado.csv"
            .ShowSave
        End With
        
        If MainMant.dlgDialog.FileName <> "" Then
                MainMant.dlgDialog.InitDir = Dir(MainMant.dlgDialog.FileName)
                SGSolo.ExportData MainMant.dlgDialog.FileName, sgFormatCharSeparatedValue, sgExportFieldNames + sgExportOverwrite
                
                ShowMsg "Se exportaron los datos a " & MainMant.dlgDialog.FileName, vbYellow, vbBlue
        End If
        
    Case Else
    
End Select

On Error GoTo 0
Exit Function

ErrAdm:
If Err.Number = 32755 Then
    ShowMsg "Se Canceló la Operación de Guardar Datos", vbRed, vbYellow, vbOKOnly
Else
    ShowMsg Err.Number & " - " & Err.Description, vbRed, vbYellow, vbOKOnly
End If
End Function

Public Function GetOption(opts As Object) As Integer
'INFO: HOWTO Determine Selected Control from Array of Option Buttons
On Error GoTo GetOptionFail
Dim opt As OptionButton
For Each opt In opts
    If opt.value Then
        GetOption = opt.Index
        Exit Function
    End If
Next

GetOptionFail:
GetOption = -1

End Function
Public Function GetDia(cFecha As String) As String
Dim oDate As Date
 '#12/31/92#
oDate = CDate(CInt(Right(cFecha, 2)) & "/" & CInt(Mid(cFecha, 5, 2)) & "/" & Mid(cFecha, 3, 2))
GetDia = UCase(WeekdayName(Weekday(oDate)))
End Function
Public Function GetFecha(cFecha As String) As String
Dim cMes As String
Dim cNombreMes As String

cMes = Mid(cFecha, 5, 2)
Select Case cMes
    Case "01"
        cNombreMes = "ENE"
    Case "02"
        cNombreMes = "FEB"
    Case "03"
        cNombreMes = "MAR"
    Case "04"
        cNombreMes = "ABR"
    Case "05"
        cNombreMes = "MAY"
    Case "06"
        cNombreMes = "JUN"
    Case "07"
        cNombreMes = "JUL"
    Case "08"
        cNombreMes = "AGO"
    Case "09"
        cNombreMes = "SEP"
    Case "10"
        cNombreMes = "OCT"
    Case "11"
        cNombreMes = "NOV"
    Case "12"
        cNombreMes = "DIC"
    Case Else
        cNombreMes = "Err"
End Select

GetFecha = Right(cFecha, 2) & "/" & cNombreMes & "/" & Left(cFecha, 4)
End Function
Public Function GetMes(cFecha As String) As String
Dim oDate As Date
 '#12/31/92#
oDate = CDate(CInt(Right(cFecha, 2)) & "/" & CInt(Mid(cFecha, 5, 2)) & "/" & Mid(cFecha, 3, 2))
GetMes = UCase(Format(oDate, "mmm"))

End Function
Public Function GetFromINI(Section As String, Key As String, Directory As String) As String
Dim strBuffer As String

On Error GoTo FileError:
    strBuffer = String(750, Chr(0))
    Key$ = LCase$(Key$)
    GetFromINI$ = Left(strBuffer, GetPrivateProfileString(Section$, ByVal Key$, "", strBuffer, Len(strBuffer), Directory$))
On Error GoTo 0
Exit Function

FileError:
    MsgBox Err.Number & ": NO SE ENCUENTRA ARCHIVO DE INICIALIZACION", vbCritical, "ERROR AL INICIAR"
    Resume Next
End Function
Public Sub WriteToINI(Section As String, Key As String, KeyValue As String, Directory As String)
On Error GoTo FileError:
    Call WritePrivateProfileString(Section$, UCase$(Key$), KeyValue$, Directory$)
On Error GoTo 0
Exit Sub

FileError:
    MsgBox Err.Number & ": NO SE ENCUENTRA ARCHIVO DE INICIALIZACION", vbCritical, "ERROR AL INICIAR"
    Resume Next
End Sub
Public Sub ActivatePrevInstance()

Dim OldTitle As String
Dim PrevHndl As Long
Dim result As Long

'Save the title of the application.
OldTitle = App.Title

'Rename the title of this application so FindWindow
'will not find this application instance.
App.Title = "Instancia de App No Deseada"

'Attempt to get window handle using VB4 class name.
PrevHndl = FindWindow("ThunderRTMain", OldTitle)

'Check for no success.
If PrevHndl = 0 Then
   'Attempt to get window handle using VB5 class name.
   PrevHndl = FindWindow("ThunderRT5Main", OldTitle)
End If

'Check if found
If PrevHndl = 0 Then
    'Attempt to get window handle using VB6 class name
    PrevHndl = FindWindow("ThunderRT6Main", OldTitle)
End If

'Check if found
If PrevHndl = 0 Then
   'No previous instance found.
   Exit Sub
End If

'Get handle to previous window.
PrevHndl = GetWindow(PrevHndl, GW_HWNDPREV)

'Restore the program.
result = OpenIcon(PrevHndl)

'Activate the application.
result = SetForegroundWindow(PrevHndl)

'End the application.
End
End Sub
Public Function SoloAvg(ParamArray x())
On Error Resume Next
    s = 0
    For i = LBound(x) To UBound(x)
        s = s + x(i)
    Next
    Avg = s / (UBound(x) - LBound(x) + 1)
End Function

Public Sub OpenImpresora()
'RC = Sys_Pos.Coptr1.Open("TM-U950P")
'ESTOY USANDO EL Logical Device Name QUE ES MAS PORTABLE, YA QUE}
'EL DEVICE NAME ME ESTABA DANDO PROBLEMAS
'BAJANDO LA VELOCIDAD A 2400 BPS, YA QUE EXISTE LA POSIBILIDAD
'DE PERDIDA DE DATOS CUANDO ES MAS ALTA

rc = Sys_Pos.Coptr1.Open(GetFromINI("SoloPosDisp", "Facturacion", DATA_PATH & "soloini.ini"))
If rc <> OPOS_SUCCESS Then
    MsgBox "LA GAVETA DE DINERO NO ESTA CONECTADA", vbCritical, "ENCIENDA LA IMPRESORA"
End If

rc = Sys_Pos.ImpresoraBarra.Open(GetFromINI("SoloPosDisp", "Bar", App.Path & "\soloini.ini"))
If rc <> OPOS_SUCCESS Then
    MsgBox "LA IMPRESORA DE BARRA NO ESTA CONECTADA", vbCritical, "ENCIENDA LA IMPRESORA"
End If
End Sub
Public Sub OpenImpresoraCocina()
rc = Sys_Pos.ImprCocina.Open(GetFromINI("SoloPosDisp", "Cocina", App.Path & "\soloini.ini"))
If rc <> OPOS_SUCCESS Then
    MsgBox "FAVOR RECUERDE ENCENDER LA IMPRESORA DE COCINA", vbCritical, "ENCIENDA LA IMPRESORA DE COCINA"
End If
End Sub
Public Sub OpenCajaRegistradora()
'rc = Sys_Pos.Cocash1.Open(NOM_GAV_DINERO = GetFromINI("SoloPosDisp", "Gaveta", App.Path & "\soloini.ini"))
rc = Sys_Pos.Cocash1.Open(GetFromINI("SoloPosDisp", "Gaveta", App.Path & "\soloini.ini"))
If rc <> OPOS_SUCCESS Then
    MsgBox "LA GAVETA DE DINERO NO ESTA CONECTADA", vbCritical, "ENCIENDA LA IMPRESORA/REVISE CONEXION"
End If
End Sub

Public Sub ClaimImpresora()
Dim nIntentos As Integer
'OtraVez:
rc = Sys_Pos.Coptr1.Claim(500) 'COMANDO CLAIM ES LO QUE MAS TARDA AL ABRIR IMPRESORA
If rc <> OPOS_SUCCESS Then
    MsgBox "LA IMPRESORA DE FACTURACION NO ESTA CONECTADA o NO ESTA ENCENDIDA", vbInformation, "ERROR NUMERO : " & rc
End If

rc = Sys_Pos.ImpresoraBarra.Claim(500) 'COMANDO CLAIM ES LO QUE MAS TARDA AL ABRIR IMPRESORA
If rc <> OPOS_SUCCESS Then
    MsgBox "LA IMPRESORA DEL BAR NO ESTA CONECTADA o NO ESTA ENCENDIDA", vbInformation, "ERROR NUMERO : " & rc
End If
End Sub

Public Sub ClaimImpresoraCocina()
rc = Sys_Pos.ImprCocina.Claim(500) 'COMANDO CLAIM ES LO QUE MAS TARDA AL ABRIR IMPRESORA
If rc <> OPOS_SUCCESS Then MsgBox "LA IMPRESORA DE COCINA NO ESTA CONECTADA o NO ESTA ENCENDIDA", vbInformation, "ERROR NUMERO : " & rc
End Sub
Public Sub GetISC()
On Error Resume Next
rsISC.Open "SELECT ISC FROM ISC", msConn, adOpenStatic, adLockOptimistic
iISC = rsISC!ISC
rsISC.Close
On Error GoTo 0
End Sub
Public Sub PutISC(nISC As Single)
On Error Resume Next
rsISC.Open "SELECT DIARIO FROM ISC", msConn, adOpenDynamic, adLockOptimistic
rsISC!DIARIO = rsISC!DIARIO + nISC
rsISC.Update
rsISC.Close
On Error GoTo 0
End Sub
Public Function RemoveNull(cCadena As Variant) As String
'18/8/2005
Dim tempCadena As String
Dim i As Integer

For i = 1 To Len(cCadena)
    If Asc(Mid(cCadena, i, 1)) < 32 Then
    Else
        tempCadena = tempCadena & Mid(cCadena, i, 1)
    End If
Next
RemoveNull = tempCadena
End Function
Public Function ShowMsg(cMsg As String, Optional oFontColor As Long, Optional oBackColor As Long, Optional Botones As Integer) As msgResult
'MENSAJE_DEL_SISTEMA = cMsg
' BOTONES: vbYes Y vbYesNo
Dim TheForm As Form

Load Mensaje

Select Case Botones
    Case vbOKOnly
        Mensaje.cmdAceptar.Enabled = True
    Case vbYesNo
        Mensaje.cmdAceptar.Enabled = True
        Mensaje.cmdCancelar.Enabled = True
        Mensaje.cmdAceptar.Visible = True
        Mensaje.cmdCancelar.Visible = True
    Case Else
        Mensaje.cmdAceptar.Enabled = True
End Select

If oBackColor = 0 Then oBackColor = Mensaje.BackColor
If oBackColor = oFontColor Then oBackColor = Mensaje.BackColor

If oFontColor = vbWhite Or oFontColor = 0 Then
Else
    Mensaje.BackColor = oBackColor
    Mensaje.lbMensaje.BackColor = oBackColor
    Mensaje.lbMensaje.ForeColor = oFontColor
End If

Mensaje.lbMensaje.Caption = cMsg
Mensaje.Show 1


ShowMsg = ButtonNumber
Unload Mensaje
End Function

'---------------------------------------------------------------------------------------
' Procedure : GetDeptoNombre
' Author    : hsequeira
' Date      : 21/01/2023
' Purpose   : GET ID DEPTO FROM STRING
'---------------------------------------------------------------------------------------
'
Public Function GetDeptoNombre(cCadena As String) As Long
Dim rsFunDEPTO As ADODB.Recordset
Dim cSQL As String

Set rsFunDEPTO = New ADODB.Recordset
cSQL = "SELECT CODIGO FROM DEPTO WHERE DESCRIP = '" & cCadena & "'"
rsFunDEPTO.Open cSQL, msConn, adOpenStatic, adLockOptimistic
If Not rsFunDEPTO.EOF Then
    GetDeptoNombre = rsFunDEPTO!CODIGO
Else
    GetDeptoNombre = 0
End If
rsFunDEPTO.Close
Set rsFunDEPTO = Nothing
End Function
