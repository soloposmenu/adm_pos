Attribute VB_Name = "Modulo00"
Option Explicit
Public msConn As ADODB.Connection
Public msPED As ADODB.Connection
'-----------  PARA BASE DE DATOS LOCAL  ----------------------
''''''''''''''''''''-------- Public msConnLoc As Connection    'Conexion BDTRANSLOCAL
'Public rsLoc00 As Recordset
'------------------------------------------------------------
Public rs As Recordset
Public rs00 As Recordset
Public rs01 As Recordset
Public rs02 As Recordset
Public rs03 As Recordset
Public rs04 As Recordset
Public rs05 As Recordset
Public rs06 As Recordset
Public rs07 As Recordset
'Public rs08 As Recordset
'Public rs09 As Recordset
'---------------------
''Public rp01 As Recordset
''Public rp02 As Recordset
''Public rp03 As Recordset
'''---------------------
Public ra01 As Recordset
Public ra02 As ADODB.Recordset
Public ra03 As ADODB.Recordset
Public ra04 As ADODB.Recordset
Public ra05 As ADODB.Recordset    'PLU
''Public ra06 As Recordset    'contend_02
Public ra07 As ADODB.Recordset    'Info.Cont/PLU
Public rstmp1 As ADODB.Recordset
Public raadm As ADODB.Recordset   'Datos de la empresa
Public rsPlu As Recordset
''Public rsPlatos As Recordset    'MESAS CON CUENTAS
'---------------------
Public npNumCaj As Integer  'Numero de Cajero
Public cNomCaj As String
Public nMesero As Integer   'Numero de Mesero
Public cCaja As Integer        'Numero de Caja
Public nMesa As Integer       'Numero de Mesa
Public nCambio As Currency  'Vuelto a Entregar
Public CajLin As Integer ' Numero de Linea de producto
'----- Variables para Msgbox
Public BoxResp, BoxPreg, BoxTit
Public nCantidad As Long 'Cantidad Marcada
Public SBTot As Currency
Public nDesc01 As Long  'Descuento Predeterminado
Public nDesc02 As Long
Public nFlag As Integer 'Bandera para MesaBarra
Public nMesaBarra As Integer    'Mesa que no pide Mesero
Public OkAnul As Integer    'Anulacion Producto
Public OKGlobal As Integer  'Descuento Global
Public OKProp As Integer    '10% de propina
Public OKDesc As Integer    'OK Aplic. Descuento
Public OKCancelar As Integer
Public txtInfo As String
Public DescPreCta As Single
Public nCta As Integer  'Numero de Cuenta Activa. Si no hay cuentas
                        'el Default es 0
Public nCliNum As Integer   'NUMERO DE CLIENTE
Public nVeriSalida As Integer
Public Const COCINA_01 = 3
'INFO: 07OCT2011
Public Const COCINA_02 = 3          '===>> CAMBIANDO DE 6 A 3
Public Const BARRA_01 = 2           '===>> CAMBIANDO DE 4 A 2
Public Const BARRA_02 = 2           '===>> CAMBIANDO DE 5 A 2
Public TXT_OPEN_DEPT As String
Public ValOpenDept As Single
'-----------------------
Public cNomMesero As String
Public cNombreCliente As String
Public cRUCCliente As String        'no estaba definido. SE DEFINE EN 17OCT2013
Public Declare Function RegOpenKeyEx Lib "advapi32" Alias "RegOpenKeyExA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal dwReserved As Long, ByVal samDesired As Long, phkResult As Long) As Long
Public Declare Function RegQueryValueEx Lib "advapi32" Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName$, ByVal lpdwReserved As Long, lpdwType As Long, lpData As Any, lpcbData As Long) As Long
Public Declare Function RegCloseKey Lib "advapi32" (ByVal hKey As Long) As Long
Public Const HKEY_CURRENT_CONFIG As Long = &H80000005
Public nMulti As Integer
Public MAX_DESCUENTO As Integer
Public PRINT_OPC As Boolean
'IMPRESION DE TIQUETES
Public TICKET_NUM As Long
Public TICKET_OK As Boolean
'INFO: OPOS DEVICES
Public NOM_PRN_FACTURA As String
Public NOM_PRN_COCINA As String
Public NOM_PRN_BAR As String
Public NOM_GAV_DINERO As String
Public nSeleccionCocina As Integer
'========================
'INFO: APLICA LA RESTRICCION DE SEGURIDAD, ASI, PREGUNTA LA CLAVE DE UN SUPERVISOR PARA DAR DESCUENTOS
'17AGO2014
Public HaySeguridad As Boolean
'INFO:REVISA LA BASE DE DATOS, PARA VER SI HAY PEDIDOS DE DOMICILO
'12ENE2017
Public nHayDomicilio As Integer
'22NOV2017
Public bLogo As Boolean

Public Function GetUnidConsumo(nUnidConsumo As Long, rsTABLA As ADODB.Recordset) As String
On Error Resume Next
rsTABLA.MoveFirst
On Error GoTo 0
rsTABLA.Find "ID = " & nUnidConsumo
If Not rsTABLA.EOF Then
    GetUnidConsumo = Left(rsTABLA!DESCRIP, 3)
Else
    GetUnidConsumo = ""
End If
End Function

Public Function FormatPrecio(precio As Single) As Currency
FormatPrecio = FormatCurrency(IIf(IsNull(precio), 0#, precio), 2)
End Function
Public Function FormatTexto(texto As String, Largo As Integer) As String
If Largo <= Len(texto) Then
   FormatTexto = Mid$(texto, 1, Largo)
Else
   FormatTexto = texto + Space(Largo - Len(texto))
End If
End Function
'Public Function OLD_EscribeLog(cTexto As String)
''Escribe en el archivo de Log de Administracion
''ELIMINADO EL 04FEB2010
'On Error Resume Next
'Open ADMIN_LOG For Append As #1
'Print #1, Format(Time & Date, "GENERAL DATE") & Chr(9) & _
'    "Usuario : " & rs!numero & Chr(9) & cTexto
'Close #1
'On Error GoTo 0
'End Function

Public Function OLD_EscribeLog(cTexto As String)
'FROM SOLOMIX. 16NOV2023
'Escribe en el archivo de Log de Administracion
Dim nFreefile As Byte
nFreefile = FreeFile()
'Open Environ("WINDIR") & "\ADMLOG.SOL" For Append As #nFreefile
Open App.Path & "\ADMLOG.SOL" For Append As #nFreefile
Print #nFreefile, Format(Time & Date, "GENERAL DATE") & Chr(9) & cTexto
    '"Usuario : " & rs!numero & Chr(9) & cTexto
Close #nFreefile
End Function

Public Function EscribeLog(cTexto As String)
'Escribe en el archivo de Log de Administracion
'INFO: 09/AGO/2007 CREANDO COMO ESTRUCTURA DE DBMS, YA QUE
'WEPOS ES COMO XP EN RELACION A LOS DERECHOS DE
'ESCRITURA SOBRE LOS FOLDERS

Dim cSQL As String, cSQL2 As String
Dim cTexto2 As String

cTexto2 = ""

If Left(cTexto, 37) = "USUARIO y PANTALLA NO ESTAN DEFINIDOS" Then
    'INFO: 01FEB2011 -  ESTO NO INTERESA QUE SALGA EN EL LOG
    Exit Function
End If


If Len(Trim(cTexto)) > 120 Then
    cTexto2 = Mid(cTexto, 121, 120)
    cTexto = Left(cTexto, 120)
End If

cSQL = "INSERT INTO LOG (FECHA, HORA, DESCRIPCION) VALUES ('"
cSQL = cSQL & Format(Date, "YYYYMMDD") & "','"
cSQL = cSQL & Format(Time, "HH:MM:SS") & "','"
cSQL = cSQL & cTexto & "')"

On Error GoTo ErrorLog:

'INFO: 02FEB2011
MesasPED "OPEN"
msPED.Execute cSQL
MesasPED "CLOSE"


If cTexto2 = "" Then
Else
    'INFO: 02ENE2012
    cSQL = "INSERT INTO LOG (FECHA, HORA, DESCRIPCION) VALUES ('"
    cSQL = cSQL & Format(Date, "YYYYMMDD") & "','"
    cSQL = cSQL & Format(Time, "HH:MM:SS") & "','"
    cSQL = cSQL & cTexto2 & "')"
    
    On Error GoTo ErrorLog:
    
    MesasPED "OPEN"
    msPED.Execute cSQL
    MesasPED "CLOSE"
End If

On Error GoTo 0
Exit Function

ErrorLog:
If Err.Number = -2147217865 Then
    cSQL2 = "CREATE TABLE LOG "
    cSQL2 = cSQL2 & "(FECHA TEXT(26), HORA TEXT(8), "
    cSQL2 = cSQL2 & "DESCRIPCION TEXT(120)) "
    msPED.Execute cSQL2
    
    cSQL2 = ""
    cSQL2 = "CREATE INDEX FECHA_HORA ON LOG (FECHA,HORA) "
    msPED.Execute cSQL2
    
    Resume
Else
    'INFO: NO MOSTRAR ERROR CORREGIR LUEGO
    'MsgBox Err.Number & " - " & Err.Description & vbCrLf & cSQL & vbCrLf & "Longitud: " & Len(cSQL), vbCritical, "ERROR EN LOG FILE"
End If
End Function


Public Function IsFormLoaded(FormToCheck As Form) As Integer
    Dim Y As Integer

    For Y = 0 To Forms.Count - 1
        If Forms(Y) Is FormToCheck Then
            IsFormLoaded = True
            Exit Function
        End If
    Next
    IsFormLoaded = False
End Function
'''Public Function Seleccion_Impresora_Default()
''''INFO: OCT-2009, YA NO ES NECESARIO
''''SELECCION DE LA IMPRESORA SOLO SOFTWARE/PARA IMPRESION DE REPORTES
'''Dim x As Printer
'''Dim nSolo As Integer
'''
'''On Error GoTo Fallo:
'''nSolo = 0
'''For Each x In Printers
'''    If x.DeviceName = DEFAULT_PRINTER Then
'''        Set Printer = x
'''        Exit For
'''    End If
'''    nSolo = nSolo + 1
'''Next
'''On Error GoTo 0
'''Exit Function
'''
'''Fallo:
'''    MsgBox "Su impresora no es la Correcta. Contacte a " & App.CompanyName & _
'''           " Podrá continuar la aplicacion, pero posiblemente tendrá Errores de Impresión", vbCritical, BoxTit
'''    Exit Function
'''
'''End Function
''''Public Sub Seleccion_Impresora()
'''''INFO: OCT-2009, YA NO ES NECESARIO EN ADMINISTRACION
'''''SELECCION DE LA IMPRESORA SOLO SOFTWARE/PARA IMPRESION DE REPORTES
''''Dim x As Printer
''''Dim cSolo As String
''''Dim nSolo As Integer
''''
''''On Error GoTo Fallo:
''''cSolo = "SOLOPRN"
''''nSolo = 0
''''For Each x In Printers
''''    If x.DeviceName = cSolo Then
''''        Set Printer = Printers(nSolo)
''''        Exit For
''''    End If
''''    nSolo = nSolo + 1
''''Next
''''On Error GoTo 0
''''Exit Sub
''''
''''Fallo:
''''    MsgBox "Su impresora no es la Correcta. Contacte a " & App.CompanyName & _
''''           " Podrá continuar la aplicacion, pero posiblemente tendrá Errores de Impresión", vbCritical, BoxTit
''''    Exit Sub
''''End Sub
'''Public Sub Seleccion_Impresora_TMU950()
''''SELECCION DE LA IMPRESORA SOLO SOFTWARE/PARA IMPRESION DE REPORTES
'''Dim x As Printer
'''Dim cSolo As String
'''Dim nSolo As Integer
'''
'''On Error GoTo Fallo:
'''cSolo = "SOLO_950"
'''nSolo = 0
'''For Each x In Printers
'''    If x.DeviceName = cSolo Then
'''        Set Printer = Printers(nSolo)
'''        Exit For
'''    End If
'''    nSolo = nSolo + 1
'''Next
'''On Error GoTo 0
'''Exit Sub
'''
'''Fallo:
'''    MsgBox "Su impresora no es la Correcta. Contacte a " & App.CompanyName & _
'''           " Podrá continuar la aplicacion, pero posiblemente tendrá Errores de Impresión", vbCritical, BoxTit
'''    Exit Sub
'''End Sub
Public Sub BorraTransHost()

Dim objCat As New ADOX.Catalog
Dim objTabla As ADOX.Table

objCat.ActiveConnection = msConn
msConn.BeginTrans
For Each objTabla In objCat.Tables
    If objTabla.Type = "TABLE" And Mid(objTabla.Name, 1, 1) = "T" Then
        
        On Error GoTo ErrDeleteHost:
        msConn.Execute "DELETE FROM " & objTabla.Name
        On Error GoTo 0
        
    End If
Next
msConn.CommitTrans
Set objTabla = Nothing
Set objCat = Nothing
Exit Sub

ErrDeleteHost:
Dim OBJERR As Error
''''Open App.Path & "\ERRLOG.TXT" For Append As #1
''''    Print #1, "HOST..........." & Date & " - " & Time
''''    Print #1, "Tabla : " & objTabla.Name
''''    For Each OBJERR In msConn.Errors
''''        Print #1, OBJERR.Description
''''    Next
    
    For Each OBJERR In msConn.Errors
        Call EscribeLog("ERROR.BorraTransHost, Tabla(" & objTabla.Name & ") " & OBJERR.Description)
        'Print #1, OBJERR.Description
    Next
    
    Resume Next
''''Close #1

End Sub
'''Public Sub Seleccion_SlipPrinter()
''''SELECCION SLIP PRINTER
'''Dim x As Printer
'''Dim cSolo As String
'''Dim nSolo As Integer
'''
'''On Error GoTo Fallo:
'''cSolo = "SOLOSLIP"
'''nSolo = 0
'''For Each x In Printers
'''    If x.DeviceName = cSolo Then
'''        'EscribeLog ("Slip.Impresora Seleccionada" & DEFAULT_PRINTER)
'''        Set Printer = x
'''        Exit For
'''    End If
'''    nSolo = nSolo + 1
'''Next
'''On Error GoTo 0
'''Exit Sub
'''
'''Fallo:
'''    MsgBox "Existe un problema con el Impresor de Notas (SLIP PRINTER)", vbCritical, BoxTit
'''    Resume Next
'''End Sub

'''Public Function GetCurrPrinter() As String
'''GetCurrPrinter = RegGetString$(HKEY_CURRENT_CONFIG, "System\CurrentControlSet\Control\Print\Printers", "Default")
'''End Function
Public Function RegGetString$(hInKey As Long, ByVal subKey$, ByVal valname$)
    Dim retVal$, hSubKey As Long, dwType As Long, SZ As Long
    Dim R As Long
    Dim v As String
    retVal$ = ""
    Const KEY_ALL_ACCESS As Long = &HF0063
    Const ERROR_SUCCESS As Long = 0
    Const REG_SZ As Long = 1
    R = RegOpenKeyEx(hInKey, subKey$, 0, KEY_ALL_ACCESS, hSubKey)
    If R <> ERROR_SUCCESS Then GoTo Quit_Now
    SZ = 256: v$ = String$(SZ, 0)
    R = RegQueryValueEx(hSubKey, valname$, 0, dwType, ByVal v$, SZ)
    If R = ERROR_SUCCESS And dwType = REG_SZ Then
        retVal$ = Left$(v$, SZ - 1)
    Else
        retVal$ = "--Not String--"
    End If
    If hInKey = 0 Then R = RegCloseKey(hSubKey)
Quit_Now:
    RegGetString$ = retVal$
End Function
'''Public Sub Seleccion_KitchenPrinter()
''''SELECCION IMPRESORA DE COCINA
'''Dim x As Printer
'''Dim cSolo As String
'''Dim nSolo As Integer
'''
'''On Error GoTo Fallo:
'''cSolo = "SOLOCOCINA"
'''nSolo = 0
'''For Each x In Printers
'''    If x.DeviceName = cSolo Then
'''        Set Printer = Printers(nSolo)
'''        Exit For
'''    End If
'''    nSolo = nSolo + 1
'''Next
'''On Error GoTo 0
'''Exit Sub
'''
'''Fallo:
'''    MsgBox "Existe un problema para conextarse con el impresor de la COCINA", vbCritical, BoxTit
'''    Resume Next
'''End Sub
''''Public Sub Seleccion_BarraPrinter()
'''''SELECCION SLIP PRINTER
''''Dim x As Printer
''''Dim cSolo As String
''''Dim nSolo As Integer
''''
''''On Error GoTo Fallo:
''''cSolo = "SOLOBARRA"
''''nSolo = 0
''''For Each x In Printers
''''    If x.DeviceName = cSolo Then
''''        Set Printer = Printers(nSolo)
''''        Exit For
''''    End If
''''    nSolo = nSolo + 1
''''Next
''''On Error GoTo 0
''''Exit Sub
''''
''''Fallo:
''''    MsgBox "Existe un problema con el impresor de la BARRA", vbCritical, BoxTit
''''    Resume Next
''''End Sub
Public Sub CargaFormasPago(RSPAGOS As Recordset, RSPROPINAS As Recordset, Forma As Form)
Dim MiTop As Integer, MiLeft As Integer, StayLeft As Integer
Dim numplu As Integer
Dim sqltext As String
Dim i As Integer
Dim iErr As Integer

iErr = 0

On Error GoTo ErrAdm:
Set RSPAGOS = New Recordset
Set RSPROPINAS = New Recordset

sqltext = "SELECT * FROM pagos WHERE CODIGO <> 999 AND CODIGO <> 99 ORDER BY CODIGO"
RSPAGOS.Open sqltext, msConn, adOpenStatic, adLockOptimistic

'INFO: SEPT 2010. AGREGANDO CHEQUE COMO OPCION DE PROPINA.
'sqltext = "SELECT * FROM pagos WHERE TIPO = 'TJ' ORDER BY CODIGO"
sqltext = "SELECT * FROM pagos WHERE TIPO = 'TJ' OR TIPO = 'CH' ORDER BY CODIGO"
RSPROPINAS.Open sqltext, msConn, adOpenStatic, adLockOptimistic

For i = 1 To 12
    Load Forma.cmdFPagos(i)
Next

For i = 1 To 9
    Load Forma.cmdPropina(i)
Next

MiTop = 360: StayLeft = 120
MiLeft = 0: numplu = 0
'MiTop = Forma.cmdFPagos(0).Top
'StayLeft = Forma.cmdFPagos(0).Left

'codigo,tipo,descrip
Do Until RSPAGOS.EOF
    If numplu < 1 Then
        Forma.cmdFPagos(numplu).Caption = RSPAGOS!DESCRIP
        Forma.cmdFPagos(numplu).Tag = RSPAGOS!CODIGO
    Else
        If Not IsObject(Forma.cmdFPagos(numplu)) Then
           Load Forma.cmdFPagos(numplu)
        End If
        Forma.cmdFPagos(numplu).Visible = True
        Forma.cmdFPagos(numplu).Caption = RSPAGOS!DESCRIP
        Forma.cmdFPagos(numplu).Tag = RSPAGOS!CODIGO
        Forma.cmdFPagos(numplu).Left = MiLeft + StayLeft
        Forma.cmdFPagos(numplu).Top = MiTop
        'StayLeft = Forma.cmdFPagos(0).Left
        StayLeft = 120
    End If
    numplu = numplu + 1
    MiLeft = MiLeft + 1440
    If numplu = 4 Or numplu = 8 Or numplu = 12 Then
        MiTop = MiTop + 800
        MiLeft = 0
    End If
    If numplu = 12 Then Exit Do
    RSPAGOS.MoveNext
Loop

MiTop = 360: StayLeft = 120
'MiTop = Forma.cmdPropina(0).Top
'StayLeft = Forma.cmdPropina(0).Left
MiLeft = 0: numplu = 0

Do Until RSPROPINAS.EOF
    If numplu < 1 Then
        Forma.cmdPropina(numplu).Caption = RSPROPINAS!DESCRIP
        Forma.cmdPropina(numplu).Tag = RSPROPINAS!CODIGO
    Else
        If Not IsObject(Forma.cmdPropina(numplu)) Then
           Load Forma.cmdPropina(numplu)
        End If
        Forma.cmdPropina(numplu).Visible = True
        Forma.cmdPropina(numplu).Caption = RSPROPINAS!DESCRIP
        Forma.cmdPropina(numplu).Tag = RSPROPINAS!CODIGO
        Forma.cmdPropina(numplu).Left = MiLeft + StayLeft
        Forma.cmdPropina(numplu).Top = MiTop
        'StayLeft = Forma.cmdPropina(0).Left
        StayLeft = 120
    End If
    numplu = numplu + 1
    MiLeft = MiLeft + 1440
    If numplu = 4 Or numplu = 8 Or numplu = 12 Then
        MiTop = MiTop + 800
        MiLeft = 0
    End If
    If numplu = 9 Then Exit Do
    RSPROPINAS.MoveNext
Loop

On Error GoTo 0

Exit Sub

ErrAdm:
If iErr < 4 Then
    iErr = iErr + 1
    Resume
Else
    ShowMsg Err.Number & Space(2) & Err.Description & vbCrLf & "ANOTE EL MENSAJE DE ERROR", vbRed, vbYellow
    ShowMsg "INTENTE LO SIGUIENTE: SALGA Y ENTRE AL PROGRAMA DE FACTURACION NUEVAMENTE", vbRed, vbYellow
    Exit Sub
End If
End Sub

Public Sub StatMesa(nMesaToBlock As Integer, nStatusLock As Byte)
'msConn.BeginTrans
msConn.Execute "UPDATE MESAS SET LOCK = " & nStatusLock & " WHERE NUMERO = " & nMesaToBlock
'msConn.CommitTrans
End Sub
Public Sub CargaFormasPagoSimple(RSPAGOS As Recordset, Forma As Form)
Dim MiTop As Integer, MiLeft As Integer, StayLeft As Integer
Dim numplu As Integer
Dim sqltext As String
Dim i As Integer
Dim iErr As Integer

iErr = 0

On Error GoTo ErrAdm:
Set RSPAGOS = New Recordset

sqltext = "SELECT * FROM pagos WHERE CODIGO <> 999 AND CODIGO <> 99 ORDER BY CODIGO"
RSPAGOS.Open sqltext, msConn, adOpenStatic, adLockOptimistic

For i = 1 To 12
    Load Forma.cmdFPagos(i)
Next

'MiTop = 360: StayLeft = 120
MiTop = 2280: StayLeft = 240
MiLeft = 0: numplu = 0

'codigo,tipo,descrip
Do Until RSPAGOS.EOF
    If numplu < 1 Then
        Forma.cmdFPagos(numplu).Caption = RSPAGOS!DESCRIP
        Forma.cmdFPagos(numplu).Tag = RSPAGOS!CODIGO
    Else
        If Not IsObject(Forma.cmdFPagos(numplu)) Then
           Load Forma.cmdFPagos(numplu)
        End If
        Forma.cmdFPagos(numplu).Visible = True
        Forma.cmdFPagos(numplu).Caption = RSPAGOS!DESCRIP
        Forma.cmdFPagos(numplu).Tag = RSPAGOS!CODIGO
        Forma.cmdFPagos(numplu).Left = MiLeft + StayLeft
        Forma.cmdFPagos(numplu).Top = MiTop
        StayLeft = 240
    End If
    numplu = numplu + 1
    MiLeft = MiLeft + 1440
    If numplu = 4 Or numplu = 8 Or numplu = 12 Then
        MiTop = MiTop + 800
        MiLeft = 0
    End If
    If numplu = 12 Then Exit Do
    RSPAGOS.MoveNext
Loop

'MiTop = 360: StayLeft = 120
MiTop = 2280: StayLeft = 240
MiLeft = 0: numplu = 0

On Error GoTo 0

Exit Sub

ErrAdm:
If iErr < 4 Then
    iErr = iErr + 1
    Resume
Else
    MsgBox Err.Number & Space(2) & Err.Description, vbCritical, "ANOTE EL MENSAJE DE ERROR"
    ShowMsg "HAGA LO SIGUIENTE: " & vbCrLf & " SALGA Y ENTRE AL MODULO DE FACTURACION NUEVAMENTE", vbRed, vbYellow
    Exit Sub
End If
End Sub

Public Function RoundToNearest(Amt As Double, RoundAmt As Variant, Direction As Integer) As Double
'Direcction = 1 Hacia Arriba
'Direcction = 0 Hacia Abajo
Dim temp As Single
Dim iPass As Integer
iPass = 0

On Error GoTo ErrAdm:

temp = Amt / RoundAmt
If Int(temp) = temp Then
   RoundToNearest = Amt
Else
   If Direction = 0 Then
      temp = Int(temp)
   Else
      temp = Int(temp) + 1
   End If
   RoundToNearest = temp * RoundAmt
End If
On Error GoTo 0
Exit Function

ErrAdm:
iPass = iPass + 1
If iPass > 3 Then
    Resume Next
Else
    Resume
End If
End Function

Public Sub CargaCargos(RSPROPINAS As Recordset, Forma As Form)
Dim MiTop As Integer, MiLeft As Integer, StayLeft As Integer
Dim numplu As Integer
Dim sqltext As String
Dim i As Integer
Dim iErr As Integer
Dim nCargosValor As Currency

iErr = 0

On Error Resume Next
nCargosValor = GetFromINI("Facturacion", "CargosValor", App.Path & "\soloini.ini")
On Error GoTo 0

On Error GoTo ErrAdm:

Set RSPROPINAS = New Recordset

sqltext = "SELECT * FROM pagos WHERE TIPO = 'CA' ORDER BY DESCRIP"
RSPROPINAS.Open sqltext, msConn, adOpenStatic, adLockOptimistic

For i = 1 To 12
    Load Forma.cmdCARGOS(i)
Next

'MiTop = 360: StayLeft = 120
MiTop = 5880: StayLeft = 240
MiLeft = 0: numplu = 0

'codigo,tipo,descrip
Do Until RSPROPINAS.EOF
    If numplu < 1 Then
        Forma.cmdCARGOS(numplu).Caption = RSPROPINAS!DESCRIP
        Forma.cmdCARGOS(numplu).Tag = RSPROPINAS!CODIGO
        Forma.cmdCARGOS(numplu).ToolTipText = Format(nCargosValor, "CURRENCY")
    Else
        If Not IsObject(Forma.cmdCARGOS(numplu)) Then
           Load Forma.cmdCARGOS(numplu)
        End If
        Forma.cmdCARGOS(numplu).Visible = True
        Forma.cmdCARGOS(numplu).Caption = RSPROPINAS!DESCRIP
        Forma.cmdCARGOS(numplu).Tag = RSPROPINAS!CODIGO
        Forma.cmdCARGOS(numplu).ToolTipText = Format(nCargosValor, "CURRENCY")
        Forma.cmdCARGOS(numplu).Left = MiLeft + StayLeft
        Forma.cmdCARGOS(numplu).Top = MiTop
        StayLeft = 240
    End If
    numplu = numplu + 1
    MiLeft = MiLeft + 1440
    If numplu = 4 Or numplu = 8 Or numplu = 12 Then
        MiTop = MiTop + 800
        MiLeft = 0
    End If
    If numplu = 12 Then Exit Do
    RSPROPINAS.MoveNext
Loop

'MiTop = 360: StayLeft = 120
MiTop = 5880: StayLeft = 240
MiLeft = 0: numplu = 0

On Error GoTo 0

Exit Sub

ErrAdm:
If iErr < 4 Then
    iErr = iErr + 1
    Resume
Else
    MsgBox Err.Number & Space(2) & Err.Description, vbCritical, "ANOTE EL MENSAJE DE ERROR"
    ShowMsg "HAGA LO SIGUIENTE: " & vbCrLf & " SALGA Y ENTRE AL MODULO DE FACTURACION NUEVAMENTE", vbRed, vbYellow
    Exit Sub
End If
End Sub



'---------------------------------------------------------------------------------------
' Procedure : NOTA_CREDITO_PUT
' Author    : hsequeira
' Date      : 08/04/2017
' Purpose   : GUARDA LA INFO DE LAS NOTAS DE CREDITO
'---------------------------------------------------------------------------------------
'
Public Function NOTA_CREDITO_PUT(nnCantidad As Long, nnValor As Single) As Boolean
Dim nnLocalCounter As Long
Dim nnLocalValor As Single


On Error Resume Next
If RegRead("HKCU\Software\SoloSoftware\SoloMix\NCCounter") = "" Then
    'SI ES LA PRIMERA CORTESIA DEL DIA, ESCRIBE LO QUE RECIBE LA FUNCION
    
    RegWrite "HKCU\Software\SoloSoftware\SoloMix\NCCounter", nnCantidad
    RegWrite "HKCU\Software\SoloSoftware\SoloMix\NCValor", (nnCantidad * nnValor)
Else
    'SI, YA HAY CORTESIAS, ENTONCES SUMA LOS DATOS RECIBIDOS
    nnLocalCounter = CLng(RegRead("HKCU\Software\SoloSoftware\SoloMix\NCCounter"))
    nnLocalValor = CSng(RegRead("HKCU\Software\SoloSoftware\SoloMix\NCValor"))
    nnLocalCounter = nnLocalCounter + nnCantidad
    nnLocalValor = nnLocalValor + (nnCantidad * nnValor)
    
    RegWrite "HKCU\Software\SoloSoftware\SoloMix\NCCounter", nnLocalCounter
    RegWrite "HKCU\Software\SoloSoftware\SoloMix\NCValor", nnLocalValor
End If
On Error GoTo 0
End Function


'---------------------------------------------------------------------------------------
' Procedure : NOTA_CREDITO_GET
' Author    : hsequeira
' Date      : 08/04/2017
' Purpose   : LEE LAS NOTAS DE CREDITO DEL PERIODO
'---------------------------------------------------------------------------------------
'
Public Function NOTA_CREDITO_GET() As String
Dim nnLocalCounter As Long
Dim nnLocalValor As Single

'-----------------------------------------------------------------------------------------------------------------
'On Error Resume Next ==> Para que no reviente el programa.
'-----------------------------------------------------------------------------------------------------------------

On Error Resume Next
If RegRead("HKCU\Software\SoloSoftware\SoloMix\NCCounter") = "" Then
    'NO HAY CORTESIAS MARCADAS EN ESTE PERIODO
    NOTA_CREDITO_GET = ""
Else
    'SI, YA HAY CORTESIAS, ENTONCES SUMA LOS DATOS RECIBIDOS
    nnLocalCounter = CLng(RegRead("HKCU\Software\SoloSoftware\SoloMix\NCCounter"))
    nnLocalValor = CSng(RegRead("HKCU\Software\SoloSoftware\SoloMix\NCValor"))
    'CORTESIA_GET = "CORTESIA:" & nnLocalCounter & Space(6) & Format(nnLocalValor, "STANDARD")
    NOTA_CREDITO_GET = "N.CREDITO:" & Format(nnLocalCounter, "@@@@@@") & _
                                    Format(Format(nnLocalValor, "STANDARD"), "@@@@@@@@@@@@")
    RegWrite "HKCU\Software\SoloSoftware\SoloMix\NCCounter", ""
    RegWrite "HKCU\Software\SoloSoftware\SoloMix\NCValor", ""
End If
On Error GoTo 0
End Function


'|||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||
'|||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||

'---------------------------------------------------------------------------------------
' Procedure : GetTAXFE
' Author    : hsequeira
' Date      : 29/08/2023
' Purpose   : OBTIENE EL TAXID SEGUN EL TAX DEL PRODUCTO
'---------------------------------------------------------------------------------------
'
Public Function GetTAXFE(nTAX As Integer) As String
Select Case nTAX
    Case 0
        GetTAXFE = "00"
    Case 7
        GetTAXFE = "01"
    Case 10
        GetTAXFE = "02"
    Case 15
        GetTAXFE = "03"
End Select
End Function
'|||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||
'|||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||
