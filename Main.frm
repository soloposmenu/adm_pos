VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{FBFD55C6-C23C-11D3-B65D-004005E66149}#1.0#0"; "swiftprint.ocx"
Begin VB.Form Main 
   BackColor       =   &H00B39665&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Repetición de Reporte Z"
   ClientHeight    =   7410
   ClientLeft      =   390
   ClientTop       =   780
   ClientWidth     =   8700
   ClipControls    =   0   'False
   Icon            =   "Main.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7410
   ScaleWidth      =   8700
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton PLU 
      Caption         =   "PLU"
      Height          =   495
      Left            =   480
      TabIndex        =   5
      Top             =   6120
      Width           =   1095
   End
   Begin VB.CommandButton PrintLocalPrinter 
      BackColor       =   &H0000FFFF&
      Height          =   495
      Left            =   840
      Picture         =   "Main.frx":030A
      Style           =   1  'Graphical
      TabIndex        =   4
      ToolTipText     =   "Envia REPORTE a la Impresora (8.5x11)"
      Top             =   5280
      Width           =   735
   End
   Begin VB.ListBox List1 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Lucida Console"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4560
      Left            =   2760
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   2400
      Width           =   5895
   End
   Begin MSComctlLib.ListView LV 
      Height          =   1935
      Left            =   120
      TabIndex        =   0
      Top             =   360
      Width           =   8535
      _ExtentX        =   15055
      _ExtentY        =   3413
      View            =   3
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   0
   End
   Begin SwiftPrintLib.SwiftPrint spDoc 
      Left            =   1800
      Top             =   5880
      _Version        =   65536
      _ExtentX        =   1508
      _ExtentY        =   661
      _StockProps     =   0
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00B39665&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   2655
      Index           =   1
      Left            =   120
      TabIndex        =   3
      Top             =   2400
      Width           =   2055
   End
   Begin VB.Label Label1 
      BackColor       =   &H00B39665&
      Caption         =   "Seleccione un reporte a Imprimir bajo la Columna Contador Z y haga DOBLE CLICK"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   7335
   End
   Begin VB.Menu mnuZ 
      Caption         =   "Principal"
      Visible         =   0   'False
      Begin VB.Menu mnuGo 
         Caption         =   "Imprimir Reporte (Z)"
      End
   End
End
Attribute VB_Name = "Main"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private msConn As New ADODB.Connection
Private rsZCounter As New ADODB.Recordset
Private nZToPrint As String
Private cZFecha As String
Private cZHora As String
Private cZTotalAnterior As Double
Private cZTotalNuevo As Double
Private cZGranTotal As Double
Private cZITBMS As Double
Private rs00 As New ADODB.Recordset
Private Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
Private Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpstring As Any, ByVal lpFileName As String) As Long

Private Declare Function PrintDlg Lib "COMDLG32.DLL" () As Integer

Private Function GetFromINI(Section As String, Key As String, Directory As String) As String
Dim strBuffer As String

On Error GoTo FileError:
    strBuffer = String(750, Chr(0))
    Key$ = LCase$(Key$)
    GetFromINI$ = Left(strBuffer, GetPrivateProfileString(Section$, ByVal Key$, "", strBuffer, Len(strBuffer), Directory$))
On Error GoTo 0
Exit Function

FileError:
    MsgBox Err.Number + ": NO SE ENCUENTRA ARCHIVO DE INICIALIZACION", vbCritical, "ERROR AL INICIAR"
    Resume Next
End Function
Private Sub WriteToINI(Section As String, Key As String, KeyValue As String, Directory As String)
On Error GoTo FileError:
    Call WritePrivateProfileString(Section$, UCase$(Key$), KeyValue$, Directory$)
On Error GoTo 0
Exit Sub

FileError:
    MsgBox Err.Number + ": NO SE ENCUENTRA ARCHIVO DE INICIALIZACION", vbCritical, "ERROR AL INICIAR"
    Resume Next
End Sub
Private Sub Form_Load()
Dim nFila As Integer

Show

Label1(1).Caption = "Haga Click con el boton derecho del Mouse para enviar a Imprimir a la Impresora de Facturación" & vbCrLf & vbCrLf & _
    "CUANDO TERMINE, VAYA AL REPORTE Z EN VENTAS, Y SELECCIONE EN EL MENU 'Repetir Z de Archivo'"

cDataPath = GetFromINI("General", "ProveedorDatos", App.Path & "\soloini.ini")

cDataPath = cDataPath & ";Jet OLEDB:Database Password=master24"

msConn.Open cDataPath

rs00.Open "SELECT * FROM ORGANIZACION ", msConn, adOpenForwardOnly, adLockReadOnly
rsZCounter.Open "SELECT * FROM Z_COUNTER ORDER BY ID DESC", msConn, adOpenStatic, adLockOptimistic
LV.ListItems.Clear
LV.ColumnHeaders.Clear
LV.ColumnHeaders.Add , , "Contador Z"
LV.ColumnHeaders.Add , , "Fecha"
LV.ColumnHeaders.Add , , "Hora"
LV.ColumnHeaders.Add , , "Se Imprimio?"
LV.ColumnHeaders.Add , , "Total Anterior"
LV.ColumnHeaders.Add , , "Total Nuevo"
LV.ColumnHeaders.Add , , "Venta del Dia"
'LV.ColumnHeaders.Add , , "ITBMS"

LV.ColumnHeaders.Item(1).Width = 1000
LV.ColumnHeaders.Item(2).Width = 1100
LV.ColumnHeaders.Item(3).Width = 800
LV.ColumnHeaders.Item(4).Width = 1200
LV.ColumnHeaders.Item(5).Width = 1300
LV.ColumnHeaders.Item(6).Width = 1300
LV.ColumnHeaders.Item(7).Width = 1350
LV.ColumnHeaders.Item(7).Alignment = lvwColumnRight
nFila = 1
Do While Not rsZCounter.EOF
    LV.ListItems.Add , , rsZCounter!CONTADOR
    LV.ListItems.Item(nFila).ListSubItems.Add , , Right(rsZCounter!FECHA, 2) & "/" & Mid(rsZCounter!FECHA, 5, 2) & "/" & Left(rsZCounter!FECHA, 4)
    LV.ListItems.Item(nFila).ListSubItems.Add , , Format(rsZCounter!HORA, "00:00")
    LV.ListItems.Item(nFila).ListSubItems.Add , , IIf(rsZCounter!PRINT_OK = -1, "SI", "NO")
    LV.ListItems.Item(nFila).ListSubItems.Add , , Format(IIf(IsNull(rsZCounter!TOTAL_ANTERIOR), "0.00", rsZCounter!TOTAL_ANTERIOR), "0.00")
    LV.ListItems.Item(nFila).ListSubItems.Add , , Format(IIf(IsNull(rsZCounter!TOTAL_NUEVO), "0.00", rsZCounter!TOTAL_NUEVO), "0.00")
    LV.ListItems.Item(nFila).ListSubItems.Add , , Format(Format(IIf(IsNull(rsZCounter!TOTAL_NUEVO - rsZCounter!TOTAL_ANTERIOR), "0.00", rsZCounter!TOTAL_NUEVO - rsZCounter!TOTAL_ANTERIOR), "0.00"), "CURRENCY")
    'LV.ListItems.Item(nFila).ListSubItems.Add , , Format(IIf(IsNull(rsZCounter!GRAN_TOTAL), "0.00", rsZCounter!GRAN_TOTAL), "0.00")
    'INFO: AGREGANDO EL ITBMS
    On Error Resume Next
    LV.ListItems.Item(nFila).ListSubItems.Add , , Format(IIf(IsNull(rsZCounter!ITBMS), "0.00", rsZCounter!ITBMS), "0.00")
    On Error GoTo 0

    nFila = nFila + 1
    rsZCounter.MoveNext
Loop
rsZCounter.Close

End Sub


Private Sub List1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
'Dim vResp As Variant
'If Button = 2 Then
'    vResp = MsgBox("¿ Desea enviar el Reporte (Z) a la imrpresora ?", vbQuestion + vbYesNo, "Impresión de Reporte")
'    Call Print2File
'End If
If List1.ListCount - 1 <> 0 Then PopupMenu mnuZ
End Sub
Private Sub Print2File()
Dim FACTURA_FILE As String
Dim DATA_PATH  As String
Dim cDataPath As String
Dim i As Integer
Dim nFreefile As Integer

Me.MousePointer = vbHourglass
FACTURA_FILE = "REPITEZ.TXT"
DATA_PATH = GetFromINI("General", "DirectorioDatos", App.Path & "\soloini.ini")
cFactFile = DATA_PATH + "\" & FACTURA_FILE

nFreefile = FreeFile()
Open cFactFile For Output As nFreefile
For i = 0 To List1.ListCount - 1
    Print #nFreefile, List1.List(i)
Next
Close nFreefile
Me.MousePointer = vbDefault
End Sub
Private Sub LV_DblClick()
Dim vResp As Variant
Dim rsTempo As New ADODB.Recordset
Dim cSQL As String

On Error Resume Next
vResp = MsgBox("¿ Desea Imprimir el Reporte (Z) # " & LV.SelectedItem.Text & " ?", vbQuestion + vbYesNo, "Impresión de Reporte")
If vResp = vbYes Then
    nZToPrint = LV.SelectedItem.Text
    cZFecha = LV.SelectedItem.ListSubItems(1).Text
    cZHora = LV.SelectedItem.ListSubItems(2).Text
    cZTotalAnterior = Val(LV.SelectedItem.ListSubItems(4).Text)
    cZTotalNuevo = Val(LV.SelectedItem.ListSubItems(5).Text)
    cZGranTotal = Val(LV.SelectedItem.ListSubItems(6).Text)
    On Error Resume Next
    cZITBMS = Val(LV.SelectedItem.ListSubItems(7).Text)
    On Error GoTo 0
    List1.Clear
'    List1.ListIndex = 0
    List1.Refresh
    Me.MousePointer = vbHourglass
    
    Call RepCajZ
    
    Me.MousePointer = vbDefault
End If
On Error GoTo 0
End Sub

Private Sub RepCajZ()
'REPORTE DE CAJEROS - TERMINAL - DEPARTAMENTOS
Dim rsVta_Z As Recordset
Dim rsPgo_Z As Recordset
Dim rsTran As Recordset
Dim RSPAGOS As Recordset    'Pagos
Dim rsAjustes As Recordset
Dim rsCajeros As Recordset
Dim rsProp As Recordset
Dim rsLocalDepto As Recordset
Dim rsLocTerminal As Recordset
Dim rsSuperGrp As Recordset
Dim rsHash As Recordset
Dim cSQL As String
Dim nSumVta As Double
Dim nSumTotal As Double
Dim nTrans As Integer
Dim MiLen1 As Integer
Dim Milen2 As Integer
Dim nErrInd As Integer
Dim errCounter As Integer

nSumVta = 0
nSumTotal = 0
nErrInd = 0

Set rsHash = New Recordset
Set rsVta_Z = New Recordset
Set rsPgo_Z = New Recordset
Set rsTran = New Recordset
Set RSPAGOS = New Recordset
Set rsAjustes = New Recordset
Set rsCajeros = New Recordset
Set rsProp = New Recordset
Set rsLocalDepto = New Recordset

'ABRE DEPARTAMENTOS
rsLocalDepto.Open "SELECT CODIGO,DESCRIP,CORTO FROM DEPTO", msConn, adOpenDynamic, adLockOptimistic

cSQL = "SELECT DISTINCT CAJERO FROM HIST_TR WHERE Z_COUNTER = '" & nZToPrint & "' ORDER BY CAJERO"
rsCajeros.Open cSQL, msConn, adOpenStatic, adLockOptimistic

If rsCajeros.EOF Then
    MsgBox "IMPRIMIENDO EL REPORTE, AUNQUE NO EXISTEN TRANSACCIONES", vbInformation, BoxTit
End If

On Error Resume Next    'NUEVO PARA CORRER REPORTE EN ZZZZ
rsCajeros.MoveFirst
rsCajeros.MoveLast

If rsCajeros.RecordCount = 0 Or rsCajeros!CAJERO = "" Then
    MsgBox "IMPRIMIENDO EL REPORTE, AUNQUE NO EXISTEN TRANSACCIONES", vbInformation, BoxTit
End If
On Error GoTo 0         'NUEVO PARA CORRER REPORTE EN ZZZZ

On Error GoTo AjustaMilen:

''--ProgBar.Value = 5
cSQL = "SELECT * FROM pagos WHERE CODIGO <> 999 ORDER BY CODIGO"
RSPAGOS.Open cSQL, msConn, adOpenStatic, adLockOptimistic

''--ProgBar.Value = 10
'''''''ReportesEscribeLog ("Reporte Z - Inicio de Reporte ")
On Error Resume Next
rsCajeros.MoveFirst
On Error GoTo 0
Do Until rsCajeros.EOF

    ''--ProgBar.Value = 20
    
    'ABRE CADA UNO DE LOS CAJEROS QUE HAN TRABAJADO DESDE LA ULTIMA (Z)
    rsTran.Open "SELECT distinct NUM_TRANS FROM HIST_TR " & _
            " WHERE CAJERO = " & rsCajeros!CAJERO & " AND " & _
            "Z_COUNTER = '" & nZToPrint & "'", msConn, adOpenStatic, adLockOptimistic

    rsTran.MoveFirst
    rsTran.MoveLast
    nTrans = rsTran.RecordCount
    rsTran.MoveFirst
    
    cSQL = "SELECT b.nombre,b.apellido,a.cajero,b.z_c, sum(a.precio) as Ventas " & _
            " FROM HIST_TR as a, cajeros as b " & _
            " WHERE a.cajero = " & rsCajeros!CAJERO & _
            " AND b.numero = " & rsCajeros!CAJERO & _
            " AND Z_COUNTER = '" & nZToPrint & "'" & _
            " GROUP BY a.cajero,b.nombre,b.apellido,b.z_c "

    'VALOR EN VENTAS DEL CAJERO CON rsVta_Z
    rsVta_Z.Open cSQL, msConn, adOpenStatic, adLockOptimistic

    If rsVta_Z.RecordCount = 0 Then
        'MsgBox "EL CAJERO NO TIENE VENTAS, NO SE IMPRIMIRA REPORTE EN Z"
        rsCajeros.MoveNext
    End If
    
    'TODOS LOS PAGOS RECIBIDOS
    'ESTOY SACANDO EL DESCUENTO GLOBAL (99) DE AQUI PARA PONERLO EN
    'LOS AJUSTES
    cSQL = "SELECT a.cajero,a.tipo_pago,SUM(a.monto) AS Valor, " & _
            " COUNT(a.tipo_pago) as Z_COUNT " & _
            " FROM HIST_TR_PAGO as a " & _
            " WHERE a.cajero = " & rsCajeros!CAJERO & _
            " AND a.tipo_pago <> 99 " & _
            " AND a.Z_COUNTER = '" & nZToPrint & "'" & _
            " GROUP BY a.cajero,a.TIPO_PAGO"
    rsPgo_Z.Open cSQL, msConn, adOpenStatic, adLockOptimistic
    
    cSQL = "SELECT MID(a.TIPO,1,2) AS CORREC,a.DESCUENTO, " & _
        " COUNT(a.lin) as Z_COUNT, SUM(a.precio) as valor " & _
        " FROM HIST_TR as a " & _
        " WHERE A.CAJERO = " & rsCajeros!CAJERO & _
        " AND MID(A.TIPO,1,1) <> ' ' " & _
        " AND Z_COUNTER = '" & nZToPrint & "'" & _
        " GROUP BY MID(a.TIPO,1,2),a.DESCUENTO"
    rsAjustes.Open cSQL, msConn, adOpenStatic, adLockOptimistic

    List1.AddItem cZFecha & Space(2) & cZHora
    List1.AddItem Space(2)
    List1.AddItem "REPORTE DE CAJEROS (Z)"
    List1.AddItem "COPIA DEL ORIGINAL"
    List1.AddItem Space(2)
    List1.AddItem rs00!DESCRIP
    List1.AddItem "RUC:" & rs00!RUC
    List1.AddItem "SERIAL:" & rs00!SERIAL
    List1.AddItem Space(2)
    '''- SEGUN HACIENDA Y TESORO -'''Printer.Print "CONTADOR TRANS : " & (rs00!TRANS + 1)
    List1.AddItem "CONTADOR Z : XXX"
    List1.AddItem "CAJERO : " & rsVta_Z!NOMBRE & ", " & rsVta_Z!APELLIDO
    List1.AddItem Space(2)
    
    Do Until rsVta_Z.EOF
        MiLen1 = Len(nTrans)
        Milen2 = Len(Format(rsVta_Z!VENTAS, "STANDARD"))
        List1.AddItem "VENTA DEL DIA:" & Space(4 - MiLen1) & nTrans & Space(11 - Milen2) & Format(rsVta_Z!VENTAS, "STANDARD")
        nSumTotal = nSumTotal + rsVta_Z!VENTAS
        rsVta_Z.MoveNext
    Loop
    
    List1.AddItem Space(2)
    List1.AddItem "TOTALES DE CAJA"
    ''''''''Printer.FontUnderline = False
    List1.AddItem Space(2)

    Do Until rsPgo_Z.EOF
        RSPAGOS.MoveFirst
        RSPAGOS.Find "CODIGO = " & rsPgo_Z!TIPO_PAGO
        If Not RSPAGOS.EOF Then
            MiLen1 = Len(rsPgo_Z!Z_COUNT)
            Milen2 = Len(Format(rsPgo_Z!VALOR, "STANDARD"))
            List1.AddItem FormatTexto(RSPAGOS!DESCRIP, 13) & Space(4 - MiLen1) & rsPgo_Z!Z_COUNT & Space(13 - Milen2) & Format(rsPgo_Z!VALOR, "STANDARD")
        Else
            MiLen1 = 1
            Milen2 = Len(Format(0#, "standard"))
            List1.AddItem "OTRO PAGO    " & Space(4 - MiLen1) & 0 & Space(13 - MiLen1) & Format(0#, "standard")
        End If
        nSumVta = nSumVta + rsPgo_Z!VALOR
        rsPgo_Z.MoveNext
    Loop

    List1.AddItem Space(2)
    MiLen1 = Len(Format(nSumVta, "currency"))
    List1.AddItem "SUBTOTAL: " & Space(20 - MiLen1) & Format(nSumVta, "currency")
    List1.AddItem "------------------------------"
    
    cSQL = "SELECT TIPO_PAGO,COUNT(TIPO_PAGO) AS Z_COUNT, " & _
            " SUM(MONTO) AS VALOR FROM HIST_TR_PROP " & _
            " WHERE CAJERO = " & rsCajeros!CAJERO & _
            " AND Z_COUNTER = '" & nZToPrint & "'" & _
            " GROUP BY TIPO_PAGO "
    rsProp.Open cSQL, msConn, adOpenStatic, adLockOptimistic
    
    Do Until rsProp.EOF
        RSPAGOS.MoveFirst
        RSPAGOS.Find "CODIGO = " & rsProp!TIPO_PAGO
        If Not RSPAGOS.EOF Then
            MiLen1 = Len(rsProp!Z_COUNT)
            Milen2 = Len(Format(rsProp!VALOR, "STANDARD"))
            List1.AddItem "Propina " & FormatTexto(RSPAGOS!DESCRIP, 5) & Space(4 - MiLen1) & rsProp!Z_COUNT & Space(13 - Milen2) & Format(rsProp!VALOR, "STANDARD")
        End If
        rsProp.MoveNext
    Loop
    
    rsProp.Close
    
    List1.AddItem Space(2)
    '''''Printer.FontUnderline = True
    List1.AddItem "AJUSTES"
    '''''Printer.FontUnderline = False
    List1.AddItem Space(2)
    
    Do Until rsAjustes.EOF
        MiLen1 = Len(rsAjustes!Z_COUNT)
        Milen2 = Len(Format(rsAjustes!VALOR, "STANDARD"))
        If rsAjustes!CORREC = "EC" Then
            List1.AddItem "CORECCION " & Space(8 - MiLen1) & rsAjustes!Z_COUNT & Space(12 - Milen2) & Format(rsAjustes!VALOR, "STANDARD")
        ElseIf rsAjustes!CORREC = "VO" Then
            List1.AddItem "ANULACION " & Space(8 - MiLen1) & rsAjustes!Z_COUNT & Space(12 - Milen2) & Format(rsAjustes!VALOR, "STANDARD")
        ElseIf rsAjustes!CORREC = "DC" Then
            List1.AddItem "DESCUENTO " & Format(rsAjustes!DESCUENTO, "0.00") & Space(4 - MiLen1) & rsAjustes!Z_COUNT & Space(12 - Milen2) & Format(rsAjustes!VALOR, "STANDARD")
        End If
        rsAjustes.MoveNext
    Loop

    'PREPARA INFO PARA DESCUENTO GLOBAL
    rsAjustes.Close
    cSQL = "SELECT a.cajero,a.tipo_pago,SUM(a.monto) AS Valor, " & _
            " COUNT(a.tipo_pago) as Z_COUNT " & _
            " FROM HIST_TR_PAGO as a " & _
            " WHERE a.cajero = " & rsCajeros!CAJERO & _
            " AND a.tipo_pago = 99 " & _
            " AND Z_COUNTER = '" & nZToPrint & "'" & _
            " GROUP BY a.cajero,a.TIPO_PAGO"
    
    rsAjustes.Open cSQL, msConn, adOpenStatic, adLockOptimistic
    
    If Not rsAjustes.EOF Then
        RSPAGOS.MoveFirst
        RSPAGOS.Find "CODIGO = " & rsAjustes!TIPO_PAGO
        If Not RSPAGOS.EOF Then
            MiLen1 = Len(rsAjustes!Z_COUNT)
            Milen2 = Len(Format(rsAjustes!VALOR, "STANDARD"))
            List1.AddItem FormatTexto(RSPAGOS!DESCRIP, 13) & Space(5 - MiLen1) & rsAjustes!Z_COUNT & Space(12 - Milen2) & Format(rsAjustes!VALOR, "STANDARD")
        End If
    End If
    
    'msConn.BeginTrans
    '''- SEGUN HACIENDA Y TESORO -'''msConn.Execute "UPDATE ORGANIZACION SET TRANS = TRANS + 1"
    'Incrementa el contador en Z y Resetea el de X a 0
    'msConn.Execute "UPDATE CAJEROS SET Z_C = Z_C + 1, X_C = 0 " & _
                   " WHERE NUMERO = " & rsCajeros!CAJERO
    'msConn.CommitTrans
    
    rsCajeros.MoveNext
    If Not rsCajeros.EOF = True Then
        For i = 1 To 10
            List1.AddItem Space(2)
        Next
        ''''---Coptr1.CutPaper 100
    End If
    
    rsVta_Z.Close
    rsPgo_Z.Close
    rsTran.Close
    rsAjustes.Close
    nSumVta = 0
Loop

nSumVta = 0
nSumTotal = 0
''--ProgBar.Value = 40
'-------------------   EL CAJERO TERMINAL  -------------------

''For i = 1 To 10
''    List1.AddItem Space(2)
''Next
''''---Coptr1.CutPaper 100

Set rsLocTerminal = New Recordset
rsLocTerminal.Open "SELECT Z_C FROM CAJEROS WHERE NUMERO = 999", msConn, adOpenStatic, adLockOptimistic

rsTran.Open "SELECT distinct NUM_TRANS FROM HIST_TR " & _
        " WHERE Z_COUNTER = '" & nZToPrint & "'", msConn, adOpenStatic, adLockOptimistic

nTrans = rsTran.RecordCount

cSQL = "SELECT sum(a.precio) as Ventas FROM HIST_TR as a " & _
       " WHERE Z_COUNTER = '" & nZToPrint & "'"
'VALOR EN VENTAS DEL CAJERO CON rsVta_Z
rsVta_Z.Open cSQL, msConn, adOpenStatic, adLockOptimistic

'''- SEGUN HACIENDA Y TESORO -'''rsHash.Open "SELECT SUM(ABS(A.PRECIO)) AS HASH_DIA FROM TRANSAC AS A", msConn, adOpenStatic, adLockOptimistic
rsHash.Open "SELECT SUM(A.PRECIO) AS HASH_DIA " & _
    " FROM HIST_TR AS A " & _
    " WHERE A.PRECIO > 0 " & _
    " AND Z_COUNTER = '" & nZToPrint & "'", msConn, adOpenStatic, adLockOptimistic

cSQL = "SELECT a.tipo_pago,SUM(a.monto) AS Valor, " & _
        " COUNT(a.tipo_pago) as Z_COUNT " & _
        " FROM HIST_TR_PAGO as a " & _
        " WHERE a.tipo_pago <> 99 " & _
        " AND Z_COUNTER = '" & nZToPrint & "'" & _
        " GROUP BY a.TIPO_PAGO"
rsPgo_Z.Open cSQL, msConn, adOpenStatic, adLockOptimistic

'cSQL = "SELECT MID(a.TIPO,1,2) AS CORREC,DESCUENTO, "
        '" COUNT(a.lin) as Z_COUNT, SUM(a.precio) as valor "
'2 de Nov 1999
'COUNT(a.lin) as Z_COUNT, SUM(a.precio_unit) as valor
'5 de Nov
cSQL = "SELECT MID(a.TIPO,1,2) AS CORREC,DESCUENTO, " & _
        " COUNT(a.lin) as Z_COUNT, SUM(a.precio) as valor " & _
        " FROM HIST_TR as a " & _
        " WHERE MID(A.TIPO,1,1) <> ' ' " & _
        " AND Z_COUNTER = '" & nZToPrint & "'" & _
        " GROUP BY MID(a.TIPO,1,2),DESCUENTO"

rsAjustes.Open cSQL, msConn, adOpenStatic, adLockOptimistic

List1.AddItem cZFecha & Space(2) & cZHora
List1.AddItem Space(2)
List1.AddItem "REPORTE DE TERMINAL (Z)"
List1.AddItem Space(2)
List1.AddItem rs00!DESCRIP
List1.AddItem "RUC:" & rs00!RUC
List1.AddItem "SERIAL:" & rs00!SERIAL
List1.AddItem Space(2)
'''- SEGUN HACIENDA Y TESORO -'''Printer.Print "CONTADOR TRANS : " & (rs00!TRANS + 1)
List1.AddItem "CONTADOR Z : " & (nZToPrint)
List1.AddItem "CAJERO : REPORTE/TERMINAL"
List1.AddItem Space(2)

Do Until rsVta_Z.EOF
    MiLen1 = Len(nTrans)
    Milen2 = Len(Format(rsVta_Z!VENTAS, "STANDARD"))
    List1.AddItem "VENTA DEL DIA:" & Space(5 - MiLen1) & nTrans & Space(10 - Milen2) & Format(rsVta_Z!VENTAS, "STANDARD")
    nSumTotal = nSumTotal + IIf(IsNull(rsVta_Z!VENTAS), 0, rsVta_Z!VENTAS)
    rsVta_Z.MoveNext
Loop

''--ProgBar.Value = 45
List1.AddItem Space(2)
''''''''Printer.FontUnderline = True
'Printer.Print "DESGLOSE DE INGRESOS"
List1.AddItem "TOTALES DE CAJA"
''''''''Printer.FontUnderline = False
List1.AddItem Space(2)

Do Until rsPgo_Z.EOF
    RSPAGOS.MoveFirst
    RSPAGOS.Find "CODIGO = " & rsPgo_Z!TIPO_PAGO
    If Not RSPAGOS.EOF Then
        MiLen1 = Len(rsPgo_Z!Z_COUNT)
        Milen2 = Len(Format(rsPgo_Z!VALOR, "STANDARD"))
        List1.AddItem FormatTexto(RSPAGOS!DESCRIP, 13) & Space(4 - MiLen1) & rsPgo_Z!Z_COUNT & Space(13 - Milen2) & Format(rsPgo_Z!VALOR, "STANDARD")
        'PRINT#1, rsPagos!descrip & Chr(9) & rsPgo_Z!Z_COUNT & Chr(9) & Format(rsPgo_Z!VALOR, "STANDARD")
    Else
        MiLen1 = 1
        Milen2 = Len(Format(0#, "standard"))
        List1.AddItem "OTRO PAGO    " & Space(4 - MiLen1) & 0 & Space(13 - MiLen1) & Format(0#, "standard")
        'PRINT#1, rsPagos!descrip & Chr(9) & 0 & Chr(9) & Format(0#, "standard")
    End If
    nSumVta = nSumVta + rsPgo_Z!VALOR
    rsPgo_Z.MoveNext
Loop

''--ProgBar.Value = 50
List1.AddItem Space(2)
MiLen1 = Len(Format(nSumVta, "currency"))
List1.AddItem "SUBTOTAL: " & Space(20 - MiLen1) & Format(nSumVta, "currency")
List1.AddItem "------------------------------"

'INFO: 2 AGO 2010
'PONIENDO EL ITBM DEL SOLOINI.INI

'List1.AddItem "ITBMS  (5%):" & Space(18 - MiLen1) & Format(cZITBMS, "CURRENCY")
List1.AddItem "ITBMS  (" & GetFromINI("Administracion", "PorcentajeImpuesto", App.Path & "\Soloini.ini") & "%):" & Space(18 - MiLen1) & Format(cZITBMS, "CURRENCY")
List1.AddItem Space(2)

cSQL = "SELECT TIPO_PAGO,COUNT(TIPO_PAGO) AS Z_COUNT, " & _
        " SUM(MONTO) AS VALOR " & _
        " FROM HIST_TR_PROP " & _
        " WHERE Z_COUNTER = '" & nZToPrint & "'" & _
        " GROUP BY TIPO_PAGO "
rsProp.Open cSQL, msConn, adOpenStatic, adLockOptimistic

Do Until rsProp.EOF
    RSPAGOS.MoveFirst
    RSPAGOS.Find "CODIGO = " & rsProp!TIPO_PAGO
    If Not RSPAGOS.EOF Then
        MiLen1 = Len(rsProp!Z_COUNT)
        Milen2 = Len(Format(rsProp!VALOR, "STANDARD"))
        List1.AddItem "Propina " & FormatTexto(RSPAGOS!DESCRIP, 5) & Space(4 - MiLen1) & rsProp!Z_COUNT & Space(13 - Milen2) & Format(rsProp!VALOR, "STANDARD")
        'PRINT#1, "Propina " & Mid(rsPagos!descrip, 1, 5) & Chr(9) & rsProp!X_COUNT & Chr(9) & Format(rsProp!VALOR, "STANDARD")
    End If
    rsProp.MoveNext
Loop
rsProp.Close

''--ProgBar.Value = 60
List1.AddItem Space(2)
'''''''''Printer.FontUnderline = True
List1.AddItem "AJUSTES"
''''''''Printer.FontUnderline = False
List1.AddItem Space(2)

Do Until rsAjustes.EOF
    MiLen1 = Len(rsAjustes!Z_COUNT)
    Milen2 = Len(Format(rsAjustes!VALOR, "STANDARD"))
    If rsAjustes!CORREC = "EC" Then
        List1.AddItem "CORECCION " & Space(8 - MiLen1) & rsAjustes!Z_COUNT & Space(12 - Milen2) & Format(rsAjustes!VALOR, "STANDARD")
    ElseIf rsAjustes!CORREC = "VO" Then
        List1.AddItem "ANULACION " & Space(8 - MiLen1) & rsAjustes!Z_COUNT & Space(12 - Milen2) & Format(rsAjustes!VALOR, "STANDARD")
    ElseIf rsAjustes!CORREC = "DC" Then
        List1.AddItem "DESCUENTO " & Format(rsAjustes!DESCUENTO, "0.00") & Space(4 - MiLen1) & rsAjustes!Z_COUNT & Space(12 - Milen2) & Format(rsAjustes!VALOR, "STANDARD")
    End If
    rsAjustes.MoveNext
Loop

'IMPRESION DE PAGOS y ABONOS
'Call ImpresionPagos_Abonos(1)

'PREPARA INFO PARA DESCUENTO GLOBAL
rsAjustes.Close
cSQL = "SELECT a.tipo_pago,SUM(a.monto) AS Valor, " & _
        " COUNT(a.tipo_pago) as Z_COUNT " & _
        " FROM HIST_TR_PAGO as a " & _
        " WHERE a.tipo_pago = 99 " & _
        " AND A.Z_COUNTER = '" & nZToPrint & "'" & _
        " GROUP BY a.TIPO_PAGO"

rsAjustes.Open cSQL, msConn, adOpenStatic, adLockOptimistic

If Not rsAjustes.EOF Then
    RSPAGOS.MoveFirst
    RSPAGOS.Find "CODIGO = " & rsAjustes!TIPO_PAGO
    If Not RSPAGOS.EOF Then
        MiLen1 = Len(rsAjustes!Z_COUNT)
        Milen2 = Len(Format(rsAjustes!VALOR, "STANDARD"))
        List1.AddItem FormatTexto(RSPAGOS!DESCRIP, 13) & Space(5 - MiLen1) & rsAjustes!Z_COUNT & Space(12 - Milen2) & Format(rsAjustes!VALOR, "STANDARD")
    End If
End If

''--ProgBar.Value = 70

ssVtatot = rs00!VTA_TOT
ssHashTot = rs00!tot_hash

''''msConn.BeginTrans
'''''''- SEGUN HACIENDA Y TESORO -'''msConn.Execute "UPDATE ORGANIZACION SET TRANS = TRANS + 1, " & _
''''        " VTA_TOT = VTA_TOT + " & nSumTotal & _
''''        ", TOT_HASH = TOT_HASH + " & IIf(IsNull(rsHash!HASH_DIA), 0, rsHash!HASH_DIA)
''''msConn.Execute "UPDATE ORGANIZACION SET VTA_TOT = VTA_TOT + " & nSumTotal & _
''''        ", TOT_HASH = TOT_HASH + " & IIf(IsNull(rsHash!HASH_DIA), 0, rsHash!HASH_DIA)
''''msConn.Execute "UPDATE CAJEROS SET Z_C = Z_C + 1, X_C = 0 " & _
''''        " WHERE NUMERO = 999"
''''msConn.CommitTrans

rsVta_Z.Close
rsPgo_Z.Close
rsTran.Close
rsAjustes.Close

rs00.Requery    'DESPUES DEL COMMIT, LOS DATOS DEBEN DE ESTAR EN EL SERVIDOR

List1.AddItem Space(2)
List1.AddItem "VENTAS ACUMULADAS"
MiLen1 = Len(Format(nSumTotal, "currency"))
List1.AddItem "HOY        : " & Format(nSumTotal, "currency")

List1.AddItem "TOTAL ANT. : " & Format(cZTotalAnterior, "currency")
List1.AddItem "TOTAL NUEVO: " & Format(cZTotalNuevo, "currency")
List1.AddItem Space(2)
'Printer.Print "HASH ANT.  : " & Format(ssHashTot, "currency")
'--- SEGUN HACIENDA Y TESORO -''Printer.Print "HASH NUEVO : " & Format(rs00!tot_hash, "currency")
List1.AddItem "GRAN TOTAL  : " & Format(cZGranTotal, "currency")

For i = 1 To 10
    List1.AddItem Space(2)
Next
''''---Coptr1.CutPaper 100

'------------- DEPARTAMENTAL --------------

''--ProgBar.Value = 80
'rsTran.Open "SELECT DEPTO,COUNT(DEPTO) AS X_COUNT,SUM(PRECIO) AS VALOR "
rsTran.Open "SELECT DEPTO,SUM(CANT) AS X_COUNT,SUM(PRECIO) AS VALOR " & _
        " FROM HIST_TR " & _
        " WHERE Z_COUNTER = '" & nZToPrint & "'" & _
        " GROUP BY DEPTO " & _
        " ORDER BY DEPTO", msConn, adOpenStatic, adLockOptimistic

'''- SEGUN HACIENDA Y TESORO -'''msConn.BeginTrans
'''- SEGUN HACIENDA Y TESORO -'''msConn.Execute "UPDATE ORGANIZACION SET TRANS = TRANS + 1"
'''- SEGUN HACIENDA Y TESORO -'''msConn.CommitTrans

List1.AddItem cZFecha & Space(2) & cZHora
List1.AddItem Space(2)
List1.AddItem "REPORTE DEPARTAMENTAL (Z)"
List1.AddItem Space(2)
List1.AddItem rs00!DESCRIP
List1.AddItem "RUC:" & rs00!RUC
List1.AddItem "SERIAL:" & rs00!SERIAL
List1.AddItem Space(2)
'''- SEGUN HACIENDA Y TESORO -'''Printer.Print "CONTADOR TRANS : " & (rs00!TRANS + 1)
List1.AddItem "CONTADOR Z : " & (nZToPrint)
List1.AddItem "DEPART. REPORTE/TERMINAL"
List1.AddItem Space(2)

Dim nTotDepto As Double

nTotDepto = 0#

Do Until rsTran.EOF
    On Error Resume Next
    rsLocalDepto.MoveFirst
    On Error GoTo 0
    rsLocalDepto.Find "CODIGO = " & rsTran!depto
    If Not rsLocalDepto.EOF Then
        MiLen1 = Len(rsTran!X_COUNT)
        Milen2 = Len(Format(rsTran!VALOR, "STANDARD"))
        List1.AddItem FormatTexto(rsLocalDepto!corto, 13) & Space(4 - MiLen1) & rsTran!X_COUNT & Space(9 - Milen2) & Format(rsTran!VALOR, "STANDARD")
    Else
        'Registro no tiene un departamento valido, IGNORAR
        'MsgBox "Error en Depto. ", vbCritical, BoxTit
    End If
    nTotDepto = nTotDepto + rsTran!VALOR
    rsTran.MoveNext
Loop

List1.AddItem Space(2)
MiLen1 = Len(Format(nTotDepto, "CURRENCY"))
If MiLen1 > 11 Then MiLen1 = 11
List1.AddItem "TOTAL DEPTOS :" & Space(11 - MiLen1) & Format(nTotDepto, "CURRENCY")
List1.AddItem Space(2)

For i = 1 To 10
    List1.AddItem Space(2)
Next
''''---Coptr1.CutPaper 100

rsTran.Close
''--ProgBar.Value = 90

'''msConn.BeginTrans
''''Tambien actualiza el conteo de Z's y X's
''''''- SEGUN HACIENDA Y TESORO -'''msConn.Execute "UPDATE ORGANIZACION SET TRANS = TRANS + 1,X_CDEP = 0,Z_CDEP = Z_CDEP + 1"
'''msConn.Execute "UPDATE ORGANIZACION SET X_CDEP = 0,Z_CDEP = Z_CDEP + 1"
'''msConn.CommitTrans

'--------------- SUPER GRUPOS ------------------------

Dim ccc As String

Set rsSuperGrp = New Recordset

ccc = "SELECT A.GRUPO,A.DESCRIP,SUM(C.PRECIO) AS VENTAS" & _
    " FROM SUPER_GRP AS A,SUPER_DET AS B, HIST_TR AS C" & _
    " Where C.Z_COUNTER = '" & nZToPrint & "' AND " & _
    " A.GRUPO = B.GRUPO And B.DEPTO = C.DEPTO " & _
    " GROUP BY A.GRUPO,A.DESCRIP" & _
    " ORDER BY A.DESCRIP"

rsSuperGrp.Open ccc, msConn, adOpenStatic, adLockOptimistic

'''- SEGUN HACIENDA Y TESORO -'''msConn.BeginTrans
'''- SEGUN HACIENDA Y TESORO -'''msConn.Execute "UPDATE ORGANIZACION SET TRANS = TRANS + 1"
'''- SEGUN HACIENDA Y TESORO -'''msConn.CommitTrans

List1.AddItem cZFecha & Space(2) & cZHora
List1.AddItem Space(2)
List1.AddItem "REPORTE DE GRUPOS (Z)"
List1.AddItem Space(2)
List1.AddItem rs00!DESCRIP
List1.AddItem "RUC:" & rs00!RUC
List1.AddItem "SERIAL:" & rs00!SERIAL
List1.AddItem Space(2)
'''- SEGUN HACIENDA Y TESORO -'''Printer.Print "CONTADOR TRANS : " & (rs00!TRANS + 1)
List1.AddItem "CONTADOR Z : " & (nZToPrint)
List1.AddItem "DEPART. GRUPO/TERMINAL"
List1.AddItem Space(2)

nTotDepto = 0#

Do Until rsSuperGrp.EOF
    Milen2 = Len(Format(rsSuperGrp!VENTAS, "STANDARD"))
    List1.AddItem FormatTexto(rsSuperGrp!DESCRIP, 13) & Space(13 - Milen2) & Format(rsSuperGrp!VENTAS, "STANDARD")
    nTotDepto = nTotDepto + rsSuperGrp!VENTAS
    rsSuperGrp.MoveNext
Loop

List1.AddItem Space(2)
MiLen1 = Len(Format(nTotDepto, "CURRENCY"))
List1.AddItem "TOTAL GRUPOS:" & Space(13 - MiLen1) & Format(nTotDepto, "CURRENCY")
List1.AddItem Space(2)

For i = 1 To 10
    List1.AddItem Space(2)
Next

'rsTran.Close

On Error GoTo 0
If nErrInd = 0 Then
    ''''''''''''''''''''If ON_LINE = True Then BorraLocal
    MsgBox "REPORTE EN 'Z' DE TERMINAL ESTA LISTO", vbInformation, BoxTit
Else
    MsgBox "HA OCURRIDO MAS DE UN ERROR EN EL REPORTE (Z). CONTACTE A SOLO SOFTWARE", vbCritical, BoxTit
End If
''--ProgBar.Value = 0
''''''''''ReportesEscribeLog ("Reporte Z - Actualiza Inventario")
'''ActualizaInvent ' ACTUALIZACION DE INVENTARIO
'''''''ReportesEscribeLog ("Reporte Z - Fin del Reporte")
Exit Sub

AjustaMilen:
errCounter = errCounter + 1
Milen2 = 11
If errCounter < 4 Then
    Resume
Else
    '3021
    If Err.Number <> 3021 Then
        'EL ERROR 3021, ES EL EOF DEL ARCHIVO DE CAJEROS,
        'ESTO OCURRE CUANDO NO HAY TRANSACCIONES, SI HAY OTRO TIPO
        'DE ERROR EL SUPERVISOR LO VERA EN LA PANTALLA
        MsgBox "# " & Err.Number & " ----> " & Err.Description, vbCritical, "ANOTE LOS DATOS EN PANTALLA: " & Err.Source
        'MsgBox "EXISTE UN PROBLEMA DE IMPRESION. UNA VEZ TERMINADO REVISE EL LISTADO Y VERIFIQUE LOS DATOS", vbCritical, "LA IMPRESION DEL REPORTE TIENE PROBLEMAS"
    End If
    'Resume Next
End If

ErrorZZZ:
Dim ADOError As Error
For Each ADOError In msConn.Errors
    sError = sError & ADOError.Number & " - " & ADOError.Description + vbCrLf
Next ADOError
MsgBox "ERROR EN EL REPORTE Z.ANOTE EL NUMERO/DESCRIPCION Y CONTACTE A SOLO SOFTWARE", vbCritical
MsgBox sError, vbCritical, "MENSAJE DE ERROR"

nErrInd = 1
'Resume Next
End Sub
Private Function FormatTexto(texto As String, Largo As Integer) As String
If Largo <= Len(texto) Then
   FormatTexto = Mid$(texto, 1, Largo)
Else
   FormatTexto = texto + Space(Largo - Len(texto))
End If
End Function

Private Sub ImpresionPagos_Abonos(nOpcion As Byte)
'SI OPCION ES 0, ENTONCES ES REPORTE X
'SI OPCION ES 1, ENTONCES ES REPORTE Z
Dim rsPagAbon As New Recordset

rsPagAbon.Open "SELECT TIPO, SUM (MONTO) AS VALOR FROM TMP_VTA_TRANS GROUP BY TIPO", msConn, adOpenStatic, adLockOptimistic
If Not rsPagAbon.EOF Then
    List1.AddItem "============================="
    List1.AddItem "====== PAGOS y ABONOS ======="
Else
    rsPagAbon.Close
    Set rsPagAbon = Nothing
    Exit Sub
End If
rsPagAbon.MoveFirst
Do Until rsPagAbon.EOF
    List1.AddItem IIf(rsPagAbon!TIPO = 0, "ABONOS", "PAGOS ") & _
            Space(5) & Format(rsPagAbon!VALOR, "CURRENCY")
    rsPagAbon.MoveNext
Loop
List1.AddItem "============================="
rsPagAbon.Close
Set rsPagAbon = Nothing
If nOpcion = 1 Then
    msConn.Execute "INSERT INTO VTA_TRANS SELECT * FROM TMP_VTA_TRANS"
    msConn.Execute "DELETE * FROM TMP_VTA_TRANS"
End If
End Sub

Private Sub mnuGo_Click()
Call Print2File
End Sub

Private Function Show_Z_File(cPath As String) As Boolean
'Call Show_Z_File(GetFromINI("General", "DirectorioDatos", App.Path & "\soloini.ini"))
Dim nFreefile As Integer
Dim cFactFile As String
Dim cCadena As String

cFactFile = cPath & "\queue\RZ_10_May_2012_20_25.txt"

nFreefile = FreeFile()
Open cFactFile For Input As #nFreefile
Do Until EOF(1)
    Line Input #nFreefile, A$
    cCadena = cCadena & A$ & vbCrLf
    'If Left(a$, 1) = "*" Then
    '    DATA_PATH = Mid(a$, 3, Len(a$) - 2)
    'Else
    '    cDataPath = a$
    'End If
Loop
txtFile.Text = cCadena
Close #nFreefile

End Function

Private Sub PLU_Click()
Dim iLin  As Integer
Dim nAcum As Single
Dim X As Integer
Dim nPagina As Integer
Dim lastlineskipped As Boolean
Dim nFreefile As Long
Dim cSQL As String
Dim rsPLU As New ADODB.Recordset
Dim nlargo As Integer
Dim nlargo2 As Integer


'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
cSQL = "SELECT PLU, DESCRIP, SUM(CANT) AS UNIDADES, "
cSQL = cSQL & " SUM(PRECIO) AS VENTAS "
cSQL = cSQL & " FROM TRANSAC"
cSQL = cSQL & " WHERE VALID "
cSQL = cSQL & " GROUP BY PLU, DESCRIP "
cSQL = cSQL & " HAVING SUM(PRECIO) >= 0 "
cSQL = cSQL & " ORDER BY 4 DESC"
rsPLU.Open cSQL, msConn, adOpenStatic

nFreefile = FreeFile()
Open App.Path & "\PLU.TXT" For Output As #nFreefile

Print #nFreefile, Format(Date, "SHORT DATE") & "---" & Format(Time(), "HH:MM")
Print #nFreefile, ""
Print #nFreefile, "MAS VENDIDOS HOY"
Print #nFreefile, ""
Print #nFreefile, "======================================"


Do While Not rsPLU.EOF
    nlargo = 6 - Len(rsPLU!UNIDADES)
    nlargo2 = 10 - Len(Format(rsPLU!VENTAS, "STANDARD"))
    Print #nFreefile, FormatTexto(rsPLU!DESCRIP, 20) & _
            Space(nlargo) & "(" & rsPLU!UNIDADES & ")" & _
            Space(nlargo2) & Format(rsPLU!VENTAS, "STANDARD")
    rsPLU.MoveNext
Loop
Print #nFreefile, "======================================"

Close #nFreefile
rsPLU.Close

'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
cSQL = "SELECT A.TIPO_PAGO, B.DESCRIP,  SUM(A.MONTO) AS COBRADO "
cSQL = cSQL & " FROM TRANSAC_PAGO AS A, PAGOS AS B WHERE "
cSQL = cSQL & " A.TIPO_PAGO = B.CODIGO "
cSQL = cSQL & "GROUP BY A.TIPO_PAGO, B.DESCRIP "
rsPLU.Open cSQL, msConn, adOpenStatic

nFreefile = FreeFile()
Open App.Path & "\GAVETA.TXT" For Output As #nFreefile

Print #nFreefile, Format(Date, "SHORT DATE") & "---" & Format(Time(), "HH:MM")
Print #nFreefile, ""
Print #nFreefile, "COBRADO HASTA EL MOMENTO"
Print #nFreefile, ""
Print #nFreefile, "======================================"


Do While Not rsPLU.EOF
    Print #nFreefile, FormatTexto(rsPLU!DESCRIP, 15) & FormatTexto(Format(rsPLU!COBRADO, "STANDARD"), 8)
    rsPLU.MoveNext
Loop
Print #nFreefile, "======================================"

Close #nFreefile
rsPLU.Close

'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
'''''''''cSQL = "SELECT A.G_DESCRIP, SUM(A.G_VALOR * B.CANT)  AS CANTIDAD, "
'''''''''cSQL = cSQL & " SUM(ABS(B.CANT) * B.PRECIO_UNIT) AS VENTAS, 0 AS VTA_NETA "
'''''''''cSQL = cSQL & " FROM G_GRUPOS AS A, TRANSAC AS B"
'''''''''cSQL = cSQL & " WHERE A.G_PLU = B.PLU"
'''''''''cSQL = cSQL & " AND '%' NOT IN (B.DESCRIP) "
'''''''''cSQL = cSQL & " AND B.DESCRIP NOT LIKE '%DESCUENTO%' "
'''''''''cSQL = cSQL & " AND B.DESCRIP NOT LIKE  '%@@%' "
'''''''''cSQL = cSQL & " GROUP BY A.G_DESCRIP"
'''''''''rsPLU.Open cSQL, msConn, adOpenStatic
'''''''''
'''''''''nFreefile = FreeFile()
'''''''''Open App.Path & "\GRANDES.TXT" For Output As #nFreefile
'''''''''
'''''''''Print #nFreefile, Format(Date, "SHORT DATE") & "---" & Format(Time(), "HH:MM")
'''''''''Print #nFreefile, ""
'''''''''Print #nFreefile, "GRANDES GRUPOS"
'''''''''Print #nFreefile, ""
'''''''''Print #nFreefile, "======================================"
'''''''''
'''''''''
'''''''''Do While Not rsPLU.EOF
'''''''''    Print #nFreefile, FormatTexto(rsPLU!G_DESCRIP, 20) & " (" & FormatTexto(Format(rsPLU!CANTIDAD), 4) & ")  " & FormatTexto(Format(rsPLU!VENTAS, "STANDARD"), 8)
'''''''''    'Print #nFreefile, FormatTexto(rsPLU!DESCRIP, 15) & FormatTexto(Format(rsPLU!COBRADO, "STANDARD"), 8)
'''''''''    rsPLU.MoveNext
'''''''''Loop
'''''''''Print #nFreefile, "======================================"
'''''''''
'''''''''Close #nFreefile
'''''''''rsPLU.Close


'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
cSQL = "SELECT DESCRIP, SUM(CANT) AS  CANTIDAD, SUM(PRECIO) AS DESCUENTOS "
cSQL = cSQL & " FROM TRANSAC "
cSQL = cSQL & " WHERE LEFT(TIPO,3) = 'DC-' AND VALID "
cSQL = cSQL & " GROUP BY DESCRIP"
rsPLU.Open cSQL, msConn, adOpenStatic

nFreefile = FreeFile()
Open App.Path & "\DESCUENTO.TXT" For Output As #nFreefile

Print #nFreefile, Format(Date, "SHORT DATE") & "---" & Format(Time(), "HH:MM")
Print #nFreefile, ""
Print #nFreefile, "DESCUENTOS, CORRECCIONES y ANULACIONES"
Print #nFreefile, ""
Print #nFreefile, "=============DESCUENTOS==============="


Do While Not rsPLU.EOF
    Print #nFreefile, FormatTexto(rsPLU!DESCRIP, 20) & " (" & FormatTexto(Format(rsPLU!CANTIDAD), 4) & ")  " & FormatTexto(Format(rsPLU!DESCUENTOS, "STANDARD"), 8)
    'Print #nFreefile, FormatTexto(rsPLU!DESCRIP, 15) & FormatTexto(Format(rsPLU!COBRADO, "STANDARD"), 8)
    rsPLU.MoveNext
Loop
Print #nFreefile, "======================================"
Print #nFreefile, ""
Print #nFreefile, "=============CORRECIONES=============="

rsPLU.Close

'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
cSQL = "SELECT FECHA, SUM(CANT) AS  CANTIDAD, SUM(PRECIO) AS CORRECCIONES "
cSQL = cSQL & " FROM TRANSAC "
cSQL = cSQL & " WHERE LEFT(TIPO,3) = 'EC-' AND VALID "
cSQL = cSQL & " GROUP BY FECHA"
rsPLU.Open cSQL, msConn, adOpenStatic

Do While Not rsPLU.EOF
    Print #nFreefile, FormatTexto("CORRECCIONES", 20) & " (" & FormatTexto(Format(rsPLU!CANTIDAD), 4) & ")  " & FormatTexto(Format(rsPLU!CORRECCIONES, "STANDARD"), 8)
    'Print #nFreefile, FormatTexto(rsPLU!DESCRIP, 15) & FormatTexto(Format(rsPLU!COBRADO, "STANDARD"), 8)
    rsPLU.MoveNext
Loop
Print #nFreefile, "======================================"
Print #nFreefile, ""
Print #nFreefile, "=============ANULACIONES=============="

rsPLU.Close
cSQL = "SELECT FECHA, SUM(CANT) AS  CANTIDAD, SUM(PRECIO) AS ANULACIONES "
cSQL = cSQL & " FROM TRANSAC "
cSQL = cSQL & " WHERE LEFT(TIPO,3) = 'VO-' AND VALID "
cSQL = cSQL & " GROUP BY FECHA"
rsPLU.Open cSQL, msConn, adOpenStatic

Do While Not rsPLU.EOF
    Print #nFreefile, FormatTexto("ANULACIONES", 20) & " (" & FormatTexto(Format(rsPLU!CANTIDAD), 4) & ")  " & FormatTexto(Format(rsPLU!ANULACIONES, "STANDARD"), 8)
    'Print #nFreefile, FormatTexto(rsPLU!DESCRIP, 15) & FormatTexto(Format(rsPLU!COBRADO, "STANDARD"), 8)
    rsPLU.MoveNext
Loop
Print #nFreefile, "======================================"

Close #nFreefile
rsPLU.Close

'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
cSQL = "SELECT FECHA, FORMAT(HORA,'##:##') AS HORA,  "
cSQL = cSQL & " FORMAT((TOTAL_NUEVO - TOTAL_ANTERIOR),'CURRENCY') AS VENTAS, "
cSQL = cSQL & " FORMAT(ITBMS, 'CURRENCY') AS ITBMS "
cSQL = cSQL & " FROM Z_COUNTER  "
cSQL = cSQL & " WHERE LEFT(FECHA,6) = '202206' ORDER BY ID"
rsPLU.Open cSQL, msConn, adOpenStatic

nFreefile = FreeFile()
Open App.Path & "\VENTA_MENSUAL.TXT" For Output As #nFreefile

Print #nFreefile, Format(Date, "SHORT DATE") & "---" & Format(Time(), "HH:MM")
Print #nFreefile, ""
Print #nFreefile, "VENTAS DIARIAS MES ACTUAL"
Print #nFreefile, ""
Print #nFreefile, "DIA  HORA     VENTAS      IMPUESTO"
Print #nFreefile, "=================================="

Do While Not rsPLU.EOF
    Print #nFreefile, Right(rsPLU!FECHA, 2) & " (" & rsPLU!HORA & ")  " & _
            FormatTexto(rsPLU!VENTAS, 15) & FormatTexto(rsPLU!ITBMS, 8)
    rsPLU.MoveNext
Loop

Print #nFreefile, "=================================="

Close #nFreefile
rsPLU.Close


'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
cSQL = "SELECT A.MESERO AS MES, B.NOMBRE, B.APELLIDO,  COUNT(A.NUM_TRANS) AS TRANS,  "
cSQL = cSQL & " SUM(A.PRECIO) AS VENTAS  FROM HIST_TR AS A LEFT JOIN MESEROS AS B  "
cSQL = cSQL & " ON A.MESERO = B.NUMERO  "
cSQL = cSQL & " WHERE A.FECHA BETWEEN '20190601' AND "
cSQL = cSQL & " '20190630' "
cSQL = cSQL & " AND '%' NOT IN (A.DESCRIP)  AND A.DESCRIP NOT LIKE '%DESCUENTO%'  "
cSQL = cSQL & " AND A.DESCRIP NOT LIKE  '%@@%'  GROUP BY A.MESERO, B.NOMBRE, B.APELLIDO"
cSQL = cSQL & " ORDER BY SUM(A.PRECIO) DESC "
rsPLU.Open cSQL, msConn, adOpenStatic

nFreefile = FreeFile()
Open App.Path & "\MESEROS_MES_ACTUAL.TXT" For Output As #nFreefile

Print #nFreefile, Format(Date, "SHORT DATE") & "---" & Format(Time(), "HH:MM")
Print #nFreefile, ""
Print #nFreefile, "VENTAS x MESEROS (MES ACTUAL)"
Print #nFreefile, ""
Print #nFreefile, "NOMBRE                TRANS       VENTAS"
Print #nFreefile, "========================================"

Do While Not rsPLU.EOF
    If IsNull(rsPLU!NOMBRE) Then GoTo MoveProximo:
    nlargo = 6 - Len(rsPLU!TRANS)
    nlargo2 = 12 - Len(Format(rsPLU!VENTAS, "STANDARD"))
    Print #nFreefile, FormatTexto(rsPLU!NOMBRE & " (APELLIDO) ", 20) & Space(nlargo) & _
            "(" & rsPLU!TRANS & ")" & Space(nlargo2) & Format(rsPLU!VENTAS, "STANDARD")

MoveProximo:
    rsPLU.MoveNext
Loop

Print #nFreefile, "========================================"

Close #nFreefile
rsPLU.Close

'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
cSQL = "SELECT A.MESERO AS MES, B.NOMBRE, B.APELLIDO,  COUNT(A.NUM_TRANS) AS TRANS,  "
cSQL = cSQL & " SUM(A.PRECIO) AS VENTAS  "
cSQL = cSQL & " FROM TRANSAC AS A LEFT JOIN MESEROS AS B  "
cSQL = cSQL & " ON A.MESERO = B.NUMERO  "
cSQL = cSQL & " WHERE '%' NOT IN (A.DESCRIP)  "
cSQL = cSQL & " AND A.DESCRIP NOT LIKE '%DESCUENTO%'  "
cSQL = cSQL & " AND A.DESCRIP NOT LIKE  '%@@%'  "
cSQL = cSQL & " GROUP BY A.MESERO, B.NOMBRE, B.APELLIDO"
cSQL = cSQL & " ORDER BY SUM(A.PRECIO) DESC "
rsPLU.Open cSQL, msConn, adOpenStatic

nFreefile = FreeFile()
Open App.Path & "\MESEROS_HOY.TXT" For Output As #nFreefile

Print #nFreefile, Format(Date, "SHORT DATE") & "---" & Format(Time(), "HH:MM")
Print #nFreefile, ""
Print #nFreefile, "VENTAS x MESEROS (HOY)"
Print #nFreefile, ""
Print #nFreefile, "NOMBRE                TRANS     VENTAS"
Print #nFreefile, "======================================"

Do While Not rsPLU.EOF
    ''If IsNull(rsPLU!NOMBRE) Then GoTo MoveProximo:
    nlargo = 13 - Len(Format(rsPLU!VENTAS, "STANDARD"))
    Print #nFreefile, FormatTexto(rsPLU!NOMBRE & " (APELLIDO) ", 20) & "    " & _
            rsPLU!TRANS & Space(nlargo) & Format(rsPLU!VENTAS, "STANDARD")
    rsPLU.MoveNext
Loop

Print #nFreefile, "======================================"

Close #nFreefile
rsPLU.Close


'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
cSQL = "SELECT HOUR(HORA) AS HORA, MAX(HOUR(HORA)) & ':59' AS FIN,  SUM(PRECIO) AS VENTAS,"
cSQL = cSQL & " SUM(CANT) AS ITEMS, MAX(DESCRIP) AS M_DESCRIP "
cSQL = cSQL & " FROM HIST_TR  "
cSQL = cSQL & " WHERE  FECHA >= '20220101' AND FECHA <= '20220618' "
cSQL = cSQL & " AND '%' NOT IN (DESCRIP)  AND DESCRIP NOT LIKE '%DESCUENTO%'  "
cSQL = cSQL & " AND DESCRIP NOT LIKE  '%@@%'  "
cSQL = cSQL & " GROUP BY  HOUR(HORA)"

rsPLU.Open cSQL, msConn, adOpenStatic

nFreefile = FreeFile()
Open App.Path & "\VENTAS_X_HORA.TXT" For Output As #nFreefile

Print #nFreefile, Format(Date, "SHORT DATE") & "---" & Format(Time(), "HH:MM")
Print #nFreefile, ""
Print #nFreefile, "VENTAS x HORA (AÑO ACTUAL)"
Print #nFreefile, ""
Print #nFreefile, "INICIO  FIN         VENTAS       ITEMS"
Print #nFreefile, "======================================"

Do While Not rsPLU.EOF
    nlargo = 17 - Len(Format(rsPLU!VENTAS, "STANDARD"))
    nlargo2 = 9 - Len(Format(rsPLU!ITEMS, "###,##0"))
    'If IsNull(rsPLU!NOMBRE) Then GoTo MoveProximo:
    Print #nFreefile, Format(rsPLU!HORA, "00") & "     " & Format(rsPLU!HORA, "00") & ":59" & _
            Space(nlargo) & Format(rsPLU!VENTAS, "STANDARD") & Space(nlargo2) & Format(rsPLU!ITEMS, "###,##0")
'    Print #nFreefile, FormatTexto(rsPLU!NOMBRE & " " & rsPLU!APELLIDO, 20) & "  " & _
            FormatTexto(rsPLU!TRANS, 6) & "    " & Format(rsPLU!VENTAS, "CURRENCY")
 
'MoveProximo:
    rsPLU.MoveNext
Loop

Print #nFreefile, "======================================"

Close #nFreefile
rsPLU.Close

'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~


cSQL = "SELECT FORMAT(HOUR(HORA),'00') AS HORA, "
cSQL = cSQL & " SUM(PRECIO) AS VENTAS "
cSQL = cSQL & " FROM TRANSAC "
cSQL = cSQL & " GROUP BY FORMAT(HOUR(HORA),'00')"

rsPLU.Open cSQL, msConn, adOpenStatic

nFreefile = FreeFile()
Open App.Path & "\VENTAS_HORA_HOY.TXT" For Output As #nFreefile

Print #nFreefile, Format(Date, "SHORT DATE") & "---" & Format(Time(), "HH:MM")
Print #nFreefile, ""
Print #nFreefile, "VENTAS x HORA (HOY)"
Print #nFreefile, ""
Print #nFreefile, "HORA         VENTAS"
Print #nFreefile, "==================="

Do While Not rsPLU.EOF
    nlargo = 17 - Len(Format(rsPLU!VENTAS, "STANDARD"))
    Print #nFreefile, rsPLU!HORA & Space(nlargo) & Format(rsPLU!VENTAS, "STANDARD")
    rsPLU.MoveNext
Loop

Print #nFreefile, "==================="

Close #nFreefile
rsPLU.Close

'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

cSQL = "SELECT MESA, MIN(HORA) AS APERTURA, SUM(PRECIO) AS VENTAS "
cSQL = cSQL & " FROM TMP_TRANS "
cSQL = cSQL & " WHERE VALID "
cSQL = cSQL & " GROUP BY MESA"

rsPLU.Open cSQL, msConn, adOpenStatic

nFreefile = FreeFile()
Open App.Path & "\MESAS_ABIERTAS.TXT" For Output As #nFreefile

Print #nFreefile, Format(Date, "SHORT DATE") & "---" & Format(Time(), "HH:MM")
Print #nFreefile, ""
Print #nFreefile, "MESAS ABIERTAS (SIN COBRAR)"
Print #nFreefile, ""
Print #nFreefile, "MESA        APERTURA      VENTAS"
Print #nFreefile, "================================"

Do While Not rsPLU.EOF
    'If IsNull(rsPLU!NOMBRE) Then GoTo MoveProximo:
    
    nlargo = 12 - Len(Format(rsPLU!VENTAS, "STANDARD"))
    Print #nFreefile, Format(rsPLU!MESA, "000") & "     " & _
            Format(rsPLU!APERTURA, "@@@@@@@@@@@@") & _
            Space(nlargo) & _
            Format(rsPLU!VENTAS, "STANDARD")
'    Print #nFreefile, FormatTexto(rsPLU!NOMBRE & " " & rsPLU!APELLIDO, 20) & "  " & _
            FormatTexto(rsPLU!TRANS, 6) & "    " & Format(rsPLU!VENTAS, "CURRENCY")
 
'MoveProximo:
    rsPLU.MoveNext
Loop

Print #nFreefile, "================================"

Close #nFreefile
rsPLU.Close

'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~


cSQL = "SELECT WEEKDAY(FORMAT(A.FECHA,'####-##-##')) as DIAN, "
cSQL = cSQL & " MAX(B.DIA_NOMBRE) AS DIA, "
cSQL = cSQL & " CDbl(SUM(A.PRECIO)) AS VENTAS "
cSQL = cSQL & " FROM HIST_TR AS A, DIA AS B "
cSQL = cSQL & " WHERE LEFT(A.FECHA,6) = '202205' AND A.VALID "
cSQL = cSQL & " AND WEEKDAY(FORMAT(A.FECHA,'####-##-##')) = B.DIA_NUM "
cSQL = cSQL & " GROUP BY WEEKDAY(FORMAT(A.FECHA,'####-##-##'))"
rsPLU.Open cSQL, msConn, adOpenStatic

Call toJSON(cSQL, "ventasdiasemana_mes")

nFreefile = FreeFile()
Open App.Path & "\VENTA_WEEKDAY.TXT" For Output As #nFreefile

Print #nFreefile, Format(Date, "SHORT DATE") & "---" & Format(Time(), "HH:MM")
Print #nFreefile, ""
Print #nFreefile, "VENTAS SEMANA (MES ACTUAL)"
Print #nFreefile, ""
Print #nFreefile, "DIA       VENTAS"
Print #nFreefile, "================"

Do While Not rsPLU.EOF
    nlargo = 13 - Len(Format(rsPLU!VENTAS, "STANDARD"))
    Print #nFreefile, rsPLU!DIA & Space(nlargo) & Format(rsPLU!VENTAS, "STANDARD")
    rsPLU.MoveNext
Loop

Print #nFreefile, "================"

Close #nFreefile
rsPLU.Close

'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

cSQL = "SELECT A.DEPTO, A.CODIGO, "
cSQL = cSQL & "IIF(A.CODIGO=C.CODIGO,C.CONTENEDOR,0) AS ENVASE, "
cSQL = cSQL & "A.DESCRIP, FORMAT(A.PRECIO1,'STANDARD') AS PRECIO, "
cSQL = cSQL & "A.DISPONIBLE, "
cSQL = cSQL & "FORMAT(C.PRECIO,'STANDARD') AS PRECIO_ENV "
'INFO: 10MAY2012
cSQL = cSQL & "INTO LOLO1 "
cSQL = cSQL & "FROM PLU as A LEFT JOIN CONTEND_02 as C ON A.CODIGO = C.CODIGO "

msConn.BeginTrans
msConn.Execute cSQL
msConn.CommitTrans

cSQL = "SELECT A.DEPTO, A.CODIGO, A.ENVASE AS ENV,"
cSQL = cSQL & "IIF(A.ENVASE = B.CONTENEDOR,B.DESCRIP,'Sin Envase') AS ENVASE,"
'INFO: 10MAY2012
cSQL = cSQL & " iif(A.DISPONIBLE, A.DESCRIP, '~' & A.DESCRIP) AS DESCRIP, "
'cSQL = cSQL & "IIF(A.DISPONIBLE=TRUE, A.DESCRIP, '~' + A.DESCRIP) AS DESCRIP,"
cSQL = cSQL & "A.PRECIO, A.PRECIO_ENV "
cSQL = cSQL & "FROM LOLO1 as A LEFT JOIN CONTENED AS B "
cSQL = cSQL & " ON A.ENVASE = B.CONTENEDOR "
'cSQL = cSQL & "ORDER BY B.DESCRIP, A.DESCRIP"
cSQL = cSQL & "ORDER BY 4, 5"
rsPLU.Open cSQL, msConn, adOpenStatic, adLockOptimistic

Do While Not rsPLU.EOF
    'nlargo = 13 - Len(Format(rsPLU!VENTAS, "STANDARD"))
    Debug.Print rsPLU!ENVASE & Space(5) & rsPLU!DESCRIP & Space(5) & Format(rsPLU!PRECIO_ENV, "STANDARD")
    rsPLU.MoveNext
Loop

rsPLU.Close
msConn.BeginTrans
msConn.Execute "DROP TABLE LOLO1"
msConn.CommitTrans

'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

Set rsPLU = Nothing

End Sub

Private Function GetDia(nDia As Integer) As String
Select Case nDia
    Case 1: GetDia = "LUN"
    Case 2: GetDia = "MAR"
    Case 3: GetDia = "MIE"
    Case 4: GetDia = "JUE"
    Case 5: GetDia = "VIE"
    Case 6: GetDia = "SAB"
    Case 7: GetDia = "DOM"
End Select
End Function
Private Function GetDepto(nDepto As Integer) As String

Select Case nDia
    Case 1: GetDia = "LUN"
    Case 2: GetDia = "MAR"
    Case 3: GetDia = "MIE"
    Case 4: GetDia = "JUE"
    Case 5: GetDia = "VIE"
    Case 6: GetDia = "SAB"
    Case 7: GetDia = "DOM"
End Select
End Function


'---------------------------------------------------------------------------------------
' Procedure : toJSON
' Author     : hsequeira
' Date        : 26/10/2021
' Date        :  3/07/2022 -> SE ESCLARECE PRESENTACION EN EL FILE JSON
' Purpose   : PASA UN RECORDSET A JSON
' 8DIC2021 : HEADER AND FOOTER
'---------------------------------------------------------------------------------------
'
Function toJSON(PassTblQry As String, Optional newName As String)
' EXPORT JSON FILE FROM TABLE OR QUERY
'Dim mydb As Database, rs As Recordset
'Dim VarField(255), VarFieldType(255)
'Dim fld As DAO.Field, cValorCampo As String
'Dim db As DAO.Database
'
'db.op
'Set Db = CurrentDb
   On Error GoTo toJSON_Error

'Set db = msConn

Dim cSQL As String
Dim rsTabla As ADODB.Recordset
Dim TipoCampo As Integer
Dim BuscaCadenaError As Integer

Set rsTabla = New ADODB.Recordset

If newName = "" Then newName = "QUERY " & Format(Now(), "YYYY-MM-DD HHMM")

If Left(PassTblQry, 6) = "SELECT" Then
    cSQL = PassTblQry
    fn = CurDir(App.Path) & "\" & newName & ".json"   ' define export current folder query date/time
    'jsonTableName = PassTblQry   // info 21jun200
    jsonTableName = "sql_query"
Else
    cSQL = MakeSQL(PassTblQry)
    fn = CurDir(App.Path) & "\" & newName & ".json"  ' define export current folder query date/time
End If

rsTabla.Open cSQL, msConn, adOpenStatic, _
        IIf(vbTrue, adLockReadOnly, adLockOptimistic)

'fn = CurrentProject.Path & "\" & PassTblQry & " " & Format(Now(), "YYYY-MM-DD HHMM") & ".json" ' define export current folder query date/time
Open fn For Output As #1    ' output to text file
'Recs = DCount("*", PassTblQry) ' record count
'Set rs = db.OpenRecordset("Select * from [" & PassTblQry & "]")
Nonulls = True ' set NoNulls = true to remove all null values within output ELSE set to false
fieldcount = 0
' Save field count, fieldnames, and type into array
For Each fld In rsTabla.Fields
    fieldcount = fieldcount + 1
    '''VarField(fieldcount) = fld.Name
    'Debug.Print VarField(fieldcount)
    '''''''''''VarFieldType(fieldcount) = "TEXT"
    Select Case fld.Type
        Case 4, 5, 6, 7 ' fieldtype 4=long, 5=Currency, 6=Single, 7-Double
            '''''''''''VarFieldType(fieldcount) = "NUMBER"
    End Select
Next
Set fld = Nothing

'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
' Print #1, "{ "                                                          ' HEADER JSON
' Print #1, Chr(34) & jsonTableName & Chr(34) & " : ["   ' Start; JSON; dataset
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
Print #1, "[" 'Start; JSON; dataset

'Print #1, " : [" ' start JSON dataset
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
'' Print #1, "[" ' start JSON dataset
' build JSON dataset from table/query data passed
Do While Not rsTabla.EOF
    '~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
    Print #1, Chr(9) & "{"  ' START JSON record
    '~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
    ' build JSON record from table/query record using fieldname and type arrays
    'For nCampoActual = 1 To fieldcount
    For nCampoActual = 0 To fieldcount - 1
        '''''''''''TipoCampo = VarFieldType(nCampoActual)
        'TipoCampo = VarFieldType(nCampoActual)
        TipoCampo = rsTabla.Fields(nCampoActual).Type
        Select Case TipoCampo
            Case 2      'SMALL INT
                QuoteID = ""
            Case 3      'INTEGER
                QuoteID = ""
            Case 5      'DOUBLE
                QuoteID = ""
            Case 6      'CURRENCY
                QuoteID = Chr(34)
                'QuoteID = ""
            Case 11     'True o False // debe ser true o false en JSON
                QuoteID = ""
            Case 202    'STRING o VARCHAR
                QuoteID = Chr(34)
            Case Else
                QuoteID = Chr(34)
        End Select
        
        'If TipoCampo = 3 Or 11 Then QuoteID = "" Else QuoteID = Chr(34)     ' No quote for numbers
        'QuoteID = Chr(34) ' double quote for text
        '''''''''''If IsNull(rs(VarField(nCampoActual)).Value) Then  ' deal with null values
        '''''''''''    cValorCampo = "Null": QuoteID = ""   ' no quote for nulls
        '''''''''''    If Nonulls = True Then cValorCampo = "": QuoteID = Chr(34)                       ' null text to empty quotes
        '''''''''''    If Nonulls = True And TipoCampo = "NUMBER" Then cValorCampo = "0": QuoteID = ""      ' null number to zero without quotes
        '''''''''''    Else
        '''''''''''    cValorCampo = Trim(rs(VarField(nCampoActual)).Value)
        '''''''''''End If
        
        If IsNull(rsTabla.Fields(nCampoActual).Value) Then  ' deal with null values
            cValorCampo = "Null": QuoteID = ""   ' no quote for nulls
            If Nonulls = True Then cValorCampo = "": QuoteID = Chr(34)                       ' null text to empty quotes
            If Nonulls = True And TipoCampo = 3 Then cValorCampo = "0": QuoteID = ""      ' null number to zero without quotes
            Else
                Select Case TipoCampo
                    Case 11
                        Select Case rsTabla.Fields(nCampoActual).Value
                            Case True
                                cValorCampo = "true"
                            Case False
                                cValorCampo = "false"
                        End Select
                    Case Else
                        cValorCampo = Trim(rsTabla.Fields(nCampoActual).Value)
                End Select
        End If
        
        BuscaCadenaError = InStr(cValorCampo, "'")
        If BuscaCadenaError > 0 Then
            cValorCampo = Replace(cValorCampo, "'", "-")
        End If
        
        cValorCampo = Replace(cValorCampo, Chr(34), "'") ' replace double quote with single quote
        cValorCampo = Replace(cValorCampo, Chr(8), "")   ' remove backspace
        cValorCampo = Replace(cValorCampo, Chr(10), "")  ' remove line feed
        cValorCampo = Replace(cValorCampo, Chr(12), "")  ' remove form feed
        cValorCampo = Replace(cValorCampo, Chr(13), "")  ' remove carriage return
        cValorCampo = Replace(cValorCampo, Chr(9), "   ")  ' replace tab with spaces
        
        '''''''''''jsonRow = Chr(34) & VarField(nCampoActual) & Chr(34) & ":" & QuoteID & cValorCampo & QuoteID
        'Debug.Print TipoCampo & " - " & rsTabla.Fields(nCampoActual).Name & " - " & cValorCampo
        jsonRow = Chr(34) & LCase(rsTabla.Fields(nCampoActual).Name) & Chr(34) & ":" & QuoteID & cValorCampo & QuoteID
        If nCampoActual < fieldcount - 1 Then jsonRow = jsonRow & "," ' add comma if not last field
        
        Print #1, Chr(9) & Chr(9) & jsonRow
    
    Next nCampoActual
    '~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
    '~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
    Print #1, Chr(9) & "}";  ' END JSON record
    '~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
    '~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
    rsTabla.MoveNext
    If Not rsTabla.EOF Then
        Print #1, "," ' add comma if not last record
        Else
        Print #1, ""
    End If
Loop

'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
Print #1, "]"  'FOOTER JSON FILE
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
'''Print #1, Chr(34) & "Info" & Chr(34) & " : {"
'''Print #1, Chr(34) & "Fecha" & Chr(34) & " : " & Chr(34) & Date & Chr(34) & ","
'''Print #1, Chr(34) & "Hora" & Chr(34) & " : " & Chr(34) & Time & Chr(34) & ","
'''Print #1, Chr(34) & "Autor" & Chr(34) & " : " & Chr(34) & App.EXEName & " - " & App.CompanyName & Chr(34)
'''Print #1, "}"
'''
'''Print #1, "}"  ' close JSON dataset
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
Close #1

   On Error GoTo 0
   Exit Function

toJSON_Error:

    ShowMsg "Error " & Err.Number & vbCrLf & " (" & Err.Description & ") in procedure toJSON of Form Browser"
    Resume Next

End Function


Private Function MakeSQL(cTabla As String) As String
Dim cSQL As String
Select Case cTabla
    Case "HIST_INVENT"
        cSQL = "SELECT A.ID, B.DESCRIP, A.BOD1_01, A.BOD2_01, A.COSTO_01, A.TOTAL_01, A.BOD1_02, A.BOD2_02, A.COSTO_02, A.TOTAL_02    ,A.BOD1_03 ,A.BOD2_03 ,A.COSTO_03    ,A.TOTAL_03    ,A.BOD1_04 ,A.BOD2_04 ,A.COSTO_04,A.TOTAL_04"
        cSQL = cSQL & ",A.BOD1_05, A.BOD2_05, A.COSTO_05, A.TOTAL_05, A.BOD1_06, A.BOD2_06, A.COSTO_06    ,A.TOTAL_06    ,A.BOD1_07 ,A.BOD2_07, A.COSTO_07, A.TOTAL_07, A.BOD1_08, A.BOD2_08, A.COSTO_08, A.TOTAL_08, A.BOD1_09 "
        cSQL = cSQL & ",A.BOD2_09 ,A.COSTO_09, A.TOTAL_09, A.BOD1_10, A.BOD2_10, A.COSTO_10, A.TOTAL_10, A.BOD1_11, A.BOD2_11, A.COSTO_11, A.TOTAL_11, A.BOD1_12, A.BOD2_12, A.COSTO_12, A.TOTAL_12"
        cSQL = cSQL & " FROM HIST_INVENT AS A, INVENT AS B "
        cSQL = cSQL & " WHERE A.ID = B.ID "
        MakeSQL = cSQL
    Case Else
        MakeSQL = "SELECT * FROM " & cTabla
End Select
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

