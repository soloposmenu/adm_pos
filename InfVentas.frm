VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.Form InfVentas 
   BackColor       =   &H00B39665&
   Caption         =   "Informe de Ventas e Impuesto (No Incluye Descuentos Globales)"
   ClientHeight    =   7005
   ClientLeft      =   60
   ClientTop       =   1305
   ClientWidth     =   12075
   Icon            =   "InfVentas.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   7005
   ScaleWidth      =   12075
   Begin VB.CommandButton Command1 
      BackColor       =   &H0000FFFF&
      Height          =   495
      Left            =   1680
      Picture         =   "InfVentas.frx":6852
      Style           =   1  'Graphical
      TabIndex        =   6
      ToolTipText     =   "Envia Seleccion a la Impresora"
      Top             =   4920
      Width           =   735
   End
   Begin VB.PictureBox Picture1 
      Height          =   495
      Left            =   360
      ScaleHeight     =   435
      ScaleWidth      =   1155
      TabIndex        =   5
      Top             =   5880
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.CommandButton cmdPrint 
      Height          =   495
      Left            =   120
      TabIndex        =   4
      Top             =   4920
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.ListBox ListAnno 
      Height          =   645
      Left            =   1320
      TabIndex        =   0
      Top             =   240
      Width           =   975
   End
   Begin MSComctlLib.ListView LV 
      Height          =   3255
      Left            =   120
      TabIndex        =   1
      Top             =   1320
      Width           =   2295
      _ExtentX        =   4048
      _ExtentY        =   5741
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   14737632
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   0
   End
   Begin MSComctlLib.ListView LVVentas 
      Height          =   6735
      Left            =   2520
      TabIndex        =   2
      Top             =   120
      Width           =   9495
      _ExtentX        =   16748
      _ExtentY        =   11880
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   14737632
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   0
   End
   Begin VB.Label Label1 
      BackColor       =   &H00B39665&
      Caption         =   "Seleccione Mes"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   120
      TabIndex        =   7
      Top             =   1080
      Width           =   1575
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackColor       =   &H00B39665&
      Caption         =   "Seleccione Año"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   120
      TabIndex        =   3
      Top             =   360
      Width           =   975
   End
End
Attribute VB_Name = "InfVentas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'NOVIEMBRE DE 2010
Private Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Private Declare Function GetDC Lib "user32" (ByVal hwnd As Long) As Long
Private nPagina As Integer
Private iLin As Integer
Private IsFAST As Boolean
Private Function Seguridad() As String
'SETUP DE SEGURIDAD DEL SISTEMA
Dim cSeguridad As String

cSeguridad = GetSecuritySetting(npNumCaj, Me.Name)
Select Case cSeguridad
    Case "CEMV"
        'INFO: NO HAY RESTRICCIONES
    Case "CMV"
        'INFO: NO HAY RESTRICCIONES
    Case "CV"
        'INFO: NO HAY RESTRICCIONES
    Case "V"
        'INFO: NO HAY RESTRICCIONES
    Case "N"
        ListAnno.Enabled = False
End Select
End Function

Private Sub PrintTit()

If nPagina = 0 Then
    MainMant.spDoc.WindowTitle = "Impresión de " & Me.Caption
    MainMant.spDoc.FirstPage = 1
    MainMant.spDoc.PageOrientation = SPOR_PORTRAIT
    MainMant.spDoc.Units = SPUN_LOMETRIC
End If
MainMant.spDoc.Page = nPagina + 1

MainMant.spDoc.TextOut 300, 200, Format(Date, "long date") & "  " & Time
MainMant.spDoc.TextOut 300, 250, "Página : " & nPagina + 1
MainMant.spDoc.TextOut 300, 350, rs00!DESCRIP
MainMant.spDoc.TextOut 300, 450, Me.Caption
MainMant.spDoc.TextOut 300, 500, "PERIODO : " & LV.SelectedItem.Text & " / " & ListAnno.Text

MainMant.spDoc.TextOut 300, 650, "FECHA"
MainMant.spDoc.TextOut 530, 650, "HORA"
MainMant.spDoc.TextOut 750, 650, "ANTERIOR"
MainMant.spDoc.TextOut 1180, 650, "NUEVO"
'MainMant.spDoc.TextOut 1350, 650, "VTAS BRUTAS"
MainMant.spDoc.TextOut 1350, 650, "VTAS NETAS"
MainMant.spDoc.TextOut 1630, 650, "IMPUESTO"
MainMant.spDoc.TextOut 1850, 650, "EXONER."
MainMant.spDoc.TextOut 300, 700, String(145, "-")

iLin = 750
nPagina = nPagina + 1
End Sub


Private Sub Command1_Click()
Dim iCtr As Integer 'Contador de Linea
Dim iCol, iFil As Integer 'Contador de Columnas
Dim cText As String
Dim ispace As Integer
Dim iLen As Integer
Dim sSubTot As Single
Dim i As Integer

sSubTot = 0#: iLin = 8: nPagina = 0

On Error GoTo ErrorPrn:

MainMant.spDoc.DocBegin
PrintTit
EscribeLog ("Admin." & "Impresion de Listado: " & Me.Caption & " AÑO" & ListAnno.Text & " MES: " & LV.SelectedItem.Text)
'
'MainMant.spDoc.TextOut 300, 650, "FECHA"
'MainMant.spDoc.TextOut 530, 650, "HORA"
'MainMant.spDoc.TextOut 750, 650, "ANTERIOR"
'MainMant.spDoc.TextOut 1180, 650, "NUEVO"
'MainMant.spDoc.TextOut 1300, 650, "VTAS BRUTAS"
'MainMant.spDoc.TextOut 1700, 650, "IMPUESTO"
'MainMant.spDoc.TextOut 1900, 650, "EXONER."

For i = 1 To LVVentas.ListItems.Count
    MainMant.spDoc.TextAlign = SPTA_LEFT
    MainMant.spDoc.TextOut 300, iLin, LVVentas.ListItems(i).Text
    MainMant.spDoc.TextOut 550, iLin, LVVentas.ListItems(i).ListSubItems(1).Text
    MainMant.spDoc.TextAlign = SPTA_RIGHT
    MainMant.spDoc.TextOut 950, iLin, LVVentas.ListItems(i).ListSubItems(2).Text
    MainMant.spDoc.TextOut 1300, iLin, LVVentas.ListItems(i).ListSubItems(3).Text
    MainMant.spDoc.TextOut 1600, iLin, LVVentas.ListItems(i).ListSubItems(4).Text
    On Error Resume Next
    MainMant.spDoc.TextOut 1820, iLin, LVVentas.ListItems(i).ListSubItems(5).Text
    MainMant.spDoc.TextOut 1990, iLin, LVVentas.ListItems(i).ListSubItems(6).Text
    On Error GoTo 0
    On Error GoTo ErrorPrn:
    MainMant.spDoc.TextAlign = SPTA_LEFT
    iLin = iLin + 50
    If iLin > 2400 Then
        MainMant.spDoc.TextAlign = SPTA_LEFT
        PrintTit
    End If
Next i

On Error GoTo 0
iLin = iLin + 100
MainMant.spDoc.TextAlign = SPTA_LEFT
MainMant.spDoc.DoPrintPreview

Exit Sub

ErrorPrn:
    ShowMsg "¡ Ocurre algún Error con la Impresora, Intente Conecterla !", vbRed, vbYellow
    
End Sub

Private Sub Form_Load()
Dim rsISCYear As ADODB.Recordset

Set rsISCYear = New ADODB.Recordset

On Error GoTo ErrAdm:
rsISCYear.Open "SELECT ISC_YEAR FROM ISC ORDER BY ISC_YEAR DESC", msConn, adOpenStatic, adLockReadOnly

If rsISCYear.EOF Then
    rsISCYear.Close
    Set rsISCYear = Nothing
    Call Load_Year
    Call Seguridad
    Exit Sub
End If

Do While Not rsISCYear.EOF
    ListAnno.AddItem rsISCYear!ISC_YEAR
    rsISCYear.MoveNext
Loop
rsISCYear.Close
Set rsISCYear = Nothing
On Error GoTo 0

Call Seguridad

Exit Sub

ErrAdm:
ShowMsg "Favor ejecutar el programa de ventas (Version 5.0 o Superior) en la estación de Caja y despues regrese al Informe de Ventas." & vbCrLf & _
" Tambien actualice la Base de datos con el Updater version 5.0", vbYellow, vbRed
'ListAnno.AddItem "2011"
'ListAnno.AddItem "2010"
'rsISCYear.Close
Set rsISCYear = Nothing
End Sub
'---------------------------------------------------------------------------------------
' Procedimiento : Load_Year
' Autor       : hsequeira
' Fecha       : 13/10/2014
' Proposito   : SELECCIONA LOS AÑOS QUE HAY EN LA TABLE Z_COUNTER
'---------------------------------------------------------------------------------------
'
Private Sub Load_Year()
Dim rsISCYear As ADODB.Recordset

IsFAST = True

Set rsISCYear = New ADODB.Recordset

rsISCYear.Open "SELECT DISTINCT LEFT(FECHA,4) AS ISC_YEAR FROM Z_COUNTER ORDER BY 1 DESC", msConn, adOpenStatic, adLockReadOnly
If rsISCYear.EOF Then
    rsISCYear.Close
    Set rsISCYear = Nothing
    Exit Sub
End If

Do While Not rsISCYear.EOF
    ListAnno.AddItem rsISCYear!ISC_YEAR
    rsISCYear.MoveNext
Loop
rsISCYear.Close
Set rsISCYear = Nothing
On Error GoTo 0


End Sub
Private Sub Form_Resize()
   On Error GoTo Form_Resize_Error

    LVVentas.Height = InfVentas.Height - 900
    LVVentas.Width = InfVentas.Width - 3000

   On Error GoTo 0
   Exit Sub

Form_Resize_Error:
If Err.Number = 380 Then
Else
    ShowMsg "Error " & Err.Number & " (" & Err.Description & ") en Pantalla de Informe de Ventas e Impuestos"
End If
End Sub

Private Sub ListAnno_Click()
Dim rsISC As ADODB.Recordset
Dim i As Integer

LVVentas.ListItems.Clear
LVVentas.ColumnHeaders.Clear

LV.ListItems.Clear
LV.ColumnHeaders.Clear

LV.ColumnHeaders.Add , , "Mes"
LV.ColumnHeaders.Add , , "Impuesto"

LV.ColumnHeaders(1).Width = 800
LV.ColumnHeaders(2).Alignment = lvwColumnRight

Set rsISC = New ADODB.Recordset
rsISC.Open "SELECT * FROM ISC WHERE ISC_YEAR =" & CLng(ListAnno.Text), msConn, adOpenStatic, adLockOptimistic

    For i = 1 To 12
        LV.ListItems.Add , , GetMes(ListAnno.Text & Mid(rsISC.Fields(i).Name, 4, 2) & "01")
        If IsFAST Then
            LV.ListItems.Item(i).ListSubItems.Add , , Format(0, "CURRENCY")
        Else
            LV.ListItems.Item(i).ListSubItems.Add , , Format(rsISC.Fields(i).value, "CURRENCY")
        End If
    Next
    LV.ListItems.Add , , "TODO"
rsISC.Close
Set rsISC = Nothing
End Sub

Private Sub LV_Click()
Dim cSQL As String
Dim cMesAño As String
Dim rsZCounter As ADODB.Recordset
Dim i As Integer
Dim nVentas As Currency
Dim nImpuesto As Currency
Dim nExo As Currency, nExoAcum As Currency
Dim bAll As Boolean
Dim nDias As Long

On Error GoTo ErrAdm:

MesasPED "OPEN"

'LV.SelectedItem.Index   'EL INDICE DEL ITEM SELECCIONADO
If LV.SelectedItem.Index = 13 Then
    'INFO: SELECCIONA TODO EL AÑO
    bAll = True
    cMesAño = ListAnno.Text
Else
    bAll = False
    cMesAño = ListAnno.Text & Format(LV.SelectedItem.Index, "00")
End If

cSQL = "SELECT ID, CONTADOR, FECHA, HORA, TOTAL_ANTERIOR, TOTAL_NUEVO,"
cSQL = cSQL & " (TOTAL_NUEVO - TOTAL_ANTERIOR) AS VENTAS, ITBMS, CONTADOR"
cSQL = cSQL & " FROM Z_COUNTER "

If bAll Then
    cSQL = cSQL & " WHERE LEFT(FECHA,4) = '" & cMesAño & "'"
Else
    cSQL = cSQL & " WHERE LEFT(FECHA,6) = '" & cMesAño & "'"
End If
cSQL = cSQL & " ORDER BY ID"

Set rsZCounter = New ADODB.Recordset
rsZCounter.Open cSQL, msConn, adOpenStatic, adLockOptimistic

LVVentas.ListItems.Clear
LVVentas.ColumnHeaders.Clear
LVVentas.ColumnHeaders.Add , , "Fecha"
LVVentas.ColumnHeaders.Add , , "Hora"
LVVentas.ColumnHeaders.Add , , "Anterior"
LVVentas.ColumnHeaders.Add , , "Nuevo"
'LVVentas.ColumnHeaders.Add , , "Ventas Brutas"
LVVentas.ColumnHeaders.Add , , "Ventas Netas"
LVVentas.ColumnHeaders.Add , , "Impuesto"
LVVentas.ColumnHeaders.Add , , "Exoneración"
LVVentas.ColumnHeaders.Add , , "Contador"

'LVVentas.ColumnHeaders(1).Width = 400
LVVentas.ColumnHeaders(2).Width = 700
LVVentas.ColumnHeaders(3).Width = 1300
LVVentas.ColumnHeaders(4).Width = 1300
LVVentas.ColumnHeaders(6).Width = 1100
LVVentas.ColumnHeaders(7).Width = 1100
LVVentas.ColumnHeaders(8).Width = 900

LVVentas.ColumnHeaders(3).Alignment = lvwColumnRight
LVVentas.ColumnHeaders(4).Alignment = lvwColumnRight
LVVentas.ColumnHeaders(5).Alignment = lvwColumnRight
LVVentas.ColumnHeaders(6).Alignment = lvwColumnRight
LVVentas.ColumnHeaders(7).Alignment = lvwColumnRight
LVVentas.ColumnHeaders(8).Alignment = lvwColumnRight

i = 1
nDias = 0
Do While Not rsZCounter.EOF
        LVVentas.ListItems.Add , , GetFecha(rsZCounter!FECHA)
        LVVentas.ListItems.Item(i).ListSubItems.Add , , Format(rsZCounter!HORA, "00:00")
        LVVentas.ListItems.Item(i).ListSubItems.Add , , Format(rsZCounter!TOTAL_ANTERIOR, "STANDARD")
        LVVentas.ListItems.Item(i).ListSubItems.Add , , Format(rsZCounter!TOTAL_NUEVO, "STANDARD")
        LVVentas.ListItems.Item(i).ListSubItems.Add , , Format(rsZCounter!Ventas, "STANDARD")
        nVentas = nVentas + rsZCounter!Ventas
        LVVentas.ListItems.Item(i).ListSubItems.Add , , Format(rsZCounter!ITBMS, "STANDARD")
        nImpuesto = nImpuesto + IIf(IsNull(rsZCounter!ITBMS), 0, rsZCounter!ITBMS)
        
        'INFO: OBTIENE LA EXONERACION DEL Z_COUNTER QUE SE ESTA EJECUTANDO
        nExo = GetExoneracion(rsZCounter!CONTADOR)
        
        nExoAcum = nExoAcum + nExo
        LVVentas.ListItems.Item(i).ListSubItems.Add , , Format(nExo, "STANDARD")
        LVVentas.ListItems.Item(i).ListSubItems.Add , , rsZCounter!CONTADOR
        
        i = i + 1
        nDias = nDias + 1
        rsZCounter.MoveNext
        
Loop
LVVentas.ListItems.Add , , "TOTALES"
LVVentas.ListItems.Item(i).ListSubItems.Add , , Space(1)
LVVentas.ListItems.Item(i).ListSubItems.Add , , Space(1)
LVVentas.ListItems.Item(i).ListSubItems.Add , , "(" & nDias & ") Cierres Z"
LVVentas.ListItems.Item(i).ListSubItems.Add , , Format(nVentas, "CURRENCY")
LVVentas.ListItems.Item(i).ListSubItems.Add , , Format(nImpuesto, "STANDARD")
LVVentas.ListItems.Item(i).ListSubItems.Add , , Format(nExoAcum, "STANDARD")

LVVentas.ListItems.Add , , ""
LVVentas.ListItems.Item(i + 1).ListSubItems.Add , , Space(1)
LVVentas.ListItems.Item(i + 1).ListSubItems.Add , , Space(1)
LVVentas.ListItems.Item(i + 1).ListSubItems.Add , , "Promedio"
'INFO: PROTEGE CONTRA OVERFLOW
On Error Resume Next
LVVentas.ListItems.Item(i + 1).ListSubItems.Add , , Format(nVentas / nDias, "CURRENCY")
On Error GoTo 0
LVVentas.ListItems.Item(i + 1).ListSubItems.Add , , Space(1)



rsZCounter.Close
Set rsZCounter = Nothing

MesasPED "CLOSE"

On Error GoTo 0
Exit Sub

ErrAdm:
If msPED.State = adStateOpen Then MesasPED "CLOSE"
End Sub

Private Function GetExoneracion(nRZ As Long) As Single
'INFO: 20DIC2010
'REGRESA LA EXONERACION MARCADA y QUE YA FUE REPORTADA AL REPORTE Z
Dim nExoneracion As Single
Dim rsExo As ADODB.Recordset

nExoneracion = 0#
Set rsExo = New ADODB.Recordset

rsExo.Open "SELECT SUM(MONTO) AS EXONERACION FROM I_TRANS WHERE Z_COUNTER = " & nRZ, msPED, adOpenStatic, adLockOptimistic
If rsExo.EOF Then
    nExoneracion = 0#
Else
    If IsNull(rsExo!EXONERACION) Then
        nExoneracion = 0#
    Else
        nExoneracion = rsExo!EXONERACION
    End If
End If

rsExo.Close
Set rsExo = Nothing
GetExoneracion = nExoneracion
End Function
Private Sub LVVentas_DblClick()
Dim a As Long
Dim b As Long
Dim i As Long
Dim cHeader As String
Dim cData As String

If vbYes = ShowMsg("¿ DESEA EXPORTAR LOS DATOS ?", vbWhite, vbBlue, vbYesNo) Then

    nFileNumber = FreeFile()
    On Error GoTo ErrAdm:
    
    For i = 1 To LVVentas.ColumnHeaders.Count
        cHeader = cHeader & LVVentas.ColumnHeaders(i).Text & vbTab
    Next
    'cHeader = cHeader
    'cHeader = Mid(cHeader, 3, Len(cHeader))
    iCols = LVVentas.ColumnHeaders.Count
    
    'Open App.Path & "\INFVENTAS.txt" For Output As #nFileNumber
    Open DATA_PATH & "INFVENTAS.txt" For Output As #nFileNumber
    
    Print #nFileNumber, cHeader
    
    For a = 1 To LVVentas.ListItems.Count
        For b = 1 To iCols
            If b = 1 Then
                cData = LVVentas.ListItems.Item(a) & vbTab
            Else
                If LVVentas.ColumnHeaders(b).Width > 0.1 Then
                    cData = cData & LVVentas.ListItems(a).ListSubItems(b - 1).Text & vbTab
                End If
            End If
        Next
        'cData = vbTab & cData
        'cData = Mid(cData, 3, Len(cData))
        Print #nFileNumber, cData
        cData = ""
    Next
    Close #nFileNumber
    
    'ShowMsg "Se exportaron los datos a: " & App.Path & "\INFVENTAS.txt"
    ShowMsg "Se exportaron los datos a: " & DATA_PATH & "INFVENTAS.txt"
    On Error GoTo 0
End If
Exit Sub

ErrAdm:
    If Err.Number = 35600 Then
        Resume Next
    Else
        ShowMsg Err.Number & " - " & Err.Description, vbYellow, vbRed
    End If
    Close #nFileNumber
End Sub
