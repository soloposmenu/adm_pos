VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{65E121D4-0C60-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCHRT20.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{43135020-B751-4DDD-AB4A-B6D8A342216E}#1.0#0"; "sg20o.ocx"
Begin VB.Form conVtaWeek 
   BackColor       =   &H00B39665&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Consulta de Ventas por dia de la Semana"
   ClientHeight    =   7500
   ClientLeft      =   165
   ClientTop       =   330
   ClientWidth     =   11370
   Icon            =   "conVtaWeek.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7500
   ScaleWidth      =   11370
   ShowInTaskbar   =   0   'False
   Begin VB.CheckBox IncluyeDecuentos 
      BackColor       =   &H00B39665&
      Caption         =   "EXCLUYE DESCUENTOS x PRODUCTO"
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
      Left            =   120
      TabIndex        =   13
      Top             =   120
      Value           =   1  'Checked
      Width           =   3855
   End
   Begin MSChart20Lib.MSChart MSChart1 
      Height          =   4335
      Left            =   120
      OleObjectBlob   =   "conVtaWeek.frx":0442
      TabIndex        =   4
      Top             =   3000
      Width           =   6975
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid Lista 
      Height          =   1815
      Left            =   3480
      TabIndex        =   6
      ToolTipText     =   "Descuentos y Acompañantes no estan incluidos"
      Top             =   720
      Width           =   3615
      _ExtentX        =   6376
      _ExtentY        =   3201
      _Version        =   393216
      Rows            =   0
      Cols            =   4
      FixedRows       =   0
      ScrollBars      =   0
      _NumberOfBands  =   1
      _Band(0).Cols   =   4
   End
   Begin VB.PictureBox Picture1 
      Height          =   255
      Left            =   720
      ScaleHeight     =   195
      ScaleWidth      =   315
      TabIndex        =   5
      Top             =   1440
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.CommandButton Command3 
      BackColor       =   &H0000FFFF&
      Height          =   495
      Left            =   1800
      Picture         =   "conVtaWeek.frx":1EA8
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   1800
      Width           =   975
   End
   Begin MSComCtl2.DTPicker txtFecIni 
      Height          =   375
      Left            =   1320
      TabIndex        =   0
      Top             =   480
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   661
      _Version        =   393216
      Format          =   132055041
      CurrentDate     =   36970
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Ejecutar Consulta"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   120
      TabIndex        =   2
      Top             =   1800
      Width           =   1455
   End
   Begin MSComCtl2.DTPicker txtFecFin 
      Height          =   375
      Left            =   1320
      TabIndex        =   1
      Top             =   1080
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   661
      _Version        =   393216
      Format          =   132055041
      CurrentDate     =   36970
   End
   Begin DDSharpGridOLEDB2.SGGrid DD_PEDDETALLE 
      Height          =   7035
      Left            =   7200
      TabIndex        =   9
      ToolTipText     =   "Haga Doble click para Exportar"
      Top             =   360
      Width           =   4095
      _cx             =   7223
      _cy             =   12409
      DataMember      =   ""
      DataMode        =   1
      AutoFields      =   -1  'True
      Enabled         =   -1  'True
      GridBorderStyle =   1
      ScrollBars      =   3
      FlatScrollBars  =   0
      ScrollBarTrack  =   0   'False
      DataRowCount    =   2
      BeginProperty HeadingFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DataColCount    =   2
      HeadingRowCount =   1
      HeadingColCount =   0
      TextAlignment   =   0
      WordWrap        =   0   'False
      Ellipsis        =   1
      HeadingBackColor=   12632256
      HeadingForeColor=   -2147483630
      HeadingTextAlignment=   0
      HeadingWordWrap =   0   'False
      HeadingEllipsis =   1
      GridLines       =   1
      HeadingGridLines=   2
      GridLinesColor  =   -2147483633
      HeadingGridLinesColor=   -2147483632
      EvenOddStyle    =   1
      ColorEven       =   -2147483628
      ColorOdd        =   14737632
      UserResizeAnimate=   1
      UserResizing    =   3
      RowHeightMin    =   0
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      UserDragging    =   2
      UserHiding      =   0
      CellPadding     =   15
      CellBkgStyle    =   1
      CellBackColor   =   -2147483643
      CellForeColor   =   -2147483640
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      FocusRect       =   2
      FocusRectColor  =   0
      FocusRectLineWidth=   1
      TabKeyBehavior  =   0
      EnterKeyBehavior=   1
      NavigationWrapMode=   1
      SkipReadOnly    =   0   'False
      DefaultColWidth =   1219
      DefaultRowHeight=   255
      CellsBorderColor=   0
      CellsBorderVisible=   -1  'True
      RowNumbering    =   0   'False
      EqualRowHeight  =   0   'False
      EqualColWidth   =   0   'False
      HScrollHeight   =   0
      VScrollWidth    =   0
      Format          =   "General"
      Appearance      =   2
      FitLastColumn   =   0   'False
      SelectionMode   =   2
      MultiSelect     =   0
      AllowAddNew     =   0   'False
      AllowDelete     =   0   'False
      AllowEdit       =   0   'False
      ScrollBarTips   =   0
      CellTips        =   0
      CellTipsDelay   =   1000
      SpecialMode     =   0
      OutlineLines    =   1
      CacheAllRecords =   -1  'True
      ColumnClickSort =   -1  'True
      PreviewPaneColumn=   ""
      PreviewPaneType =   0
      PreviewPanePosition=   2
      PreviewPaneSize =   2000
      GroupIndentation=   241
      InactiveSelection=   1
      AutoScroll      =   -1  'True
      AutoResize      =   1
      AutoResizeHeadings=   -1  'True
      OLEDragMode     =   0
      OLEDropMode     =   0
      Caption         =   ""
      ScrollTipColumn =   ""
      MaxRows         =   4194304
      MaxColumns      =   8192
      NewRowPos       =   1
      CustomBkgDraw   =   0
      AutoGroup       =   -1  'True
      GroupByBoxVisible=   0   'False
      GroupByBoxText  =   "Arrastre el Titulo de la columna aqui para agrupar por esa columna"
      AlphaBlendEnabled=   0   'False
      DragAlphaLevel  =   206
      AutoSearch      =   0
      AutoSearchDelay =   2000
      OutlineIcons    =   1
      CellTipsDisplayLength=   3000
      StylesCollection=   $"conVtaWeek.frx":21B2
      ColumnsCollection=   $"conVtaWeek.frx":3FE1
      ValueItems      =   $"conVtaWeek.frx":4957
   End
   Begin MSComctlLib.ProgressBar ProgBar 
      Height          =   300
      Left            =   120
      TabIndex        =   11
      Top             =   2560
      Width           =   6975
      _ExtentX        =   12303
      _ExtentY        =   529
      _Version        =   393216
      Appearance      =   1
   End
   Begin VB.Label Label1 
      BackColor       =   &H00B39665&
      Caption         =   "DIA              VENTAS       ITEMS"
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
      Index           =   3
      Left            =   3495
      TabIndex        =   12
      Top             =   525
      Width           =   2895
   End
   Begin VB.Label Label1 
      BackColor       =   &H00B39665&
      Caption         =   "Ventas x Hora (Según Factura)"
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
      Index           =   2
      Left            =   7680
      TabIndex        =   10
      Top             =   120
      Width           =   3495
   End
   Begin VB.Label Label1 
      BackColor       =   &H00B39665&
      Caption         =   "Fecha Final"
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
      Index           =   1
      Left            =   120
      TabIndex        =   8
      Top             =   1200
      Width           =   1095
   End
   Begin VB.Label Label1 
      BackColor       =   &H00B39665&
      Caption         =   "Fecha Inicial"
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
      Index           =   0
      Left            =   120
      TabIndex        =   7
      Top             =   600
      Width           =   1215
   End
End
Attribute VB_Name = "conVtaWeek"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'Private rs As New ADODB.Recordset
Private nDias(7) As String
Private nVentas(7) As Double
Private nITEMS(7) As Long
Private nDiaCounter(7) As Long

Private Sub Command2_Click()
Dim cT As String
Dim dFe As Date
Dim i, MiRow As Integer
Dim dF1 As String
Dim dF2 As String
Dim cSQL As String
Dim rsHORA As ADODB.Recordset
Dim rsDIAS As ADODB.Recordset
Dim iLoop As Long

Set rsHORA = New ADODB.Recordset
Set rsDIAS = New ADODB.Recordset

dF1 = Format(txtFecIni, "YYYYMMDD")
dF2 = Format(txtFecFin, "YYYYMMDD")

'NO INCLUYE LOS DESCUENTOS
cT = "SELECT FECHA, SUM(PRECIO) AS VENTAS, SUM(CANT) AS ITEMS "
cT = cT & " FROM HIST_TR "
cT = cT & " WHERE  FECHA >= '" & dF1 & "'"
cT = cT & " AND FECHA <= '" & dF2 & "'"
cT = cT & " AND '%' NOT IN (DESCRIP) "
If IncluyeDecuentos.value = 1 Then
    cT = cT & " AND DESCRIP NOT LIKE '%DESCUENTO%' "
End If
cT = cT & " AND DESCRIP NOT LIKE  '%@@%' "
cT = cT & " GROUP BY FECHA"

'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
'cSQL = "SELECT HOUR(HORA) AS HORA, MAX(HOUR(HORA)+1) AS FIN, "
'INFO: 16FEB2013
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
cSQL = "SELECT HOUR(HORA) AS HORA, MAX(FORMAT(HOUR(HORA),'00')) & ':59' AS FIN, "
cSQL = cSQL & " SUM(PRECIO) AS VENTAS, SUM(CANT) AS ITEMS, MAX(DESCRIP) AS M_DESCRIP"
cSQL = cSQL & " FROM HIST_TR "
cSQL = cSQL & " WHERE  FECHA >= '" & dF1 & "'"
cSQL = cSQL & " AND FECHA <= '" & dF2 & "'"
cSQL = cSQL & " AND '%' NOT IN (DESCRIP) "
If IncluyeDecuentos.value = 1 Then
    cSQL = cSQL & " AND DESCRIP NOT LIKE '%DESCUENTO%' "
End If
cSQL = cSQL & " AND DESCRIP NOT LIKE  '%@@%' "
cSQL = cSQL & " GROUP BY  HOUR(HORA)"

Me.MousePointer = vbHourglass

'If rsDIAS.State = adStateOpen Then rsDIAS.Close
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
rsDIAS.Open cT, msConn, adOpenStatic, adLockOptimistic
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

On Error Resume Next
ProgBar.value = 0
ProgBar.Max = rsDIAS.RecordCount + 1
On Error GoTo 0

Do Until rsDIAS.EOF
    'MsgBox Format(rs!FECHA, "####-##-##") & "---" & rs!VENTAS
    dFe = Format(rsDIAS!FECHA, "####-##-##")
    ProgBar.value = ProgBar.value + 1
    ProgBar.Refresh
    Call AsignaVentas(Weekday(dFe), rsDIAS!VENTAS, rsDIAS!items)
    rsDIAS.MoveNext
Loop
ProgBar.value = 0

With MSChart1
    .chartType = VtChChartType2dBar
    '.chartType = VtChChartType3dBar
    .RowCount = 7
    .ColumnCount = 2
    '.Title.TextLength = 50
    '.TitleText = "Ventas Semanales. Periodo " & txtFecIni & " - " & txtFecFin
    .Title.Text = Space(5) & "Ventas Semanales. Periodo " & txtFecIni & " - " & txtFecFin & Space(5)
End With

Lista.Clear: Lista.Rows = 0: Lista.Refresh
MiRow = 1
ProgBar.value = 0
ProgBar.Max = 7
For MiRow = 1 To 7
    ProgBar.value = MiRow
    ProgBar.Refresh
    MSChart1.Row = MiRow
    MSChart1.Column = 1
    MSChart1.RowLabel = nDias(MiRow)
    MSChart1.Data = Format(nVentas(MiRow), "STANDARD")
    MSChart1.Column = 2
    MSChart1.Data = nITEMS(MiRow)
    
    Lista.AddItem nDias(MiRow) & vbTab & _
                              Format(nVentas(MiRow), "STANDARD") & vbTab & _
                              Format(nITEMS(MiRow), "###,###") & vbTab & nDiaCounter(MiRow)
Next
For i = 1 To 7
    nVentas(i) = 0#
    nITEMS(i) = 0
    nDiaCounter(i) = 0
Next

'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
Dim arreglo(0, 0)
DD_PEDDETALLE.LoadArray arreglo

ProgBar.value = 1
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
rsHORA.Open cSQL, msConn, adOpenStatic, adLockOptimistic
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
ProgBar.value = 2

If rsHORA.EOF Then
    'NO HAY DATOS PARA MOSTRAR
    DD_PEDDETALLE.Columns.RemoveAll True
    rsHORA.Close
    rsDIAS.Close
    Set rsHORA = Nothing
    Set rsDIAS = Nothing
    ProgBar.value = 0
    Me.MousePointer = vbDefault
    Exit Sub
End If

'INFO: 16FEB2013
With DD_PEDDETALLE
   .Columns.RemoveAll True
   .DataMode = sgUnbound

   .LoadArray rsHORA.GetRows()
   ProgBar.value = 3

   ' define each column from the recordsets' fields collection
   For iLoop = 1 To rsHORA.Fields.Count
      .Columns(iLoop).Caption = rsHORA.Fields(iLoop - 1).Name
      .Columns(iLoop).DBField = rsHORA.Fields(iLoop - 1).Name
      .Columns(iLoop).Key = rsHORA.Fields(iLoop - 1).Name
   Next iLoop
   ProgBar.value = 4
End With

With DD_PEDDETALLE
    
    .ColumnClickSort = True
    .EvenOddStyle = sgEvenOddRows
    .ColorEven = vbWhite
    .ColorOdd = &HE0E0E0
    .Columns(1).Width = 600
    .Columns(2).Width = 600
    .Columns(3).Width = 1200
    .Columns(4).Width = 800
    .Columns(5).Width = 0
    .Columns(5).Hidden = True

    .Columns(1).Style.TextAlignment = sgAlignLeftCenter
    .Columns(2).Style.TextAlignment = sgAlignLeftCenter
    .Columns(3).Style.TextAlignment = sgAlignRightCenter
    .Columns(4).Style.TextAlignment = sgAlignRightCenter

    .Columns(1).Style.Format = "00"
    .Columns(2).Style.Format = "Short Time"
    .Columns(3).Style.Format = "Standard"

End With
ProgBar.value = 5

'INFO: 16FEB2013
rsHORA.Close
rsDIAS.Close
Set rsHORA = Nothing
Set rsDIAS = Nothing

ProgBar.value = 7
Me.MousePointer = vbDefault

Call Seguridad

End Sub

Private Function AsignaVentas(d As Integer, n As Double, nnITEMS As Long)
nVentas(d) = nVentas(d) + n
nITEMS(d) = nITEMS(d) + nnITEMS
nDiaCounter(d) = nDiaCounter(d) + 1
End Function

Private Function AsignaDias()
Dim i As Integer
For i = 1 To 7
        If i = 1 Then nDias(i) = "Domingo  "
        If i = 2 Then nDias(i) = "Lunes       "
        If i = 3 Then nDias(i) = "Martes     "
        If i = 4 Then nDias(i) = "Miercoles"
        If i = 5 Then nDias(i) = "Jueves     "
        If i = 6 Then nDias(i) = "Viernes    "
        If i = 7 Then nDias(i) = "Sabado    "
Next
End Function

Private Sub Command3_Click()
Dim i As Integer
Dim cCad As String
Dim sSubTot As Double
Dim nPagina As Integer
Dim iLin As Long, iFil As Long, iCol As Long
Dim nDia_AVG As Long
Dim nVenta_AVG As Double

If Lista.Rows = 0 Then
    ShowMsg "NO HAY DATOS PARA IMPRIMIR", vbRed, vbYellow
    Exit Sub
End If

Me.MousePointer = vbHourglass

MainMant.spDoc.DocBegin
MainMant.spDoc.TextAlign = SPTA_LEFT

MainMant.spDoc.WindowTitle = "Impresión de " & Me.Caption
MainMant.spDoc.FirstPage = 1
MainMant.spDoc.PageOrientation = SPOR_PORTRAIT
MainMant.spDoc.Units = SPUN_LOMETRIC

MainMant.spDoc.TextOut 300, 200, Format(Date, "long date") & "  " & Time
MainMant.spDoc.TextOut 300, 250, "Página : " & nPagina + 1
MainMant.spDoc.TextOut 300, 350, rs00!DESCRIP
MainMant.spDoc.TextOut 300, 400, MSChart1.Title.Text & Space(5) & Label1(2).Caption
'MainMant.spDoc.TextOut 300, 550, "PERIODO : Desde " & txtFecIni & " Hasta " & txtFecFin

MainMant.spDoc.TextOut 300, 500, "DIA"
MainMant.spDoc.TextOut 700, 500, "VENTAS"
MainMant.spDoc.TextOut 950, 500, "Promedio"
MainMant.spDoc.TextOut 300, 550, "--------------------------------------------------------------------"

iLin = 600
nPagina = nPagina + 1

For i = 1 To 7
    Lista.Row = i - 1
    Lista.Col = 0
    Lista.Col = 3
    nDia_AVG = IIf(Lista.Text = "0", 1, Lista.Text)
    MainMant.spDoc.TextOut 300, iLin, nDias(i)
    MainMant.spDoc.TextOut 500, iLin, "(" & Lista.Text & ")"
    MainMant.spDoc.TextAlign = SPTA_RIGHT
    Lista.Col = 1
    sSubTot = sSubTot + Val(Format(Lista.Text, "#.##"))
    nVenta_AVG = Lista.Text
    MainMant.spDoc.TextOut 830, iLin, Format(Lista.Text, "STANDARD")
    MainMant.spDoc.TextOut 1080, iLin, Format((nVenta_AVG / nDia_AVG), "STANDARD")
    MainMant.spDoc.TextAlign = SPTA_LEFT
    iLin = iLin + 50
Next
iLin = iLin + 100
MainMant.spDoc.TextAlign = SPTA_LEFT
MainMant.spDoc.TextOut 300, iLin, "Sub Total del Periodo = " & Format(sSubTot, "Currency")

'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
MainMant.spDoc.TextOut 1200, 500, "RANGO"
MainMant.spDoc.TextOut 1550, 500, "VENTAS"
MainMant.spDoc.TextOut 1800, 500, "ITEMS"
MainMant.spDoc.TextOut 1200, 550, "-------------------------------------------------------------"

iLin = 600
sSubTot = 0
For iFil = 1 To DD_PEDDETALLE.RowCount - 1
    
    DD_PEDDETALLE.Row = iFil

    For iCol = 0 To DD_PEDDETALLE.ColCount - 1
        Select Case iCol
           Case 0, 1, 2, 3
                DD_PEDDETALLE.Col = iCol
                Select Case iCol
                    Case 0
                        MainMant.spDoc.TextAlign = SPTA_LEFT
                        MainMant.spDoc.TextOut 1200, iLin, DD_PEDDETALLE.Text & ":00 ~"
                    Case 1
                        MainMant.spDoc.TextAlign = SPTA_LEFT
                        MainMant.spDoc.TextOut 1325, iLin, DD_PEDDETALLE.Text
                    Case 2
                        MainMant.spDoc.TextAlign = SPTA_RIGHT
                        MainMant.spDoc.TextOut 1700, iLin, Format(DD_PEDDETALLE.Text, "STANDARD")
                        sSubTot = sSubTot + Val(Format(DD_PEDDETALLE.Text, "#.##"))
                    Case 3
                        MainMant.spDoc.TextAlign = SPTA_RIGHT
                        MainMant.spDoc.TextOut 1900, iLin, Format(DD_PEDDETALLE.Text, "General Number")
                End Select
                MainMant.spDoc.TextAlign = SPTA_LEFT
            End Select
    Next
    iLin = iLin + 50
Next
MainMant.spDoc.TextAlign = SPTA_LEFT
MainMant.spDoc.TextOut 1200, iLin + 50, "Sub Total del Periodo = " & Format(sSubTot, "Currency")
'INFO
iLin = iLin + 100
MainMant.spDoc.TextAlign = SPTA_LEFT
MainMant.spDoc.TextOut 1200, iLin, "Ventas x Hora = " & Format(sSubTot / (DD_PEDDETALLE.RowCount - 1), "Currency")

'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
MainMant.spDoc.TextAlign = SPTA_LEFT

Dim nFileId As Long
On Error Resume Next
    Kill App.Path & "\" & "mi_imagen.wmf"
On Error GoTo 0
MSChart1.EditCopy
SavePicture Clipboard.GetData(vbCFMetafile), App.Path & "\" & "mi_imagen.wmf"
MainMant.spDoc.LoadImage App.Path & "\mi_imagen.wmf", nFileId
MainMant.spDoc.PlaceImage nFileId, 83, iLin + 100, 2045, 2630, SPIA_BESTFIT

Me.MousePointer = vbDefault

MainMant.spDoc.DoPrintPreview
sSubTot = 0
Exit Sub

ErrChart:
    Me.MousePointer = vbDefault
    MsgBox "Error de Impresión (Papel o Impresora). " & Err.Number & " -- " & Err.Description
    Resume Next
End Sub
Private Sub DD_PEDDETALLE_DblClick()
    'INFO: 16FEB2013
    Call ExportToExcelOrCSVList(DD_PEDDETALLE)
End Sub

Private Sub Form_Load()
txtFecIni = Format(Date, "SHORT DATE")
txtFecFin = Format(Date, "SHORT DATE")
'MSChart1.Title.TextLength = 50
Call AsignaDias

Call Seguridad

End Sub
Private Function Seguridad() As String
'SETUP DE SEGURIDAD DEL SISTEMA
Dim cSeguridad As String

cSeguridad = GetSecuritySetting(npNumCaj, Me.Name)
Select Case cSeguridad
    Case "CEMV"        'Crear - Eliminar - Modificar - Ver
        'INFO: NO HAY RESTRICCIONES
    Case "CMV"        'Crear - Modificar - Ver"
        'INFO: NO HAY RESTRICCIONES
    Case "CV"        'Crear - Ver
        'INFO: NO HAY RESTRICCIONES
    Case "V"        'Ver solamente
        Command3.Enabled = False
    Case "N"        'SIN DERECHOS
        txtFecIni.Enabled = False: txtFecFin.Enabled = False: Command2.Enabled = False: Command3.Enabled = False
End Select
End Function

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
'If rs.State = adStateOpen Then rs.Close
Unload Me
End Sub
Private Sub txtFecFin_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then Command2.SetFocus
End Sub

Private Sub txtFecIni_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then txtFecFin.SetFocus
End Sub
