VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{65E121D4-0C60-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCHRT20.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{43135020-B751-4DDD-AB4A-B6D8A342216E}#1.0#0"; "sg20o.ocx"
Begin VB.Form ConDept 
   BackColor       =   &H00B39665&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "CONSULTA DE VENTAS POR DEPARTAMENTO"
   ClientHeight    =   7350
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   10980
   Icon            =   "ConDept.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7350
   ScaleWidth      =   10980
   StartUpPosition =   1  'CenterOwner
   Begin DDSharpGridOLEDB2.SGGrid DD_PEDDETALLE 
      Height          =   4755
      Left            =   360
      TabIndex        =   16
      Top             =   1800
      Width           =   10335
      _cx             =   18230
      _cy             =   8387
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
      StylesCollection=   $"ConDept.frx":0442
      ColumnsCollection=   $"ConDept.frx":2271
      ValueItems      =   $"ConDept.frx":2BE7
   End
   Begin DDSharpGridOLEDB2.SGGrid oDD_PEDDETALLE 
      Height          =   1575
      Left            =   9480
      TabIndex        =   15
      Top             =   4920
      Width           =   1215
      _cx             =   2143
      _cy             =   2778
      DataMember      =   ""
      DataMode        =   1
      AutoFields      =   -1  'True
      Enabled         =   -1  'True
      GridBorderStyle =   1
      ScrollBars      =   3
      FlatScrollBars  =   0
      ScrollBarTrack  =   0   'False
      DataRowCount    =   0
      BeginProperty HeadingFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DataColCount    =   0
      HeadingRowCount =   1
      HeadingColCount =   1
      TextAlignment   =   0
      WordWrap        =   0   'False
      Ellipsis        =   1
      HeadingBackColor=   -2147483633
      HeadingForeColor=   -2147483630
      HeadingTextAlignment=   0
      HeadingWordWrap =   0   'False
      HeadingEllipsis =   1
      GridLines       =   1
      HeadingGridLines=   2
      GridLinesColor  =   -2147483633
      HeadingGridLinesColor=   -2147483632
      EvenOddStyle    =   0
      ColorEven       =   -2147483628
      ColorOdd        =   -2147483624
      UserResizeAnimate=   1
      UserResizing    =   3
      RowHeightMin    =   0
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      UserDragging    =   2
      UserHiding      =   2
      CellPadding     =   15
      CellBkgStyle    =   1
      CellBackColor   =   -2147483643
      CellForeColor   =   -2147483640
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      FocusRect       =   1
      FocusRectColor  =   0
      FocusRectLineWidth=   1
      TabKeyBehavior  =   0
      EnterKeyBehavior=   0
      NavigationWrapMode=   1
      SkipReadOnly    =   0   'False
      DefaultColWidth =   1200
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
      ColumnClickSort =   0   'False
      PreviewPaneColumn=   ""
      PreviewPaneType =   0
      PreviewPanePosition=   2
      PreviewPaneSize =   2000
      GroupIndentation=   225
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
      GroupByBoxText  =   "Drag a column header here to group by that column"
      AlphaBlendEnabled=   0   'False
      DragAlphaLevel  =   206
      AutoSearch      =   0
      AutoSearchDelay =   2000
      OutlineIcons    =   1
      CellTipsDisplayLength=   3000
      StylesCollection=   $"ConDept.frx":2C87
      ColumnsCollection=   $"ConDept.frx":4A5A
      ValueItems      =   $"ConDept.frx":4F6F
   End
   Begin VB.OptionButton opcTipo 
      BackColor       =   &H00B39665&
      Caption         =   "Por Tipo"
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
      Left            =   5400
      TabIndex        =   14
      ToolTipText     =   "Obtener Ventas por Reporte Z"
      Top             =   240
      Width           =   1095
   End
   Begin VB.OptionButton opcTipo 
      BackColor       =   &H00B39665&
      Caption         =   "Por Fecha"
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
      Left            =   3840
      TabIndex        =   13
      ToolTipText     =   "Obtener Ventas por Fechas Seleccionadas"
      Top             =   240
      Value           =   -1  'True
      Width           =   1215
   End
   Begin MSComctlLib.ListView LVZ 
      Height          =   830
      Left            =   6500
      TabIndex        =   12
      Top             =   150
      Width           =   2415
      _ExtentX        =   4260
      _ExtentY        =   1455
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   0   'False
      HideSelection   =   -1  'True
      Checkboxes      =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   0
      NumItems        =   0
   End
   Begin VB.PictureBox Picture1 
      Height          =   375
      Left            =   3000
      ScaleHeight     =   315
      ScaleWidth      =   435
      TabIndex        =   10
      Top             =   0
      Visible         =   0   'False
      Width           =   495
   End
   Begin MSComCtl2.DTPicker txtFecIni 
      Height          =   345
      Left            =   1560
      TabIndex        =   0
      Top             =   105
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   609
      _Version        =   393216
      Format          =   53149697
      CurrentDate     =   36431
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H0000FFFF&
      Height          =   495
      Left            =   240
      Picture         =   "ConDept.frx":5354
      Style           =   1  'Graphical
      TabIndex        =   9
      ToolTipText     =   "Envia Seleccion a la Impresora"
      Top             =   6720
      Width           =   735
   End
   Begin MSChart20Lib.MSChart MSChart1 
      Height          =   4575
      Left            =   360
      OleObjectBlob   =   "ConDept.frx":565E
      TabIndex        =   8
      Top             =   1800
      Visible         =   0   'False
      Width           =   10335
   End
   Begin MSComctlLib.TabStrip TabStrip1 
      Height          =   5415
      Left            =   240
      TabIndex        =   4
      Top             =   1200
      Width           =   10575
      _ExtentX        =   18653
      _ExtentY        =   9551
      MultiRow        =   -1  'True
      Separators      =   -1  'True
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   3
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Reporte Tabular"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Gráfico Diario del Periodo"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab3 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Gráfico Mensual del Periodo"
            ImageVarType    =   2
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.CommandButton cmdSalir 
      Caption         =   "Sa&lir"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   555
      Left            =   9600
      TabIndex        =   7
      Top             =   6720
      Width           =   1215
   End
   Begin VB.CheckBox Top10 
      BackColor       =   &H00B39665&
      Caption         =   "10 Mas Frecuentes"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   3120
      TabIndex        =   2
      Top             =   600
      Width           =   2055
   End
   Begin VB.CommandButton cmdEjec 
      Caption         =   "&Ejecutar Consulta"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   520
      Left            =   9120
      TabIndex        =   3
      Top             =   360
      Width           =   1695
   End
   Begin MSComCtl2.DTPicker txtFecFin 
      Height          =   345
      Left            =   1560
      TabIndex        =   1
      Top             =   600
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   609
      _Version        =   393216
      Format          =   53149697
      CurrentDate     =   36430
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   5280
      Picture         =   "ConDept.frx":7466
      ToolTipText     =   "Exportar Datos"
      Top             =   6720
      Width           =   480
   End
   Begin VB.Shape Borde1 
      BorderColor     =   &H0000FFFF&
      BorderWidth     =   2
      Height          =   975
      Index           =   1
      Left            =   5280
      Top             =   75
      Width           =   3735
   End
   Begin VB.Shape Borde1 
      BorderColor     =   &H0000FFFF&
      BorderWidth     =   2
      Height          =   975
      Index           =   0
      Left            =   240
      Top             =   75
      Width           =   5055
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00B39665&
      Caption         =   "Selección Reporte Z"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   2
      Left            =   5280
      TabIndex        =   11
      Top             =   600
      Width           =   1215
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
      Height          =   255
      Index           =   1
      Left            =   360
      TabIndex        =   6
      Top             =   720
      Width           =   1215
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
      Height          =   255
      Index           =   0
      Left            =   360
      TabIndex        =   5
      Top             =   240
      Width           =   1215
   End
End
Attribute VB_Name = "ConDept"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rsConsulta01 As Recordset
Private mintCurFrame As Integer ' Current Frame visible
Private nPagina As Integer
Private iLin As Integer
Dim nDepSel As Integer
Dim c1DepSel As String, c2DepSel As String
Private Sub GetReportesZ()
Dim cSQL As String
Dim rsZetas As ADODB.Recordset
Dim iLinea As Integer
Dim nZetasINI As Long

On Error GoTo ErrAdm:
Set rsZetas = New ADODB.Recordset

nZetasINI = CLng(GetFromINI("Administracion", "MaxZ", App.Path & "\soloini.ini"))

cSQL = "SELECT TOP " & nZetasINI & " VAL(CONTADOR) AS CONTADOR, FECHA FROM Z_COUNTER ORDER BY VAL(CONTADOR) DESC "
rsZetas.Open cSQL, msConn, adOpenStatic, adLockOptimistic

LVZ.ListItems.Clear
LVZ.ColumnHeaders.Clear

LVZ.ColumnHeaders.Add , , "Z#"
LVZ.ColumnHeaders.Add , , "Fecha"

'LV.ColumnHeaders.Item(1).Alignment = lvwColumnRight
LVZ.ColumnHeaders.Item(1).Alignment = lvwColumnLeft
LVZ.ColumnHeaders.Item(2).Alignment = lvwColumnRight
LVZ.ColumnHeaders.Item(1).Width = 700
LVZ.ColumnHeaders.Item(2).Width = 1250
iLinea = 1
Do While Not rsZetas.EOF
    LVZ.ListItems.Add , , rsZetas!CONTADOR
    
    LVZ.ListItems.Item(iLinea).ListSubItems.Add , , GetFecha(rsZetas!FECHA)
    iLinea = iLinea + 1
    rsZetas.MoveNext
Loop
On Error GoTo 0
rsZetas.Close

ErrAdm:
Set rsZetas = Nothing
End Sub

Private Sub ImprimeChart()
On Error GoTo ErrorImprimeChart:
    Picture1.Visible = True
    MSChart1.EditCopy
    Picture1.Picture = Clipboard.GetData()
    Printer.PaintPicture Picture1.Picture, 0, 3000
    Picture1.Visible = False
    Printer.EndDoc
On Error GoTo 0
Exit Sub

ErrorImprimeChart:
    MsgBox "Error de Impresión (Papel o Cable). " & Err.Description, vbExclamation, "Favor Revisar la Impresora"
    Resume Next
End Sub
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
MainMant.spDoc.TextOut 300, 450, ConDept.Caption

If ConDept.opcTipo(0).value = True Then
    MainMant.spDoc.TextOut 300, 550, "PERIODO : Desde " & txtFecIni & " Hasta " & txtFecFin
Else
    'MainMant.spDoc.TextOut 300, 550, "PERIODO : REPORTE Z # " & LVZ.SelectedItem.Text
    'INFO: 29OCT2010

    MainMant.spDoc.TextOut 300, 550, "REPORTE(S) Z # INCLUIDOS: " & cmdEjec.Tag
End If

If Not MSChart1.Visible Then
    MainMant.spDoc.TextOut 300, 650, "CODIGO"
    MainMant.spDoc.TextOut 500, 650, "DESCRIPCION"
    MainMant.spDoc.TextOut 950, 650, "Unidades"
    MainMant.spDoc.TextOut 1300, 650, "Ventas"
    MainMant.spDoc.TextOut 1450, 650, "Ventas Netas"
    MainMant.spDoc.TextOut 1700, 650, "Descuento"
    MainMant.spDoc.TextOut 1900, 650, "Porcentaje"
    MainMant.spDoc.TextOut 300, 700, "-----------------------------------------------------------------------------------------------------------------------------------------------------"
End If

iLin = 750
nPagina = nPagina + 1
End Sub

Private Sub DoLinea2()
Dim rsGrafico As Recordset
Dim sqltxt As String
Dim dF1 As String
Dim dF2 As String

dF1 = Format(txtFecIni, "YYYYMMDD")
dF2 = Format(txtFecFin, "YYYYMMDD")

If nDepSel = 0 Then
    MsgBox "DEBE SELECCIONAR UN DEPARTAMENTO", vbInformation, BoxTit
    Exit Sub
End If

If Top10.value = 1 Then cTop10 = " TOP 10 " Else cTop10 = ""

Set rsGrafico = New Recordset

sqltxt = "SELECT " & cTop10 & " A.DEPTO,MID(A.FECHA,5,2) AS MES, " & _
        " A.CANT AS UNIDADES, " & _
        " format(A.PRECIO,'standard') AS VENTAS" & _
        " INTO LOLO FROM HIST_TR AS A " & _
        " WHERE A.DEPTO = " & nDepSel & _
        " AND A.FECHA BETWEEN '" & dF1 & "'" & _
        " AND '" & dF2 & "'"

msConn.BeginTrans
msConn.Execute sqltxt
msConn.CommitTrans

sqltxt = "SELECT A.DEPTO,A.MES, " & _
        " SUM(A.UNIDADES) AS CANT, " & _
        " format(SUM(A.VENTAS),'standard') AS PRECIO" & _
        " FROM LOLO AS A " & _
        " GROUP BY A.DEPTO,A.MES " & _
        " ORDER BY A.MES"

rsGrafico.Open sqltxt, msConn, adOpenStatic, adLockOptimistic

With MSChart1
    .chartType = VtChChartType2dLine
    .RowCount = rsGrafico.RecordCount
    .TitleText = "Ventas Mensuales de " & c2DepSel
End With
MiRow = 1
Do Until rsGrafico.EOF
    MSChart1.row = MiRow
    MSChart1.Column = 1
    MSChart1.RowLabel = rsGrafico!MES
    MSChart1.Data = Format(rsGrafico!precio, "standard")
    MSChart1.Column = 2
    MSChart1.Data = Format(rsGrafico!CANT, "standard")
    MiRow = MiRow + 1
    rsGrafico.MoveNext
Loop
rsGrafico.Close
msConn.Execute "DROP TABLE LOLO"
End Sub
Private Sub DoLinea()
Dim rsGrafico As Recordset
Dim sqltxt As String
Dim cTop10 As String
Dim dF1 As String
Dim dF2 As String

dF1 = Format(txtFecIni, "YYYYMMDD")
dF2 = Format(txtFecFin, "YYYYMMDD")

If nDepSel = 0 Then
    MsgBox "DEBE SELECCIONAR UN DEPARTAMENTO", vbInformation, BoxTit
    Exit Sub
End If

If Top10.value = 1 Then cTop10 = " TOP 10 " Else cTop10 = ""

Set rsGrafico = New Recordset
sqltxt = "SELECT " & cTop10 & " A.DEPTO,A.FECHA,SUM(A.CANT) AS UNIDADES, " & _
        " format(SUM(A.PRECIO),'standard') AS VENTAS" & _
        " FROM HIST_TR AS A " & _
        " WHERE A.DEPTO = " & nDepSel & _
        " AND A.FECHA BETWEEN '" & dF1 & "'" & _
        " AND '" & dF2 & "'" & _
        " GROUP BY A.DEPTO,A.FECHA " & _
        " ORDER BY A.FECHA "
rsGrafico.Open sqltxt, msConn, adOpenStatic, adLockOptimistic

With MSChart1
    .chartType = VtChChartType2dLine
    .RowCount = rsGrafico.RecordCount
    .TitleText = "Ventas Diarias de " & c2DepSel
End With
MiRow = 1
Do Until rsGrafico.EOF
    MSChart1.row = MiRow
    MSChart1.RowLabel = Mid(rsGrafico!FECHA, 7, 2) & "/" & Mid(rsGrafico!FECHA, 5, 2)
    MSChart1.Column = 1
    MSChart1.Data = Format(rsGrafico!VENTAS, "standard")
    MSChart1.Column = 2
    MSChart1.Data = Format(rsGrafico!UNIDADES, "standard")
    MiRow = MiRow + 1
    rsGrafico.MoveNext
Loop
rsGrafico.Close

End Sub
Private Sub cmdEjec_Click()
Dim sqltxt As String
Dim cTop10 As String
Dim dF1 As String
Dim dF2 As String
Dim rsTmp01 As Recordset
Dim nTotVal As Double       ''INFO: 16FEB2013
Dim i As Byte
Dim iZZ As Integer  'Z Counter
Dim jFalseCounter As Byte
Dim cSQL As String, cZetas As String

'INFO: 27OCT2010
Dim nMinZ As Long   'REPORTE Z INICIAL
Dim nMaxZ As Long   'REPORTE Z FINAL
Dim cArrayZ() As String

On Error GoTo ErrAdm:

For i = 0 To opcTipo.Count - 1
    Select Case opcTipo(i).value
        Case True
            Exit For
        Case False
            jFalseCounter = jFalseCounter + 1
        Case Else
    End Select
    'If opcTipo(i).Value = True Then Exit For
Next

' if 2 = jFalseCounter then NONE selected, then DEFAULT TO DATE OPTION
If jFalseCounter = 2 Then i = 0

If i = 0 Then
    'NORMAL NORMAL POR FECHA
Else
    'REPORTE POR Z#
    'INFO: ABRIL2008
    'CHECK TO SEE IF A Z HAS BEEN SELECTED
    For iZZ = 1 To LVZ.ListItems.Count
        If LVZ.ListItems(iZZ).Checked = True Then
            cZetas = cZetas & LVZ.ListItems(iZZ).Text & "','"
        End If
    Next
    On Error Resume Next
    If cZetas = "" Then
        ShowMsg "Seleccione un reporte (Z)" & vbCrLf & "Debe seleccionar al menos un Reporte Z"
        Exit Sub
    Else
        cZetas = "'" & Mid(cZetas, 1, Len(cZetas) - 2)
    End If

    'INFO: 27OCT2010
    cmdEjec.Tag = cZetas
    cArrayZ = Split(Replace(cZetas, "'", ""), ",")
    'cmdEjec.Tag = Array(cArrayZ)
    
    On Error GoTo 0
    On Error GoTo ErrAdm:
End If

'-------------------------------
''''MSHFDepto.Visible = True
MSChart1.Visible = False
'-------------------------------

'-ProgBar.Value = 5
If Top10.value = 1 Then cTop10 = " TOP 10 " Else cTop10 = ""

If (txtFecFin - txtFecIni) > 31 Then
    Call LoadReporteAnual
    Exit Sub
End If

    
dF1 = Format(txtFecIni, "YYYYMMDD")
dF2 = Format(txtFecFin, "YYYYMMDD")

Me.MousePointer = vbHourglass

Set rsTmp01 = New Recordset
cSQL = "SELECT SUM(PRECIO) AS VALOR "
cSQL = cSQL & " FROM HIST_TR "
If i = 0 Then
    cSQL = cSQL & " WHERE FECHA Between '" & dF1 & "'"
    cSQL = cSQL & " AND '" & dF2 & "'"
Else
    'cSQL = cSQL & " WHERE Z_COUNTER = '" & Val(LVZ.SelectedItem.Text) & "'"
    'INFO: 16ABR2008. MULTIPLES Z
    '''cSQL = cSQL & " WHERE Z_COUNTER IN (" & cZetas & ")"
    'INFO: VALIDAR LA Z CONTRA VALORES NUMERICOS EN VEZ DE TEXTO (27OCT2010)
    cSQL = cSQL & " WHERE VAL(Z_COUNTER) BETWEEN  " & GetNumber_FromArray(cArrayZ, GetMin)
    cSQL = cSQL & " AND " & GetNumber_FromArray(cArrayZ, Getmax)
End If

rsTmp01.Open cSQL, msConn, adOpenStatic, adLockOptimistic
'-ProgBar.Value = 10

'INFO: 12FEB2011. ESTANDARIZACION DE GRID PARA QUE SEA COMO EL DE PLU
Dim arreglo(0, 0)
DD_PEDDETALLE.LoadArray arreglo

If IsNull(rsTmp01!VALOR) Then
    ''''Set MSHFDepto.DataSource = Nothing
    ''''MSHFDepto.Clear
    '-ProgBar.Value = 0
    Me.MousePointer = vbDefault
    ShowMsg "NO EXISTE INFORMACION PARA MOSTRAR. SELECCIONE OTRA(S) FECHA(S)"
    'DD_PEDDETALLE.Rows.RemoveAll True
    'DD_PEDDETALLE.Columns.RemoveAll True
    Exit Sub
End If

'DD_PEDDETALLE.Rows.RemoveAll True
'DD_PEDDETALLE.Columns.RemoveAll True

nTotVal = rsTmp01!VALOR
rsTmp01.Close
'-ProgBar.Value = 20
'''MSHFDepto.Clear
msConn.BeginTrans
'NO INCLUYE DESCUENTOS
sqltxt = "SELECT " & cTop10 & " A.DEPTO,B.DESCRIP,"
sqltxt = sqltxt & " B.CORTO,SUM(A.CANT) AS UNIDADES, "
sqltxt = sqltxt & " SUM(A.PRECIO) AS P_ORDEN, "
sqltxt = sqltxt & " format(SUM(A.PRECIO),'STANDARD') AS VENTAS"
sqltxt = sqltxt & " INTO LOLO "
sqltxt = sqltxt & " FROM HIST_TR AS A LEFT JOIN DEPTO AS B ON A.DEPTO = B.CODIGO "
If i = 0 Then
    sqltxt = sqltxt & " WHERE A.FECHA Between '" & dF1 & "'"
    sqltxt = sqltxt & " AND '" & dF2 & "'"
Else
    'sqltxt = sqltxt & " WHERE A.Z_COUNTER = '" & Val(LVZ.SelectedItem.Text) & "'"
    'INFO: 16ABR2008. MULTIPLES Z
    'sqltxt = sqltxt & " WHERE Z_COUNTER IN (" & cZetas & ")"
    'INFO: VALIDAR LA Z CONTRA VALORES NUMERICOS EN VEZ DE TEXTO (27OCT2010)
    sqltxt = sqltxt & " WHERE VAL(A.Z_COUNTER) BETWEEN  " & GetNumber_FromArray(cArrayZ, GetMin)
    sqltxt = sqltxt & " AND " & GetNumber_FromArray(cArrayZ, Getmax)
End If
'INFO REEMPLAZANDO x ('%' NOT IN A.DESCRIP) 18/MAY/2009
'sqltxt = sqltxt & " AND MID(A.DESCRIP,LEN(TRIM(A.DESCRIP)),1) <> '%' "
sqltxt = sqltxt & " AND '%' NOT IN (A.DESCRIP) "
sqltxt = sqltxt & " AND A.DESCRIP NOT LIKE '%DESCUENTO%' "
sqltxt = sqltxt & " AND A.DESCRIP NOT LIKE  '%@@%' "
sqltxt = sqltxt & " GROUP BY A.DEPTO,B.DESCRIP,B.CORTO "
sqltxt = sqltxt & " ORDER BY 5 DESC "
        
        '" AND '%' NOT IN (A.DESCRIP) "
        '" AND '@' NOT IN (A.DESCRIP) "

DoEvents
msConn.Execute sqltxt
'-ProgBar.Value = 30
sqltxt = "SELECT " & cTop10 & " A.DEPTO,B.DESCRIP,"
sqltxt = sqltxt & " B.CORTO,SUM(A.CANT) AS UNIDADES, "
sqltxt = sqltxt & " SUM(A.PRECIO) AS P_ORDEN, "
sqltxt = sqltxt & " format(SUM(A.PRECIO),'standard') AS VENTAS"
sqltxt = sqltxt & " INTO LOLO2 "
sqltxt = sqltxt & " FROM HIST_TR AS A LEFT JOIN DEPTO AS B ON A.DEPTO = B.CODIGO "

'============================================================
'INFO: 12FEB2011 AQUI TAMBIEN SE DEBE FILTRAR POR EL Z_COUNTER, NO SE ESTABA HACIENDO
If i = 0 Then
    sqltxt = sqltxt & " WHERE A.FECHA Between '" & dF1 & "'"
    sqltxt = sqltxt & " AND '" & dF2 & "'"
Else
    sqltxt = sqltxt & " WHERE VAL(A.Z_COUNTER) BETWEEN  " & GetNumber_FromArray(cArrayZ, GetMin)
    sqltxt = sqltxt & " AND " & GetNumber_FromArray(cArrayZ, Getmax)
End If
'============================================================

sqltxt = sqltxt & " GROUP BY A.DEPTO,B.DESCRIP,B.CORTO "
sqltxt = sqltxt & " ORDER BY 5 DESC "
msConn.Execute sqltxt
'-ProgBar.Value = 40
msConn.CommitTrans

'-ProgBar.Value = 50
sqltxt = "SELECT a.DEPTO,a.DESCRIP,"
sqltxt = sqltxt & " a.CORTO,A.UNIDADES, "
sqltxt = sqltxt & " a.P_ORDEN, "
sqltxt = sqltxt & " a.Ventas, "
sqltxt = sqltxt & " b.Ventas as Ventas_Net, "
sqltxt = sqltxt & " format(A.VENTAS - B.VENTAS,'STANDARD') AS DESCTO, "
sqltxt = sqltxt & " format(a.ventas / " & nTotVal & ",'Percent') AS Porcentaje "
sqltxt = sqltxt & " FROM LOLO AS A LEFT JOIN LOLO2 AS B "
sqltxt = sqltxt & " ON A.DEPTO = B.DEPTO "

rsConsulta01.Open sqltxt, msConn, adOpenStatic, adLockOptimistic
'-ProgBar.Value = 60
''''Set MSHFDepto.DataSource = rsConsulta01
'''''-ProgBar.Value = 70
''''If MSHFDepto.Rows < 1 Then
''''    Set MSHFDepto.DataSource = Nothing
''''End If

''INFO: GRID NUEVO
''DD_PEDDETALLE.Tag = cTable
'DD_PEDDETALLE.DataMode = sgBound
'Set DD_PEDDETALLE.DataSource = rsConsulta01

With DD_PEDDETALLE
    '.Columns.RemoveAll True
    '.Rows.RemoveAll True
   ' used for print preview
'   Set .PrintSettings.Viewer = ARViewer21
   ' ensure grid is in unbound mode
'   Dim sLayout As String
'   sLayout = .GetLayoutString(sgLayoutXML)
'   .DataMode = sgUnbound
'   .LoadLayoutString sLayout, sgLayoutXML
   ' use the GetRows() method on the recordset object
   ' to load all the records into the grid...
    
    'INFO: 12FEB2011
   .Columns.RemoveAll True
   .DataMode = sgUnbound

   .LoadArray rsConsulta01.GetRows()
   ' define each column from the recordsets' fields collection
   For iLoop = 1 To rsConsulta01.Fields.Count
      .Columns(iLoop).Caption = rsConsulta01.Fields(iLoop - 1).Name
      .Columns(iLoop).DBField = rsConsulta01.Fields(iLoop - 1).Name
      .Columns(iLoop).Key = rsConsulta01.Fields(iLoop - 1).Name
   Next iLoop
End With

With DD_PEDDETALLE
    .ColumnClickSort = True
    .EvenOddStyle = sgEvenOddRows
    .ColorEven = vbWhite
    .ColorOdd = &HE0E0E0
    .Columns(1).Width = 800
    .Columns(2).Width = 2200
    .Columns(3).Width = 1700
    .Columns(4).Width = 1000
    .Columns(5).Width = 0
    .Columns(6).Width = 1000
    .Columns(7).Width = 1000
    .Columns(8).Width = 800
    .Columns(9).Width = 1100
    .Columns(1).Style.TextAlignment = sgAlignRightCenter
    .Columns(4).Style.TextAlignment = sgAlignRightCenter
    .Columns(6).Style.TextAlignment = sgAlignRightCenter
    .Columns(7).Style.TextAlignment = sgAlignRightCenter
    .Columns(8).Style.TextAlignment = sgAlignRightCenter
    .Columns(9).Style.TextAlignment = sgAlignRightCenter
    
    .Columns(1).SortType = sgSortTypeNumber
    .Columns(4).SortType = sgSortTypeNumber
    .Columns(6).SortType = sgSortTypeNumber
    .Columns(7).SortType = sgSortTypeNumber
    .Columns(8).SortType = sgSortTypeNumber
    '.Columns(9).SortType = sgSortTypeNumber
    
End With

rsConsulta01.Close
'''With MSHFDepto
'''    .ColWidth(0) = 800: .ColWidth(1) = 2200: .ColWidth(2) = 1700:
'''    .ColWidth(3) = 1000: .ColWidth(4) = 0: .ColWidth(5) = 1000
'''    .ColWidth(6) = 1000: .ColWidth(7) = 1000
'''    .ColAlignment(5) = flexAlignRightCenter
'''    .ColAlignment(6) = flexAlignRightCenter
'''    .ColAlignment(7) = flexAlignRightCenter
'''End With
'-ProgBar.Value = 80
msConn.BeginTrans
msConn.Execute "DROP TABLE LOLO"
msConn.Execute "DROP TABLE LOLO2"
msConn.CommitTrans
Me.MousePointer = vbDefault
'-ProgBar.Value = 0
On Error GoTo 0
Exit Sub

ErrAdm:
Me.MousePointer = vbDefault
If Err.Number = 91 Then
    EscribeLog ("ConDept.LA OPCION DE REPORTES POR REPORTE Z, NO ESTA HABILITADA")
    ShowMsg "Error en Reporte" & vbCrLf & "LA OPCION DE REPORTES POR REPORTE Z, NO ESTA HABILITADA", vbYellow, vbRed
ElseIf Err.Number = -2147217900 Or Err.Number = -2147213302 Then
    'info: EL GRID NUEVO ESTA BOUND. ASI QUE NO SE PUEDE SOLTAR AUN LA TABLA.
    'O LA LLAVE DE LA COLUMNA SE ESTA REPITIENDO
    Resume Next
Else
    ShowMsg Err.Number & " - " & Err.Description, vbRed, vbYellow, vbOKOnly
End If
Resume Next
End Sub

Private Sub cmdSalir_Click()
Unload Me
Set ConDept = Nothing
End Sub

Private Sub Command1_Click()


'Call SoloPrintPreview.SoloPrintPreviewFunction(DD_PEDDETALLE, "")
'Exit Sub
'Para La impresion de MSHFGrid  a la impresora, la propiedad
'CLIP del Grid te trae toda la informacion, pero debido a las
'diferentes longitudes del texto y numeros no son cuadros
'perfectos. Asi que la impresion del Grid va a modo manual
'CUANDO UN DIM DECLARA MAS DE UNA VARIABLE, HAY QUE INICIALIZARLAS

Dim iCtr As Integer 'Contador de Linea
Dim iCol, iFil As Integer 'Contador de Columnas
Dim cText As String
Dim ispace As Integer
Dim iLen As Integer
Dim sSubTot As Double       ''INFO: 16FEB2013
'INFO: 10FEB2017. AGREGAR TOTALES NETOS y DE DESCUENTOS
Dim sNETO As Double
Dim sDESCUENTO As Double

sSubTot = 0#

'For i = 1 To DD_PEDDETALLE.Columns.Count - 1
'    DD_PEDDETALLE.Columns(i).Width = 800
'Next
'Exit Sub

On Error GoTo ErrorPrn:
nPagina = 0
If MSChart1.Visible = True Then
    '///////////Seleccion_Impresora_Default
    MainMant.spDoc.DocBegin
    PrintTit
    ImprimeChart
    '///////////Seleccion_Impresora
    Exit Sub
End If

EscribeLog ("Impresión de Departamentos de Venta: " & ConDept.Caption & " PERIODO : Desde " & txtFecIni & " Hasta " & txtFecFin)
MainMant.spDoc.DocBegin
PrintTit    'Rutina de Titulos

'-ProgBar.Value = 10

'For iFil = 0 To MSHFDepto.Rows - 1
For iFil = 0 To DD_PEDDETALLE.RowCount - 1

    MainMant.spDoc.TextAlign = SPTA_LEFT
    If iLin > 2400 Then
        MainMant.spDoc.TextAlign = SPTA_LEFT
        PrintTit
    End If
    DD_PEDDETALLE.row = iFil
    '''For iCol = 0 To MSHFDepto.Cols - 1
    'For iCol = 0 To DD_PEDDETALLE.ColCount - 1
    For iCol = 0 To DD_PEDDETALLE.ColCount - 1
        Select Case iCol
           Case 0, 1, 3, 5, 6, 7, 8
            'Case 1, 2, 4, 6, 7, 8
                '''MSHFDepto.Col = iCol
                DD_PEDDETALLE.Col = iCol
                MainMant.spDoc.TextAlign = SPTA_LEFT
'                If iFil = 0 Then
'                Else
                    '''If IsNumeric(MSHFDepto.Text) Then ispace = 10 Else ispace = 25
                    If IsNumeric(DD_PEDDETALLE.Text) Then ispace = 10 Else ispace = 25
                    Select Case iCol
                    Case 0
                        MainMant.spDoc.TextAlign = SPTA_LEFT
                        MainMant.spDoc.TextOut 300, iLin, DD_PEDDETALLE.Text
                        '''MSHFDepto.Text
                    Case 1
                        MainMant.spDoc.TextAlign = SPTA_LEFT
                        'MainMant.spDoc.TextOut 500, iLin, FormatTexto(MSHFDepto.Text, ispace)
                        MainMant.spDoc.TextOut 450, iLin, FormatTexto(DD_PEDDETALLE.Text, ispace)
                    Case 3
                        MainMant.spDoc.TextAlign = SPTA_RIGHT
                        'MainMant.spDoc.TextOut 1100, iLin, Format(MSHFDepto.Text, "General Number")
                        MainMant.spDoc.TextOut 1100, iLin, Format(DD_PEDDETALLE.Text, "General Number")
                    Case 5
                        MainMant.spDoc.TextAlign = SPTA_RIGHT
                        'MainMant.spDoc.TextOut 1300, iLin, Format(MSHFDepto.Text, "standard")
                        MainMant.spDoc.TextOut 1400, iLin, Format(DD_PEDDETALLE.Text, "standard")
                        'sSubTot = sSubTot + Format(MSHFDepto.Text, "standard")
                        'INFO: USANDO FUNCION VAL(), YA QUE EL DD_PEDDETALLE.Text = ""
                        'sSubTot = sSubTot + Format(Val(DD_PEDDETALLE.Text), "STANDARD")
                        If DD_PEDDETALLE.Text = "" Then
                        Else
                            sSubTot = sSubTot + CSng(DD_PEDDETALLE.Text)
                        End If
                        
                    Case 6
                        MainMant.spDoc.TextAlign = SPTA_RIGHT
                        'MainMant.spDoc.TextOut 1550, iLin, Format$(Format$(MSHFDepto.Text, "standard"), "@@@@@@@@@@")
                        MainMant.spDoc.TextOut 1650, iLin, Format$(Format$(DD_PEDDETALLE.Text, "standard"), "@@@@@@@@@@")
                        If DD_PEDDETALLE.Text = "" Then
                        Else
                            sNETO = sNETO + CSng(DD_PEDDETALLE.Text)
                        End If
                    Case 7
                        MainMant.spDoc.TextAlign = SPTA_RIGHT
                        'MainMant.spDoc.TextOut 1550, iLin, Format$(Format$(MSHFDepto.Text, "standard"), "@@@@@@@@@@")
                        MainMant.spDoc.TextOut 1850, iLin, Format$(Format$(DD_PEDDETALLE.Text, "standard"), "@@@@@@@@@@")
                        If DD_PEDDETALLE.Text = "" Then
                        Else
                            sDESCUENTO = sDESCUENTO + CSng(DD_PEDDETALLE.Text)
                        End If
                    Case 8
                        MainMant.spDoc.TextAlign = SPTA_RIGHT
                        
                        'cText = cText & Format$(Format$(MSHFDepto.Text, "percent"), "@@@@@@@@@@@@")
                        'MainMant.spDoc.TextOut 1800, iLin, Format$(Format$(MSHFDepto.Text, "percent"), "@@@@@@@@@@@@")
                        cText = cText & Format$(Format$(DD_PEDDETALLE.Text, "percent"), "@@@@@@@@@@@@")
                        MainMant.spDoc.TextOut 2050, iLin, Format$(Format$(DD_PEDDETALLE.Text, "percent"), "@@@@@@@@@@@@")
                    End Select
'                End If
            End Select
    Next
    iLin = iLin + 50
    '-If ProgBar.Value < 100 Then '-ProgBar.Value = '-ProgBar.Value + 5
Next

'-ProgBar.Value = 100

iLin = iLin + 100
MainMant.spDoc.TextAlign = SPTA_LEFT
MainMant.spDoc.TextOut 500, iLin, "Sub Total VENTAS del Periodo: " & Format(sSubTot, "Currency")
iLin = iLin + 50
MainMant.spDoc.TextOut 500, iLin, "Sub Total DESCUENTOS: " & Format(sDESCUENTO, "Currency")
iLin = iLin + 50
MainMant.spDoc.TextOut 500, iLin, "Sub Total VENTAS NETAS: " & Format(sNETO, "Currency")
'MainMant.spDoc.TextOut 500, iLin, "Sub Total NETO del Periodo : " & Format(sSubTot, "Currency")

'Open "c:\mifile.txt" For Input As #1
'Printer.Print
'Do Until EOF(1)
'    Line Input #1, a$
'    Printer.Print a$
'Loop
'Close #1

'MainMant.spDoc.AboutBox
MainMant.spDoc.DoPrintPreview
'MainMant.spDoc.DoPrintPreview
On Error GoTo 0

Call Seguridad

Exit Sub

100:
Exit Sub
ErrorPrn:
    ShowMsg "¡ Ocurre algún Error con la Impresora, Intente Conectarla !", , , vbOKOnly
    Resume
End Sub


Private Sub DD_PEDDETALLE_Click()
On Error Resume Next
nDepSel = Val(DD_PEDDETALLE.Text)
c2DepSel = DD_PEDDETALLE.Rows.Current.Cells(2).value
c1DepSel = c2DepSel
On Error GoTo 0
End Sub
'
'Private Sub DD_PEDDETALLE_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
'If Button = 2 Then
'    PopupMenu MainMant.MenuSharpGrid
'End If
'End Sub

Private Sub Form_Load()
Set rsConsulta01 = New Recordset
txtFecIni = Format(Date, "SHORT DATE")
txtFecFin = Format(Date, "SHORT DATE")
nDepSel = 0

Call GetReportesZ

Call Seguridad

'txtFecFin.SetFocus
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
        Command1.Enabled = False
        Image1.Enabled = False
    Case "N"        'SIN DERECHOS
        txtFecIni.Enabled = False: txtFecFin.Enabled = False: opcTipo(0).Enabled = False: opcTipo(1).Enabled = False
        LVZ.Enabled = False: cmdEjec.Enabled = False
        DD_PEDDETALLE.Enabled = False
        Command1.Enabled = False
        Image1.Enabled = False
End Select
End Function

Private Sub Image1_Click()
Call ExportToExcelOrCSVList(DD_PEDDETALLE)
End Sub

Private Sub LVZ_ItemCheck(ByVal Item As MSComctlLib.ListItem)
opcTipo(1).value = True
End Sub

'''Private Sub MSHFDepto_EnterCell()
'''If Len(MSHFDepto.Text) = 0 Then Exit Sub
'''If MSHFDepto.Rows = 1 Then Exit Sub
'''MSHFDepto.Col = 0
'''nDepSel = MSHFDepto.Text
'''MSHFDepto.Col = 2
'''c2DepSel = MSHFDepto.Text
'''MSHFDepto.Col = 1
'''c1DepSel = MSHFDepto.Text
'''End Sub
Private Sub TabStrip1_Click()
If TabStrip1.SelectedItem.Index = 1 Then
    '''MSHFDepto.Visible = True
    DD_PEDDETALLE.Visible = True
    MSChart1.Visible = False
ElseIf TabStrip1.SelectedItem.Index = 2 Then
    '''MSHFDepto.Visible = False
    DD_PEDDETALLE.Visible = False
    MSChart1.Visible = True
    DoLinea
ElseIf TabStrip1.SelectedItem.Index = 3 Then
    '''MSHFDepto.Visible = False
    DD_PEDDETALLE.Visible = False
    MSChart1.Visible = True
    DoLinea2
End If
End Sub

Private Sub txtFecFin_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then cmdEjec.SetFocus
End Sub

Private Sub txtFecIni_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then txtFecFin.SetFocus
End Sub

Private Sub txtFecIni_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then txtFecIni.SetFocus
End Sub

Private Sub txtFecFin_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then cmdEjec.SetFocus
End Sub


'---------------------------------------------------------------------------------------
' Procedure : LoadReporteAnual
' Author    : hsequeira
' Date      : 20/01/2023
' Purpose   : muestra los datos en formato anual
'------------------------------ ---------------------------------------------------------
'
Private Sub LoadReporteAnual()
Dim cSQL As String
Dim rsReporteAnual As ADODB.Recordset
Dim nDeptoBreak As Long
Dim cDeptoDescrip As String
Dim cData As String
Dim ventameses(12) As Single
Dim i As Integer
Dim dF1 As String
Dim dF2 As String

   On Error GoTo LoadReporteAnual_Error

dF1 = Format(txtFecIni, "YYYYMMDD")
dF2 = Format(txtFecFin, "YYYYMMDD")

cSQL = "SELECT A.CODIGO, A.DESCRIP,  LEFT(B.FECHA,6) AS FECHA, "
cSQL = cSQL & " SUM(ABS(B.CANT) * B.PRECIO_UNIT) AS VTA_NETA  "
cSQL = cSQL & " FROM DEPTO AS A, HIST_TR AS B "
cSQL = cSQL & " WHERE A.CODIGO = B.DEPTO "
cSQL = cSQL & " AND B.FECHA BETWEEN '" & dF1 & "'" & " AND '" & dF2 & "'"
cSQL = cSQL & " AND '%' NOT IN (B.DESCRIP)  AND B.DESCRIP NOT LIKE  '%@@%'  "
cSQL = cSQL & " GROUP BY A.CODIGO, A.DESCRIP, LEFT(B.FECHA,6)"
cSQL = cSQL & " ORDER BY  2,3"
Set rsReporteAnual = New ADODB.Recordset

rsReporteAnual.Open cSQL, msConn, adOpenStatic, adLockOptimistic

If rsReporteAnual.EOF Then
    Me.MousePointer = vbNormal
    MsgBox "NO EXISTEN DATOS PARA ESTAS FECHAS", vbCritical
    rsReporteAnual.Close
    Set rsReporteAnual = Nothing
    Exit Sub
End If

DD_PEDDETALLE.DataMode = sgUnbound
DD_PEDDETALLE.DataRowCount = 0
DD_PEDDETALLE.DataColCount = 13
DD_PEDDETALLE.RedrawEnabled = False
DD_PEDDETALLE.AutoResizeHeadings = True

DD_PEDDETALLE.Columns(1).Caption = "Revenues": DD_PEDDETALLE.Columns(1).Width = 2100
DD_PEDDETALLE.Columns(1).Style.TextAlignment = sgAlignLeftCenter
DD_PEDDETALLE.Columns(2).Caption = "Period 1":   DD_PEDDETALLE.Columns(2).Style.TextAlignment = sgAlignRightCenter
DD_PEDDETALLE.Columns(3).Caption = "Period 2":   DD_PEDDETALLE.Columns(3).Style.TextAlignment = sgAlignRightCenter
DD_PEDDETALLE.Columns(4).Caption = "Period 3":   DD_PEDDETALLE.Columns(4).Style.TextAlignment = sgAlignRightCenter
DD_PEDDETALLE.Columns(5).Caption = "Period 4":   DD_PEDDETALLE.Columns(5).Style.TextAlignment = sgAlignRightCenter

DD_PEDDETALLE.Columns(6).Caption = "Period 5":   DD_PEDDETALLE.Columns(6).Style.TextAlignment = sgAlignRightCenter
DD_PEDDETALLE.Columns(7).Caption = "Period 6":   DD_PEDDETALLE.Columns(7).Style.TextAlignment = sgAlignRightCenter
DD_PEDDETALLE.Columns(8).Caption = "Period 7":   DD_PEDDETALLE.Columns(8).Style.TextAlignment = sgAlignRightCenter
DD_PEDDETALLE.Columns(9).Caption = "Period 8":   DD_PEDDETALLE.Columns(9).Style.TextAlignment = sgAlignRightCenter
DD_PEDDETALLE.Columns(10).Caption = "Period 9":   DD_PEDDETALLE.Columns(10).Style.TextAlignment = sgAlignRightCenter
DD_PEDDETALLE.Columns(11).Caption = "Period 10":   DD_PEDDETALLE.Columns(11).Style.TextAlignment = sgAlignRightCenter
DD_PEDDETALLE.Columns(12).Caption = "Period 11":   DD_PEDDETALLE.Columns(12).Style.TextAlignment = sgAlignRightCenter
DD_PEDDETALLE.Columns(13).Caption = "Period 12":   DD_PEDDETALLE.Columns(13).Style.TextAlignment = sgAlignRightCenter
    
Do While Not rsReporteAnual.EOF
    
    nDeptoBreak = rsReporteAnual!CODIGO
    cDeptoDescrip = rsReporteAnual!DESCRIP
    Do While nDeptoBreak = rsReporteAnual!CODIGO
        ventameses(Int(Right(rsReporteAnual!FECHA, 2))) = rsReporteAnual!VTA_NETA
        rsReporteAnual.MoveNext
         If rsReporteAnual.EOF Then Exit Do
    Loop
    'Debug.Print cDeptoDescrip
    'cData = rsReporteAnual!G_DESCRIP & "|" & rsReporteAnual!VTA_NETA & "|" & String(4, Chr(126))
    cData = cDeptoDescrip & "|" & Format(ventameses(1), "STANDARD") & "|" & Format(ventameses(2), "STANDARD") & "|" & _
        Format(ventameses(3), "STANDARD") & "|" & Format(ventameses(4), "STANDARD") & "|" & _
        Format(ventameses(5), "STANDARD") & "|" & Format(ventameses(6), "STANDARD") & "|" & _
        Format(ventameses(7), "STANDARD") & "|" & Format(ventameses(8), "STANDARD") & "|" & _
        Format(ventameses(9), "STANDARD") & "|" & Format(ventameses(10), "STANDARD") & "|" & _
        Format(ventameses(11), "STANDARD") & "|" & Format(ventameses(12), "STANDARD") & "|"
    DD_PEDDETALLE.Rows.Add sgFormatCharSeparatedValue, cData, "|"
    'rsReporteAnual.MoveNext
    For i = 1 To 12
        ventameses(i) = 0#
    Next
    If rsReporteAnual.EOF Then
        Exit Do
    End If
Loop

For i = 2 To 13
    DD_PEDDETALLE.Columns(i).Width = 1100
Next

DD_PEDDETALLE.EvenOddStyle = sgEvenOddRows
DD_PEDDETALLE.AllowAddNew = False
DD_PEDDETALLE.ColumnClickSort = False
DD_PEDDETALLE.RedrawEnabled = True

   On Error GoTo 0
   Exit Sub

LoadReporteAnual_Error:

    ShowMsg "Error " & Err.Number & " (" & Err.Description & ") in procedure LoadReporteAnual of Form ConDept"

End Sub


