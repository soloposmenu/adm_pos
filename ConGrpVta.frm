VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{65E121D4-0C60-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCHRT20.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{43135020-B751-4DDD-AB4A-B6D8A342216E}#1.0#0"; "sg20o.ocx"
Begin VB.Form ConGrpVta 
   BackColor       =   &H00B39665&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "CONSULTA DE VENTAS POR GRUPO"
   ClientHeight    =   7680
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10980
   Icon            =   "ConGrpVta.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7680
   ScaleWidth      =   10980
   StartUpPosition =   1  'CenterOwner
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
      TabIndex        =   15
      ToolTipText     =   "Obtener Ventas por Fechas Seleccionadas"
      Top             =   240
      Value           =   -1  'True
      Width           =   1215
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
   Begin DDSharpGridOLEDB2.SGGrid DD_PEDDETALLE 
      Height          =   4605
      Left            =   360
      TabIndex        =   12
      Top             =   1920
      Width           =   10335
      _cx             =   18230
      _cy             =   8123
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
      StylesCollection=   $"ConGrpVta.frx":0442
      ColumnsCollection=   $"ConGrpVta.frx":2215
      ValueItems      =   $"ConGrpVta.frx":272A
   End
   Begin MSChart20Lib.MSChart MSChart1 
      Height          =   4575
      Left            =   360
      OleObjectBlob   =   "ConGrpVta.frx":2B0F
      TabIndex        =   11
      Top             =   1920
      Visible         =   0   'False
      Width           =   10335
   End
   Begin MSComctlLib.ProgressBar ProgBar 
      Height          =   255
      Left            =   240
      TabIndex        =   10
      Top             =   6600
      Width           =   10575
      _ExtentX        =   18653
      _ExtentY        =   450
      _Version        =   393216
      Appearance      =   1
   End
   Begin VB.PictureBox Picture1 
      Height          =   375
      Left            =   3000
      ScaleHeight     =   315
      ScaleWidth      =   435
      TabIndex        =   8
      Top             =   240
      Visible         =   0   'False
      Width           =   495
   End
   Begin MSComCtl2.DTPicker txtFecIni 
      Height          =   345
      Left            =   1560
      TabIndex        =   0
      Top             =   225
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   609
      _Version        =   393216
      Format          =   139001857
      CurrentDate     =   36431
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H0000FFFF&
      Height          =   495
      Left            =   240
      Picture         =   "ConGrpVta.frx":4917
      Style           =   1  'Graphical
      TabIndex        =   9
      ToolTipText     =   "Envia Seleccion a la Impresora"
      Top             =   6960
      Width           =   735
   End
   Begin MSComctlLib.TabStrip TabStrip1 
      Height          =   5295
      Left            =   240
      TabIndex        =   4
      Top             =   1320
      Width           =   10575
      _ExtentX        =   18653
      _ExtentY        =   9340
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
      Top             =   6960
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
      Left            =   3000
      TabIndex        =   2
      Top             =   720
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
      Top             =   480
      Width           =   1695
   End
   Begin MSComCtl2.DTPicker txtFecFin 
      Height          =   345
      Left            =   1560
      TabIndex        =   1
      Top             =   720
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   609
      _Version        =   393216
      Format          =   139001857
      CurrentDate     =   36430
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFDepto 
      Height          =   1335
      Left            =   360
      TabIndex        =   13
      Top             =   1680
      Width           =   9375
      _ExtentX        =   16536
      _ExtentY        =   2355
      _Version        =   393216
      FixedCols       =   0
      FocusRect       =   2
      HighLight       =   2
      SelectionMode   =   1
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
   End
   Begin MSComctlLib.ListView LVZ 
      Height          =   825
      Left            =   6495
      TabIndex        =   16
      Top             =   240
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
      ForeColor       =   &H0000FFFF&
      Height          =   375
      Index           =   2
      Left            =   5400
      TabIndex        =   17
      Top             =   570
      Width           =   975
   End
   Begin VB.Shape Borde1 
      BorderColor     =   &H0000FFFF&
      BorderWidth     =   2
      Height          =   1000
      Index           =   0
      Left            =   240
      Top             =   120
      Width           =   5055
   End
   Begin VB.Shape Borde1 
      BorderColor     =   &H0000FFFF&
      BorderWidth     =   2
      Height          =   1000
      Index           =   1
      Left            =   5280
      Top             =   120
      Width           =   3735
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   5280
      Picture         =   "ConGrpVta.frx":4C21
      ToolTipText     =   "Exportar Datos"
      Top             =   6960
      Width           =   480
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
      ForeColor       =   &H0000FFFF&
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
      ForeColor       =   &H0000FFFF&
      Height          =   255
      Index           =   0
      Left            =   360
      TabIndex        =   5
      Top             =   360
      Width           =   1215
   End
End
Attribute VB_Name = "ConGrpVta"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private rsConsulta01 As Recordset
Private mintCurFrame As Integer ' Current Frame visible
Private nPagina As Integer
Private iLin As Integer
Dim nDepSel As Integer
Dim c1DepSel As String, c2DepSel As String
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
MainMant.spDoc.TextOut 300, 450, ConGrpVta.Caption

If ConDept.opcTipo(0).value = True Then
    MainMant.spDoc.TextOut 300, 550, "PERIODO : Desde " & txtFecIni & " Hasta " & txtFecFin
Else
    'INFO: 14NOV2012
    MainMant.spDoc.TextOut 300, 550, "REPORTE(S) Z # INCLUIDOS: " & cmdEjec.Tag
End If
'If Not MSChart1.Visible Then
    MainMant.spDoc.TextOut 300, 650, "GRUPO"
    MainMant.spDoc.TextOut 1100, 650, "Ventas"
    MainMant.spDoc.TextOut 1350, 650, "Ventas Netas"
    MainMant.spDoc.TextOut 1650, 650, "Porcentaje"
    MainMant.spDoc.TextOut 300, 700, "--------------------------------------------------------------------------------------------------------------------------------"
'End If

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
    MSChart1.Row = MiRow
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
    MSChart1.Row = MiRow
    MSChart1.RowLabel = Mid(rsGrafico!FECHA, 7, 2) & "/" & Mid(rsGrafico!FECHA, 5, 2)
    MSChart1.Column = 1
    MSChart1.Data = Format(rsGrafico!Ventas, "standard")
    MSChart1.Column = 2
    MSChart1.Data = Format(rsGrafico!UNIDADES, "standard")
    MiRow = MiRow + 1
    rsGrafico.MoveNext
Loop
rsGrafico.Close

End Sub
Private Sub cmdEjec_Click()
Dim cSQL As String
Dim cTop10 As String
Dim dF1 As String
Dim dF2 As String
Dim rsTmp01 As Recordset
Dim nTotVal As Double       'INFO: 16FEB2013
'INFO: 14NOV2012
Dim i As Byte
Dim iZZ As Integer  'Z Counter
Dim jFalseCounter As Byte
Dim cZetas As String

On Error GoTo ErrAdm:

'---------------------------------------------------------------------------------------------
'INFO: 14NOV2012
'---------------------------------------------------------------------------------------------
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
        MsgBox "Debe seleccionar al menos un Reporte Z", vbInformation, "Seleccione un reporte (Z)"
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
''---------------------------------------------------------------------------------------------

'-------------------------------
DD_PEDDETALLE.Visible = True
'MSChart1.Visible = False
'-------------------------------

'ProgBar.value = 5
If Top10.value = 1 Then cTop10 = " TOP 10 " Else cTop10 = ""
    
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
    'INFO: VALIDAR LA Z CONTRA VALORES NUMERICOS EN VEZ DE TEXTO (27OCT2010)
    cSQL = cSQL & " WHERE VAL(Z_COUNTER) BETWEEN  " & GetNumber_FromArray(cArrayZ, GetMin)
    cSQL = cSQL & " AND " & GetNumber_FromArray(cArrayZ, Getmax)
End If

rsTmp01.Open cSQL, msConn, adOpenStatic, adLockOptimistic
'ProgBar.value = 10
If IsNull(rsTmp01!VALOR) Then
    '''Set MSHFDepto.DataSource = Nothing
    '''MSHFDepto.Clear
    '-ProgBar.Value = 0
    Me.MousePointer = vbDefault
    ShowMsg "NO EXISTE INFORMACION PARA MOSTRAR. SELECCIONE OTRA(S) FECHA(S)"
    'DD_PEDDETALLE.Rows.RemoveAll True
    Exit Sub
End If

nTotVal = rsTmp01!VALOR
rsTmp01.Close
ProgBar.value = 20
'''MSHFDepto.Clear

'NO INCLUYE DESCUENTOS
cSQL = "SELECT " & cTop10 & " A.DEPTO,B.DESCRIP,"
cSQL = cSQL & " B.CORTO,SUM(A.CANT) AS UNIDADES, "
cSQL = cSQL & " SUM(A.PRECIO) AS P_ORDEN, "
cSQL = cSQL & " format(SUM(A.PRECIO),'standard') AS VENTAS"
cSQL = cSQL & " INTO LOLO "
cSQL = cSQL & " FROM HIST_TR AS A LEFT JOIN DEPTO AS B ON A.DEPTO = B.CODIGO "
If i = 0 Then
    cSQL = cSQL & " WHERE A.FECHA Between '" & dF1 & "'"
    cSQL = cSQL & " AND '" & dF2 & "'"
Else
    cSQL = cSQL & " WHERE VAL(A.Z_COUNTER) BETWEEN  " & GetNumber_FromArray(cArrayZ, GetMin)
    cSQL = cSQL & " AND " & GetNumber_FromArray(cArrayZ, Getmax)
End If
'INFO: 14NOV2012. REVISION QUE EXCLUYA DESCUENTOS
'cSQL = cSQL & " AND MID(A.DESCRIP,LEN(TRIM(A.DESCRIP)),1) <> '%' "
'cSQL = cSQL & " AND '%' NOT IN (A.DESCRIP) "
cSQL = cSQL & " AND '%' NOT IN (A.DESCRIP) "
cSQL = cSQL & " AND A.DESCRIP NOT LIKE '%DESCUENTO%' "
cSQL = cSQL & " AND A.DESCRIP NOT LIKE  '%@@%' "
cSQL = cSQL & " GROUP BY A.DEPTO,B.DESCRIP,B.CORTO "
cSQL = cSQL & " ORDER BY 5 DESC "

DoEvents

msConn.BeginTrans
msConn.Execute cSQL
msConn.CommitTrans

ProgBar.value = 30
cSQL = "SELECT " & cTop10 & " A.DEPTO,B.DESCRIP,"
cSQL = cSQL & " B.CORTO,SUM(A.CANT) AS UNIDADES, "
cSQL = cSQL & " SUM(A.PRECIO) AS P_ORDEN, "
cSQL = cSQL & " format(SUM(A.PRECIO),'standard') AS VENTAS"
cSQL = cSQL & " INTO LOLO2 "
cSQL = cSQL & " FROM HIST_TR AS A LEFT JOIN DEPTO AS B ON A.DEPTO = B.CODIGO "
If i = 0 Then
    cSQL = cSQL & " WHERE A.FECHA Between '" & dF1 & "'"
    cSQL = cSQL & " AND '" & dF2 & "'"
Else
    cSQL = cSQL & " WHERE VAL(A.Z_COUNTER) BETWEEN  " & GetNumber_FromArray(cArrayZ, GetMin)
    cSQL = cSQL & " AND " & GetNumber_FromArray(cArrayZ, Getmax)
End If
cSQL = cSQL & " GROUP BY A.DEPTO,B.DESCRIP,B.CORTO "
cSQL = cSQL & " ORDER BY 5 DESC "

msConn.BeginTrans
msConn.Execute cSQL
ProgBar.value = 40
msConn.CommitTrans

ProgBar.value = 50
cSQL = "SELECT a.DEPTO,a.DESCRIP,"
cSQL = cSQL & " a.CORTO,A.UNIDADES, "
cSQL = cSQL & " a.P_ORDEN, "
cSQL = cSQL & " a.Ventas, "
cSQL = cSQL & " b.Ventas as Ventas_Net "
cSQL = cSQL & " INTO LOLO3 "
cSQL = cSQL & " FROM LOLO AS A LEFT JOIN LOLO2 AS B "
cSQL = cSQL & " ON A.DEPTO = B.DEPTO "

msConn.BeginTrans
msConn.Execute cSQL
msConn.CommitTrans

ProgBar.value = 60
cSQL = "SELECT C.DESCRIP AS GRUPO,"
cSQL = cSQL & " FORMAT(SUM(A.VENTAS),'STANDARD') AS VENTAS, "
cSQL = cSQL & " FORMAT(SUM(A.VENTAS_NET),'STANDARD') AS VENTAS_NETAS, "
cSQL = cSQL & " format(SUM(A.VENTAS) / " & nTotVal & ",'Percent') AS Porcentaje "
cSQL = cSQL & " FROM LOLO3 AS A, SUPER_DET AS B, SUPER_GRP AS C "
cSQL = cSQL & " WHERE A.DEPTO = B.DEPTO AND B.GRUPO = C.GRUPO "
cSQL = cSQL & " GROUP BY C.DESCRIP "
cSQL = cSQL & " ORDER BY 1 "

rsConsulta01.Open cSQL, msConn, adOpenStatic, adLockOptimistic
ProgBar.value = 70
'''Set MSHFDepto.DataSource = rsConsulta01

If rsConsulta01.EOF Then
    Me.MousePointer = vbDefault
    ShowMsg "NO EXISTE INFORMACION PARA MOSTRAR. SELECCIONE OTRA(S) FECHA(S)"
    rsConsulta01.Close
    DD_PEDDETALLE.Rows.RemoveAll True
    GoTo ErrAdm:
End If

With DD_PEDDETALLE
   .Columns.RemoveAll True
   .Rows.RemoveAll True
   .DataMode = sgUnbound
   .LoadArray rsConsulta01.GetRows()
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
    .Columns(1).Width = 1700
    .Columns(2).Width = 1300
    .Columns(3).Width = 1200
    .Columns(4).Width = 1200
    '.Columns(1).Style.TextAlignment = sgAlignRightCenter
    .Columns(2).Style.TextAlignment = sgAlignRightCenter
    .Columns(3).Style.TextAlignment = sgAlignRightCenter
    .Columns(4).Style.TextAlignment = sgAlignRightCenter
End With

ProgBar.value = 80
'''If MSHFDepto.Rows < 1 Then
'''    Set MSHFDepto.DataSource = Nothing
'''End If
Me.Refresh
rsConsulta01.Close
'''With MSHFDepto
'''    .ColWidth(0) = 1700: .ColWidth(1) = 1300: .ColWidth(2) = 1200:
'''    '.ColAlignment(0) = flexAlignRightCenter
'''    .ColAlignment(1) = flexAlignRightCenter
'''    .ColAlignment(2) = flexAlignRightCenter
''''    .ColAlignment(7) = flexAlignRightCenter
'''End With
ProgBar.value = 90

msConn.BeginTrans
msConn.Execute "DROP TABLE LOLO"
msConn.Execute "DROP TABLE LOLO2"
msConn.Execute "DROP TABLE LOLO3"
msConn.CommitTrans

Me.MousePointer = vbDefault
ProgBar.value = 0
On Error GoTo 0
Exit Sub

ErrAdm:
Select Case Err.Number
    Case 0
        msConn.BeginTrans
        msConn.Execute "DROP TABLE LOLO"
        msConn.Execute "DROP TABLE LOLO2"
        msConn.Execute "DROP TABLE LOLO3"
        msConn.CommitTrans
        ProgBar.value = 0
    Case 3021
        ShowMsg "NO HAY REGISTROS PARA MOSTAR. FAVOR SALIR DEL PROGRAMA y VOLVER A ENTRAR", vbRed, vbYellow
        ProgBar.value = 0
    Case Else
        ShowMsg Err.Number & " - " & Err.Description, vbYellow, vbRed
End Select
Me.MousePointer = vbDefault
End Sub

Private Sub cmdSalir_Click()
Unload Me
End Sub

Private Sub Command1_Click()
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
Dim sSubTot As Double              'INFO: 16FEB2013

sSubTot = 0#

On Error GoTo ErrorPrn:
nPagina = 0
'If MSChart1.Visible = True Then
'    Seleccion_Impresora_Default
'    MainMant.spDoc.DocBegin
'    PrintTit
'    ImprimeChart
'    Seleccion_Impresora
'    Exit Sub
'End If

EscribeLog ("Impresión de Grupos de Venta: " & ConGrpVta.Caption & " PERIODO : Desde " & txtFecIni & " Hasta " & txtFecFin)
MainMant.spDoc.DocBegin
PrintTit    'Rutina de Titulos

'-ProgBar.Value = 10

For iFil = 0 To DD_PEDDETALLE.RowCount - 1
    If iLin > 2400 Then PrintTit
    'If ProgBar.Value < 100 Then
        '-ProgBar.Value = '-ProgBar.Value + 10
    'End If
    DD_PEDDETALLE.Row = iFil
    On Error Resume Next
    For iCol = 0 To DD_PEDDETALLE.ColCount - 1
        Select Case iCol
            '''Case 0, 1, 2
            'Case 1, 2, 3
            Case 0, 1, 2, 3
                DD_PEDDETALLE.Col = iCol
                MainMant.spDoc.TextAlign = SPTA_LEFT
''''''''                If iFil = 1 Then
''''''''                Else
                    If IsNumeric(DD_PEDDETALLE.Text) Then ispace = 10 Else ispace = 25
                    Select Case iCol
                    Case 0
                        MainMant.spDoc.TextAlign = SPTA_LEFT
                        MainMant.spDoc.TextOut 300, iLin, DD_PEDDETALLE.Text
                    Case 1
                        MainMant.spDoc.TextAlign = SPTA_RIGHT
                        MainMant.spDoc.TextOut 1200, iLin, Format$(Format$(DD_PEDDETALLE.Text, "standard"), "@@@@@@@@@@")
                        sSubTot = sSubTot + Format(DD_PEDDETALLE.Text, "standard")
                    Case 2
                        MainMant.spDoc.TextAlign = SPTA_RIGHT
                        MainMant.spDoc.TextOut 1550, iLin, Format$(Format$(DD_PEDDETALLE.Text, "standard"), "@@@@@@@@@@")
                        'sSubTot = sSubTot + Format(DD_PEDDETALLE.Text, "standard")
                    Case 3
                        MainMant.spDoc.TextAlign = SPTA_RIGHT
                        MainMant.spDoc.TextOut 1800, iLin, Format$(Format$(DD_PEDDETALLE.Text, "percent"), "@@@@@@@@@@@@")
                    End Select
''''''''                End If
            End Select
    Next
    On Error GoTo 0
    iLin = iLin + 50
    '-If ProgBar.Value < 100 Then '-ProgBar.Value = '-ProgBar.Value + 5
Next

'-ProgBar.Value = 100

iLin = iLin + 100
MainMant.spDoc.TextAlign = SPTA_LEFT
MainMant.spDoc.TextOut 500, iLin, "Sub Total del Periodo = " & Format(sSubTot, "Currency")

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
    'MsgBox "¡ Ocurre algún Error con la Impresora, Intente Conecterla !", vbExclamation, BoxTit
    ShowMsg " ¡ Ocurre algún Error con la Impresora, Intente Conecterla !", vbRed, vbYellow
    Resume 100
End Sub

Private Sub Form_Load()
Set rsConsulta01 = New Recordset
txtFecIni = Format(Date, "SHORT DATE")
txtFecFin = Format(Date, "SHORT DATE")
nDepSel = 0
'txtFecFin.SetFocus

Call GetReportesZ

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
        Image1.Enabled = False
    Case "N"        'SIN DERECHOS
        txtFecIni.Enabled = False: txtFecFin.Enabled = False
        cmdEjec.Enabled = False
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

'Private Sub MSHFDepto_EnterCell()
'If Len(MSHFDepto.Text) = 0 Then Exit Sub
'If MSHFDepto.Rows = 1 Then Exit Sub
'MSHFDepto.Col = 0
'nDepSel = MSHFDepto.Text
'MSHFDepto.Col = 2
'c2DepSel = MSHFDepto.Text
'MSHFDepto.Col = 1
'c1DepSel = MSHFDepto.Text
'End Sub

Private Sub TabStrip1_Click()
If TabStrip1.SelectedItem.Index = 1 Then
    DD_PEDDETALLE.Visible = True
    MSChart1.Visible = False
ElseIf TabStrip1.SelectedItem.Index = 2 Then
    DD_PEDDETALLE.Visible = False
    MSChart1.Visible = True
    DoLinea
ElseIf TabStrip1.SelectedItem.Index = 3 Then
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

