VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{65E121D4-0C60-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCHRT20.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{43135020-B751-4DDD-AB4A-B6D8A342216E}#1.0#0"; "sg20o.ocx"
Begin VB.Form conPlu 
   BackColor       =   &H00B39665&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "CONSULTA DE VENTAS POR PRODUCTO"
   ClientHeight    =   8490
   ClientLeft      =   165
   ClientTop       =   4455
   ClientWidth     =   11415
   Icon            =   "conPlu.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8490
   ScaleWidth      =   11415
   StartUpPosition =   2  'CenterScreen
   Begin DDSharpGridOLEDB2.SGGrid DD_PEDDETALLE 
      Height          =   4750
      Left            =   3240
      TabIndex        =   25
      Top             =   2160
      Width           =   7935
      _cx             =   13996
      _cy             =   8378
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
      StylesCollection=   $"conPlu.frx":0442
      ColumnsCollection=   $"conPlu.frx":2271
      ValueItems      =   $"conPlu.frx":2BE7
   End
   Begin VB.PictureBox Picture2 
      BackColor       =   &H00B39665&
      BorderStyle     =   0  'None
      Height          =   735
      Left            =   8160
      ScaleHeight     =   735
      ScaleWidth      =   3195
      TabIndex        =   19
      Top             =   120
      Width           =   3200
      Begin VB.OptionButton Option4 
         BackColor       =   &H00B39665&
         Caption         =   "Por  DEPTO"
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
         Left            =   1560
         TabIndex        =   23
         Top             =   360
         Width           =   1575
      End
      Begin VB.OptionButton Option2 
         BackColor       =   &H00B39665&
         Caption         =   "Por  Nombre"
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
         Left            =   0
         TabIndex        =   22
         Top             =   360
         Width           =   1455
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H00B39665&
         Caption         =   "Por Ventas"
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
         Left            =   0
         TabIndex        =   21
         Top             =   120
         Width           =   1335
      End
      Begin VB.OptionButton Option3 
         BackColor       =   &H00B39665&
         Caption         =   "Por  Unidades"
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
         Left            =   1560
         TabIndex        =   20
         Top             =   120
         Width           =   1575
      End
   End
   Begin VB.CheckBox chkSinPrecio 
      BackColor       =   &H00C0FFC0&
      Caption         =   "Mostrar productos sin precio"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   6720
      TabIndex        =   18
      ToolTipText     =   "Muestra en la consulta los productos con ventas en 0.00"
      Top             =   1290
      Width           =   2775
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
      Left            =   3000
      TabIndex        =   17
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
      Left            =   4440
      TabIndex        =   15
      ToolTipText     =   "Obtener Ventas por Reporte Z"
      Top             =   240
      Width           =   1095
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H0000FFFF&
      Height          =   495
      Left            =   120
      Picture         =   "conPlu.frx":2C87
      Style           =   1  'Graphical
      TabIndex        =   13
      ToolTipText     =   "Envia Seleccion a la Impresora"
      Top             =   7800
      Width           =   735
   End
   Begin VB.TextBox Text2 
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   6480
      TabIndex        =   12
      Text            =   "Sub-Total"
      Top             =   7065
      Width           =   975
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  'Right Justify
      Enabled         =   0   'False
      Height          =   285
      Left            =   7560
      TabIndex        =   11
      Top             =   6960
      Width           =   1215
   End
   Begin VB.ListBox List1 
      Height          =   4740
      Left            =   240
      TabIndex        =   10
      Top             =   2160
      Width           =   3015
   End
   Begin MSChart20Lib.MSChart MSChart1 
      Height          =   4575
      Left            =   240
      OleObjectBlob   =   "conPlu.frx":2F91
      TabIndex        =   4
      Top             =   2160
      Visible         =   0   'False
      Width           =   10695
   End
   Begin VB.PictureBox Picture1 
      Height          =   255
      Left            =   2760
      ScaleHeight     =   195
      ScaleWidth      =   315
      TabIndex        =   9
      Top             =   120
      Visible         =   0   'False
      Width           =   375
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
      Left            =   9600
      TabIndex        =   3
      Top             =   960
      Width           =   1695
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
      Height          =   375
      Left            =   2880
      TabIndex        =   2
      Top             =   720
      Width           =   1335
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
      Left            =   10080
      TabIndex        =   6
      Top             =   7800
      Width           =   1215
   End
   Begin MSComCtl2.DTPicker txtFecIni 
      Height          =   345
      Left            =   1440
      TabIndex        =   0
      Top             =   240
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   609
      _Version        =   393216
      Format          =   157876225
      CurrentDate     =   36418
   End
   Begin MSComCtl2.DTPicker txtFecFin 
      Height          =   345
      Left            =   1440
      TabIndex        =   1
      Top             =   720
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   609
      _Version        =   393216
      Format          =   157876225
      CurrentDate     =   36418
   End
   Begin MSComctlLib.TabStrip TabStrip1 
      Height          =   6135
      Left            =   120
      TabIndex        =   5
      Top             =   1560
      Width           =   11175
      _ExtentX        =   19711
      _ExtentY        =   10821
      MultiRow        =   -1  'True
      ShowTips        =   0   'False
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
   Begin MSComctlLib.ProgressBar ProgBar 
      Height          =   180
      Left            =   120
      TabIndex        =   14
      Top             =   1320
      Width           =   6495
      _ExtentX        =   11456
      _ExtentY        =   318
      _Version        =   393216
      Appearance      =   1
   End
   Begin MSComctlLib.ListView LVZ 
      Height          =   825
      Left            =   5560
      TabIndex        =   24
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
   Begin VB.Image Image1 
      Height          =   480
      Left            =   5760
      Picture         =   "conPlu.frx":4D99
      ToolTipText     =   "Exportar Datos"
      Top             =   7920
      Width           =   480
   End
   Begin VB.Shape Borde1 
      BorderColor     =   &H0000FFFF&
      BorderWidth     =   2
      Height          =   1095
      Index           =   0
      Left            =   120
      Top             =   120
      Width           =   4215
   End
   Begin VB.Shape Borde1 
      BorderColor     =   &H0000FFFF&
      BorderWidth     =   2
      Height          =   1095
      Index           =   1
      Left            =   4320
      Top             =   120
      Width           =   3735
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
      Left            =   4320
      TabIndex        =   16
      Top             =   765
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
      Left            =   240
      TabIndex        =   8
      Top             =   360
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
      Left            =   240
      TabIndex        =   7
      Top             =   840
      Width           =   1095
   End
End
Attribute VB_Name = "conPlu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private rsConsulta01 As Recordset
Private mintCurFrame As Integer ' Current Frame visible
Private rsConDept As Recordset
Private nPagina As Integer
Private iLin As Integer
Private nconDepto As Integer
Private nPluSel As Integer
Private c1PluSel As String, c2PluSel As String
Private cOrdSel As String
Private Function GetDepto(nDepto As Integer) As String
Dim rsFunDEPTO As ADODB.Recordset
Dim cSQL As String

Set rsFunDEPTO = New ADODB.Recordset
cSQL = "SELECT DESCRIP FROM DEPTO WHERE CODIGO = " & nDepto
rsFunDEPTO.Open cSQL, msConn, adOpenStatic, adLockOptimistic
If Not rsFunDEPTO.EOF Then
    GetDepto = rsFunDEPTO!DESCRIP
Else
    GetDepto = "DEPARTAMENTO ELIMINADO"
End If
rsFunDEPTO.Close
Set rsFunDEPTO = Nothing
End Function

Private Sub GetReportesZ()
Dim cSQL As String
Dim rsZetas As ADODB.Recordset
Dim iLinea As Integer
Dim nZetasINI As Long

On Error GoTo ErrAdm:
Set rsZetas = New ADODB.Recordset

nZetasINI = CLng(GetFromINI("Administracion", "MaxZ", App.Path & "\soloini.ini"))

cSQL = "SELECT TOP " & nZetasINI & " VAL(CONTADOR) AS CONTADOR, FECHA FROM Z_COUNTER ORDER BY VAL(CONTADOR) DESC "
'cSQL = "SELECT TOP 100 CONTADOR, FECHA FROM Z_COUNTER ORDER BY CONTADOR DESC "
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
On Error GoTo ErrChart:
    Picture1.Visible = True
    MSChart1.EditCopy
    Picture1.Picture = Clipboard.GetData()
    Printer.PaintPicture Picture1.Picture, 0, 3000
    Picture1.Visible = False
    Printer.EndDoc
On Error GoTo 0
Exit Sub

ErrChart:
    MsgBox "Error de Impresión (Papel o Conexión). " & Err.Description
    Resume Next
End Sub
Private Sub PrintTit()
Dim cOrderTXT As String

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
Select Case cOrdSel
    Case " 5 DESC " 'POR VENTAS
        cOrderTXT = Space(10) & " (Ordenado por Ventas)"
    Case " 3 ASC " 'POR DESCRICION
        cOrderTXT = Space(10) & " (Ordenado por Descripción)"
    Case " 4 DESC " 'POR CANTIDADES VENDIDAS
        cOrderTXT = Space(10) & " (Ordenado por Cantidades Vendidas)"
    Case " 9 ASC, 3 ASC " 'POR DEPARTAMENTO
        cOrderTXT = Space(10) & " (Ordenado por Departamento)"
    Case Else
        cOrderTXT = ""
End Select
MainMant.spDoc.TextOut 300, 450, conPlu.Caption & cOrderTXT

If conPlu.opcTipo(0).value = True Then
    MainMant.spDoc.TextOut 300, 500, "PERIODO : Desde " & txtFecIni & " Hasta " & txtFecFin
Else
    MainMant.spDoc.TextOut 300, 500, "PERIODO : REPORTE Z # " & cmdEjec.Tag
End If

If chkSinPrecio.value = vbChecked Then
    MainMant.spDoc.TextOut 300, 550, "Muestra productos con ventas en " & Format(0#, "CURRENCY") & "    " & "Departamento Seleccionado: " & Trim(Left(List1.Text, 50))
Else
    MainMant.spDoc.TextOut 300, 550, "Departamento Seleccionado: " & Trim(Left(List1.Text, 50))
End If

If Not MSChart1.Visible Then
    MainMant.spDoc.TextOut 300, 650, "CODIGO"
    MainMant.spDoc.TextOut 500, 650, "DESCRIPCION"
    MainMant.spDoc.TextOut 1000, 650, "Unidades"
    MainMant.spDoc.TextOut 1300, 650, "Ventas"
    MainMant.spDoc.TextOut 1450, 650, "Ventas Netas"
    MainMant.spDoc.TextOut 1750, 650, "Descuento"
    MainMant.spDoc.TextOut 300, 700, "------------------------------------------------------------------------------------------------------------------------------------------"
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

If nPluSel = 0 Then
    MsgBox "DEBE SELECCIONAR UN PRODUCTO", vbInformation, BoxTit
    Exit Sub
End If

If Top10.value = 1 Then cTop10 = " TOP 10 " Else cTop10 = ""

Set rsGrafico = New Recordset

sqltxt = "SELECT " & cTop10 & " A.PLU,MID(A.FECHA,5,2) AS MES, " & _
        " A.CANT AS UNIDADES, " & _
        " format(A.PRECIO,'standard') AS VENTAS" & _
        " INTO LOLO FROM HIST_TR AS A " & _
        " WHERE A.PLU = " & nPluSel & _
        " AND A.FECHA >= '" & dF1 & "'" & _
        " AND A.FECHA <= '" & dF2 & "'"

msConn.BeginTrans
msConn.Execute sqltxt
msConn.CommitTrans

sqltxt = "SELECT A.PLU,A.MES, " & _
        " SUM(A.UNIDADES) AS CANT, " & _
        " format(SUM(A.VENTAS),'standard') AS PRECIO" & _
        " FROM LOLO AS A " & _
        " GROUP BY A.PLU,A.MES " & _
        " ORDER BY A.MES"

rsGrafico.Open sqltxt, msConn, adOpenStatic, adLockOptimistic

With MSChart1
    .chartType = VtChChartType2dLine
    .RowCount = rsGrafico.RecordCount
    .TitleText = "Ventas Mensuales de " & c2PluSel
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

If nPluSel = 0 Then
    MsgBox "DEBE SELECCIONAR UN PRODUCTO", vbInformation, BoxTit
    Exit Sub
End If

If Top10.value = 1 Then cTop10 = " TOP 10 " Else cTop10 = ""

Set rsGrafico = New Recordset
sqltxt = "SELECT " & cTop10 & " A.PLU,A.FECHA,SUM(A.CANT) AS UNIDADES, " & _
        " format(SUM(A.PRECIO),'standard') AS VENTAS" & _
        " FROM HIST_TR AS A " & _
        " WHERE A.PLU = " & nPluSel & _
        " AND A.FECHA >= '" & dF1 & "'" & _
        " AND A.FECHA <= '" & dF2 & "'" & _
        " GROUP BY A.PLU,A.FECHA " & _
        " ORDER BY A.FECHA "
rsGrafico.Open sqltxt, msConn, adOpenStatic, adLockOptimistic

With MSChart1
    .chartType = VtChChartType2dLine
    .RowCount = rsGrafico.RecordCount
    .TitleText = "Ventas Diarias de " & c2PluSel
End With
MiRow = 1
Do Until rsGrafico.EOF
    MSChart1.Row = MiRow
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
Dim cEldepto As String
Dim nTotal As Double        ''INFO: 16FEB2013
Dim i As Long
Dim j As Long
Dim iZZ As Integer  'Z Counter
Dim jFalseCounter As Byte
Dim cSQL As String, cZetas As String
Dim iLoop As Integer

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
        MsgBox "Debe seleccionar al menos un Reporte Z", vbInformation, "Seleccione un reporte (Z)"
        Exit Sub
    Else
        cZetas = "'" & Mid(cZetas, 1, Len(cZetas) - 2)
    End If
    cmdEjec.Tag = cZetas
    cArrayZ = Split(Replace(cZetas, "'", ""), ",")
    On Error GoTo 0
End If

'-------------------------------
''''MSHFDepto.Visible = True
DD_PEDDETALLE.Visible = True
MSChart1.Visible = False
'-------------------------------

If Top10.value = 1 Then cTop10 = " TOP 10 " Else cTop10 = ""
'If nconDepto = 0 Then cEldepto = "" Else cEldepto = " AND A.DEPTO = " & nconDepto
'INFO ERROR 3159 DE BOOKMARK, NO LO PUEDE HACER CON EL SIGNO DE IGUAL, CAMBIANDO A LIKE
If nconDepto = 0 Then cEldepto = "" Else cEldepto = " AND A.DEPTO LIKE " & nconDepto
dF1 = Format(txtFecIni, "YYYYMMDD")
dF2 = Format(txtFecFin, "YYYYMMDD")
nTotal = 0

Me.MousePointer = vbHourglass
'MSHFDepto.Clear
'SELECCIONA PLU's NO TOMA ENCUENTA DESCUENTOS
sqltxt = "SELECT " & cTop10 & " A.DEPTO,A.PLU AS CODIGO,"
sqltxt = sqltxt & " iif(IsNull(B.CONTENEDOR),0,B.CONTENEDOR) as Envase,"
sqltxt = sqltxt & " iif(IsNull(B.DESCRIP),'','-' + B.DESCRIP) AS Descrip, "
sqltxt = sqltxt & " SUM(A.CANT) AS UNID, "
sqltxt = sqltxt & " SUM(A.PRECIO) AS P_PRECIO, "
sqltxt = sqltxt & " FORMAT(SUM(A.PRECIO),'STANDARD') AS VENTAS "
sqltxt = sqltxt & " INTO LOLO1 "
sqltxt = sqltxt & " FROM HIST_TR AS A LEFT JOIN CONTENED AS B "
sqltxt = sqltxt & " ON A.ENVASE =B.CONTENEDOR "
If i = 0 Then
    sqltxt = sqltxt & " WHERE A.FECHA >= '" & dF1 & "'"
    sqltxt = sqltxt & " AND A.FECHA <= '" & dF2 & "'"
    sqltxt = sqltxt & cEldepto
Else
    'INFO: EL PROGRAMA REVIENTA SI HAY VENTAS EN EL DEPARTAMENTO ABIERTO (18/MAY/2009)
    'sqltxt = sqltxt & " WHERE A.Z_COUNTER = '" & Val(LVZ.SelectedItem.Text) & "'"
    'INFO: Z REPORT (19ABR2008)
    'sqltxt = sqltxt & " WHERE A.Z_COUNTER IN (" & cZetas & ") "
    
    'INFO: Z REPORT CON MAXIMOS Y MINIMOS (21/MAY/2009)
    'INFO: VALIDAR LA Z CONTRA VALORES NUMERICOS EN VEZ DE TEXTO (27OCT2010)
    sqltxt = sqltxt & " WHERE VAL(A.Z_COUNTER) BETWEEN  " & GetNumber_FromArray(cArrayZ, GetMin)
    sqltxt = sqltxt & " AND " & GetNumber_FromArray(cArrayZ, Getmax)
    sqltxt = sqltxt & cEldepto
End If
'ES NECESARIA LA SIGUIENTE LINEA PARA VER LOS DESCUENTOS.
'Y RESTARLOS CORRECTAMENTE.
'INFO: 13/JUL/2009
sqltxt = sqltxt & " AND MID(A.DESCRIP,LEN(TRIM(A.DESCRIP)),1) <> '%' "
sqltxt = sqltxt & " AND '%' NOT IN (A.DESCRIP) "
sqltxt = sqltxt & " GROUP BY A.DEPTO,A.PLU,B.CONTENEDOR,B.DESCRIP "

If chkSinPrecio.value = vbChecked Then
    'sqltxt = sqltxt & " HAVING SUM(A.CANT) > 0 "
Else
    'INFO: AQUI SOLAMENTE MUESTRA PRODUCTOS CON VENTAS.
    'NO APARECERAN LAS CORTESIAS o PRODUCTOS SIN VENTAS
    'POR EJEMPLO: SI UN PRODUCTO VALE 0.00 Y SE VENDE 100 VECES, ESE
    'PRODUCTO NO APARECERA EN EL LISTADO
    sqltxt = sqltxt & " HAVING SUM(A.PRECIO) > 0.00 and SUM(A.CANT) > 0 "
End If

sqltxt1 = "SELECT " & cTop10 & " A.DEPTO,A.PLU AS CODIGO,"
sqltxt1 = sqltxt1 & " iif(IsNull(B.CONTENEDOR),0,B.CONTENEDOR) as Envase,"
sqltxt1 = sqltxt1 & " iif(IsNull(B.DESCRIP),'','-' + B.DESCRIP) AS Descrip, "
sqltxt1 = sqltxt1 & " SUM(A.CANT) AS UNID, "
sqltxt1 = sqltxt1 & " SUM(A.PRECIO) AS P_PRECIO, "
sqltxt1 = sqltxt1 & " FORMAT(SUM(A.PRECIO),'STANDARD') AS VENTAS "
sqltxt1 = sqltxt1 & " INTO LOLO2 "
sqltxt1 = sqltxt1 & " FROM HIST_TR AS A LEFT JOIN CONTENED AS B "
sqltxt1 = sqltxt1 & " ON A.ENVASE =B.CONTENEDOR "

If i = 0 Then
    sqltxt1 = sqltxt1 & " WHERE A.FECHA >= '" & dF1 & "'"
    sqltxt1 = sqltxt1 & " AND A.FECHA <= '" & dF2 & "'"
    sqltxt1 = sqltxt1 & cEldepto
Else
    'sqltxt1 = sqltxt1 & " WHERE A.Z_COUNTER = '" & Val(LVZ.SelectedItem.Text) & "'"
    'INFO: Z REPORT (19ABR2008)
    'sqltxt1 = sqltxt1 & " WHERE A.Z_COUNTER IN (" & cZetas & ") "
    
    'INFO: Z REPORT CON MAXIMOS Y MINIMOS (21/MAY/2009)
    'INFO: 27OCT2010. CAMBIANDO LA COMPARACION DE DATOS EN EL REQUEST DE BETWEEN
    'YA QUE AL HACERLO POR CADENA DE TEXTO SE INCLUIAN OTROS VALORES, POR EJEMPO
    'SI JALO LA Z "60" AND "59", TAMBIEN SE VENIA LA Z_6
    sqltxt1 = sqltxt1 & " WHERE VAL(A.Z_COUNTER) BETWEEN  " & GetNumber_FromArray(cArrayZ, GetMin)
    sqltxt1 = sqltxt1 & " AND " & GetNumber_FromArray(cArrayZ, Getmax)
    
    sqltxt1 = sqltxt1 & cEldepto
End If
sqltxt1 = sqltxt1 & " GROUP BY A.DEPTO,A.PLU,B.CONTENEDOR,B.DESCRIP "

If chkSinPrecio.value = vbChecked Then
    'sqltxt1 = sqltxt1 & " HAVING SUM(A.CANT) > 0 "
Else
    'INFO: AQUI SOLAMENTE MUESTRA PRODUCTOS CON VENTAS.
    'NO APARECERAN LAS CORTESIAS o PRODUCTOS SIN VENTAS
    'POR EJEMPLO: SI UN PRODUCTO VALE 0.00 Y SE VENDE 100 VECES, ESE
    'PRODUCTO NO APARECERA EN EL LISTADO
    sqltxt1 = sqltxt1 & " HAVING SUM(A.PRECIO) > 0.00 and SUM(A.CANT) > 0 "
End If

''Debug.Print "SQLTXT = " & sqltxt
''Debug.Print "SQLTXT1 = " & sqltxt1

msConn.BeginTrans
msConn.Execute sqltxt
msConn.Execute sqltxt1
msConn.CommitTrans

msConn.BeginTrans
msConn.Execute "SELECT A.*,B.VENTAS AS VENT_DESC " & _
        " INTO LOLO3 " & _
        " FROM LOLO1 AS A LEFT JOIN LOLO2 AS B " & _
        " ON A.DEPTO = B.DEPTO " & _
        " AND A.CODIGO = B.CODIGO " & _
        " AND A.ENVASE = B.ENVASE "

msConn.CommitTrans

If cOrdSel = " 9 ASC, 3 ASC " Then
    ProgBar.value = 30
    sqltxt = "SELECT A.CODIGO,A.ENVASE, "
    sqltxt = sqltxt & " (B.DESCRIP + A.DESCRIP) AS Descrip, "
    sqltxt = sqltxt & " A.UNID, A.P_PRECIO, A.VENTAS, A.VENT_DESC, "
    sqltxt = sqltxt & " FORMAT((A.VENTAS - A.VENT_DESC),'STANDARD') AS DESCUENT, A.DEPTO "
    sqltxt = sqltxt & " INTO LOLO4 "
    sqltxt = sqltxt & " FROM LOLO3 AS A LEFT JOIN PLU AS B "
    sqltxt = sqltxt & " ON A.CODIGO = B.CODIGO "
    sqltxt = sqltxt & " ORDER BY " & cOrdSel
    
    msConn.BeginTrans
    msConn.Execute sqltxt
    msConn.CommitTrans
      
    ProgBar.value = 40
    sqltxt = "SELECT A.*, B.DESCRIP AS DEPARTAMENTO"
    sqltxt = sqltxt & " FROM LOLO4 AS A LEFT JOIN DEPTO AS B "
    sqltxt = sqltxt & " ON A.DEPTO = B.CODIGO "
    sqltxt = sqltxt & " ORDER BY 10 ASC, 3 ASC"
Else
    ProgBar.value = 30
    sqltxt = "SELECT A.CODIGO,A.ENVASE, "
    sqltxt = sqltxt & " (B.DESCRIP + A.DESCRIP) AS Descrip, "
    sqltxt = sqltxt & " A.UNID, A.P_PRECIO, A.VENTAS, A.VENT_DESC, "
    sqltxt = sqltxt & " FORMAT((A.VENTAS - A.VENT_DESC),'STANDARD') AS DESCUENT, A.DEPTO "
    sqltxt = sqltxt & " FROM LOLO3 AS A LEFT JOIN PLU AS B "
    sqltxt = sqltxt & " ON A.CODIGO = B.CODIGO "
    sqltxt = sqltxt & " ORDER BY " & cOrdSel
End If

rsConsulta01.Open sqltxt, msConn, adOpenStatic, adLockOptimistic
ProgBar.value = 50

Dim arreglo(0, 0)
DD_PEDDETALLE.LoadArray arreglo

If rsConsulta01.EOF Then
    Text1.Enabled = True
    Text1 = Format(0#, "standard")
    Text1.Enabled = False
    '''Set MSHFDepto.DataSource = Nothing
    '''MSHFDepto.Clear
    '''MSHFDepto.Refresh
    rsConsulta01.Close
    msConn.BeginTrans
    msConn.Execute "DROP TABLE LOLO1"
    msConn.Execute "DROP TABLE LOLO2"
    msConn.Execute "DROP TABLE LOLO3"
    msConn.CommitTrans
    
    If cOrdSel = " 9 ASC, 3 ASC " Then
        msConn.BeginTrans
        msConn.Execute "DROP TABLE LOLO4"
        msConn.CommitTrans
    End If
    
    Me.MousePointer = vbDefault
    ProgBar.value = 0
    'DD_PEDDETALLE.Rows.RemoveAll True
    ShowMsg "NO EXISTE INFORMACION PARA MOSTRAR. SELECCIONE OTRA(S) FECHA(S)", , , vbOKOnly
    Exit Sub
Else
    '''Set MSHFDepto.DataSource = rsConsulta01
End If

With DD_PEDDETALLE
   .Columns.RemoveAll True
'   .Rows.RemoveAll True
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
    .Columns(1).Width = 800:    'CODIGO
    .Columns(2).Width = 0:      'ENVASE
    .Columns(3).Width = 2800:   'DESCRIPCION
    .Columns(4).Width = 700:    'UNIDADES
    .Columns(5).Width = 0:      'P_PRECIO
    .Columns(6).Width = 1100:   'VENTAS
    .Columns(7).Width = 1000:   'VENTAS_DESC
    .Columns(8).Width = 1100     'DESCUENTO
    .Columns(9).Width = 0       'DEPTO
    On Error Resume Next
        .Columns(10).Width = 0       'DEPTO DESCRIP
    On Error GoTo 0
    On Error GoTo ErrAdm:
    .Columns(1).Style.TextAlignment = sgAlignRightCenter
    .Columns(4).Style.TextAlignment = sgAlignRightCenter
    .Columns(6).Style.TextAlignment = sgAlignRightCenter
    .Columns(7).Style.TextAlignment = sgAlignRightCenter
    .Columns(8).Style.TextAlignment = sgAlignRightCenter
    
    'INFO: FEB2011
    .Columns(1).SortType = sgSortTypeNumber
    .Columns(4).SortType = sgSortTypeNumber
    .Columns(6).SortType = sgSortTypeNumber
    .Columns(7).SortType = sgSortTypeNumber
    .Columns(8).SortType = sgSortTypeNumber
End With

rsConsulta01.MoveFirst
Do Until rsConsulta01.EOF
    nTotal = nTotal + rsConsulta01!VENTAS
    rsConsulta01.MoveNext
Loop
ProgBar.value = 70
Text1.Enabled = True
Text1 = Format(nTotal, "standard")
Text1.Enabled = False
rsConsulta01.Close
ProgBar.value = 90

msConn.Execute "DROP TABLE LOLO1"
msConn.Execute "DROP TABLE LOLO2"
msConn.Execute "DROP TABLE LOLO3"

    If cOrdSel = " 9 ASC, 3 ASC " Then
        msConn.BeginTrans
        msConn.Execute "DROP TABLE LOLO4"
        msConn.CommitTrans
    End If

ProgBar.value = 0
Me.MousePointer = vbDefault
On Error GoTo 0

Call Seguridad

Exit Sub

ErrAdm:
Me.MousePointer = vbDefault
If Err.Number = 91 Then
    EscribeLog ("Admin." & "conPLU.Seleccion de Productos. LA OPCION DE REPORTES POR REPORTE Z, NO ESTA HABILITADA")
    MsgBox "LA OPCION DE REPORTES POR REPORTE Z, NO ESTA HABILITADA", vbCritical, "Error en Reporte"
ElseIf Err.Number = -2147217900 Or Err.Number = -2147213302 Then
    'info: EL GRID NUEVO ESTA BOUND. ASI QUE NO SE PUEDE SOLTAR AUN LA TABLA.
    'O LA LLAVE DE LA COLUMNA SE ESTA REPITIENDO
    Resume Next
Else
    EscribeLog ("Admin." & "ConPLU.Seleccion de Productos " & Err.Number & " - " & Err.Description)
    MsgBox Err.Number & " - " & Err.Description, vbCritical, "Error en Reporte de Productos"
End If
End Sub

Private Sub cmdSalir_Click()
Unload Me
End Sub
Private Sub Command1_Click()
Dim iCol, iFil As Integer 'Contador de Columnas
Dim ispace As Integer
Dim sVentaNeta As Double        ''INFO: 16FEB2013
Dim sVentaBruta As Double        ''INFO: 16FEB2013
Dim sDescuentos As Double        ''INFO: 16FEB2013
Dim nDepto As Integer
Dim nPrevCol As Integer
Dim bPrintDepto As Boolean
Dim nSubNETOTotDepto As Double        ''INFO: 16FEB2013
Dim nSubBRUTOTotDepto As Double        ''INFO: 16FEB2013
Dim nSubUNIDDepto As Double        ''INFO: 16FEB2013
Dim cDeptoDescript As String

sVentaNeta = 0#: iLin = 8: nPagina = 0
sVentaBruta = 0#
sDescuentos = 0#

If cOrdSel = " 9 ASC, 3 ASC " Then bPrintDepto = True Else bPrintDepto = False

On Error GoTo ErrorPrn:

'      MSChart1.EditCopy
'      Picture1.Picture = Clipboard.GetData()

If MSChart1.Visible = True Then
    nPagina = 0
    '///////////Seleccion_Impresora_Default
    PrintTit
    ImprimeChart
    '///////////Seleccion_Impresora
    Exit Sub
End If

MainMant.spDoc.DocBegin
PrintTit
EscribeLog ("Admin." & "Impresion de Listado: " & Me.Caption & " PERIODO : Desde " & txtFecIni & " Hasta " & txtFecFin)

ProgBar.value = 10

'nPrevCol = MSHFDepto.Col
'MSHFDepto.Col = 8
'nDepto = Val(MSHFDepto.Text)
'MSHFDepto.Col = nPrevCol
On Error GoTo 0
On Error Resume Next
''''For iFil = 0 To MSHFDepto.Rows - 1
For iFil = 0 To DD_PEDDETALLE.RowCount - 1
    'iLin = iLin + 50
    If iLin > 2400 Then
        MainMant.spDoc.TextAlign = SPTA_LEFT
        PrintTit
    End If
    If ProgBar.value < 100 Then
        ProgBar.value = ProgBar.value + 5
    End If
    DD_PEDDETALLE.Row = iFil
    For iCol = 0 To DD_PEDDETALLE.ColCount - 1
        Select Case iCol
            'Case 0, 2, 3, 5, 6, 7
            'Case 2, 3, 4, 6, 7, 8
            Case 0, 2, 3, 5, 6, 7
            'Case 0, 2, 3, 5, 6, 7
                DD_PEDDETALLE.Col = iCol
                'Debug.Print iCol & " - " & DD_PEDDETALLE.Text
                MainMant.spDoc.TextAlign = SPTA_LEFT
'                If iFil = 0 Then
'                Else
                    If IsNumeric(DD_PEDDETALLE.Text) Then ispace = 10 Else ispace = 25
                    '*************************************************************
                    If bPrintDepto Then
                        nPrevCol = DD_PEDDETALLE.Col
                        DD_PEDDETALLE.Col = 8
                        If nDepto <> Val(DD_PEDDETALLE.Text) Then
                            nDepto = Val(DD_PEDDETALLE.Text)
                            cDeptoDescript = Trim(GetDepto(nDepto))
                            If iFil <> 1 Then
                                MainMant.spDoc.TextAlign = SPTA_LEFT
                                MainMant.spDoc.TextOut 300, iLin, "(TOTALES DEL DEPARTAMENTO)"
                                MainMant.spDoc.TextAlign = SPTA_RIGHT
                                MainMant.spDoc.TextOut 1150, iLin, Format(nSubUNIDDepto, "STANDARD")
                                MainMant.spDoc.TextOut 1400, iLin, Format(nSubBRUTOTotDepto, "STANDARD")
                                MainMant.spDoc.TextOut 1650, iLin, Format(nSubNETOTotDepto, "STANDARD")
                                iLin = iLin + 100
                            End If
                            nSubBRUTOTotDepto = 0
                            nSubNETOTotDepto = 0
                            nSubUNIDDepto = 0
                            MainMant.spDoc.TextAlign = SPTA_LEFT
                            MainMant.spDoc.TextOut 300, iLin, "***** " & cDeptoDescript & " *****"
                            iLin = iLin + 50
                        End If
                        DD_PEDDETALLE.Col = nPrevCol
                    End If
                    '*************************************************************
   
                    Select Case iCol
                    Case 0  'CODIGO PLU
                        MainMant.spDoc.TextAlign = SPTA_LEFT
                        MainMant.spDoc.TextOut 300, iLin, DD_PEDDETALLE.Text
                    Case 2  'DESCRIPCION
                        MainMant.spDoc.TextAlign = SPTA_LEFT
                        MainMant.spDoc.TextOut 500, iLin, FormatTexto(DD_PEDDETALLE.Text, ispace)
                    Case 3  'UNIDADES
                        MainMant.spDoc.TextAlign = SPTA_RIGHT
                        MainMant.spDoc.TextOut 1150, iLin, Format(DD_PEDDETALLE.Text, "General Number")
                        nSubUNIDDepto = nSubUNIDDepto + Format(DD_PEDDETALLE.Text, "General Number")
                    Case 5 'VENTAS BRUTA
                        MainMant.spDoc.TextAlign = SPTA_RIGHT
                        MainMant.spDoc.TextOut 1400, iLin, Format(DD_PEDDETALLE.Text, "STANDARD")
                        On Error Resume Next
                        sVentaBruta = sVentaBruta + Format(DD_PEDDETALLE.Text, "STANDARD")
                        nSubBRUTOTotDepto = nSubBRUTOTotDepto + Format(DD_PEDDETALLE.Text, "STANDARD")
                        On Error GoTo 0
                    Case 6 'VENTA NETA
                    'Print Format$(Format$(MSHFDepto.Text, "standard"), "@@@@@@@@@")
                        MainMant.spDoc.TextAlign = SPTA_RIGHT
                        MainMant.spDoc.TextOut 1650, iLin, Format$(Format$(DD_PEDDETALLE.Text, "STANDARD"), "@@@@@@@@@@")
                        On Error Resume Next
                        sVentaNeta = sVentaNeta + Format(DD_PEDDETALLE.Text, "STANDARD")
                        nSubNETOTotDepto = nSubNETOTotDepto + Format(DD_PEDDETALLE.Text, "STANDARD")
                        On Error GoTo 0
                    Case 7  'DESCUENTOS
                        MainMant.spDoc.TextAlign = SPTA_RIGHT
                        MainMant.spDoc.TextOut 1900, iLin, Format$(Format$(DD_PEDDETALLE.Text, "STANDARD"), "@@@@@@@@@@@@@")
                        On Error Resume Next
                        sDescuentos = sDescuentos + Format(DD_PEDDETALLE.Text, "STANDARD")
                        On Error GoTo 0
                    End Select
'                End If
            End Select
        Next
        iLin = iLin + 50
Next

If bPrintDepto Then
        MainMant.spDoc.TextAlign = SPTA_LEFT
        MainMant.spDoc.TextOut 300, iLin, "(TOTALES DEL DEPARTAMENTO)"
        MainMant.spDoc.TextAlign = SPTA_RIGHT
        MainMant.spDoc.TextOut 1150, iLin, Format(nSubUNIDDepto, "STANDARD")
        MainMant.spDoc.TextOut 1400, iLin, Format(nSubBRUTOTotDepto, "STANDARD")
        MainMant.spDoc.TextOut 1650, iLin, Format(nSubNETOTotDepto, "STANDARD")
'        iLin = iLin + 100
End If

'Printer.PaintPicture Picture1.Picture, 0, 0
'Printer.EndDoc
ProgBar.value = 95
iLin = iLin + 100
MainMant.spDoc.TextAlign = SPTA_LEFT
MainMant.spDoc.TextOut 500, iLin, "Sub Total Ventas           : " & Format(sVentaBruta, "Currency")
iLin = iLin + 50
MainMant.spDoc.TextOut 500, iLin, "Sub Total NETO del Periodo : " & Format(sVentaNeta, "Currency")
iLin = iLin + 50
MainMant.spDoc.TextOut 500, iLin, "Sub Total Descuentos       : " & Format(sDescuentos, "Currency")
MainMant.spDoc.DoPrintPreview
100:
ProgBar.value = 0
On Error GoTo 0
Exit Sub
ErrorPrn:
    'MsgBox Err.Number & " - " & Err.Description & vbCrLf & _
        "¡ Ocurre algún Error con la Impresora, Intente Conectarla !", vbExclamation, BoxTit
    ShowMsg Err.Number & " - " & Err.Description & vbCrLf & "¡ Ocurre algún Error con la Impresora, Intente Conecterla !", vbRed, vbYellow
    Resume 100
End Sub




Private Sub DD_PEDDETALLE_Click()
'nPluSel = DD_PEDDETALLE.RowS(
On Error Resume Next
nPluSel = Val(DD_PEDDETALLE.Rows.Current.Cells(0).value)
c2PluSel = DD_PEDDETALLE.Rows.Current.Cells(2).value
c1PluSel = c2PluSel
On Error GoTo 0
End Sub

Private Sub Form_Load()
Set rsConsulta01 = New Recordset
txtFecIni = Format(Date, "SHORT DATE")
txtFecFin = Format(Date, "SHORT DATE")

nPluSel = 0
cOrdSel = " 5 DESC " 'POR VENTAS
Set rsConDept = New Recordset
rsConDept.Open "SELECT CODIGO,DESCRIP FROM DEPTO ORDER BY DESCRIP ", msConn, adOpenStatic, adLockOptimistic
List1.AddItem "TODOS LOS DEPARTAMENTOS"
Do Until rsConDept.EOF
    List1.AddItem rsConDept!DESCRIP & Space(70) & rsConDept!CODIGO
    rsConDept.MoveNext
Loop
nconDepto = 0
rsConDept.Close
nPagina = 1

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
        Command1.Enabled = False: Image1.Enabled = False
    Case "N"        'SIN DERECHOS
        txtFecIni.Enabled = False: txtFecFin.Enabled = False: LVZ.Enabled = False: cmdEjec.Enabled = False
        List1.Enabled = False: DD_PEDDETALLE.Enabled = False
        Command1.Enabled = False: Image1.Enabled = False
End Select
End Function

Private Sub Image1_Click()
Call ExportToExcelOrCSVList(DD_PEDDETALLE)
End Sub

Private Sub List1_Click()
Dim POSIC As Integer
'CAPTURA EL NUMERO DEL DEPTO.
POSIC = Len(List1.Text)
If IsNumeric(Val(Mid(List1.Text, POSIC - 5, 6))) Then
    nconDepto = Val(Mid(List1.Text, POSIC - 5, 6))
Else
    nconDepto = 0
End If
cmdEjec_Click
End Sub
'''Private Sub MSHFDepto_Click()
'''If Len(MSHFDepto.Text) = 0 Then Exit Sub
'''If MSHFDepto.Rows < 1 Then Exit Sub
'''MSHFDepto.Col = 0
'''If IsNumeric(MSHFDepto.Text) Then nPluSel = MSHFDepto.Text
'''MSHFDepto.Col = 2
'''c2PluSel = MSHFDepto.Text
'''MSHFDepto.Col = 0
'''End Sub

Private Sub LVZ_ItemCheck(ByVal Item As MSComctlLib.ListItem)
opcTipo(1).value = True
End Sub

Private Sub Option1_Click()
cOrdSel = " 5 DESC " 'POR VENTAS
End Sub

Private Sub Option2_Click()
cOrdSel = " 3 ASC " 'POR DESCRICION
End Sub

Private Sub Option3_Click()
cOrdSel = " 4 DESC " 'POR CANTIDADES VENDIDAS
End Sub

Private Sub Option4_Click()
cOrdSel = " 9 ASC, 3 ASC " 'POR DEPARTAMENTO
End Sub

Private Sub TabStrip1_Click()
If TabStrip1.SelectedItem.Index = 1 Then
    List1.Visible = True
    DD_PEDDETALLE.Visible = True
    MSChart1.Visible = False
ElseIf TabStrip1.SelectedItem.Index = 2 Then
    List1.Visible = False
    DD_PEDDETALLE.Visible = False
    MSChart1.Visible = True
    DoLinea
ElseIf TabStrip1.SelectedItem.Index = 3 Then
    List1.Visible = False
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
If KeyAscii = 13 Then
    cmdEjec.SetFocus
End If
End Sub

