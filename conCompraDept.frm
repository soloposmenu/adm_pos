VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{43135020-B751-4DDD-AB4A-B6D8A342216E}#1.0#0"; "sg20o.ocx"
Begin VB.Form conCompraDept 
   BackColor       =   &H00B39665&
   Caption         =   "Consulta de Compras x Departamento de Inventario"
   ClientHeight    =   7350
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   10905
   Icon            =   "conCompraDept.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   7350
   ScaleWidth      =   10905
   StartUpPosition =   1  'CenterOwner
   Begin DDSharpGridOLEDB2.SGGrid DD_PEDDETALLE 
      Height          =   4965
      Left            =   360
      TabIndex        =   13
      Top             =   1440
      Width           =   9375
      _cx             =   16536
      _cy             =   8758
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
      StylesCollection=   $"conCompraDept.frx":0442
      ColumnsCollection=   $"conCompraDept.frx":2215
      ValueItems      =   $"conCompraDept.frx":272A
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
      Left            =   5760
      TabIndex        =   12
      Top             =   540
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
      Left            =   5760
      TabIndex        =   11
      Top             =   240
      Width           =   1335
   End
   Begin VB.CommandButton cmdEjec 
      Caption         =   "&Ejecutar Consulta"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   9000
      TabIndex        =   7
      Top             =   360
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
      Height          =   255
      Left            =   1320
      TabIndex        =   6
      Top             =   720
      Visible         =   0   'False
      Width           =   2055
   End
   Begin VB.CommandButton cmdSalir 
      Caption         =   "Sa&lir"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   555
      Left            =   9480
      TabIndex        =   5
      Top             =   6720
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H0000FFFF&
      Height          =   495
      Left            =   120
      Picture         =   "conCompraDept.frx":2B0F
      Style           =   1  'Graphical
      TabIndex        =   3
      ToolTipText     =   "Envia Seleccion a la Impresora"
      Top             =   6720
      Width           =   735
   End
   Begin VB.PictureBox Picture1 
      Height          =   375
      Left            =   2880
      ScaleHeight     =   315
      ScaleWidth      =   435
      TabIndex        =   1
      Top             =   0
      Visible         =   0   'False
      Width           =   495
   End
   Begin MSComctlLib.ProgressBar ProgBar 
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   6360
      Width           =   10575
      _ExtentX        =   18653
      _ExtentY        =   450
      _Version        =   393216
      Appearance      =   1
   End
   Begin MSComCtl2.DTPicker txtFecIni 
      Height          =   345
      Left            =   1320
      TabIndex        =   2
      Top             =   345
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   609
      _Version        =   393216
      Format          =   65077249
      CurrentDate     =   36431
   End
   Begin MSComctlLib.TabStrip TabStrip1 
      Height          =   5295
      Left            =   120
      TabIndex        =   4
      Top             =   1080
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
   Begin MSComCtl2.DTPicker txtFecFin 
      Height          =   345
      Left            =   3960
      TabIndex        =   8
      Top             =   360
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   609
      _Version        =   393216
      Format          =   65077249
      CurrentDate     =   36430
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   5160
      Picture         =   "conCompraDept.frx":2E19
      ToolTipText     =   "Exportar Datos"
      Top             =   6720
      Width           =   480
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
      Left            =   120
      TabIndex        =   10
      Top             =   480
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
      Left            =   2760
      TabIndex        =   9
      Top             =   480
      Width           =   1215
   End
End
Attribute VB_Name = "conCompraDept"
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
Dim cOrdSel As String
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
MainMant.spDoc.TextOut 300, 450, conCompraDept.Caption
MainMant.spDoc.TextOut 300, 550, "PERIODO : Desde " & txtFecIni & " Hasta " & txtFecFin

'If Not MSChart1.Visible Then
    MainMant.spDoc.TextOut 300, 650, "DEPARTAMENTO"
    MainMant.spDoc.TextOut 1100, 650, "Compras"
    MainMant.spDoc.TextOut 300, 700, "--------------------------------------------------------------------------------------------------------------------------------"
'End If

iLin = 750
nPagina = nPagina + 1
End Sub

Private Sub cmdEjec_Click()
Dim cSQL As String
Dim cTop10 As String
Dim dF1 As String
Dim dF2 As String
Dim rsTmp01 As ADODB.Recordset
Dim rsConsulta01 As New ADODB.Recordset
Dim nTotVal As Single

'-------------------------------
DD_PEDDETALLE.Visible = True
'MSChart1.Visible = False
'-------------------------------

ProgBar.value = 5
If Top10.value = 1 Then cTop10 = " TOP 10 " Else cTop10 = ""
    
dF1 = Format(txtFecIni, "YYYYMMDD")
dF2 = Format(txtFecFin, "YYYYMMDD")

ProgBar.value = 20
'''MSHFDepto.Clear
'NO INCLUYE DESCUENTOS
cSQL = "SELECT C.CODIGO,MAX(C.DESCRIP) AS DEPARTAMENTO, "
cSQL = cSQL & " SUM (B.COSTO_IN) AS o_COMPRAS, "
cSQL = cSQL & " format(SUM (B.COSTO_IN),'STANDARD') AS COMPRAS "
cSQL = cSQL & " FROM COMPRAS_HEAD AS A, COMPRAS_DETA AS B, DEP_INV AS C "
cSQL = cSQL & " WHERE A.FECHA Between '" & dF1 & "'"
cSQL = cSQL & " AND '" & dF2 & "'"
cSQL = cSQL & " AND A.INDICE = B.NUM_COMPRA "
cSQL = cSQL & " AND B.DEPT_INV = C.CODIGO "
cSQL = cSQL & " GROUP BY C.CODIGO "
cSQL = cSQL & " ORDER BY " & cOrdSel

ProgBar.value = 40
rsConsulta01.Open cSQL, msConn, adOpenStatic, adLockOptimistic
ProgBar.value = 70

If rsConsulta01.EOF Then
    MsgBox "NO EXISTEN DATOS PARA EL PERIODO SELECCIONADO", vbInformation, BoxTit
    rsConsulta01.Close
    Set rsConsulta01 = Nothing
    ProgBar.value = 0
    DD_PEDDETALLE.Rows.RemoveAll True
    Exit Sub
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
    .Columns(1).Width = 0: .Columns(2).Width = 2500: .Columns(3).Width = 0:
    .Columns(4).Width = 1100
    .Columns(2).Style.TextAlignment = sgAlignRightCenter
    .Columns(4).Style.TextAlignment = sgAlignRightCenter
End With

'''Set MSHFDepto.DataSource = rsConsulta01
ProgBar.value = 80
'''If MSHFDepto.Rows < 1 Then
'''    Set MSHFDepto.DataSource = Nothing
'''End If
Me.Refresh
rsConsulta01.Close
'''With MSHFDepto
'''    .ColWidth(0) = 0: .ColWidth(1) = 2500: .ColWidth(2) = 0: .ColWidth(3) = 1100:
'''    '.ColAlignment(0) = flexAlignRightCenter
'''    '.ColAlignment(1) = flexAlignRightCenter
'''    .ColAlignment(3) = flexAlignRightCenter
''''    .ColAlignment(7) = flexAlignRightCenter
'''End With
ProgBar.value = 0

Call Seguridad

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
Dim sSubTot As Single

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

EscribeLog ("Impresión de Compras x Departamento: " & ConGrpVta.Caption & " PERIODO : Desde " & txtFecIni & " Hasta " & txtFecFin)
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
            'Case 0, 1, 2, 3
            Case 1, 2, 3, 4
                DD_PEDDETALLE.Col = iCol
                MainMant.spDoc.TextAlign = SPTA_LEFT
                If iFil = 1 Then
                Else
                    If IsNumeric(DD_PEDDETALLE.Text) Then ispace = 10 Else ispace = 25
                    Select Case iCol
                    Case 2
                        MainMant.spDoc.TextAlign = SPTA_LEFT
                        MainMant.spDoc.TextOut 300, iLin, DD_PEDDETALLE.Text
                    Case 3
                        MainMant.spDoc.TextAlign = SPTA_RIGHT
                        MainMant.spDoc.TextOut 1250, iLin, Format$(Format$(DD_PEDDETALLE.Text, "STANDARD"), "@@@@@@@@@@@@")
                        sSubTot = sSubTot + Format(DD_PEDDETALLE.Text, "standard")
                    End Select
                End If
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
cOrdSel = " 2 ASC " 'POR DESCRICION

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
        txtFecIni.Enabled = False: txtFecFin.Enabled = False
        Option1.Enabled = False: Option2.Enabled = False: cmdEjec.Enabled = False
        Command1.Enabled = False: Image1.Enabled = False
End Select
End Function

Private Sub Image1_Click()
Call ExportToExcelOrCSVList(DD_PEDDETALLE)
End Sub

Private Sub Option1_Click()
cOrdSel = " 3 DESC " 'POR VENTAS
End Sub

Private Sub Option2_Click()
cOrdSel = " 2 ASC " 'POR DESCRICION
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
