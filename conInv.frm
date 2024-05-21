VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{65E121D4-0C60-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mschrt20.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{43135020-B751-4DDD-AB4A-B6D8A342216E}#1.0#0"; "sg20o.ocx"
Begin VB.Form conInv 
   BackColor       =   &H00B39665&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "CONSULTA DE CONSUMO POR PRODUCTO DE INVENTARIO"
   ClientHeight    =   8085
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11415
   Icon            =   "conInv.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8085
   ScaleWidth      =   11415
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin DDSharpGridOLEDB2.SGGrid DD_PEDDETALLE 
      Height          =   5415
      Left            =   3240
      TabIndex        =   15
      Top             =   1680
      Width           =   7935
      _cx             =   13996
      _cy             =   9551
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
      StylesCollection=   $"conInv.frx":000C
      ColumnsCollection=   $"conInv.frx":1E3B
      ValueItems      =   $"conInv.frx":27B1
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
      Left            =   9840
      TabIndex        =   8
      Top             =   7440
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
      Height          =   375
      Left            =   5280
      TabIndex        =   7
      Top             =   360
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
      Left            =   9360
      TabIndex        =   6
      Top             =   360
      Width           =   1695
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
      Left            =   7200
      TabIndex        =   5
      Top             =   660
      Width           =   1455
   End
   Begin VB.PictureBox Picture1 
      Height          =   255
      Left            =   2760
      ScaleHeight     =   195
      ScaleWidth      =   315
      TabIndex        =   4
      Top             =   0
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.OptionButton Option1 
      BackColor       =   &H00B39665&
      Caption         =   "Por Cantidad"
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
      Left            =   7200
      TabIndex        =   2
      Top             =   360
      Width           =   1575
   End
   Begin VB.ListBox List1 
      Height          =   5325
      Left            =   240
      TabIndex        =   1
      Top             =   1680
      Width           =   3015
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H0000FFFF&
      Height          =   495
      Left            =   120
      Picture         =   "conInv.frx":2851
      Style           =   1  'Graphical
      TabIndex        =   0
      ToolTipText     =   "Envia Seleccion a la Impresora"
      Top             =   7440
      Width           =   735
   End
   Begin MSChart20Lib.MSChart MSChart1 
      Height          =   4575
      Left            =   240
      OleObjectBlob   =   "conInv.frx":2B5B
      TabIndex        =   3
      Top             =   1680
      Visible         =   0   'False
      Width           =   10695
   End
   Begin MSComCtl2.DTPicker txtFecIni 
      Height          =   345
      Left            =   1320
      TabIndex        =   9
      Top             =   360
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   609
      _Version        =   393216
      Format          =   63897601
      CurrentDate     =   36418
   End
   Begin MSComCtl2.DTPicker txtFecFin 
      Height          =   345
      Left            =   3840
      TabIndex        =   10
      Top             =   360
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   609
      _Version        =   393216
      Format          =   63897601
      CurrentDate     =   36418
   End
   Begin MSComctlLib.TabStrip TabStrip1 
      Height          =   6135
      Left            =   120
      TabIndex        =   11
      Top             =   1080
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
      TabIndex        =   12
      Top             =   840
      Width           =   6855
      _ExtentX        =   12091
      _ExtentY        =   318
      _Version        =   393216
      Appearance      =   1
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   5400
      Picture         =   "conInv.frx":4963
      ToolTipText     =   "Exportar Datos"
      Top             =   7440
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
      Height          =   255
      Index           =   1
      Left            =   2760
      TabIndex        =   14
      Top             =   480
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
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   13
      Top             =   480
      Width           =   1215
   End
End
Attribute VB_Name = "conInv"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private rsConsulta01 As Recordset
Private mintCurFrame As Integer ' Current Frame visible
Private rsConDept As Recordset
Private nPagina As Integer
Private iLin As Integer
Private nconDepto As Integer
Private nPluSel As Integer
Private c1PluSel As String, c2PluSel As String
Private cOrdSel As String

Private Sub cmdEjec_Click()
Dim sqltxt As String
Dim cTop10 As String
Dim dF1 As String
Dim dF2 As String
Dim cEldepto As String
Dim nTotal As Single
Dim i As Long
Dim j As Long
Dim iLoop As Integer
'-------------------------------
DD_PEDDETALLE.Visible = True
MSChart1.Visible = False
'-------------------------------

On Error GoTo ErrAdm:
If Top10.value = 1 Then cTop10 = " TOP 10 " Else cTop10 = ""
If nconDepto = 0 Then cEldepto = "" Else cEldepto = " AND B.ID_DEPT_INV = " & nconDepto
dF1 = Format(txtFecIni, "YYYYMMDD")
dF2 = Format(txtFecFin, "YYYYMMDD")
nTotal = 0

ProgBar.value = 20

On Error Resume Next
msConn.BeginTrans
msConn.Execute "DROP TABLE LOLO6"
msConn.CommitTrans
On Error GoTo 0

On Error GoTo ErrAdm:

ProgBar.value = 30

'''''INFO: QUERY REVISADO
'RELACION ENTRE VENTAS y EL ENLACE DE LOS PRODUCTOS DE VENTAS CON EL INVCENTARIO
sqltxt = "SELECT " & cTop10
sqltxt = sqltxt & " B.ID_PROD_INV, MAX(B.ID_DEPT_INV) AS DEPTO, max(B.DESCRIP) AS PROD_INVENT,"
sqltxt = sqltxt & " sum(A.CANT * B.CANT) AS VENTAS, MAX(C.DESCRIP) AS MEDIDA"
sqltxt = sqltxt & " Into LOLO6"
sqltxt = sqltxt & " FROM HIST_TR AS A, PLU_INVENT AS B, UNID_CONSUMO AS C"
sqltxt = sqltxt & " WHERE A.FECHA >= '" & dF1 & "'"
sqltxt = sqltxt & " AND A.FECHA <= '" & dF2 & "'" & cEldepto
sqltxt = sqltxt & " AND MID(A.DESCRIP,LEN(TRIM(A.DESCRIP)),1) <> '%' "
sqltxt = sqltxt & " AND '%' NOT IN (A.DESCRIP) "
sqltxt = sqltxt & " AND A.PLU = B.ID_PLU AND B.ID_UNID_CONSUMO = C.ID"
sqltxt = sqltxt & " AND MID(A.DESCRIP,LEN(TRIM(A.DESCRIP)),1) <> '%'  AND '%' NOT IN (A.DESCRIP)"
sqltxt = sqltxt & " GROUP BY B.ID_PROD_INV"
sqltxt = sqltxt & " HAVING sum(A.CANT * B.CANT) > 0 "
sqltxt = sqltxt & " ORDER BY " & cOrdSel

msConn.BeginTrans
msConn.Execute sqltxt
msConn.CommitTrans

ProgBar.value = 40

sqltxt = "SELECT ID_PROD_INV, DEPTO, PROD_INVENT, "
sqltxt = sqltxt & " FORMAT(VENTAS,'STANDARD') AS REF_VENTAS, MEDIDA, "
sqltxt = sqltxt & " FORMAT(C.EXIST2,'STANDARD') AS EN_STOCK, FORMAT(C.MERMA/100,'PERCENT') as [(%) MERMA],"
sqltxt = sqltxt & " FORMAT(IIF(C.EXIST2 < 0, C.EXIST2, C.EXIST2 - ABS(C.EXIST2 * CSNG(C.MERMA/100))),'STANDARD') AS STOCK_REAL  "

'sqltxt = sqltxt & " FORMAT(C.EXIST2 - ABS(C.EXIST2 * CSNG(C.MERMA/100)),'STANDARD') AS STOCK_REAL"
'LMerma = LExist2 - Abs((LExist2 * CSng(Text1(10).Text) / 100))
sqltxt = sqltxt & " FROM "
sqltxt = sqltxt & " LOLO6 As A, DEP_INV  As B, INVENT As C"
sqltxt = sqltxt & " WHERE A.DEPTO = B.CODIGO And A.ID_PROD_INV = C.ID"
sqltxt = sqltxt & " ORDER BY 3"

rsConsulta01.Open sqltxt, msConn, adOpenStatic, adLockOptimistic
ProgBar.value = 60

If rsConsulta01.EOF Then
''''    Set MSHFDepto.DataSource = Nothing
''''    MSHFDepto.Clear
''''    MSHFDepto.Refresh
    'DD_PEDDETALLE.Rows.RemoveAll True
    Dim arreglo(0, 0)
    DD_PEDDETALLE.LoadArray arreglo
    rsConsulta01.Close
    ProgBar.value = 0
    Exit Sub
Else
''''    Set MSHFDepto.DataSource = rsConsulta01
End If

With DD_PEDDETALLE
'   .Columns.RemoveAll True
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
    .Columns(1).Width = 0: .Columns(2).Width = 0: .Columns(3).Width = 2500:
    .Columns(4).Width = 1200: .Columns(5).Width = 800: .Columns(6).Width = 1100:
    .Columns(7).Width = 1050: .Columns(8).Width = 1100:
    .Columns(1).Style.TextAlignment = sgAlignRightCenter
    .Columns(4).Style.TextAlignment = sgAlignRightCenter
    .Columns(6).Style.TextAlignment = sgAlignRightCenter
    .Columns(7).Style.TextAlignment = sgAlignRightCenter
    .Columns(8).Style.TextAlignment = sgAlignRightCenter
    '.Columns(8).Style.TextAlignment = sgAlignRightCenter
End With

ProgBar.value = 70
rsConsulta01.Close
ProgBar.value = 80


ProgBar.value = 90
msConn.BeginTrans
msConn.Execute "DROP TABLE LOLO6"
msConn.CommitTrans

ProgBar.value = 100
ProgBar.value = 0
On Error GoTo 0
Exit Sub

ErrAdm:
If Err.Number = -2147217900 Or Err.Number = -2147213302 Then
    'info: EL GRID NUEVO ESTA BOUND. ASI QUE NO SE PUEDE SOLTAR AUN LA TABLA.
    'O LA LLAVE DE LA COLUMNA SE ESTA REPITIENDO
Else
    MsgBox Err.Description, vbCritical, "ERROR EN SELECCION DE DATOS"
End If
Resume Next
End Sub

Private Sub cmdSalir_Click()
Unload Me
End Sub

Private Sub Command1_Click()
Dim iCtr As Integer 'Contador de Linea
Dim iCol, iFil As Integer 'Contador de Columnas
Dim cText As String
Dim ispace As Integer
Dim iLen As Integer
Dim sSubTot As Single

sSubTot = 0#: iLin = 8: nPagina = 0

On Error GoTo ErrorPrn:

If MSChart1.Visible = True Then
    nPagina = 0
    '///////////Seleccion_Impresora_Default
    PrintTit
    'ImprimeChart
    '///////////Seleccion_Impresora
    Exit Sub
End If

MainMant.spDoc.DocBegin
PrintTit
EscribeLog ("Admin." & "Impresion de Listado: " & Me.Caption & " PERIODO : Desde " & txtFecIni & " Hasta " & txtFecFin)

ProgBar.value = 10
For iFil = 0 To DD_PEDDETALLE.RowCount - 1
'For iFil = 1 To DD_PEDDETALLE.RowCount - 1
    If iLin > 2400 Then
        MainMant.spDoc.TextAlign = SPTA_LEFT
        PrintTit
    End If
    If ProgBar.value < 100 Then
        ProgBar.value = ProgBar.value + 5
    End If
    DD_PEDDETALLE.Row = iFil
    For iCol = 1 To DD_PEDDETALLE.ColCount - 1
    'For iCol = 0 To DD_PEDDETALLE.Cols
        Select Case iCol
            Case 2, 3, 4, 5, 6, 7
                DD_PEDDETALLE.Col = iCol
                'INFO: COLUMNAS EMPIEZAN EN CERO
                'SELECT ID_PROD_INV, DEPTO, PROD_INVENT, VENTAS, MEDIDA, C.EXIST2 AS EN_STOCK"
''                If iFil = 1 Then
''                Else
                    Select Case iCol
                        Case 2  'DESCRIPCION
                        'cText = cText & Format(Mid(DD_PEDDETALLE.Text, 1, 20), "!@@@@@@@@@@@@@@@@@@@@")
                            MainMant.spDoc.TextAlign = SPTA_LEFT
                            MainMant.spDoc.TextOut 300, iLin, DD_PEDDETALLE.Text
                        Case 3 'CANTIDAD
                        'cText = cText & Space(5) & Format(Format(DD_PEDDETALLE.Text, "#######.00"), "@@@@@@@@@@")
                            MainMant.spDoc.TextAlign = SPTA_RIGHT
                            MainMant.spDoc.TextOut 1000, iLin, Format(DD_PEDDETALLE.Text, "##,##0.00")
                        Case 4  'UNIDAD DE CONSUMO
                        'cText = cText & Space(1) & Format(Mid(DD_PEDDETALLE.Text, 1, 8), "!@@@@@@@@")
                            MainMant.spDoc.TextAlign = SPTA_LEFT
                            MainMant.spDoc.TextOut 1050, iLin, DD_PEDDETALLE.Text
                        Case 5  'EXIST2 EN INVENTARIO
                        'cText = cText & Space(5) & Format(Format(DD_PEDDETALLE.Text, "#######.00"), "@@@@@@@@@@")
                            MainMant.spDoc.TextAlign = SPTA_RIGHT
                            MainMant.spDoc.TextOut 1450, iLin, Format(DD_PEDDETALLE.Text, "STANDARD")
                            'sSubTot = sSubTot + Format(DD_PEDDETALLE.Text, "standard")
                        Case 6      '%MERMA
                            MainMant.spDoc.TextAlign = SPTA_RIGHT
                            MainMant.spDoc.TextOut 1650, iLin, Format(DD_PEDDETALLE.Text, "PERCENT")
                        Case 7
                            MainMant.spDoc.TextAlign = SPTA_RIGHT
                            MainMant.spDoc.TextOut 1910, iLin, Format(DD_PEDDETALLE.Text, "STANDARD")
                    End Select
''                End If
            End Select
        Next
        iLin = iLin + 50
Next
ProgBar.value = 95
iLin = iLin + 100
MainMant.spDoc.TextAlign = SPTA_LEFT
'MainMant.spDoc.TextOut 500, iLin, "Sub Total del Periodo = " & Format(sSubTot, "Currency")
MainMant.spDoc.DoPrintPreview

100:
ProgBar.value = 0
On Error GoTo 0

Call Seguridad

Exit Sub

ErrorPrn:
    'MsgBox "¡ Ocurre algún Error con la Impresora, Intente Conecterla !", vbExclamation, BoxTit
    ShowMsg Err.Description & " ¡ Ocurre algún Error con la Impresora, Intente Conecterla !", vbRed, vbYellow
    Resume
End Sub

Private Sub Form_Load()
Set rsConsulta01 = New Recordset
txtFecIni = Format(Date, "SHORT DATE")
txtFecFin = Format(Date, "SHORT DATE")
nPluSel = 0
cOrdSel = " 3 ASC " 'POR PRODUCTO
Set rsConDept = New Recordset
rsConDept.Open "SELECT CODIGO,DESCRIP FROM DEP_INV " & _
            " ORDER BY DESCRIP ", msConn, adOpenStatic, adLockOptimistic
List1.AddItem "TODOS LOS DEPARTAMENTOS"
Do Until rsConDept.EOF
    List1.AddItem rsConDept!DESCRIP & Space(70) & rsConDept!CODIGO
    rsConDept.MoveNext
Loop
nconDepto = 0
rsConDept.Close
nPagina = 1

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
        txtFecIni.Enabled = False: txtFecFin.Enabled = False: cmdEjec.Enabled = False
        List1.Enabled = False: DD_PEDDETALLE.Enabled = True
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

Private Sub Option1_Click()
cOrdSel = " 3 DESC " 'POR DESCRIPCION
End Sub

Private Sub Option2_Click()
cOrdSel = " 4 ASC " 'POR VENTAS
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
MainMant.spDoc.TextOut 300, 500, "PERIODO : Desde " & txtFecIni & " Hasta " & txtFecFin
MainMant.spDoc.TextOut 300, 550, "Departamento Seleccionado: " & Left(List1.Text, 50)

If Not MSChart1.Visible Then
    MainMant.spDoc.TextOut 300, 650, "DESCRIPCION"
    'MainMant.spDoc.TextOut 850, 650, "Cantidad " ==> INFO: TEXTO ERRONEO EN EL TITULO.(13/JUL/2009)
    MainMant.spDoc.TextOut 800, 650, "REF.VENTA"
    MainMant.spDoc.TextOut 1100, 650, "MED."
    MainMant.spDoc.TextOut 1320, 650, "EXIST"
    MainMant.spDoc.TextOut 1500, 650, "MERMA [%]"
    MainMant.spDoc.TextOut 1750, 650, "EXIST REAL"
    MainMant.spDoc.TextOut 300, 700, String(145, "-")
End If

iLin = 750
nPagina = nPagina + 1
End Sub

