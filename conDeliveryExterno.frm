VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{43135020-B751-4DDD-AB4A-B6D8A342216E}#1.0#0"; "sg20o.ocx"
Begin VB.Form conDeliveryExterno 
   BackColor       =   &H00B39665&
   Caption         =   "VENTAS DE DELIVERY EXTERNO"
   ClientHeight    =   7680
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   5370
   Icon            =   "conDeliveryExterno.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   7680
   ScaleWidth      =   5370
   StartUpPosition =   2  'CenterScreen
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
      Height          =   615
      Left            =   3480
      TabIndex        =   7
      Top             =   120
      Width           =   1335
   End
   Begin VB.CommandButton Command3 
      BackColor       =   &H0000FFFF&
      Caption         =   "IMPRIMIR"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   3480
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   1080
      Width           =   1335
   End
   Begin DDSharpGridOLEDB2.SGGrid DD_DOMICILIO 
      Height          =   5715
      Left            =   120
      TabIndex        =   1
      ToolTipText     =   "Incluye Facturas donde se marco una entrega a Domicilio"
      Top             =   1800
      Width           =   4695
      _cx             =   8281
      _cy             =   10081
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
         Name            =   "Verdana"
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
      StylesCollection=   $"conDeliveryExterno.frx":08CA
      ColumnsCollection=   $"conDeliveryExterno.frx":26FA
      ValueItems      =   $"conDeliveryExterno.frx":3070
   End
   Begin MSComCtl2.DTPicker txtFecIni 
      Height          =   345
      Left            =   1680
      TabIndex        =   3
      Top             =   120
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   609
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   156631041
      CurrentDate     =   36431
   End
   Begin MSComCtl2.DTPicker txtFecFin 
      Height          =   345
      Left            =   1680
      TabIndex        =   4
      Top             =   600
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   609
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   156631041
      CurrentDate     =   36430
   End
   Begin VB.Label Label1 
      BackColor       =   &H00B39665&
      Caption         =   "Fecha Final"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
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
      TabIndex        =   6
      Top             =   720
      Width           =   1335
   End
   Begin VB.Label Label1 
      BackColor       =   &H00B39665&
      Caption         =   "Fecha Inicial"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
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
      TabIndex        =   5
      Top             =   240
      Width           =   1335
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00B39665&
      Caption         =   "VENTAS DE DELIVERY (Incluye Impuesto)"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   495
      Index           =   2
      Left            =   120
      TabIndex        =   2
      Top             =   1200
      Width           =   2535
   End
End
Attribute VB_Name = "conDeliveryExterno"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private nPagina As Integer
Private iLin As Integer
Private cEmpresa As String
Private nDomiPLU As Integer

Private Sub cmdEjec_Click()
Dim dF1 As String, dF2 As String

dF1 = Format(txtFecIni, "YYYYMMDD")
dF2 = Format(txtFecFin, "YYYYMMDD")

Call InfoDomicilio(dF1, dF2)

End Sub

Private Sub Command3_Click()
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

'For i = 1 To DD_DOMICILIO.Columns.Count - 1
'    DD_DOMICILIO.Columns(i).Width = 800
'Next
'Exit Sub

On Error GoTo ErrorPrn:
nPagina = 0

MainMant.spDoc.DocBegin
MainMant.spDoc.TextAlign = SPTA_LEFT
PrintTit    'Rutina de Titulos

'-ProgBar.Value = 10

'For iFil = 0 To MSHFDepto.Rows - 1
For iFil = 0 To DD_DOMICILIO.RowCount - 1

    MainMant.spDoc.TextAlign = SPTA_LEFT
    If iLin > 2400 Then
        MainMant.spDoc.TextAlign = SPTA_LEFT
        PrintTit
    End If
    DD_DOMICILIO.Row = iFil
    '''For iCol = 0 To MSHFDepto.Cols - 1
    'For iCol = 0 To DD_DOMICILIO.ColCount - 1
    For iCol = 0 To DD_DOMICILIO.ColCount - 1
        Select Case iCol
           Case 0, 1, 2
                    DD_DOMICILIO.Col = iCol
                    MainMant.spDoc.TextAlign = SPTA_LEFT
                    If IsNumeric(DD_DOMICILIO.Text) Then ispace = 10 Else ispace = 25
                    Select Case iCol
                    Case 0
                        MainMant.spDoc.TextAlign = SPTA_LEFT
                        MainMant.spDoc.TextOut 300, iLin, DD_DOMICILIO.Text
                        '''MSHFDepto.Text
                    Case 1
                        MainMant.spDoc.TextAlign = SPTA_RIGHT
                        MainMant.spDoc.TextOut 850, iLin, Format(DD_DOMICILIO.Text, "STANDARD")
                    Case 2
                        MainMant.spDoc.TextAlign = SPTA_LEFT
                        MainMant.spDoc.TextOut 950, iLin, DD_DOMICILIO.Text
                    End Select
                        
'                End If
            End Select
    Next
    iLin = iLin + 50
    '-If ProgBar.Value < 100 Then '-ProgBar.Value = '-ProgBar.Value + 5
Next

'-ProgBar.Value = 100

MainMant.spDoc.DoPrintPreview
'spDoc.DoPrintPreview
On Error GoTo 0

Exit Sub

100:
Exit Sub
ErrorPrn:
    ShowMsg "¡ Ocurre algún Error con la Impresora, Intente Conectarla !", , , vbOKOnly
    Resume
End Sub

Private Sub Form_Load()
txtFecIni = Format(Date, "SHORT DATE")
txtFecFin = Format(Date, "SHORT DATE")

nDomiPLU = Val(GetFromINI("Facturacion", "DomicilioPLU", App.Path & "\soloini.ini"))

End Sub

Private Function InfoDomicilio(cFINI As String, cFFIN As String)
Dim cSQL As String, cSQL2 As String
Dim rsTMP As ADODB.Recordset
Dim rsTMP2 As ADODB.Recordset
Dim rsDOMI As ADODB.Recordset
Dim iFila As Long
Dim nNUM() As Long
Dim cFecha As String
Dim nVentas() As String
Dim i As Integer
Dim j As Integer
Dim nTotalVentas As Single
Dim nRegistros As Integer
Dim nPedidos As Integer
Dim cNumTrans  As String
Dim nFreefile As Integer        'PARA LEER LOS VALORES DE LA STAR
Dim cFile As String, cReturn As String

   On Error GoTo InfoDomicilio_Error

        On Error Resume Next
        cFile = App.Path & "\LOG.txt"
        Kill cFile
        On Error GoTo 0

nFreefile = FreeFile()
Open cFile For Output As #nFreefile

Me.MousePointer = vbHourglass

Set rsDOMI = New ADODB.Recordset
Set rsTMP = New ADODB.Recordset
Set rsTMP2 = New ADODB.Recordset

cSQL = "SELECT DISTINCT A.NUM_TRANS, A.FECHA FROM HIST_TR AS A "
cSQL = cSQL & " WHERE A.PLU = " & nDomiPLU
cSQL = cSQL & " AND A.FECHA BETWEEN '" & cFINI & "'  AND '" & cFFIN & "'"
cSQL = cSQL & " AND A.VALID AND A.PRECIO > 0"
cSQL = cSQL & " ORDER BY A.FECHA"

rsTMP.Open cSQL, msConn, adOpenStatic, adLockOptimistic

'ReDim nNUM(rsTMP.RecordCount)
'ReDim cFecha(rsTMP.RecordCount)
j = txtFecFin - txtFecIni
ReDim nVentas(2, j) ' FECHA, VENTAS, PEDIDOS

If Not rsTMP.EOF Then cFecha = rsTMP!FECHA Else cFecha = ""

Do While Not rsTMP.EOF
    Do While cFecha = rsTMP!FECHA
        cNumTrans = cNumTrans & rsTMP!NUM_TRANS & ","
        rsTMP.MoveNext
        If rsTMP.EOF Then
            Exit Do
        End If
    Loop
    cNumTrans = Left(cNumTrans, Len(cNumTrans) - 1)
    Print #nFreefile, Now & Space(1) & cNumTrans
    'Debug.Print cNumTrans
    nPedidos = GetRegistros(cNumTrans)
    'cSQL2 = "SELECT SUM(PRECIO) AS VENTAS FROM HIST_TR WHERE NUM_TRANS IN (" & cNumTrans & ") "
    cSQL2 = "SELECT SUM(SUB_TOTAL + ITBM) AS VENTAS FROM TRANSAC_FISCAL WHERE DOC_SOLO IN (" & cNumTrans & ") "
    rsTMP2.Open cSQL2, msConn, adOpenStatic, adLockOptimistic
    
    If Not rsTMP2.EOF Then
        nVentas(0, i) = Format(cFecha, "####-##-##")
        nVentas(1, i) = Str(rsTMP2!VENTAS)
        nVentas(2, i) = "(" & nPedidos & ") Pedidos"
        nPedidos = 0
        nRegistros = nRegistros + 1
        nTotalVentas = nTotalVentas + rsTMP2!VENTAS
    End If
    i = i + 1
    
    rsTMP2.Close
    cNumTrans = ""
    If rsTMP.EOF Then Exit Do
    cFecha = rsTMP!FECHA
Loop

ReDim Preserve nVentas(2, nRegistros)

DD_DOMICILIO.LoadArray nVentas

With DD_DOMICILIO
    .Columns(1).Caption = "FECHA"
    .Columns(2).Caption = "VENTAS"
    .Columns(3).Caption = "PEDIDOS"
    .ColumnClickSort = True
    .EvenOddStyle = sgEvenOddRows
    .ColorEven = vbWhite
    .ColorOdd = &HE0E0E0
    .Columns(1).Width = 1600
    .Columns(2).Width = 1100
    .Columns(3).Width = 1500
    .Columns(1).Style.TextAlignment = sgAlignLeftCenter
    .Columns(2).Style.TextAlignment = sgAlignRightCenter
    .Columns(2).Style.Format = "Standard"
    .Columns(2).SortType = sgSortTypeNumber
    .Columns(3).Style.TextAlignment = sgAlignLeftCenter
End With

cData = "TOTAL PERIODO"
cData = cData & ";" & Format(nTotalVentas, "STANDARD")
DD_DOMICILIO.Rows.Add sgFormatCharSeparatedValue, cData, ";"

Me.MousePointer = vbNormal
Close #nFreefile

   On Error GoTo 0
   Exit Function

InfoDomicilio_Error:
    If Err.Number = 9 Then
        ShowMsg "FECHAS INVALIDAS", vbYellow, vbRed
    Else
        ShowMsg "Error " & Err.Number & " (" & Err.Description & ") in procedure InfoDomicilio of Form1"
    End If
    Close #nFreefile
End Function

Private Function GetRegistros(cCadena As String) As Integer
    Dim asciiToSearchFor As Integer
    Dim Count As Integer
    asciiToSearchFor = Asc(",")
    For i = 1 To Len(cCadena)
        If Asc(Mid$(cCadena, i, 1)) = asciiToSearchFor Then Count = Count + 1
    Next
    GetRegistros = Count + 1
End Function


Private Sub PrintTit()

If nPagina = 0 Then
    MainMant.spDoc.WindowTitle = "Impresión de Ventas de Delivery Externo"
    MainMant.spDoc.FirstPage = 1
    MainMant.spDoc.PageOrientation = SPOR_PORTRAIT
    MainMant.spDoc.Units = SPUN_LOMETRIC
End If

MainMant.spDoc.Page = nPagina + 1

MainMant.spDoc.TextOut 300, 200, Format(Date, "long date") & "  " & Time
MainMant.spDoc.TextOut 300, 250, "Página : " & nPagina + 1
MainMant.spDoc.TextOut 300, 350, "VENTAS DE DELIVERY (" & cEmpresa & ")"
MainMant.spDoc.TextOut 300, 400, "PERIODO : Desde " & txtFecIni & " Hasta " & txtFecFin
MainMant.spDoc.TextOut 300, 450, "-----------------------------------------------------------------------------------------------------------------------------------"
iLin = 500
nPagina = nPagina + 1
End Sub

