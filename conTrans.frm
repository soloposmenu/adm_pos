VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{43135020-B751-4DDD-AB4A-B6D8A342216E}#1.0#0"; "sg20o.ocx"
Begin VB.Form conTrans 
   BackColor       =   &H00B39665&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "CONSULTA DE TRANSACCIONES"
   ClientHeight    =   8250
   ClientLeft      =   45
   ClientTop       =   345
   ClientWidth     =   12060
   Icon            =   "conTrans.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8250
   ScaleWidth      =   12060
   Begin VB.CheckBox Top10 
      BackColor       =   &H00B39665&
      Caption         =   "Incluir Detalle de Productos"
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
      Left            =   6600
      TabIndex        =   19
      Top             =   7800
      Width           =   2775
   End
   Begin DDSharpGridOLEDB2.SGGrid DD_PEDDETALLE 
      Height          =   915
      Left            =   10680
      TabIndex        =   17
      Top             =   120
      Visible         =   0   'False
      Width           =   1095
      _cx             =   1931
      _cy             =   1614
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
         Size            =   8.25
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
      StylesCollection=   $"conTrans.frx":0442
      ColumnsCollection=   $"conTrans.frx":2271
      ValueItems      =   $"conTrans.frx":2BE7
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
      TabIndex        =   16
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
      Left            =   3000
      TabIndex        =   15
      ToolTipText     =   "Obtener Ventas por Fechas Seleccionadas"
      Top             =   240
      Value           =   -1  'True
      Width           =   1215
   End
   Begin VB.CheckBox chkPlatos 
      BackColor       =   &H00B39665&
      Caption         =   "Mostrar Platos de la Factura"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   255
      Left            =   5880
      TabIndex        =   11
      Top             =   3600
      Width           =   2775
   End
   Begin VB.CommandButton Command1 
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
      Height          =   615
      Left            =   10320
      TabIndex        =   7
      Top             =   7560
      Width           =   1695
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFTrans 
      Height          =   6495
      Left            =   60
      TabIndex        =   5
      Top             =   1800
      Width           =   5775
      _ExtentX        =   10186
      _ExtentY        =   11456
      _Version        =   393216
      FixedCols       =   0
      SelectionMode   =   1
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
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
      Height          =   615
      Left            =   8520
      TabIndex        =   2
      Top             =   120
      Width           =   1815
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
      Format          =   115277825
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
      Format          =   115277825
      CurrentDate     =   36418
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFPagos 
      Height          =   2055
      Left            =   5880
      TabIndex        =   6
      Top             =   1500
      Width           =   5055
      _ExtentX        =   8916
      _ExtentY        =   3625
      _Version        =   393216
      FixedCols       =   0
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFPlatos 
      Height          =   3615
      Left            =   5880
      TabIndex        =   10
      Top             =   3840
      Visible         =   0   'False
      Width           =   6135
      _ExtentX        =   10821
      _ExtentY        =   6376
      _Version        =   393216
      FixedCols       =   0
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
   End
   Begin MSComctlLib.ProgressBar ProgBar 
      Height          =   180
      Left            =   60
      TabIndex        =   12
      Top             =   1320
      Width           =   5175
      _ExtentX        =   9128
      _ExtentY        =   318
      _Version        =   393216
      Appearance      =   1
   End
   Begin MSComctlLib.ListView LVZ 
      Height          =   825
      Left            =   5775
      TabIndex        =   13
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
   Begin DDSharpGridOLEDB2.SGGrid DD_PEDDETALLE2 
      Height          =   915
      Left            =   10680
      TabIndex        =   18
      Top             =   1800
      Visible         =   0   'False
      Width           =   1095
      _cx             =   1931
      _cy             =   1614
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
         Size            =   8.25
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
      StylesCollection=   $"conTrans.frx":2C87
      ColumnsCollection=   $"conTrans.frx":4AB6
      ValueItems      =   $"conTrans.frx":542C
   End
   Begin VB.Image Image 
      Height          =   480
      Left            =   5880
      Picture         =   "conTrans.frx":54CC
      ToolTipText     =   "Exportar Datos"
      Top             =   7680
      Width           =   480
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
      Index           =   4
      Left            =   4440
      TabIndex        =   14
      Top             =   690
      Width           =   1215
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
      Width           =   3975
   End
   Begin VB.Label Label1 
      BackColor       =   &H00B39665&
      Caption         =   "Transacciones"
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
      Left            =   60
      TabIndex        =   9
      Top             =   1560
      Width           =   1335
   End
   Begin VB.Label Label1 
      BackColor       =   &H00B39665&
      Caption         =   "Detalle de Pago la Factura"
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
      Left            =   5880
      TabIndex        =   8
      Top             =   1275
      Width           =   2655
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
      Left            =   240
      TabIndex        =   4
      Top             =   840
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
      Left            =   240
      TabIndex        =   3
      Top             =   360
      Width           =   1215
   End
End
Attribute VB_Name = "conTrans"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private rsConTrans As Recordset
Private rsConPagos As Recordset
Private nNumTran As Long

Private Sub chkPlatos_Click()
If chkPlatos.value = 1 Then
    MSHFPlatos.Visible = True
    Dim rsConPlatos As Recordset
    Set rsConPlatos = New Recordset
'    txtsql = "SELECT A.DESCRIP AS PLATO, A.CANT, " & _
            " FORMAT(A.PRECIO,'STANDARD') AS PRECIO, " & _
            " A.Hora, (b.nombre + ',' + b.apellido) as Cajero " & _
            " FROM HIST_TR AS A LEFT JOIN CAJEROS AS b ON A.CAJERO = b.NUMERO " & _
            " WHERE NUM_TRANS = " & nNumTran
            
    'INFO: 27MAR2010
    txtsql = "SELECT A.DESCRIP AS Plato, A.CANT as Cant, "
    txtsql = txtsql & " FORMAT(A.PRECIO,'STANDARD') AS Precio, "
    txtsql = txtsql & " MID(A.FECHA_TRANS,7,2) & '-' &   MID(A.FECHA_TRANS,5,2) AS Fecha, A.HORA_TRANS AS Hora, "
    txtsql = txtsql & " (b.nombre + ',' +b.apellido) as Cajero "
    'txtsql = txtsql & " FORMAT(A.FECHA_TRANS,'####-##-##') AS Fecha, A.HORA_TRANS AS Hora, (b.nombre + ',' +b.apellido) as Cajero "
    'txtsql = txtsql & " A.Hora, (b.nombre + ',' +b.apellido) as Cajero "
    txtsql = txtsql & " FROM HIST_TR AS A LEFT JOIN CAJEROS AS b ON A.CAJERO = b.NUMERO "
    txtsql = txtsql & " WHERE NUM_TRANS = " & nNumTran
    txtsql = txtsql & " ORDER BY A.LIN"
    
    rsConPlatos.Open txtsql, msConn, adOpenStatic, adLockOptimistic
    Set MSHFPlatos.DataSource = rsConPlatos
    rsConPlatos.Close
    With MSHFPlatos
'        .ColWidth(0) = 2200: .ColWidth(1) = 600: .ColWidth(2) = 1000:
'        .ColAlignment(2) = flexAlignRightCenter
'        .ColWidth(3) = 900: .ColWidth(4) = 1500
        
        .ColWidth(0) = 2200: .ColWidth(1) = 500: .ColWidth(2) = 800:
        .ColAlignment(2) = flexAlignRightCenter
        .ColWidth(3) = 600: .ColWidth(4) = 1000: .ColWidth(5) = 1500
        
    End With
Else
    MSHFPlatos.Visible = False
End If
End Sub

Private Sub cmdEjec_Click()
Dim dF1 As String
Dim dF2 As String

'INFO: 12FEB2011
Dim iZZ As Integer  'Z Counter
Dim jFalseCounter As Byte
Dim cSQL As String, cZetas As String

Dim nMinZ As Long   'REPORTE Z INICIAL
Dim nMaxZ As Long   'REPORTE Z FINAL
Dim cArrayZ() As String

'INFO: 17OCT2017
Dim rsClone As ADODB.Recordset


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
    'INFO: FEB2011
    'CHECK TO SEE IF A Z HAS BEEN SELECTED
    For iZZ = 1 To LVZ.ListItems.Count
        If LVZ.ListItems(iZZ).Checked = True Then
            cZetas = cZetas & LVZ.ListItems(iZZ).Text & "','"
        End If
    Next
    On Error Resume Next
    If cZetas = "" Then
        ShowMsg "Debe seleccionar al menos un Reporte Z"
        Exit Sub
    Else
        cZetas = "'" & Mid(cZetas, 1, Len(cZetas) - 2)
    End If

    'INFO: 27OCT2011
    cmdEjec.Tag = cZetas
    cArrayZ = Split(Replace(cZetas, "'", ""), ",")
    
    On Error GoTo 0
    On Error GoTo ErrAdm:
End If

Me.MousePointer = vbHourglass
dF1 = Format(txtFecIni, "YYYYMMDD")
dF2 = Format(txtFecFin, "YYYYMMDD")
ProgBar.value = 50

'~~~~||||'~~~~||||'~~~~||||'~~~~||||'~~~~||||'~~~~||||'~~
'INFO: INCLUYE TRANSACCIONES FISCALES
'21NOV2011
'~~~~||||'~~~~||||'~~~~||||'~~~~||||'~~~~||||'~~~~||||'~~
'''INFO: 2DIC2012
''sqltxt = sqltxt & " MAX(B.HORA) AS Hora, "
''sqltxt = sqltxt & " MAX(A.MESA) AS Mesa, "
''sqltxt = sqltxt & " FORMAT(SUM(A.PRECIO),'STANDARD') AS Ventas "
'''~~~~||||'~~~~||||'~~~~||||'~~~~||||'~~~~||||'~~~~||||'~~
'''sqltxt = sqltxt & " From HIST_TR AS A "
''sqltxt = sqltxt & " FROM HIST_TR AS A LEFT JOIN TRANSAC_FISCAL AS B "
''sqltxt = sqltxt & " ON A.NUM_TRANS = B.DOC_SOLO "


'INFO: UPDATE 15MAR2013
sqltxt = "SELECT A.NUM_TRANS AS Trans,  A.FECHA as Fecha_Trans, "
sqltxt = sqltxt & " MAX(IIF(ISNULL(B.FISCAL),0,FORMAT(B.FISCAL,'GENERAL NUMBER'))) AS Fiscal,"
'~~~~||||'~~~~||||'~~~~||||'~~~~||||'~~~~||||'~~~~||||'~~
sqltxt = sqltxt & " MAX((MID(A.FECHA,7,2) + '/' + MID(A.FECHA,5,2) + '/' + MID(A.FECHA,1,4))) AS FECHA,  "
sqltxt = sqltxt & " MAX(iif(ISNULL(B.HORA),FORMAT(A.HORA,'HH:MM'),B.HORA)) AS Hora,"
sqltxt = sqltxt & " MAX(A.MESA) AS Msa, "
sqltxt = sqltxt & " MAX(A.MESERO) AS Msro,"
sqltxt = sqltxt & " FORMAT(SUM(A.PRECIO),'STANDARD') AS Ventas, "
'ÍNFO: 4AGO2021
sqltxt = sqltxt & " MAX(A.CAJA) AS Caja "

'sqltxt = sqltxt & " , SPACE (50) AS PAGADO, SPACE (40) AS PRODUCTO "

'~~~~||||'~~~~||||'~~~~||||'~~~~||||'~~~~||||'~~~~||||'~~
'sqltxt = sqltxt & " From HIST_TR AS A "
'|||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||
'|||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||
sqltxt = sqltxt & " FROM HIST_TR AS A LEFT JOIN TRANSAC_FISCAL AS B "
'|||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||
'|||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||
'sqltxt = sqltxt & " FROM TRANSAC AS A LEFT JOIN TRANSAC_FISCAL AS B "
sqltxt = sqltxt & " ON A.NUM_TRANS = B.DOC_SOLO "

If i = 0 Then
    sqltxt = sqltxt & " WHERE A.FECHA BETWEEN '" & dF1 & "'"
    sqltxt = sqltxt & " AND '" & dF2 & "'"
Else
    'cSQL = cSQL & " WHERE Z_COUNTER = '" & Val(LVZ.SelectedItem.Text) & "'"
    'INFO: 16ABR2008. MULTIPLES Z
    '''cSQL = cSQL & " WHERE Z_COUNTER IN (" & cZetas & ")"
    'INFO: VALIDAR LA Z CONTRA VALORES NUMERICOS EN VEZ DE TEXTO (27OCT2010)
    sqltxt = sqltxt & " WHERE VAL(A.Z_COUNTER) BETWEEN  " & GetNumber_FromArray(cArrayZ, GetMin)
    sqltxt = sqltxt & " AND " & GetNumber_FromArray(cArrayZ, Getmax)
End If
'INFO: HACERLO x Z_COUNTER y POR FECHA
'sqltxt = sqltxt & " WHERE A.FECHA >= '" & dF1 & "'"
'sqltxt = sqltxt & " AND A.FECHA <= '" & dF2 & "'"
'sqltxt = sqltxt & " AND B.FISCAL IS NOT NULL "
sqltxt = sqltxt & " GROUP BY A.NUM_TRANS, A.FECHA "
'sqltxt = sqltxt & " ORDER BY A.FECHA, A.NUM_TRANS "
sqltxt = sqltxt & " ORDER BY A.NUM_TRANS, A.FECHA "

rsConTrans.Open sqltxt, msConn, adOpenStatic, adLockOptimistic

'INFO: MEJORA YA QUE EL PROGRAMA REVENTABA
' 19MAY2014
If rsConTrans.EOF Then
    rsConTrans.Close
    ShowMsg "Ejecutar Consulta. NO hay datos disponibles para el periodo seleccionado", vbRed, vbYellow
    Me.MousePointer = vbDefault
    ProgBar.value = 0
    Exit Sub
End If

'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
'INFO: 17OCT2017
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
'~~ Set rsClone = New ADODB.Recordset
'rsClone.CursorLocation = adUseClient
'rsClone.CursorType = adOpenKeyset
'rsClone.CursorType = adOpenDynamic
'rsClone.LockType = adLockOptimistic
'rsClone.ActiveConnection = Nothing
'rsClone.Open sqltxt, msConn, adOpenDynamic
'rsClone.Fields.Append "PAGADO", adChar, 50
'rsClone.Fields.Append "PRODUCTO", adChar, 40

'On Error Resume Next
'~~ If Dir(App.Path & "\rsClone.xml") <> "" Then Kill App.Path & "\rsClone.xml"

'~~ rsConTrans.Save App.Path & "\rsClone.xml", adPersistXML
'~~ rsClone.Open App.Path & "\rsClone.xml", "Provider=MSPersist;", adOpenDynamic, adLockBatchOptimistic, adCmdFile

'~~ Do While Not rsClone.EOF
'~~     Call GetTipoPago_MontoPagado(rsClone, rsClone!TRANS)
'~~     rsClone.MoveNext
'~~ Loop
'~~ rsClone.MoveFirst
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~


'''~~~~||||'~~~~||||'~~~~||||'~~~~||||'~~~~||||'~~~~||||'~~
' PASAR AL GRID
'''~~~~||||'~~~~||||'~~~~||||'~~~~||||'~~~~||||'~~~~||||'~~
Dim arreglo(0, 0)
DD_PEDDETALLE.LoadArray arreglo

With DD_PEDDETALLE
    'INFO: 12FEB2011
   .Columns.RemoveAll True
   .DataMode = sgUnbound

   .LoadArray rsConTrans.GetRows()
   ' define each column from the recordsets' fields collection
   For iLoop = 1 To rsConTrans.Fields.Count
      .Columns(iLoop).Caption = rsConTrans.Fields(iLoop - 1).Name
      .Columns(iLoop).DBField = rsConTrans.Fields(iLoop - 1).Name
      .Columns(iLoop).Key = rsConTrans.Fields(iLoop - 1).Name
   Next iLoop
End With

With DD_PEDDETALLE
    .ColumnClickSort = False
    .EvenOddStyle = sgEvenOddRows
    .ColorEven = vbWhite
    .ColorOdd = &HE0E0E0
    .Columns(1).Width = 900
    .Columns(2).Width = 0
    .Columns(3).Width = 700
    .Columns(4).Width = 900
    .Columns(5).Width = 500
    .Columns(6).Width = 450
    .Columns(7).Width = 500
    .Columns(8).Width = 840

    '.Columns(9).Width = 1100
    .Columns(1).Style.TextAlignment = sgAlignRightCenter
    .Columns(3).Style.TextAlignment = sgAlignRightCenter
    .Columns(4).Style.TextAlignment = sgAlignRightCenter
    .Columns(6).Style.TextAlignment = sgAlignRightCenter
    .Columns(7).Style.TextAlignment = sgAlignRightCenter
    .Columns(8).Style.TextAlignment = sgAlignRightCenter
    'ÍNFO: 4AGO2021
    .Columns(9).Style.TextAlignment = sgAlignRightCenter
End With

ProgBar.value = 75
Set MSHFTrans.DataSource = rsConTrans
ProgBar.value = 100
rsConTrans.Close
With MSHFTrans
    '.ColWidth(0) = 1100: .ColWidth(1) = 700: .ColWidth(2) = 0:
    .ColWidth(0) = 900: .ColWidth(1) = 0: .ColWidth(2) = 700:
    .ColWidth(3) = 980:
    'INFO: 2DIC2012 - INCLUYE HORA
    .ColWidth(4) = 500: .ColWidth(5) = 450: .ColWidth(6) = 500:
    .ColWidth(7) = 820:
    'ÍNFO: 4AGO2021
    .ColWidth(8) = 450
    .ColAlignment(2) = flexAlignRightCenter
    .ColAlignment(4) = flexAlignRightCenter
    .ColAlignment(6) = flexAlignRightCenter
    .ColAlignment(7) = flexAlignRightCenter
    'ÍNFO: 4AGO2021
    .ColAlignment(8) = flexAlignRightCenter
End With
MSHFTrans_EnterCell
Me.MousePointer = vbDefault
ProgBar.value = 0

'Call Seguridad

Exit Sub

ErrAdm:
    If Err.Number = 3021 Then
        ShowMsg "Ejecutar Consulta. NO hay datos disponibles para el periodo seleccionado", vbRed, vbYellow
    Else
        ShowMsg "Ejecutar Consulta. " & Err.Number & " - " & Err.Description, vbRed, vbYellow
    End If
    Me.MousePointer = vbDefault
    ProgBar.value = 0
    Debug.Print sqltxt
    If rsConTrans.State = adStateOpen Then
        rsConTrans.Close
        Set rsConTrans = Nothing
    End If
    Resume Next
End Sub

Private Sub Command1_Click()
Set rsConTrans = Nothing
Set rsConPagos = Nothing
Unload Me
End Sub

Private Sub Form_Load()
Set rsConTrans = New Recordset
Set rsConPagos = New Recordset

txtFecIni = Format(Date, "SHORT DATE")
txtFecFin = Format(Date, "SHORT DATE")

Call GetReportesZ

Call Seguridad

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
        'INFO: NO HAY RESTRICCIONES
    Case "N"        'SIN DERECHOS
        txtFecIni.Enabled = False: txtFecFin.Enabled = False: cmdEjec.Enabled = False
        MSHFTrans.Enabled = False
        chkPlatos.Enabled = False
End Select
End Function

Private Sub Image_Click()
If Top10.value = 1 Then
    Call BuildSQLDetalleProductos
    Call ExportToExcelOrCSVList(DD_PEDDETALLE2)
Else
    Call ExportToExcelOrCSVList(DD_PEDDETALLE)
End If
End Sub
'---------------------------------------------------------------------------------------
' Procedure : BuildSQLDetalleProductos
' Author    : hsequeira
' Date      : 18/05/2018
' Purpose   : MUESTRA EL DETALLE DE PRODUCTOS
'---------------------------------------------------------------------------------------
'
Private Function BuildSQLDetalleProductos() As Boolean
Dim dF1 As String
Dim dF2 As String

'INFO: 12FEB2011
Dim iZZ As Integer  'Z Counter
Dim jFalseCounter As Byte
Dim cSQL As String, cZetas As String

Dim nMinZ As Long   'REPORTE Z INICIAL
Dim nMaxZ As Long   'REPORTE Z FINAL
Dim cArrayZ() As String

'INFO: 17OCT2017
Dim rsClone2 As ADODB.Recordset

Set rsClone2 = New ADODB.Recordset

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
    'INFO: FEB2011
    'CHECK TO SEE IF A Z HAS BEEN SELECTED
    For iZZ = 1 To LVZ.ListItems.Count
        If LVZ.ListItems(iZZ).Checked = True Then
            cZetas = cZetas & LVZ.ListItems(iZZ).Text & "','"
        End If
    Next
    On Error Resume Next
    If cZetas = "" Then
        ShowMsg "Debe seleccionar al menos un Reporte Z"
        Exit Function
    Else
        cZetas = "'" & Mid(cZetas, 1, Len(cZetas) - 2)
    End If

    'INFO: 27OCT2011
    cmdEjec.Tag = cZetas
    cArrayZ = Split(Replace(cZetas, "'", ""), ",")
    
    On Error GoTo 0
    On Error GoTo ErrAdm:
End If

Me.MousePointer = vbHourglass
dF1 = Format(txtFecIni, "YYYYMMDD")
dF2 = Format(txtFecFin, "YYYYMMDD")

ProgBar.value = 50

sqltxt = "SELECT A.NUM_TRANS AS Trans, A.FECHA as Fecha_Trans, A.LIN,  "
sqltxt = sqltxt & " A.DESCRIP AS Producto, A.CANT AS Cant, A.PRECIO AS Ventas2, "
sqltxt = sqltxt & " MAX(IIF(ISNULL(B.FISCAL),0,FORMAT(B.FISCAL,'GENERAL NUMBER'))) AS Fiscal,"
sqltxt = sqltxt & " MAX((MID(A.FECHA,7,2) + '/' + MID(A.FECHA,5,2) + '/' + MID(A.FECHA,1,4))) AS FECHA,  "
sqltxt = sqltxt & " MAX(iif(ISNULL(B.HORA),FORMAT(A.HORA,'HH:MM'),B.HORA)) AS Hora,"
sqltxt = sqltxt & " MAX(A.MESA) AS Msa, "
sqltxt = sqltxt & " MAX(A.MESERO) AS Msro,"
'INFO: 18MAY2018. INFO DE PRODUCTOS
'sqltxt = sqltxt & " FORMAT(SUM(A.PRECIO),'STANDARD') AS Ventas "
sqltxt = sqltxt & " FORMAT(MAX(B.SUB_TOTAL),'STANDARD') AS Ventas "
sqltxt = sqltxt & " FROM HIST_TR AS A LEFT JOIN TRANSAC_FISCAL AS B "
sqltxt = sqltxt & " ON A.NUM_TRANS = B.DOC_SOLO "

If i = 0 Then
    sqltxt = sqltxt & " WHERE A.FECHA BETWEEN '" & dF1 & "'"
    sqltxt = sqltxt & " AND '" & dF2 & "'"
Else
    sqltxt = sqltxt & " WHERE VAL(A.Z_COUNTER) BETWEEN  " & GetNumber_FromArray(cArrayZ, GetMin)
    sqltxt = sqltxt & " AND " & GetNumber_FromArray(cArrayZ, Getmax)
End If

sqltxt = sqltxt & " GROUP BY A.NUM_TRANS, A.FECHA, A.LIN, A.DESCRIP, A.CANT, A.PRECIO "
sqltxt = sqltxt & " ORDER BY A.NUM_TRANS, A.FECHA, A.LIN  "

rsClone2.Open sqltxt, msConn, adOpenStatic, adLockOptimistic

'INFO: MEJORA YA QUE EL PROGRAMA REVENTABA
' 19MAY2014
If rsClone2.EOF Then
    rsClone2.Close
    ShowMsg "Ejecutar Consulta. NO hay datos disponibles para el periodo seleccionado", vbRed, vbYellow
    Me.MousePointer = vbDefault
    ProgBar.value = 0
    Exit Function
End If

'''~~~~||||'~~~~||||'~~~~||||'~~~~||||'~~~~||||'~~~~||||'~~
' PASAR AL GRID
'''~~~~||||'~~~~||||'~~~~||||'~~~~||||'~~~~||||'~~~~||||'~~
Dim arreglo(0, 0)
DD_PEDDETALLE2.LoadArray arreglo

With DD_PEDDETALLE2
    'INFO: 12FEB2011
   .Columns.RemoveAll True
   .DataMode = sgUnbound

   .LoadArray rsClone2.GetRows()
   ' define each column from the recordsets' fields collection
   For iLoop = 1 To rsClone2.Fields.Count
      .Columns(iLoop).Caption = rsClone2.Fields(iLoop - 1).Name
      .Columns(iLoop).DBField = rsClone2.Fields(iLoop - 1).Name
      .Columns(iLoop).Key = rsClone2.Fields(iLoop - 1).Name
   Next iLoop
End With

With DD_PEDDETALLE2
    .ColumnClickSort = False
    .EvenOddStyle = sgEvenOddRows
    .ColorEven = vbWhite
    .ColorOdd = &HE0E0E0
    .Columns(1).Width = 600
    .Columns(2).Width = 0
    
    .Columns(3).Width = 450
    .Columns(4).Width = 2000
    .Columns(5).Width = 450
    
   
    .Columns(6).Width = 800
    .Columns(7).Width = 700
    .Columns(8).Width = 900
    .Columns(9).Width = 600
    .Columns(10).Width = 450
    .Columns(11).Width = 600
    .Columns(12).Width = 600
    
    .Columns(1).Style.TextAlignment = sgAlignRightCenter
    .Columns(6).Style.TextAlignment = sgAlignRightCenter
    .Columns(7).Style.TextAlignment = sgAlignRightCenter
    .Columns(8).Style.TextAlignment = sgAlignRightCenter
    .Columns(9).Style.TextAlignment = sgAlignRightCenter
    .Columns(10).Style.TextAlignment = sgAlignRightCenter
    .Columns(11).Style.TextAlignment = sgAlignRightCenter
    .Columns(12).Style.TextAlignment = sgAlignRightCenter
End With

ProgBar.value = 100
rsClone2.Close
Set rsClone2 = Nothing

Me.MousePointer = vbDefault
ProgBar.value = 0

Exit Function

ErrAdm:
    If Err.Number = 3021 Then
        ShowMsg "Ejecutar Consulta. NO hay datos disponibles para el periodo seleccionado", vbRed, vbYellow
    Else
        ShowMsg "Ejecutar Consulta. " & Err.Number & " - " & Err.Description, vbRed, vbYellow
    End If
    Me.MousePointer = vbDefault
    ProgBar.value = 0
    Debug.Print sqltxt
    If rsClone2.State = adStateOpen Then
        rsClone2.Close
        Set rsClone2 = Nothing
    End If
    Resume Next
End Function
Private Sub MSHFTrans_Click()
MSHFTrans_EnterCell
End Sub

Private Sub MSHFTrans_EnterCell()
nNumTran = Val(MSHFTrans.Text)

sqltxt = "SELECT A.NUM_TRANS, A.TIPO_PAGO, B.DESCRIP, "
sqltxt = sqltxt & " format(SUM(A.MONTO),'STANDARD') AS PAGOS "
'|||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||
'|||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||
sqltxt = sqltxt & " FROM HIST_TR_PAGO AS A LEFT JOIN PAGOS AS B "
'sqltxt = sqltxt & " FROM TRANSAC_PAGO AS A LEFT JOIN PAGOS AS B "
'|||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||
'|||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||
sqltxt = sqltxt & " ON A.TIPO_PAGO = B.CODIGO "
sqltxt = sqltxt & " WHERE A.NUM_TRANS = " & nNumTran
sqltxt = sqltxt & " GROUP BY A.NUM_TRANS, A.TIPO_PAGO, B.DESCRIP"

rsConPagos.Open sqltxt, msConn, adOpenStatic, adLockOptimistic

Set MSHFPagos.DataSource = rsConPagos
rsConPagos.Close

'|||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||
'|||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||
rsConPagos.Open "SELECT A.TIPO_PAGO,('Propina ' + B.DESCRIP) as Descrip, " & _
                " format(a.monto,'standard') as Pagos " & _
                " FROM HIST_TR_PROP as A LEFT JOIN PAGOS AS B " & _
                " ON A.TIPO_PAGO = B.CODIGO" & _
                " WHERE A.NUM_TRANS = " & nNumTran & _
                " ORDER BY A.TIPO_PAGO "
                
'rsConPagos.Open "SELECT A.TIPO_PAGO,('Propina ' + B.DESCRIP) as Descrip, " & _
                " format(a.monto,'standard') as Pagos " & _
                " FROM TRANSAC_PROP as A LEFT JOIN PAGOS AS B " & _
                " ON A.TIPO_PAGO = B.CODIGO" & _
                " WHERE A.NUM_TRANS = " & nNumTran & _
                " ORDER BY A.TIPO_PAGO "
'|||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||
'|||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||
Do Until rsConPagos.EOF
    MSHFPagos.AddItem rsConPagos!tipo_pago & Chr(9) & rsConPagos!tipo_pago & Chr(9) & _
            IIf(IsNull(rsConPagos!DESCRIP), "Falta Tipo Pago", rsConPagos!DESCRIP) & Chr(9) & Format(rsConPagos!pagos * (-1#), "standard")
    rsConPagos.MoveNext
Loop
rsConPagos.Close

With MSHFPagos
    .ColWidth(0) = 0: .ColWidth(1) = 0: .ColWidth(2) = 2300:
    .ColWidth(3) = 1300:
    .ColAlignment(3) = flexAlignRightCenter
End With

If chkPlatos.value = 1 Then
    Dim rsConPlatos As Recordset
    Set rsConPlatos = New Recordset
    
    txtsql = "SELECT A.DESCRIP AS Plato, A.CANT as Cant, "
    txtsql = txtsql & " FORMAT(A.PRECIO,'STANDARD') AS Precio, "
    txtsql = txtsql & " MID(A.FECHA_TRANS,7,2) & '-' &   MID(A.FECHA_TRANS,5,2) AS Fecha, A.HORA_TRANS AS Hora, "
    txtsql = txtsql & " (b.nombre + ',' +b.apellido) as Cajero "
    'txtsql = txtsql & " FORMAT(A.FECHA_TRANS,'####-##-##') AS Fecha, A.HORA_TRANS AS Hora, (b.nombre + ',' +b.apellido) as Cajero "
    'txtsql = txtsql & " A.Hora, (b.nombre + ',' +b.apellido) as Cajero "
    '|||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||
    '|||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||
    txtsql = txtsql & " FROM HIST_TR AS A LEFT JOIN CAJEROS AS b ON A.CAJERO = b.NUMERO "
    'txtsql = txtsql & " FROM TRANSAC AS A LEFT JOIN CAJEROS AS b ON A.CAJERO = b.NUMERO "
    '|||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||
    '|||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||
    txtsql = txtsql & " WHERE NUM_TRANS = " & nNumTran
    txtsql = txtsql & " ORDER BY A.LIN"

    rsConPlatos.Open txtsql, msConn, adOpenStatic, adLockOptimistic
    
    Set MSHFPlatos.DataSource = rsConPlatos
    rsConPlatos.Close
    With MSHFPlatos
        .ColWidth(0) = 2200: .ColWidth(1) = 500: .ColWidth(2) = 800:
        .ColAlignment(2) = flexAlignRightCenter
        .ColWidth(3) = 600: .ColWidth(4) = 1000: .ColWidth(5) = 1500
    End With
End If
End Sub
Private Sub txtFecFin_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then cmdEjec.SetFocus
End Sub

Private Sub txtFecIni_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then txtFecFin.SetFocus
End Sub
'---------------------------------------------------------------------------------------
' Procedure : GetTipoPago_MontoPagado
' Author    : hsequeira
' Date      : 17/10/2017
' Purpose   : ADICIONA AL RECORSET EL TIPO Y MONTO PAGADO
'---------------------------------------------------------------------------------------
'
Private Sub GetTipoPago_MontoPagado(rsLocal As ADODB.Recordset, ID_TRANS As Long)
Dim rsH_TR_PAGO As New ADODB.Recordset
Dim rsH_TR  As New ADODB.Recordset
Dim rs_PAGO As New ADODB.Recordset
Dim cCad1 As String
Dim cCad2 As String

rsH_TR_PAGO.Open "SELECT * FROM HIST_TR_PAGO WHERE NUM_TRANS = " & ID_TRANS & " ORDER BY NUM_TRANS, LIN", msConn, adOpenStatic
rsH_TR.Open "SELECT DESCRIP FROM HIST_TR WHERE NUM_TRANS = " & ID_TRANS & " AND VALID ORDER BY NUM_TRANS, LIN", msConn, adOpenStatic
cCad2 = rsH_TR!DESCRIP
rsH_TR.Close

Do While Not rsH_TR_PAGO.EOF
    rs_PAGO.Open "SELECT DESCRIP FROM PAGOS WHERE CODIGO = " & rsH_TR_PAGO!tipo_pago, msConn, adOpenStatic
    cCad1 = cCad1 & rs_PAGO!DESCRIP & ": " & Format(rsH_TR_PAGO!MONTO, "STANDARD") & " // "
    rs_PAGO.Close
    rsH_TR_PAGO.MoveNext
Loop
rsH_TR_PAGO.Close
rsLocal.Update "PAGADO", cCad1
rsLocal.Update "PRODUCTO", cCad2
'rsLocal!PAGADO = cCad1
'rsLocal!PRODUCTO = cCad2

End Sub
