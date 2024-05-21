VERSION 5.00
Object = "{43135020-B751-4DDD-AB4A-B6D8A342216E}#1.0#0"; "sg20o.ocx"
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   7455
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   11145
   LinkTopic       =   "Form1"
   ScaleHeight     =   7455
   ScaleWidth      =   11145
   StartUpPosition =   3  'Windows Default
   Begin DDSharpGridOLEDB2.SGGrid DD_DESCUENTO 
      Height          =   735
      Left            =   2760
      TabIndex        =   0
      Top             =   1560
      Width           =   7335
      _cx             =   12938
      _cy             =   1296
      DataMember      =   ""
      DataMode        =   1
      AutoFields      =   -1  'True
      Enabled         =   -1  'True
      GridBorderStyle =   1
      ScrollBars      =   1
      FlatScrollBars  =   0
      ScrollBarTrack  =   0   'False
      DataRowCount    =   2
      BeginProperty HeadingFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DataColCount    =   2
      HeadingRowCount =   0
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
      EvenOddStyle    =   0
      ColorEven       =   -2147483628
      ColorOdd        =   14737632
      UserResizeAnimate=   0
      UserResizing    =   0
      RowHeightMin    =   600
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      UserDragging    =   0
      UserHiding      =   0
      CellPadding     =   15
      CellBkgStyle    =   1
      CellBackColor   =   65535
      CellForeColor   =   -2147483640
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   24
         Charset         =   0
         Weight          =   700
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
      DefaultRowHeight=   500
      CellsBorderColor=   0
      CellsBorderVisible=   -1  'True
      RowNumbering    =   0   'False
      EqualRowHeight  =   0   'False
      EqualColWidth   =   0   'False
      HScrollHeight   =   0
      VScrollWidth    =   600
      Format          =   "General"
      Appearance      =   2
      FitLastColumn   =   0   'False
      SelectionMode   =   2
      MultiSelect     =   0
      AllowAddNew     =   0   'False
      AllowDelete     =   0   'False
      AllowEdit       =   -1  'True
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
      AutoResizeHeadings=   0   'False
      OLEDragMode     =   0
      OLEDropMode     =   0
      Caption         =   ""
      ScrollTipColumn =   ""
      MaxRows         =   4194304
      MaxColumns      =   8192
      NewRowPos       =   1
      CustomBkgDraw   =   0
      AutoGroup       =   0   'False
      GroupByBoxVisible=   0   'False
      GroupByBoxText  =   ""
      AlphaBlendEnabled=   0   'False
      DragAlphaLevel  =   206
      AutoSearch      =   0
      AutoSearchDelay =   2000
      OutlineIcons    =   1
      CellTipsDisplayLength=   3000
      StylesCollection=   $"Form1.frx":0000
      ColumnsCollection=   $"Form1.frx":1E1F
      ValueItems      =   $"Form1.frx":2793
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Function LoadData(JustChek As Boolean)
Dim rsTTMP As ADODB.Recordset
Dim nTotal As Single
Dim nDESCUENTO_Percent As Single
Dim nDESCUENTO_Valor As Single
Dim i As Integer
   
Set rsTTMP = New ADODB.Recordset

cSQL = "SELECT TIPO, PORCENTAJE & '% - ' & DESCRIP AS DESCUENTO, PORCENTAJE FROM DESCUENTO ORDER BY TIPO"
rsTTMP.Open cSQL, msConn, adOpenStatic, adLockOptimistic

DD_DESCUENTO.Columns.RemoveAll True

DD_DESCUENTO.DataMode = sgBound
Set DD_DESCUENTO.DataSource = rsTTMP

'DD_DESCUENTO.AutoResize = sgAutoResizeRowsAndColumns
'DD_DESCUENTO.WordWrap = True


DD_DESCUENTO.ColumnClickSort = True
DD_DESCUENTO.Columns(3).AutoSize

'DD_DESCUENTO.Columns(1).Hidden = True
DD_DESCUENTO.Columns(1).ReadOnly = True
DD_DESCUENTO.Columns(1).Width = 0
DD_DESCUENTO.Columns(2).Width = 8000
DD_DESCUENTO.Columns(3).Width = 0
'DD_DESCUENTO.Columns(3).Style.WordWrap = True
'DD_DESCUENTO.Columns(4).Width = 3800
'DD_DESCUENTO.Columns(4).Style.WordWrap = True
'DD_DESCUENTO.Columns(5).Width = 3000
'DD_DESCUENTO.Columns(5).Style.WordWrap = True

'DD_DESCUENTO.Columns(6).Width = 1900  'SECTOR
'DD_DESCUENTO.Columns(7).Width = 1200
'DD_DESCUENTO.Columns(8).Width = 1200
'DD_DESCUENTO.Columns(8).Style.WordWrap = True
'DD_DESCUENTO.Columns(9).Width = 2200
'DD_DESCUENTO.Columns(10).Width = 2200

For i = 0 To DD_DESCUENTO.Rows.Count - 1
    DD_DESCUENTO.Rows.At(i).AutoSize
Next

'DD_DESCUENTO.Rows.Find "CODIGO", sgOpEqual, GetNEWIndice("CLIENTES.CODIGO", True)

If JustChek Then
    DD_DESCUENTO.ToolTipText = "Tocar Tecla DEL para ELIMINAR" & Space(10) & "F2 Para Modificar Descripción"
Else
    '~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
    'INFO: MUESTRA EN PANTALLA LA UBICACION DONDE FUE AGREGADO EL CLIENTE
    '~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
'    Dim objRow As SGRow
'    Set objRow = DD_DESCUENTO.Rows.Find("ZONA", sgOpEqual, GetNEWIndice("ZONAS.ZONA", True))
'    DD_DESCUENTO.Row = objRow.Position
'    DD_DESCUENTO.Rows.Current.Style.BackColor = vbYellow
End If

End Function

Private Sub Form_Load()
Call LoadData(True)
End Sub
