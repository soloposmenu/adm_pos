VERSION 5.00
Object = "{43135020-B751-4DDD-AB4A-B6D8A342216E}#1.0#0"; "sg20o.ocx"
Begin VB.Form conDomiClientes 
   BackColor       =   &H00B39665&
   Caption         =   "Clientes de Domicilio"
   ClientHeight    =   8715
   ClientLeft      =   -195
   ClientTop       =   1365
   ClientWidth     =   12885
   Icon            =   "conDomiClientes.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   8715
   ScaleWidth      =   12885
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox txtNombre 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2400
      MaxLength       =   12
      TabIndex        =   1
      Top             =   7560
      Width           =   2655
   End
   Begin VB.TextBox txtApellido 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2400
      MaxLength       =   12
      TabIndex        =   2
      Top             =   8160
      Width           =   2655
   End
   Begin VB.CommandButton cmdSalir 
      Caption         =   "Sa&lir"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   555
      Left            =   11520
      TabIndex        =   3
      Top             =   8040
      Width           =   1215
   End
   Begin DDSharpGridOLEDB2.SGGrid DD_PEDDETALLE 
      Height          =   7335
      Left            =   120
      TabIndex        =   0
      ToolTipText     =   "Tocar Tecla DEL para ELIMINAR"
      Top             =   120
      Width           =   12615
      _cx             =   22251
      _cy             =   12938
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
      UserResizing    =   0
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
      DefaultRowHeight=   450
      CellsBorderColor=   0
      CellsBorderVisible=   -1  'True
      RowNumbering    =   -1  'True
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
      AllowDelete     =   -1  'True
      AllowEdit       =   -1  'True
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
      AutoGroup       =   0   'False
      GroupByBoxVisible=   0   'False
      GroupByBoxText  =   ""
      AlphaBlendEnabled=   0   'False
      DragAlphaLevel  =   206
      AutoSearch      =   0
      AutoSearchDelay =   2000
      OutlineIcons    =   1
      CellTipsDisplayLength=   3000
      StylesCollection=   $"conDomiClientes.frx":0442
      ColumnsCollection=   $"conDomiClientes.frx":2272
      ValueItems      =   $"conDomiClientes.frx":2BE8
   End
   Begin VB.Label lbPedidosCliente 
      BackColor       =   &H00B39665&
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFC0&
      Height          =   255
      Left            =   5280
      TabIndex        =   7
      Top             =   7800
      Width           =   3855
   End
   Begin VB.Label Label1 
      BackColor       =   &H00B39665&
      Caption         =   "Escriba NOMBRE"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FFFF&
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   6
      Top             =   7800
      Width           =   2055
   End
   Begin VB.Label Label1 
      BackColor       =   &H00B39665&
      Caption         =   "Escriba APELLIDO"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FFFF&
      Height          =   255
      Index           =   1
      Left            =   120
      TabIndex        =   5
      Top             =   8400
      Width           =   2055
   End
   Begin VB.Image Image 
      Height          =   480
      Left            =   10080
      Picture         =   "conDomiClientes.frx":2C88
      ToolTipText     =   "Exportar Datos"
      Top             =   8085
      Width           =   480
   End
   Begin VB.Label lb 
      BackColor       =   &H00B39665&
      Caption         =   "PROCESANDO DATOS, ESPERE ..."
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   255
      Left            =   5235
      TabIndex        =   4
      Top             =   8280
      Width           =   4695
   End
End
Attribute VB_Name = "conDomiClientes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private rsTTMP As ADODB.Recordset
Private cSQL As String
Private Sub cmdSalir_Click()
If HayConexion Then
    Call CloseDBDomicilio
End If
Unload Me
'Set conDomi = Nothing
End Sub

Private Sub DD_PEDDETALLE_BeforeDelete(CancelDelete As Boolean)
'CancelDelete = MsgBox("¿Desea Eliminar Este Registro?", _
      vbQuestion + vbYesNo + vbDefaultButton2, "Confirmar Eliminacion") = vbNo
'CancelDelete = ShowMsg("¿ DESEA ELIMINAR ESTE REGISTRO ?", vbYellow, vbRed, vbYesNo)
If ShowMsg("¿ DESEA ELIMINAR ESTE REGISTRO ?", vbYellow, vbRed, vbYesNo) = vbYes Then
    CancelDelete = False
Else
    CancelDelete = True
End If
End Sub

Private Sub DD_PEDDETALLE_Click()
    Call PedidosDelCliente
End Sub
Private Sub Form_Load()
If OpenDBDomicilio Then
    HayConexion = True
    Show
    Me.lb.Refresh
    Call LoadData(True)
    Me.lb.ForeColor = vbYellow
    Me.lb.Caption = "TOTAL DE REGISTROS: " & Format(DD_PEDDETALLE.Rows.Count - 1, "###,###")
Else
    HayConexion = False
End If
End Sub

Private Function LoadData(JustChek As Boolean)

Dim nTotal As Single
Dim nDESCUENTO_Percent As Single
Dim nDESCUENTO_Valor As Single
Dim i As Integer
    
Set rsTTMP = New ADODB.Recordset
cSQL = "SELECT FORMAT(TELEFONO,'####-####') AS TEL, EXTENSION, NOMBRE, APELLIDO, "
cSQL = cSQL & " EMAIL, DIRECCION1, DIRECCION2, DIRECCION3 "
cSQL = cSQL & " FROM CLIENTES "
cSQL = cSQL & " ORDER BY NOMBRE, APELLIDO"
rsTTMP.Open cSQL, msConnDomi, adOpenStatic, adLockOptimistic

DD_PEDDETALLE.Columns.RemoveAll True
DD_PEDDETALLE.AutoResize = sgAutoResizeRows

DD_PEDDETALLE.DataMode = sgBound
'DD_PEDDETALLE.DataMode = sgUnbound
Set DD_PEDDETALLE.DataSource = rsTTMP

'DD_PEDDETALLE.WordWrap = True

DD_PEDDETALLE.ColumnClickSort = True
'DD_PEDDETALLE.Columns(6).AutoSize
'DD_PEDDETALLE.Columns(7).AutoSize
'DD_PEDDETALLE.Columns(8).AutoSize

'DD_PEDDETALLE.Columns(1).Hidden = True
DD_PEDDETALLE.Columns(1).ReadOnly = True
DD_PEDDETALLE.Columns(2).Width = 650
DD_PEDDETALLE.Columns(3).Width = 1800
DD_PEDDETALLE.Columns(4).Width = 1800
DD_PEDDETALLE.Columns(5).Width = 2000
DD_PEDDETALLE.Columns(6).Width = 4000
DD_PEDDETALLE.Columns(7).Width = 4000
DD_PEDDETALLE.Columns(8).Width = 4000
'DD_PEDDETALLE.Columns(6).Style.WordWrap = False
'DD_PEDDETALLE.Columns(7).Style.WordWrap = False
'DD_PEDDETALLE.Columns(8).Style.WordWrap = False
'For i = 1 To DD_PEDDETALLE.Rows.Count - 1
    'DD_PEDDETALLE.Rows.At(i).AutoSize
'Next

If JustChek Then
Else
    '~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
    'INFO: MUESTRA EN PANTALLA LA UBICACION DONDE FUE AGREGADO EL CLIENTE
    '~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
'    Dim objRow As SGRow
'    Set objRow = DD_PEDDETALLE.Rows.Find("ZONA", sgOpEqual, GetNEWIndice("ZONAS.ZONA", True))
'    DD_PEDDETALLE.Row = objRow.Position
'    DD_PEDDETALLE.Rows.Current.Style.BackColor = vbYellow
End If

rsTTMP.Close

End Function

Private Sub Image_Click()
ShowMsg "Para leer este Archivo Correctamente, seleccione la Opcion (2) Para Lista, y luego Abralo en EXCEL", vbBlue, vbYellow
Call ExportToExcelOrCSVList(DD_PEDDETALLE)
End Sub

Private Sub txtApellido_Change()

If txtApellido = "" Then
    'NO HACE NADA
    cSQL = "SELECT FORMAT(TELEFONO,'####-####') AS TEL, EXTENSION, NOMBRE, APELLIDO, "
    cSQL = cSQL & " EMAIL, DIRECCION1, DIRECCION2, DIRECCION3 "
    cSQL = cSQL & " FROM CLIENTES "
    cSQL = cSQL & " ORDER BY NOMBRE, APELLIDO"
Else
    cSQL = "SELECT FORMAT(TELEFONO,'####-####') AS TEL, EXTENSION, NOMBRE, APELLIDO, "
    cSQL = cSQL & " EMAIL, DIRECCION1, DIRECCION2, DIRECCION3 "
    cSQL = cSQL & " FROM CLIENTES "
    cSQL = cSQL & " WHERE APELLIDO LIKE '%" & txtApellido & "%'"
    cSQL = cSQL & " ORDER BY APELLIDO, NOMBRE"
End If
    rsTTMP.Open cSQL, msConnDomi, adOpenStatic, adLockOptimistic
    
    DD_PEDDETALLE.AutoResize = sgAutoResizeRows
    DD_PEDDETALLE.DataMode = sgBound
    Set DD_PEDDETALLE.DataSource = rsTTMP
    
    DD_PEDDETALLE.ReBind
    
    DD_PEDDETALLE.Columns(1).ReadOnly = True
    DD_PEDDETALLE.Columns(2).Width = 650
    DD_PEDDETALLE.Columns(3).Width = 1800
    DD_PEDDETALLE.Columns(4).Width = 1800
    DD_PEDDETALLE.Columns(5).Width = 2000
    DD_PEDDETALLE.Columns(6).Width = 4000
    DD_PEDDETALLE.Columns(7).Width = 4000
    DD_PEDDETALLE.Columns(8).Width = 4000
    
    'DD_PEDDETALLE.RowHeightMin = 400
    'DD_PEDDETALLE.TextAlignment = sgAlignCenterCenter
    
    rsTTMP.Close
    Me.lb.Caption = "TOTAL DE REGISTROS: " & Format(DD_PEDDETALLE.Rows.Count - 1, "###,###")
    Me.lbPedidosCliente.Caption = ""
'End If
End Sub

Private Sub txtNombre_Change()

If txtNombre = "" Then
    'NO HACE NADA
    cSQL = "SELECT FORMAT(TELEFONO,'####-####') AS TEL, EXTENSION, NOMBRE, APELLIDO, "
    cSQL = cSQL & " EMAIL, DIRECCION1, DIRECCION2, DIRECCION3 "
    cSQL = cSQL & " FROM CLIENTES "
    cSQL = cSQL & " ORDER BY NOMBRE, APELLIDO"
Else
    cSQL = "SELECT FORMAT(TELEFONO,'####-####') AS TEL, EXTENSION, NOMBRE, APELLIDO, "
    cSQL = cSQL & " EMAIL, DIRECCION1, DIRECCION2, DIRECCION3 "
    cSQL = cSQL & " FROM CLIENTES "
    cSQL = cSQL & " WHERE NOMBRE LIKE '%" & txtNombre & "%'"
    cSQL = cSQL & " ORDER BY NOMBRE, APELLIDO"
End If
    rsTTMP.Open cSQL, msConnDomi, adOpenStatic, adLockOptimistic
    
    DD_PEDDETALLE.AutoResize = sgAutoResizeRows
    DD_PEDDETALLE.DataMode = sgBound
    Set DD_PEDDETALLE.DataSource = rsTTMP
    
    DD_PEDDETALLE.ReBind
    
    DD_PEDDETALLE.Columns(1).ReadOnly = True
    DD_PEDDETALLE.Columns(2).Width = 650
    DD_PEDDETALLE.Columns(3).Width = 1800
    DD_PEDDETALLE.Columns(4).Width = 1800
    DD_PEDDETALLE.Columns(5).Width = 2000
    DD_PEDDETALLE.Columns(6).Width = 4000
    DD_PEDDETALLE.Columns(7).Width = 4000
    DD_PEDDETALLE.Columns(8).Width = 4000
    
    'DD_PEDDETALLE.RowHeightMin = 400
    'DD_PEDDETALLE.TextAlignment = sgAlignCenterCenter
    
    rsTTMP.Close
    Me.lb.Caption = "TOTAL DE REGISTROS: " & Format(DD_PEDDETALLE.Rows.Count - 1, "###,###")
    Me.lbPedidosCliente.Caption = ""
'End If
End Sub

Private Sub PedidosDelCliente()
Dim rsPedidos As ADODB.Recordset
Dim cTelefono As String

cTelefono = Replace(DD_PEDDETALLE.Rows.Current.Cells(0).value, "-", "")

Set rsPedidos = New ADODB.Recordset
rsPedidos.Open "SELECT COUNT(*) AS PEDIDOS FROM CLIENTES_TRANS WHERE CLIENTE = '" & cTelefono & "'", msConnDomi, adOpenStatic, adLockOptimistic
If Not rsPedidos.EOF Then
    Me.lbPedidosCliente.Caption = "Pedidos del Cliente: " & rsPedidos!PEDIDOS
Else
    Me.lbPedidosCliente.Caption = ""
End If
rsPedidos.Close
Set rsPedidos = Nothing
End Sub
