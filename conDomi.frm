VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{43135020-B751-4DDD-AB4A-B6D8A342216E}#1.0#0"; "sg20o.ocx"
Begin VB.Form conDomi 
   BackColor       =   &H00B39665&
   Caption         =   "CONSULTA VENTAS DE DOMICILIO"
   ClientHeight    =   8490
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11370
   Icon            =   "conDomi.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   8490
   ScaleWidth      =   11370
   StartUpPosition =   1  'CenterOwner
   Begin VB.CheckBox chkTodos 
      BackColor       =   &H00B39665&
      Caption         =   "Mostrar todos los Clientes"
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
      Height          =   735
      Left            =   8160
      TabIndex        =   16
      Top             =   7680
      Width           =   1215
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  'Right Justify
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3960
      Locked          =   -1  'True
      TabIndex        =   15
      Top             =   7800
      Width           =   2295
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
      Left            =   9960
      TabIndex        =   14
      Top             =   7680
      Width           =   1215
   End
   Begin VB.OptionButton opcInfo 
      BackColor       =   &H00B39665&
      Caption         =   "Mostrar Detalle de Transacciones"
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
      Index           =   1
      Left            =   120
      TabIndex        =   13
      Top             =   8040
      Visible         =   0   'False
      Width           =   3855
   End
   Begin VB.OptionButton opcInfo 
      BackColor       =   &H00B39665&
      Caption         =   "Mostrar Transacciones"
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
      Index           =   0
      Left            =   120
      TabIndex        =   12
      Top             =   7560
      Width           =   3015
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
      TabIndex        =   6
      Top             =   480
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
      Left            =   3000
      TabIndex        =   5
      Top             =   720
      Visible         =   0   'False
      Width           =   2055
   End
   Begin VB.PictureBox Picture1 
      Height          =   375
      Left            =   2880
      ScaleHeight     =   315
      ScaleWidth      =   435
      TabIndex        =   3
      Top             =   120
      Visible         =   0   'False
      Width           =   495
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
      Left            =   3720
      TabIndex        =   1
      ToolTipText     =   "Obtener Ventas por Fechas Seleccionadas"
      Top             =   360
      Value           =   -1  'True
      Visible         =   0   'False
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
      Left            =   5280
      TabIndex        =   0
      ToolTipText     =   "Obtener Ventas por Reporte Z"
      Top             =   360
      Visible         =   0   'False
      Width           =   1095
   End
   Begin MSComctlLib.ListView LVZ 
      Height          =   825
      Left            =   6375
      TabIndex        =   2
      Top             =   270
      Visible         =   0   'False
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
   Begin MSComCtl2.DTPicker txtFecIni 
      Height          =   345
      Left            =   1440
      TabIndex        =   4
      Top             =   225
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   609
      _Version        =   393216
      Format          =   132710401
      CurrentDate     =   36431
   End
   Begin MSComCtl2.DTPicker txtFecFin 
      Height          =   345
      Left            =   1440
      TabIndex        =   7
      Top             =   720
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   609
      _Version        =   393216
      Format          =   132710401
      CurrentDate     =   36430
   End
   Begin DDSharpGridOLEDB2.SGGrid DD_PEDDETALLE 
      Height          =   6135
      Left            =   120
      TabIndex        =   11
      Top             =   1320
      Width           =   11055
      _cx             =   19500
      _cy             =   10821
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
      GroupByBoxVisible=   -1  'True
      GroupByBoxText  =   "Arrastre el Titulo de la columna aqui para agrupar por esa columna"
      AlphaBlendEnabled=   0   'False
      DragAlphaLevel  =   206
      AutoSearch      =   0
      AutoSearchDelay =   2000
      OutlineIcons    =   1
      CellTipsDisplayLength=   3000
      StylesCollection=   $"conDomi.frx":000C
      ColumnsCollection=   $"conDomi.frx":1E3B
      ValueItems      =   $"conDomi.frx":27B1
   End
   Begin VB.Image ImageDomiClientes 
      Height          =   480
      Left            =   7560
      Picture         =   "conDomi.frx":2851
      ToolTipText     =   "Exportar Clientes de Domicilio"
      Top             =   7800
      Width           =   480
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   6600
      Picture         =   "conDomi.frx":2C93
      ToolTipText     =   "Exportar Datos"
      Top             =   7800
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
      Left            =   240
      TabIndex        =   10
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
      TabIndex        =   9
      Top             =   840
      Width           =   1215
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
      Left            =   5160
      TabIndex        =   8
      Top             =   720
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Shape Borde1 
      BorderColor     =   &H0000FFFF&
      BorderWidth     =   2
      Height          =   975
      Index           =   0
      Left            =   120
      Top             =   195
      Width           =   5055
   End
   Begin VB.Shape Borde1 
      BorderColor     =   &H0000FFFF&
      BorderWidth     =   2
      Height          =   975
      Index           =   1
      Left            =   5160
      Top             =   195
      Width           =   3735
   End
End
Attribute VB_Name = "conDomi"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim nOpcion As Integer
Private HayConexion As Boolean

Private Sub cmdEjec_Click()
Dim dF1 As String, dF2 As String
Dim cSQL As String
Dim rsDOMI As ADODB.Recordset
Dim oTemp As String
Dim oDate As Date


If Not HayConexion Then
    ShowMsg "NO HAY CONEXION CON LA BASE DE DATOS DE DOMICILIO, NO PUEDE EJECUTAR LA CONSULTA", vbYellow, vbRed
    Exit Sub
End If

dF1 = Format(txtFecIni, "YYYYMMDD")
dF2 = Format(txtFecFin, "YYYYMMDD")

Me.MousePointer = vbHourglass

DD_PEDDETALLE.DataMode = sgBound
Set DD_PEDDETALLE.DataSource = Nothing
Set rsDOMI = New ADODB.Recordset

Text1.Text = ""

Select Case nOpcion
    Case 0
    
        msConnDomi.BeginTrans
        msConnDomi.Execute "DROP TABLE TMP_PEDPENDIENTES"
        msConnDomi.CommitTrans

        msConnDomi.BeginTrans
        
        cSQL = "SELECT A.MESA AS PEDIDO, A.CLIENTE AS TELEFONO, A.EXTENSION AS Exten, "
        cSQL = cSQL & " C.NOMBRE + SPACE(1) + C.APELLIDO AS CLIENTE,"
        cSQL = cSQL & " RIGHT(A.FECHA,2) + '/' + MID(A.FECHA,5,2) + '/' + LEFT(A.FECHA,4)  AS FECHA, "
        cSQL = cSQL & " A.HORA, DATEDIFF('n',A.HORA,FORMAT(NOW(),'HH:MM')) AS MINUTOS, "
        cSQL = cSQL & " D.DESCRIPCION_LARGA AS ZONA, A.ID_MOTO, FORMAT(A.MONTO,'CURRENCY') AS MONTO "
        cSQL = cSQL & " INTO TMP_PEDPENDIENTES "
        cSQL = cSQL & " FROM "
        cSQL = cSQL & " CLIENTES_TRANS AS A, CLIENTES AS C, ZONAS AS D "
        cSQL = cSQL & " WHERE "
        cSQL = cSQL & " A.CLIENTE = C.TELEFONO AND A.EXTENSION = C.EXTENSION "
        cSQL = cSQL & " AND C.ZONA = D.ZONA "
        'INFO: 17NOV2017
        cSQL = cSQL & " AND A.FECHA >= '" & dF1 & "' "
        cSQL = cSQL & " AND A.FECHA <= '" & dF2 & "' "
        cSQL = cSQL & " ORDER BY A.FECHA DESC, A.HORA ASC"
        
        msConnDomi.Execute cSQL
        msConnDomi.CommitTrans
        
        cSQL = ""
        
        cSQL = "SELECT A.PEDIDO, A.TELEFONO, A.EXTEN, A.MONTO, "
        cSQL = cSQL & " A.CLIENTE, A.FECHA, A.HORA, A.MINUTOS, "
        cSQL = cSQL & " B.NOMBRE + SPACE(1) + B.APELLIDO AS MOTORIZADO,"
        cSQL = cSQL & " A.ZONA "
        cSQL = cSQL & " FROM "
        cSQL = cSQL & " TMP_PEDPENDIENTES AS A LEFT JOIN MOTO AS B "
        cSQL = cSQL & " ON A.ID_MOTO = B.ID_MOTO "
        'cSQL = cSQL & " ORDER BY A.HORA"
    Case 1
        DD_PEDDETALLE.DataMode = sgBound
        Set DD_PEDDETALLE.DataSource = Nothing
        Me.MousePointer = vbDefault
        Exit Sub
        
    Case Else
        cSQL = "SELECT FECHA, MAX(SPACE(12)) AS DIA, SUM(MONTO) AS VENTAS, "
        cSQL = cSQL & " MAX(SPACE(3)) AS MES, MAX(SPACE(3)) AS DIARIO, COUNT(*) AS REPARTOS, AVG(MONTO) AS PROMEDIO "
        cSQL = cSQL & " FROM CLIENTES_TRANS"
        cSQL = cSQL & " WHERE FECHA >= '" & dF1 & "' "
        cSQL = cSQL & " AND FECHA <= '" & dF2 & "' "
        cSQL = cSQL & " GROUP BY FECHA   "
        cSQL = cSQL & " ORDER BY FECHA DESC"
        DD_PEDDETALLE.Columns(2).Style.TextAlignment = sgAlignRightCenter
        'DD_PEDDETALLE.Columns(2).Style.Format = "CURRENCY"
End Select



'rsDOMI.Open cSQL, msConnDomi, adOpenKeyset, adLockOptimistic
rsDOMI.Open cSQL, msConnDomi, adOpenDynamic, adLockPessimistic
'
'Do While Not rsDOMI.EOF
'    rsDOMI!FECHA = GetFecha(rsDOMI!FECHA)
'    rsDOMI.MoveNext
'Loop
'rsDOMI.MoveFirst

DD_PEDDETALLE.DataMode = sgBound
Set DD_PEDDETALLE.DataSource = rsDOMI
On Error Resume Next
DD_PEDDETALLE.Columns(4).Style.TextAlignment = sgAlignRightCenter

Select Case nOpcion
    Case 0
    Case 1
    Case Else
        Dim nVentas As Single
        DD_PEDDETALLE.Columns(1).Width = 0
        DD_PEDDETALLE.Columns(2).Width = 1300
        DD_PEDDETALLE.Columns(4).Width = 500
        DD_PEDDETALLE.Columns(5).Width = 1000
        For i = 1 To DD_PEDDETALLE.RowCount - 1
            DD_PEDDETALLE.Col = 0
            DD_PEDDETALLE.Row = i
          
            oTemp = DD_PEDDETALLE.CurrentCell.value
            
            DD_PEDDETALLE.Col = 1
            DD_PEDDETALLE.CurrentCell.value = GetFecha(oTemp)
            
            DD_PEDDETALLE.Col = 4
            DD_PEDDETALLE.CurrentCell.value = GetDia(oTemp)
            
            DD_PEDDETALLE.Col = 0
            oTemp = DD_PEDDETALLE.CurrentCell.value
            DD_PEDDETALLE.Col = 3
            DD_PEDDETALLE.CurrentCell.value = GetMes(oTemp)
            
            DD_PEDDETALLE.Col = 2
            nVentas = nVentas + DD_PEDDETALLE.CurrentCell.value
        Next
        'DD_PEDDETALLE.Columns(1).Style.Format = "####-##-##"
        DD_PEDDETALLE.Columns(3).Style.Format = "CURRENCY"
        DD_PEDDETALLE.Columns(7).Style.Format = "CURRENCY"
        DD_PEDDETALLE.Columns(7).Width = DD_PEDDETALLE.Columns(7).Width * 0.7
        DD_PEDDETALLE.Row = 1
        Text1.Text = Format(nVentas, "CURRENCY")
        'DD_PEDDETALLE.Columns(1).DataChanged
End Select

On Error GoTo 0

Me.MousePointer = vbDefault

End Sub

Private Sub cmdSalir_Click()
If HayConexion Then
    Call CloseDBDomicilio
End If
Unload Me
Set conDomi = Nothing
End Sub

Private Sub DD_PEDDETALLE_BeforeGroupChange(ByVal Operation As DDSharpGridOLEDB2.sgGroupOperation, ByVal GroupOrColIndex As Long, ByVal NewIndex As Long, SortOrder As DDSharpGridOLEDB2.sgSortOrder, SortType As DDSharpGridOLEDB2.sgSortType, ShowFooter As Boolean, Cancel As Boolean)
DD_PEDDETALLE.Sort -1, -1, DD_PEDDETALLE.Columns(1).Key, sgSortAscending, sgSortTypeNumber
End Sub

Private Sub DD_PEDDETALLE_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
If Button = 2 Then
    PopupMenu MainMant.MenuSharpGrid
End If
End Sub

Private Sub Form_Load()
txtFecIni = Format(Date, "SHORT DATE")
txtFecFin = Format(Date, "SHORT DATE")
nOpcion = 2
If OpenDBDomicilio Then
    HayConexion = True
    Call Seguridad
Else
    HayConexion = False
End If

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
        txtFecIni.Enabled = False: txtFecFin.Enabled = False: opcTipo(0).Enabled = False: opcTipo(1).Enabled = False
        LVZ.Enabled = False: cmdEjec.Enabled = False
        DD_PEDDETALLE.Enabled = False
        Image1.Enabled = False
End Select
End Function

Private Sub Form_Resize()
AlignGrid Me, DD_PEDDETALLE, 120, 1400, 120, 1000
End Sub

Private Sub Image1_Click()
Call ExportToExcelOrCSVList(DD_PEDDETALLE)
End Sub

Private Sub ImageDomiClientes_Click()
'INFO DE CLIENTES.
' SELECT A.TELEFONO, A.NOMBRE, A.APELLIDO, A.ZONA, A.DIRECCION1, A.DIRECCION2, COUNT(B.CLIENTE) AS PEDIDOS, SUM(B.MONTO) AS VENTAS
' FROM CLIENTES AS A LEFT JOIN CLIENTES_TRANS AS B ON A.TELEFONO = B.CLIENTE
' GROUP BY A.TELEFONO, A.NOMBRE, A.APELLIDO, A.ZONA, A.DIRECCION1, A.DIRECCION2
'================================================
'SELECT A.TELEFONO, A.NOMBRE, A.APELLIDO, A.ZONA, A.DIRECCION1, A.DIRECCION2, COUNT(B.CLIENTE) AS PEDIDOS, SUM(B.MONTO) AS VENTAS
'FROM CLIENTES AS A LEFT JOIN CLIENTES_TRANS AS B ON A.TELEFONO = B.CLIENTE
' GROUP BY A.TELEFONO, A.NOMBRE, A.APELLIDO, A.ZONA, A.DIRECCION1, A.DIRECCION2
'ORDER BY 7 DESC, 8 ASC
Dim rsDOMI As ADODB.Recordset
Dim cSQL As String
Dim dF1 As String, dF2 As String

Me.MousePointer = vbHourglass

dF1 = Format(txtFecIni, "YYYYMMDD")
dF2 = Format(txtFecFin, "YYYYMMDD")

DD_PEDDETALLE.DataMode = sgBound
Set DD_PEDDETALLE.DataSource = Nothing
Set rsDOMI = New ADODB.Recordset

Text1.Text = ""

cSQL = "SELECT A.TELEFONO, A.NOMBRE, A.APELLIDO, A.EMAIL, A.ZONA, A.DIRECCION1, A.DIRECCION2, "
cSQL = cSQL & " COUNT(B.CLIENTE) AS PEDIDOS, FORMAT(SUM(B.MONTO),'CURRENCY') AS VENTAS"
cSQL = cSQL & " FROM CLIENTES AS A LEFT JOIN CLIENTES_TRANS AS B ON A.TELEFONO = B.CLIENTE"
If conDomi.chkTodos.value = vbChecked Then
Else
    cSQL = cSQL & " WHERE B.FECHA >= '" & dF1 & "' "
    cSQL = cSQL & " AND B.FECHA <= '" & dF2 & "' "
End If
cSQL = cSQL & " GROUP BY A.TELEFONO, A.NOMBRE, A.APELLIDO, A.EMAIL, A.ZONA, A.DIRECCION1, A.DIRECCION2"
cSQL = cSQL & " ORDER BY 8 DESC, 9 ASC"

rsDOMI.Open cSQL, msConnDomi, adOpenDynamic, adLockPessimistic

DD_PEDDETALLE.DataMode = sgBound
Set DD_PEDDETALLE.DataSource = rsDOMI
'On Error Resume Next
'DD_PEDDETALLE.Columns(4).Style.TextAlignment = sgAlignRightCenter

Me.MousePointer = vbDefault

End Sub

Private Sub opcInfo_Click(Index As Integer)
nOpcion = Index
End Sub

'---------------------------------------------------------------------------------------
' Procedure : AlignGrid
' Author    : hsequeira
' Date      : 20/09/2012
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub AlignGrid(frm As Form, _
                     grid As Object, _
                     Left As Single, _
                     Top As Single, _
                     Right As Single, _
                     Bottom As Single)

   grid.Left = Left
   grid.Top = Top
   
   Dim W As Single
   W = frm.ScaleWidth - grid.Left - Right
   If W < 1000 Then W = 1000
   grid.Width = W
   
   Dim H As Single
   H = frm.ScaleHeight - grid.Top - Bottom
   If H < 1000 Then H = 1000
   grid.Height = H
End Sub

