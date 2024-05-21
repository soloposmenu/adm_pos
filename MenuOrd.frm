VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{43135020-B751-4DDD-AB4A-B6D8A342216E}#1.0#0"; "sg20o.ocx"
Begin VB.Form MenuOrd 
   BackColor       =   &H00B39665&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "ORDEN DEL MENU EN VENTAS"
   ClientHeight    =   7410
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6540
   Icon            =   "MenuOrd.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   7410
   ScaleWidth      =   6540
   StartUpPosition =   2  'CenterScreen
   Begin MSComctlLib.ListView LV 
      Height          =   615
      Left            =   4560
      TabIndex        =   4
      Top             =   480
      Visible         =   0   'False
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   1085
      View            =   3
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   0
   End
   Begin VB.CommandButton cmdOrdAlf 
      Caption         =   "Mostrar Menu Alfabeticamente"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   4560
      TabIndex        =   3
      Top             =   1680
      Width           =   1815
   End
   Begin VB.CommandButton cmdOrdVta 
      Caption         =   "Mostrar Menu Ordenado x Ventas"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   4560
      TabIndex        =   2
      Top             =   3000
      Width           =   1815
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
      Left            =   4680
      TabIndex        =   1
      Top             =   6720
      Width           =   1575
   End
   Begin DDSharpGridOLEDB2.SGGrid DD_MENU 
      Height          =   6615
      Left            =   240
      TabIndex        =   5
      Top             =   720
      Width           =   4095
      _cx             =   7223
      _cy             =   11668
      DataMember      =   ""
      DataMode        =   1
      AutoFields      =   -1  'True
      Enabled         =   -1  'True
      GridBorderStyle =   1
      ScrollBars      =   1
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
      UserDragging    =   3
      UserHiding      =   2
      CellPadding     =   15
      CellBkgStyle    =   1
      CellBackColor   =   -2147483643
      CellForeColor   =   -2147483640
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
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
      GroupIndentation=   225
      InactiveSelection=   1
      AutoScroll      =   -1  'True
      AutoResize      =   0
      AutoResizeHeadings=   -1  'True
      OLEDragMode     =   1
      OLEDropMode     =   1
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
      StylesCollection=   $"MenuOrd.frx":0442
      ColumnsCollection=   $"MenuOrd.frx":2251
      ValueItems      =   $"MenuOrd.frx":2766
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00B39665&
      Caption         =   "ESCRIBA el Orden en que desea que los Departamentos aparezcan en el Menú de Ventas"
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
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   4215
   End
End
Attribute VB_Name = "MenuOrd"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\access\SOLO.mdb;Persist Security Info=False;Jet OLEDB:Database Password=master24
Dim rsLocal As New ADODB.Recordset
Dim vOrden As Variant
Dim rsTABLA As ADODB.Recordset
Private Sub cmdOrdAlf_Click()
Dim vResp As Variant

'vResp = MsgBox("¿ Desea Mostrar el Menu Alfabeticamente ? ", vbQuestion + vbYesNo, "Orden del Menu en el Sistema de Ventas")
If ShowMsg("¿ Desea Mostrar el Menu Alfabeticamente ?", vbYellow, vbBlue, vbYesNo) = vbYes Then vResp = vbYes Else vResp = vbNo
If vResp = vbYes Then
    vOrden = "DESCRIP"
    Call LoadData
End If

Call Seguridad

End Sub

Private Sub cmdOrdVta_Click()
Dim vResp As Variant

'vResp = MsgBox("¿ Desea Mostrar el Menu por Ventas ? ", vbQuestion + vbYesNo, "Orden del Menu en el Sistema de Ventas")
If ShowMsg("¿ Desea Mostrar el Menu por Ventas ?", vbYellow, vbBlue, vbYesNo) = vbYes Then vResp = vbYes Else vResp = vbNo
If vResp = vbYes Then
    vOrden = "VALOR DESC, DESCRIP"
    Call LoadData
End If

Call Seguridad

End Sub

Private Sub Command1_Click()
Command1.Tag = "100"
rsTABLA.Close
Set rsTABLA = Nothing
Unload Me
End Sub

Private Sub DataGrid1_Click()
MsgBox "NO ESTA DISPONIBLE EN ESTE MOMENTO", vbInformation, BoxTit
On Error Resume Next
End Sub

Private Sub Form_Load()
vOrden = "DESCRIP"

Set rsTABLA = New ADODB.Recordset

Call LoadData

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
        cmdOrdAlf.Enabled = False
        cmdOrdVta.Enabled = False
    Case "N"        'SIN DERECHOS
        LV.Enabled = False
        cmdOrdAlf.Enabled = False
        cmdOrdVta.Enabled = False
End Select
End Function

Private Sub SHOWDATA()
Dim nFila As Integer
Dim cSQL As String
nFila = 1
LV.ListItems.Clear
LV.ColumnHeaders.Clear
LV.ColumnHeaders.Add , , "Descripcion"
LV.ColumnHeaders.Add , , "Orden"
LV.ColumnHeaders.Item(2).Alignment = lvwColumnRight

cSQL = "SELECT DESCRIP, ORDEN FROM DEPTO order by " & vOrden
rsLocal.Open cSQL, msConn, adOpenStatic, adLockOptimistic
'Debug.Print cSQL
Do While Not rsLocal.EOF
    LV.ListItems.Add , , rsLocal!DESCRIP
    LV.ListItems.Item(nFila).ListSubItems.Add , , rsLocal!Orden
    nFila = nFila + 1
    rsLocal.MoveNext
Loop
rsLocal.Close

'Call LoadData
End Sub

Private Function LoadData() As Boolean

Dim cSQL As String

On Error GoTo ErrAdm:

cSQL = "SELECT DESCRIP, ORDEN FROM DEPTO order by " & vOrden

'rsTABLA.Open cSQL, msConn, adOpenStatic, adLockReadOnly
rsTABLA.Open cSQL, msConn, adOpenKeyset, adLockOptimistic

With DD_MENU
    rsTABLA.MoveFirst
   '.DataMode = sgUnbound
   .DataMode = sgBound
   Set .DataSource = rsTABLA
   '.LoadArray rsTABLA.GetRows()
End With

'INFO: ARCHIVO QUE SE ESTA EXPORTANDO

DD_MENU.Columns(1).Width = 2400
DD_MENU.Columns(1).ReadOnly = True
DD_MENU.Columns(2).Width = 700

LoadData = True
On Error GoTo 0
Exit Function

ErrAdm:
If Err.Number = 3705 Then
    rsTABLA.Close
    Resume
Else
    ShowMsg "Admin. Menu Orden: " & Err.Number & " - " & Err.Description, vbYellow, vbRed
End If
End Function

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
If Command1.Tag = "100" Then
Else
    BoxPreg = "¡ Favor utilize el boton Salir !"
    BoxResp = MsgBox(BoxPreg, vbExclamation + vbOKOnly, BoxTit)
    Cancel = True
End If
End Sub

Private Sub LV_DblClick()
Dim nNewValue As Variant
Dim vResp As Variant

nNewValue = InputBox("Introduzca el nuevo Orden que desea que aparezca el departmento " & LV.SelectedItem.Text, "Cambio de Orden", LV.SelectedItem.ListSubItems.Item(1).Text)
If nNewValue = "" Then Exit Sub
If nNewValue = LV.SelectedItem.ListSubItems.Item(1).Text Then Exit Sub
vResp = MsgBox("¿ Desea actualizar el Orden ?", vbYesNo + vbQuestion, "Cambio de Orden")
If vResp = vbYes Then
    EscribeLog ("Admin." & "Cambio de Orden del Menu, " & LV.SelectedItem.Text & " (" & nNewValue & ") en  vez de (" & LV.SelectedItem.ListSubItems.Item(1).Text & ")")
    msConn.BeginTrans
    msConn.Execute "UPDATE DEPTO SET ORDEN = " & nNewValue & " WHERE DESCRIP = '" & LV.SelectedItem.Text & "'"
    msConn.CommitTrans
    Call LoadData
End If
End Sub

Private Sub LV_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    LV_DblClick
End If
End Sub

Private Sub DD_MENU_OnInit()

DD_MENU.ColumnClickSort = True
DD_MENU.EvenOddStyle = sgEvenOddRows
DD_MENU.ColorOdd = &HE0E0E0

DD_MENU.AllowEdit = True

DD_MENU.DataColCount = 2

'D_RECETAS.Columns(1).Hidden = True
DD_MENU.Columns(1).Caption = "Descripción"
DD_MENU.Columns(1).Width = 2400
DD_MENU.Columns(1).ReadOnly = True

DD_MENU.Columns(2).Caption = "Orden"
DD_MENU.Columns(2).Width = 700
DD_MENU.Columns(2).Style.TextAlignment = sgAlignRightCenter
DD_MENU.Columns(2).ReadOnly = False

End Sub

