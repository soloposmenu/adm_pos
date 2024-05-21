VERSION 5.00
Object = "{43135020-B751-4DDD-AB4A-B6D8A342216E}#1.0#0"; "sg20o.ocx"
Begin VB.Form AdmAreas 
   BackColor       =   &H00B39665&
   Caption         =   "MANTENIMIENTO DE AREAS"
   ClientHeight    =   9090
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   13395
   LinkTopic       =   "Form1"
   ScaleHeight     =   9090
   ScaleWidth      =   13395
   StartUpPosition =   3  'Windows Default
   Begin VB.ListBox lstMesasAsignadas 
      BackColor       =   &H00C0FFFF&
      BeginProperty DataFormat 
         Type            =   0
         Format          =   "0"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1033
         SubFormatType   =   0
      EndProperty
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   7305
      ItemData        =   "AdmAreas.frx":0000
      Left            =   12000
      List            =   "AdmAreas.frx":0002
      TabIndex        =   2
      Top             =   480
      Width           =   1215
   End
   Begin VB.CommandButton cmdAddArea 
      Caption         =   "&NUEVA AREA"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   120
      TabIndex        =   3
      Top             =   8040
      Width           =   3375
   End
   Begin VB.ListBox listAreas 
      BackColor       =   &H00FF8080&
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   930
      Left            =   5160
      TabIndex        =   0
      Top             =   8160
      Visible         =   0   'False
      Width           =   5055
   End
   Begin DDSharpGridOLEDB2.SGGrid DD_MESAS 
      Height          =   7215
      Left            =   5280
      TabIndex        =   1
      Top             =   960
      Width           =   6615
      _cx             =   11668
      _cy             =   12726
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
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   238
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DataColCount    =   2
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
         Size            =   12
         Charset         =   238
         Weight          =   700
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
      SelectionMode   =   0
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
      GroupByBoxText  =   "Drag a column header here to group by that column"
      AlphaBlendEnabled=   0   'False
      DragAlphaLevel  =   206
      AutoSearch      =   0
      AutoSearchDelay =   2000
      OutlineIcons    =   1
      CellTipsDisplayLength=   3000
      StylesCollection=   $"AdmAreas.frx":0004
      ColumnsCollection=   $"AdmAreas.frx":1E0A
      ValueItems      =   $"AdmAreas.frx":2C65
   End
   Begin DDSharpGridOLEDB2.SGGrid DD_AREAS 
      Height          =   7335
      Left            =   120
      TabIndex        =   4
      Top             =   360
      Width           =   5055
      _cx             =   8916
      _cy             =   12938
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
      TextAlignment   =   5
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
      ColorEven       =   16744576
      ColorOdd        =   16744576
      UserResizeAnimate=   1
      UserResizing    =   3
      RowHeightMin    =   400
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      UserDragging    =   2
      UserHiding      =   0
      CellPadding     =   15
      CellBkgStyle    =   1
      CellBackColor   =   16744576
      CellForeColor   =   16777215
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      FocusRect       =   2
      FocusRectColor  =   255
      FocusRectLineWidth=   7
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
      AutoResize      =   0
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
      GroupByBoxText  =   "Arrastre el Titulo de la columna aqui para agrupar por esa columna"
      AlphaBlendEnabled=   0   'False
      DragAlphaLevel  =   206
      AutoSearch      =   0
      AutoSearchDelay =   2000
      OutlineIcons    =   1
      CellTipsDisplayLength=   3000
      StylesCollection=   $"AdmAreas.frx":30B7
      ColumnsCollection=   $"AdmAreas.frx":4EDC
      ValueItems      =   $"AdmAreas.frx":5852
   End
   Begin VB.Label lbArea 
      BackColor       =   &H00FF8080&
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   5280
      TabIndex        =   5
      Top             =   360
      Width           =   6615
   End
End
Attribute VB_Name = "AdmAreas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim cArea As String
Dim nArea As Long
Private Function GetArea(nArea) As String
Dim rsAreas As New ADODB.Recordset

rsAreas.Open "SELECT DESCRIPCION FROM AREAS WHERE ID = " & nArea, msConn, adOpenStatic, adLockOptimistic
If rsAreas.EOF Then
Else
    GetArea = rsAreas!Descripcion
End If
rsAreas.Close
Set rsAreas = Nothing
End Function
Private Sub cmdAddArea_Click()

cArea = InputBox("NOMBRE DE AREA (20 CARACTERES MAXIMO)", "NOMBRE DESCRIPTIVO DE AREA", "")
If cArea = "" Then
    ShowMsg "DEBE ESCRIBIR UNA DESCRIPCION PARA CREAR UNA AREA", vbYellow, vbRed
Else
    cArea = Left(cArea, 20)
    msConn.BeginTrans
    msConn.Execute "INSERT INTO AREAS(DESCRIPCION) VALUES ('" & cArea & "')"
    msConn.CommitTrans
    Call LoadData(1)
End If
End Sub

Private Sub DD_AREAS_Click()
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

On Error GoTo DD_AREAS_Click_Error

cDescripAcompa = "" ' se limpia acompanante
nIDAcompa = 0
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
nArea = DD_AREAS.Rows.Current.Cells(0).value
cArea = DD_AREAS.Rows.Current.Cells(1).value
lbArea.Caption = cArea

Call LoadMesasDelArea(nArea)

 On Error GoTo 0
   Exit Sub

DD_AREAS_Click_Error:
    If Err.Number = -2147024809 Then
        ShowMsg "PRIMERO DEBE CREAR LAS AREAS", vbYellow, vbRed
    Else
        ShowMsg "Error " & Err.Number & " (" & Err.Description & ") in AdmAreas_Busqueda"
    End If
End Sub
Private Sub LoadMesasDelArea(nArea As Long)
Dim rsMesasEnArea As New ADODB.Recordset
Dim cSQL As String

cSQL = "SELECT AREA, MESA FROM AREAS_MESAS WHERE AREA = " & nArea & " ORDER BY AREA, MESA "
rsMesasEnArea.Open cSQL, msConn, adOpenStatic, adLockOptimistic
lstMesasAsignadas.Clear
Do While Not rsMesasEnArea.EOF
    lstMesasAsignadas.AddItem rsMesasEnArea!MESA
    rsMesasEnArea.MoveNext
Loop
rsMesasEnArea.Close
Set rsMesasEnArea = Nothing

End Sub
Private Sub DD_AREAS_DblClick()
nArea = DD_AREAS.Rows.Current.Cells(0).value
cArea = DD_AREAS.Rows.Current.Cells(1).value

If ShowMsg("¿ Desea Quitar AREA ?" & vbCrLf & vbCrLf & cArea, vbYellow, vbRed, vbYesNo) = vbYes Then BoxResp = vbYes Else BoxResp = vbNo
If BoxResp = vbYes Then
    
    msConn.BeginTrans
    msConn.Execute "DELETE FROM AREAS WHERE ID= " & nArea & ""
    msConn.CommitTrans
    
    msConn.BeginTrans
    msConn.Execute "DELETE FROM AREAS_MESAS WHERE AREA = " & nArea & ""
    msConn.CommitTrans
    
    lbArea.Caption = ""
    nArea = -99
    
    Call LoadData(0)
    lstMesasAsignadas.Clear
    'Call LoadMesasDelArea(nArea)
End If

End Sub

Private Sub DD_MESAS_Click()
'nPluSel = DD_MESAS.Rows.Current.Cells(0).value
''DD_MESAS.CurrentCell.value
If nArea = -99 Then
    ShowMsg "DEBE SELECCIONAR UN AREA", vbYellow, vbRed, vbOKOnly
Else
    If Verificar_MesaEnArea(nArea, DD_MESAS.CurrentCell.value) Then
    End If
End If
'c2PluSel = DD_AREAS.Rows.Current.Cells(2).value
End Sub
Private Function Verificar_MesaEnArea(nArea As Long, nMesa As Long) As Boolean
Dim rsVerificacion As New ADODB.Recordset
Dim cSQL As String

cSQL = "SELECT AREA, MESA FROM AREAS_MESAS WHERE MESA = " & nMesa

rsVerificacion.Open cSQL, msConn, adOpenStatic
If rsVerificacion.EOF Then
    msConn.BeginTrans
    msConn.Execute "INSERT INTO AREAS_MESAS VALUES (" & nArea & "," & nMesa & ",'" & Format(Date, "LONG DATE") & "')"
    msConn.CommitTrans
    Verificar_MesaEnArea = True
    lstMesasAsignadas.AddItem nMesa
Else
    If rsVerificacion!AREA = nArea Then
        Verificar_MesaEnArea = True
    Else
        ShowMsg "YA LA MESA (" & nMesa & ")" & vbCrLf & "ESTA ASIGNADA A OTRA AREA" & vbCrLf & "(" & GetArea(rsVerificacion!AREA) & ")", vbYellow, vbRed
        Verificar_MesaEnArea = False
    End If
End If
End Function
'Private Sub listAreas_DblClick()
'Dim iArea As Integer
'cArea = listAreas.Text
'iArea = listAreas.ListIndex
'If ShowMsg("¿ Desea Quitar AREA ?" & vbCrLf & vbCrLf & listAreas.Text, vbYellow, vbRed, vbYesNo) = vbYes Then BoxResp = vbYes Else BoxResp = vbNo
'If BoxResp = vbYes Then
'    listAreas.RemoveItem iArea
'
'    msConn.BeginTrans
'    msConn.Execute "DELETE FROM AREAS WHERE DESCRIPCION= '" & cArea & "'"
'    msConn.CommitTrans
'    Call LoadData(1)
'
'End If
'End Sub

Private Sub Form_Load()
nArea = -99
Call LoadData(0)
End Sub

Private Sub LoadData(nValorInicial As Integer)
Dim cSQL As String, cSQL2 As String
Dim rsAreas As ADODB.Recordset
Dim rsMesas As ADODB.Recordset
Dim i As Integer
Dim rowData As String

cSQL = "SELECT * FROM AREAS ORDER BY DESCRIPCION"
cSQL2 = "SELECT NUMERO FROM MESAS WHERE NUMERO <> -99 ORDER BY NUMERO "

Set rsAreas = New ADODB.Recordset
rsAreas.Open cSQL, msConn, adOpenStatic, adLockOptimistic

'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
'Set DD_AREAS.DataSource = Nothing
   
DD_AREAS.ReBind
DD_AREAS.DataMode = sgBound
Set DD_AREAS.DataSource = rsAreas
DD_AREAS.ReBind
DD_AREAS.RowHeightMin = 585
DD_AREAS.TextAlignment = sgAlignLeftCenter

 On Error Resume Next
 With DD_AREAS
    .ColumnClickSort = False
    .Columns(1).Width = 0:        'ID AREA
    '.Columns(1).Style.WordWrap = True
    .Columns(2).Width = 5000:        'DESCRIPCION
    .Columns(2).Style.WordWrap = True
End With
On Error GoTo 0


DD_AREAS.RowHeightMin = 585
DD_AREAS.TextAlignment = sgAlignLeftCenter
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

''''''On Error Resume Next
''''''listAreas.Clear
''''''rsAreas.MoveFirst
''''''On Error GoTo 0
''''''Do While Not rsAreas.EOF
''''''    listAreas.AddItem rsAreas!Descripcion
''''''    rsAreas.MoveNext
''''''Loop
''''''On Error Resume Next
''''''rsAreas.MoveFirst
''''''On Error GoTo 0


'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
If nValorInicial = 0 Then
    Set rsMesas = New ADODB.Recordset
    rsMesas.Open cSQL2, msConn, adOpenStatic, adLockOptimistic
    
    With DD_MESAS
      .GroupByBoxVisible = False
      .Appearance = sgFlat
      .BackColor = Me.BackColor
      .GridBorderStyle = sgNone
      .AllowEdit = False
      .FocusRect = sgFocusRectSolid
      .FocusRectColor = vbRed
      .FocusRectLineWidth = 3
      .AutoScroll = True
      .ScrollBarTrack = True
      
      '.ImageList = ImageList1
      .Columns.RemoveAll True
      .Rows.RemoveAll True
      .DefaultColWidth = 42 * Screen.TwipsPerPixelX
      .DefaultRowHeight = 42 * Screen.TwipsPerPixelY
    
        For i = 1 To 10
             With .Columns.Add
                .DataType = sgtLong
                '.Style.Borders = sgCellBorderAll
                .Style.TextAlignment = sgAlignCenterCenter
                .Style.DisplayType = sgDisplayText
                .Style.Font.Bold = True
                '.Style.PictureAlignment = sgPicAlignTile
                '.Style.BkgPictureAlignment = sgPicAlignCenterCenter
             End With
        Next
    
        For i = 1 To rsMesas.RecordCount Step 10
           rowData = Trim(i) & "," & Trim(i + 1) & "," _
              & Trim(i + 2) & "," & Trim(i + 3) & "," & Trim(i + 4) & "," _
              & Trim(i + 5) & "," & Trim(i + 6) & "," & Trim(i + 7) & "," _
              & Trim(i + 8) & "," & Trim(i + 9)
           .Rows.Add sgFormatCharSeparatedValue, rowData
        Next
    End With
    rsMesas.Close
    Set rsMesas = Nothing
End If
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

'rsAreas.Close

'Set rsAreas = Nothing
End Sub

Private Sub Form_Resize()
   On Error Resume Next
   
   ''DD_MESAS.Move 0, 0, Me.ScaleWidth, Me.ScaleHeight
End Sub

'Private Sub listAreas_Click()
'lstMesasAsignadas.Clear
'End Sub
Private Sub lstMesasAsignadas_DblClick()
Dim iMesa As Integer
iMesa = Val(lstMesasAsignadas.Text)

If ShowMsg("QUITAR LA MESA " & iMesa & "?", vbYellow, vbRed, vbYesNo) = vbYes Then BoxResp = vbYes Else BoxResp = vbNo
If BoxResp = vbYes Then
    
    msConn.BeginTrans
    msConn.Execute "DELETE FROM AREAS_MESAS WHERE MESA = " & iMesa
    msConn.CommitTrans
    Call LoadMesasDelArea(nArea)
End If

End Sub
