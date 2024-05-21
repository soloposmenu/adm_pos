VERSION 5.00
Object = "{A45D986F-3AAF-4A3B-A003-A6C53E8715A2}#1.0#0"; "ARVIEW2.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{43135020-B751-4DDD-AB4A-B6D8A342216E}#1.0#0"; "sg20o.ocx"
Begin VB.Form ScreenReport 
   BackColor       =   &H00B39665&
   Caption         =   "REPORTES POR PANTALLA"
   ClientHeight    =   8010
   ClientLeft      =   420
   ClientTop       =   450
   ClientWidth     =   11910
   LinkTopic       =   "Form1"
   ScaleHeight     =   8010
   ScaleWidth      =   11910
   Begin VB.ComboBox ComboReportes 
      BackColor       =   &H0080FFFF&
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      ItemData        =   "ScreenReport.frx":0000
      Left            =   120
      List            =   "ScreenReport.frx":0028
      Style           =   2  'Dropdown List
      TabIndex        =   11
      Top             =   480
      Width           =   4335
   End
   Begin VB.CommandButton cmdExpandir 
      Appearance      =   0  'Flat
      BackColor       =   &H0080C0FF&
      Caption         =   "Expandir"
      Default         =   -1  'True
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   10440
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   600
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.CommandButton cmdColapsar 
      Appearance      =   0  'Flat
      BackColor       =   &H0080C0FF&
      Caption         =   "Colapsar"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   10440
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   360
      Visible         =   0   'False
      Width           =   1335
   End
   Begin DDActiveReportsViewer2Ctl.ARViewer2 arv 
      Height          =   495
      Left            =   1680
      TabIndex        =   8
      Top             =   7440
      Width           =   495
      _ExtentX        =   873
      _ExtentY        =   873
      SectionData     =   "ScreenReport.frx":0128
   End
   Begin VB.CommandButton cmdEjec 
      Caption         =   "&Ejecutar Consulta"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   550
      Left            =   8160
      TabIndex        =   2
      Top             =   360
      Width           =   1935
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H0000FFFF&
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
      Left            =   120
      Picture         =   "ScreenReport.frx":0164
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   7440
      Width           =   1215
   End
   Begin VB.CommandButton Salir 
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
      Height          =   495
      Left            =   10560
      TabIndex        =   4
      Top             =   7440
      Width           =   1215
   End
   Begin DDSharpGridOLEDB2.SGGrid DD_PEDDETALLE 
      Height          =   6255
      Left            =   120
      TabIndex        =   5
      Top             =   1080
      Width           =   11655
      _cx             =   20558
      _cy             =   11033
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
      StylesCollection=   $"ScreenReport.frx":046E
      ColumnsCollection=   $"ScreenReport.frx":2241
      ValueItems      =   $"ScreenReport.frx":2756
   End
   Begin MSComCtl2.DTPicker txtFecIni 
      Height          =   345
      Left            =   4920
      TabIndex        =   0
      Top             =   480
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   609
      _Version        =   393216
      Format          =   117178369
      CurrentDate     =   36431
   End
   Begin MSComCtl2.DTPicker txtFecFin 
      Height          =   345
      Left            =   6600
      TabIndex        =   1
      Top             =   480
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   609
      _Version        =   393216
      Format          =   117178369
      CurrentDate     =   36430
   End
   Begin VB.Label Label1 
      BackColor       =   &H00B39665&
      Caption         =   "Seleccione un Reporte"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   2
      Left            =   120
      TabIndex        =   12
      Top             =   240
      Width           =   2895
   End
   Begin VB.Label Label1 
      BackColor       =   &H00B39665&
      Caption         =   "Fecha Inicial"
      BeginProperty Font 
         Name            =   "Verdana"
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
      Left            =   4920
      TabIndex        =   7
      Top             =   240
      Width           =   1575
   End
   Begin VB.Label Label1 
      BackColor       =   &H00B39665&
      Caption         =   "Fecha Final"
      BeginProperty Font 
         Name            =   "Verdana"
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
      Left            =   6600
      TabIndex        =   6
      Top             =   240
      Width           =   1455
   End
End
Attribute VB_Name = "ScreenReport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private iLin As Integer
Private nPagina As Integer
Private nOpcionSeleccionada As Integer
'OPCION DE COLORES DEL TITULOS (16ABR2011)
Private nHeadingBackColor As Variant
Private nCellBkgStyle  As Variant

Private Function GettINVENTDescription(nID As Integer, nCant As Long, nUnidades As Long) As String
'REGRESA LA CADENA CON EL NOMBRE, LA UNIDAD DE MEDIDA, LA UNIDAD DE CONSUMO
Dim rsPLUDescrip As ADODB.Recordset
Dim cSQL As String

Set rsPLUDescrip = New ADODB.Recordset
'GettINVENTDescription(rsDetalle!CODI_INV) & ";Cant: " & rsDetalle!CANT & ";Unidades: " & rsDetalle!UNIDADES
cSQL = "SELECT A.NOMBRE, B.DESCRIP AS MEDIDA, C.DESCRIP AS CONSUMO "
cSQL = cSQL & " FROM INVENT  AS A, UNIDADES AS B, UNID_CONSUMO AS C "
cSQL = cSQL & " WHERE A.ID = " & nID
cSQL = cSQL & " AND A.UNID_MEDIDA = B.ID "
cSQL = cSQL & " AND A.UNID_CONSUMO = C.ID "
rsPLUDescrip.Open cSQL, msConn, adOpenStatic, adLockReadOnly
If rsPLUDescrip.EOF Then
    GettINVENTDescription = "PRODUCTO NO ECONTRADO" & ";" & nCant & " ()" & ";" & nUnidades & " ()"
Else
    GettINVENTDescription = rsPLUDescrip!NOMBRE & ";" & nCant & " " & rsPLUDescrip!MEDIDA & ";" & nUnidades & " " & rsPLUDescrip!CONSUMO
End If
rsPLUDescrip.Close
Set rsPLUDescrip = Nothing
End Function

Private Sub cmdColapsar_Click()
DD_PEDDETALLE.CollapseAll
End Sub

Private Sub cmdEjec_Click()

If arv.Visible = True Then
    arv.Visible = False
    DD_PEDDETALLE.Visible = True
Else
End If

cmdColapsar.Visible = False
cmdExpandir.Visible = False

Select Case nOpcionSeleccionada
    Case 0      '= Eventos Administración
        Call EventosAdministracion
    Case 1      '= Eventos CAJA
        Call EventosCaja
    Case 2      '= Eventos de Autorización
        Call EventosAutorizacion
    Case 3      '= Reorden de Inventario
        Call ReodenInventario
    Case 4      '= Inventario Costo
        Call InventarioCosto
    Case 5      '= EMPAQUE / CONSUMO
        Call EmpaqueConsumo
'    Case 6      '= MENU vs INVENTARIO =====>>> QUE LO HAGAN EN EL REPORTE DE ENLACE CON INVENTARIO
        'Call Menu_VS_Inventario
    Case 6
        Call ListadoCompras
    Case 7
        cmdColapsar.Visible = True
        cmdExpandir.Visible = True
        Call Recetas_Inventario
    Case 8
        Call ListadoDescuentos
    Case 9
        Call ListadoProductosDescuentos
    Case 10
        Call ClientesEnMesa
    Case 11
        Call ListadoAcompanantes
End Select

Call Seguridad

End Sub

Private Sub Recetas_Inventario()
Dim rsEventos As ADODB.Recordset
Dim cSQL As String
Dim iLoop As Integer
'INFO: ABRIL 2011
'DD_PEDDETALLE.Rows.RemoveAll True
DD_PEDDETALLE.Columns.RemoveAll True

On Error GoTo ErrAdm:

cSQL = "SELECT A.NOMBRE & ' (Existencia: ' & FORMAT(A.EXIST2,'#.000') & ')' AS RECETA, "
cSQL = cSQL & "C.NOMBRE AS INVENTARIO, B.CANTIDAD_CONSUME AS Cant, D.DESCRIP AS Consumo, "
cSQL = cSQL & " B.COSTO AS Costo"
cSQL = cSQL & " FROM RECETAS AS A, RECETAS_INVENT AS B, INVENT AS C, UNID_CONSUMO AS D"
cSQL = cSQL & " WHERE A.ID = B.ID And B.ID_INVENT = C.ID And C.UNID_CONSUMO = d.ID"
cSQL = cSQL & " ORDER BY 1,2"

Set rsEventos = New ADODB.Recordset
rsEventos.Open cSQL, msConn, adOpenStatic, adLockOptimistic

If rsEventos.EOF Then
    'DO NOTHING
    ShowMsg "NO HAY REGISTROS PARA REPORTE DE RECETAS vs INVENTARIO"
Else

    With DD_PEDDETALLE
       .LoadArray rsEventos.GetRows()
       ' define each column from the recordsets' fields collection
       For iLoop = 1 To rsEventos.Fields.Count
          .Columns(iLoop).Caption = rsEventos.Fields(iLoop - 1).Name
          .Columns(iLoop).DBField = rsEventos.Fields(iLoop - 1).Name
          .Columns(iLoop).Key = rsEventos.Fields(iLoop - 1).Name
       Next iLoop

        .ColumnClickSort = True
        .EvenOddStyle = sgEvenOddRows
        .ColorEven = vbWhite
        .ColorOdd = &HE0E0E0
    
        .Columns(1).Width = 7400: .Columns(2).Width = 3900: .Columns(3).Width = 1200:
        .Columns(4).Width = 1300: .Columns(5).Width = 900
        .Columns(1).Style.TextAlignment = sgAlignLeftCenter
        .Columns(2).Style.TextAlignment = sgAlignLeftCenter
        .Columns(3).Style.TextAlignment = sgAlignRightCenter
        .Columns(4).Style.TextAlignment = sgAlignRightCenter
        .Columns(5).Style.TextAlignment = sgAlignRightCenter
        
        .Columns(4).SortType = sgSortTypeNumber
        .Columns(5).SortType = sgSortTypeNumber
        
         .Columns("Cant").Style.Format = "#0.000"
         .Columns("Costo").Style.Format = "#0.000"
        
         .Columns("RECETA").Hidden = True
        Dim grp1 As SGGroup, grp2 As SGGroup
        Set grp1 = .Groups.Add("RECETA", sgNoSorting, , False, False)
      
      ' Initialize 'RECETAS' group
        grp1.FetchHeaderStyle = True
        grp1.HeaderTextSource = sgGrpHdrColCaptionAndValue
        'grp1.HeaderTextSource = sgGrpHdrColCaptionAndValue
        grp1.FooterTextSource = sgGrpFooterCaption
        'grp1.FooterTextSource = sgGrpFooterFormula
        
        .RefreshGroups sgCollapseGroups
        
    End With
End If

Set rsEventos = Nothing

Exit Sub

ErrAdm:
    ShowMsg Err.Number & " - " & Err.Description
    On Error Resume Next
    If msPED.State = adStateOpen Then MesasPED "CLOSE"
    If rsEventos.State = adStateOpen Then
        rsEventos.Close
        Set rsEventos = Nothing
    End If
    On Error GoTo 0
End Sub

Private Sub cmdExpandir_Click()
DD_PEDDETALLE.ExpandAll
End Sub

Private Sub ComboReportes_Click()
'nOpcionSeleccionada = Index
nOpcionSeleccionada = ComboReportes.ListIndex
End Sub

Private Sub Command1_Click()
'MsgBox "LISTADOS UNICAMENTE SON VISIBLES POR PANTALLA", vbInformation, "Impresion de Listados no está disponible"
'MainMant.spDoc.DoPrintDirect
''''''''Dim i As Integer
''''''''
'''''''''On Error Resume Next
''''''''If MsgBox("¿ Desea imprimir este reporte ?", vbQuestion + vbYesNo, "Prepare el papel para imprimir") = vbYes Then
''''''''    nPagina = 0
''''''''    MainMant.spDoc.DocBegin
''''''''    Call PrintTit
''''''''
''''''''    For i = 0 To MSHFText.ListCount
''''''''        MainMant.spDoc.TextAlign = SPTA_LEFT
''''''''        MainMant.spDoc.TextOut 300, iLin, MSHFText.List(i)
''''''''        iLin = iLin + 50
''''''''        If iLin > 1900 Then Call PrintTit
''''''''    Next
''''''''    MainMant.spDoc.DoPrintPreview
''''''''End If
'''''''''On Error GoTo 0
Dim i As Integer

On Error GoTo ErrAdm:
    
    cmdColapsar.Visible = False
    cmdExpandir.Visible = False
    
    DD_PEDDETALLE.RedrawEnabled = True
    
    With DD_PEDDETALLE
        
         If .Visible Then
            .EvenOddStyle = sgNoEvenOdd
            .GridLines = sgGridLineNone
            nHeadingBackColor = .HeadingBackColor
            .HeadingBackColor = vbWhite
            nCellBkgStyle = .CellBkgStyle
            .CellBkgStyle = sgCellBkgNone
        End If
        'nGridBackColor = .Columns("RECETA").Style.grid.BackColor
        '.Columns("RECETA").Style.grid.BackColor = vbWhite
        '.Columns("RECETA").Style.grid.BackColor = vbWhite
        '.Columns("RECETA").Style.BackColor2 = vbWhite
        
    End With
    'DD_PEDDETALLE.Columns("RECETA").Style.BackColor = vbWhite
    'DD_PEDDETALLE.Columns("RECETA").Style.BackColor2 = vbWhite
    'DD_PEDDETALLE.Columns("RECETA").grid.BackColor = vbWhite
    'nBkgStyle = DD_PEDDETALLE.Columns("RECETA").Style.BkgStyle

    'DD_PEDDETALLE.Columns("RECETA").Style.BackColor = vbWhite
    'DD_PEDDETALLE.Columns("RECETA").Style.BkgStyle = sgCellBkgNone
    
   With DD_PEDDETALLE.PrintSettings
      
      .MarginBottom = 300
      .MarginLeft = 900
      .MarginRight = 900
      .MarginTop = 900

      .HeaderHeight = 750
      '.HeaderStyle.ForeColor = vbGreen
      .HeaderStyle.Font.Name = "Tahoma"
      .HeaderStyle.Font.Bold = True
      .HeaderStyle.Font.Size = 10
      .HeaderStyle.TextAlignment = sgAlignCenterCenter
      '.HeaderText = rs00!DESCRIP & vbCrLf & "Reporte de " & OptionRep(nOpcionSeleccionada).Caption
      .HeaderText = rs00!DESCRIP & vbCrLf & "Reporte de " & ComboReportes.Text
      
      .FooterHeight = 750
'      .FooterStyle.ForeColor = vbRed
      .FooterStyle.Font.Name = "Tahoma"
      .FooterStyle.Font.Bold = False
      .FooterStyle.Font.Size = 10
      .FooterText = "Fecha: " & Format(Date, "LONG DATE") & "           Hora: " & Format(Time, "LONG TIME")
'      '.FooterText = DD_PEDDETALLE.DataRowCount & " files"

      .MaxHorizontalPages = 1
      .CreateTOCFromGroups = True
      .RepeatColumnHeaders = True
      .RepeatFrozenColumns = False
      .RepeatFrozenRows = True
      .TranslateColors = False
      .TransparentBackground = False
   End With

   With DD_PEDDETALLE

      If .Visible Then
         Set .PrintSettings.Viewer = Me.Controls("arv").object
         
         .PrintSettings.PrintGrid
         
         .Visible = False
         
         Me.Controls("arv").Visible = True
         'For i = 0 To arv.Pages.Count - 1
         '   arv.Pages(i).Orientation = ddOLandscape
         'Next
      Else
        .EvenOddStyle = sgEvenOddRows
        .GridLines = sgGridLineFlat
        .HeadingBackColor = nHeadingBackColor
        .CellBkgStyle = nCellBkgStyle

         .Visible = True
         'For i = 0 To arv.Pages.Count - 1
         '   arv.Pages(i).Orientation = ddOPortrait
         'Next
         Me.Controls("arv").Visible = False
      End If
   End With
   
On Error GoTo 0
Exit Sub

ErrAdm:
ShowMsg Err.Number & " - " & Err.Description

End Sub

Private Sub DD_PEDDETALLE_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
If Button = vbRightButton Then
    Call ExportToExcelOrCSVList(DD_PEDDETALLE)
End If
End Sub

Private Sub Form_Load()
'SET OPTION
Dim iOpc As Integer
Dim vTemp As Variant


On Error Resume Next
 vTemp = RegRead("HKLM\Software\SoloSoftware\SoloAdmin\ReportSelection")
 If vTemp = "" Then vTemp = 0
 iOpc = CInt(vTemp)
On Error GoTo 0

On Error GoTo Form_Load_Error

'''Select Case iOpc
'''    Case 0, 4
'''        OptionRep(0).value = True
'''    Case 8
'''        OptionRep(7).value = True
'''    Case Else
'''        OptionRep(iOpc + 3).value = True
'''End Select

'SET FECHAS
txtFecIni = Format(Date, "SHORT DATE")
txtFecFin = Format(Date, "SHORT DATE")

Me.Controls("arv").Visible = False

Call Seguridad

cCuentaClientes = GetFromINI("Meseros", "CuentaClientes", DATA_PATH & "\soloini.ini")
If cCuentaClientes = "SI" Then
    Me.ComboReportes.AddItem "Clientes En Mesa"
End If

   On Error GoTo 0
   Exit Sub

Form_Load_Error:

    ShowMsg "Error " & Err.Number & " (" & Err.Description & ") En Formulario ScreenReport", vbRed, vbYellow

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
        Command1.Enabled = False
    Case "N"        'SIN DERECHOS
''        Dim i As Integer
''        For i = 0 To OptionRep.Count - 1
''            OptionRep(i).Enabled = False
''        Next
        ComboReportes.Enabled = False
        txtFecIni.Enabled = False: txtFecFin.Enabled = False: cmdEjec.Enabled = False
        Command1.Enabled = False
End Select
End Function

Private Sub Form_Resize()
    Command1.Top = Me.Height - 1120
    Salir.Top = Me.Height - 1120
    Salir.Left = Me.Width - 1600
    'AlignGrid Me, DD_PEDDETALLE, 120, 1400, 120, 740
    AlignGrid Me, DD_PEDDETALLE, 120, 1200, 120, 740
    arv.Move DD_PEDDETALLE.Left, DD_PEDDETALLE.Top, DD_PEDDETALLE.Width, DD_PEDDETALLE.Height
End Sub

'Private Sub OptionRep_Click(Index As Integer)
'nOpcionSeleccionada = Index
''0 = Eventos Administración
''1 = Eventos CAJA
''2 = Eventos de Autorización
''3= Reorden de Inventario
'End Sub

Private Sub Salir_Click()
Unload Me
End Sub

'''Private Sub PrintTit()
'''If nPagina = 0 Then
'''    MainMant.spDoc.WindowTitle = "Impresión de " & Me.Caption
'''    MainMant.spDoc.FirstPage = 1
'''    MainMant.spDoc.PageOrientation = SPOR_LANDSCAPE
'''    MainMant.spDoc.Units = SPUN_LOMETRIC
'''End If
'''MainMant.spDoc.Page = nPagina + 1
'''
'''MainMant.spDoc.TextOut 300, 200, Format(Date, "long date") & "  " & Time
'''MainMant.spDoc.TextOut 300, 250, "Página : " & nPagina + 1
'''MainMant.spDoc.TextOut 300, 300, rs00!DESCRIP
''''Informacion de Inventario. Este reporte esta guardado en la Base de Datos
''''Revision Inventario por Unidad
''''Revision Relacion Producto de Venta e Inventario"
''''Eventos del Sistema
'''Select Case nLst
'''    Case 0
'''        MainMant.spDoc.TextOut 300, 350, "Informacion de Inventario"
'''    Case 1
'''        MainMant.spDoc.TextOut 300, 350, "Revision Inventario por Unidad"
'''    Case 2
'''        MainMant.spDoc.TextOut 300, 350, "Revision Relacion Producto de Venta e Inventario"
'''    Case 3
'''        MainMant.spDoc.TextOut 300, 350, "Eventos del Sistema"
'''End Select
'''MainMant.spDoc.TextOut 300, 450, "--------------------------------------------------------------------------------------------------------------------------------"
'''
'''iLin = 500
'''nPagina = nPagina + 1
'''End Sub

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
Private Sub ListadoCompras()
Dim rsHeader As ADODB.Recordset
Dim rsDetalle As ADODB.Recordset
Dim cSQL As String
Dim iLoop As Integer
Dim dF1 As String
Dim dF2 As String
Dim cData As String
Dim nCompras As Single, nPagado As Single

'DD_PEDDETALLE.Rows.RemoveAll True
'DD_PEDDETALLE.Columns.RemoveAll True

Me.MousePointer = vbHourglass
dF1 = Format(txtFecIni, "YYYYMMDD")
dF2 = Format(txtFecFin, "YYYYMMDD")

On Error GoTo ErrAdm:

Set rsHeader = New ADODB.Recordset
Set rsDetalle = New ADODB.Recordset

cSQL = "SELECT A.NUMERO, A.COD_PROV, A.MONTO, A.USUARIO, A.PAGADO, A.FECHA, A.INDICE,"
cSQL = cSQL & " B.NOMBRE & ' - ' & B.APELLIDO & ' (' & B.EMPRESA & ')' AS PROVEEDOR,"
cSQL = cSQL & " C.NOMBRE & SPACE(1) & C.APELLIDO AS USUARIO, A.TIPO "
'cSQL = cSQL & " FROM COMPRAS_HEAD AS A LEFT JOIN USUARIOS AS C ON A.USUARIO =  C.NUMERO , PROVEEDORES AS B "
cSQL = cSQL & " FROM COMPRAS_HEAD AS A, USUARIOS AS C, PROVEEDORES AS B "
cSQL = cSQL & " WHERE A.FECHA BETWEEN '" & dF1 & "' AND '" & dF2 & "'"
cSQL = cSQL & " AND A.COD_PROV = B.CODIGO "
cSQL = cSQL & " AND A.USUARIO =  C.NUMERO "
cSQL = cSQL & " ORDER BY A.FECHA, A.INDICE"

DD_PEDDETALLE.DataMode = sgUnbound
DD_PEDDETALLE.DataRowCount = 0
DD_PEDDETALLE.DataColCount = 5
DD_PEDDETALLE.RedrawEnabled = False
DD_PEDDETALLE.AutoResizeHeadings = True

rsHeader.Open cSQL, msConn, adOpenStatic, adLockReadOnly
Do While Not rsHeader.EOF
    
    If rsHeader!TIPO = "CR" Then
        cData = "PROVEEDOR: " & rsHeader!PROVEEDOR & ";" & " " & GetFecha(rsHeader!FECHA) & ";NUM:" & rsHeader!NUMERO
        cData = cData & ";Pend.: " & Format(rsHeader!MONTO, "CURRENCY") & ";"
    Else
        cData = "PROVEEDOR: " & rsHeader!PROVEEDOR & ";" & " " & GetFecha(rsHeader!FECHA) & ";NUM:" & rsHeader!NUMERO
        cData = cData & ";" & ";" & "Pagado: " & Format(rsHeader!PAGADO, "STANDARD")
    End If
    DD_PEDDETALLE.Rows.Add sgFormatCharSeparatedValue, cData, ";"
    
    nCompras = nCompras + rsHeader!MONTO
    nPagado = nPagado + rsHeader!PAGADO
    
    If rsHeader!TIPO = "CR" Then
        cData = "USUARIO: " & rsHeader!USUARIO & ";" & ";" & ";" & "CREDITO"
    Else
        cData = "USUARIO: " & rsHeader!USUARIO & ";" & ";" & ";" & ";" & "EFECTIVO"
    End If
    DD_PEDDETALLE.Rows.Add sgFormatCharSeparatedValue, cData, ";"

    cSQL = "SELECT LINEA, CODI_INV, CANT, UNIDADES, COSTO_UNIT, ITBM, COSTO_IN"
    cSQL = cSQL & " FROM COMPRAS_DETA "
    cSQL = cSQL & " WHERE NUM_COMPRA = '" & rsHeader!indice & "'"
    cSQL = cSQL & " ORDER BY LINEA"
    
    '////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
    cData = String(30, Chr(126)) & ";" & String(10, Chr(126)) & ";" & String(10, Chr(126)) & ";" & String(13, Chr(126)) & ";" & String(13, Chr(126))
    DD_PEDDETALLE.Rows.Add sgFormatCharSeparatedValue, cData, ";"
    
    rsDetalle.Open cSQL, msConn, adOpenStatic, adLockReadOnly
    Do While Not rsDetalle.EOF
        
        cData = GettINVENTDescription(rsDetalle!CODI_INV, rsDetalle!CANT, rsDetalle!UNIDADES)
        cData = cData & ";Costo Unit: " & Format(rsDetalle!COSTO_UNIT, "#0.0000") & ";Costo Tot: " & Format(rsDetalle!COSTO_IN, "STANDARD")
        DD_PEDDETALLE.Rows.Add sgFormatCharSeparatedValue, cData, ";"
        
        rsDetalle.MoveNext
    Loop
    rsDetalle.Close
    cData = String(30, Chr(126)) & ";" & String(10, Chr(126)) & ";" & String(10, Chr(126)) & ";" & String(13, Chr(126)) & ";" & String(13, Chr(126))
    DD_PEDDETALLE.Rows.Add sgFormatCharSeparatedValue, cData, ";"
    'cData = String(32, Chr(61)) & ";" & String(10, Chr(61)) & ";" & String(10, Chr(61)) & ";" & String(14, Chr(61)) & ";" & String(14, Chr(61))
    cData = String(32, Space(1)) & ";" & String(10, Space(1)) & ";" & String(10, Space(1)) & ";" & String(14, Space(1)) & ";" & String(14, Space(1))
    DD_PEDDETALLE.Rows.Add sgFormatCharSeparatedValue, cData, ";"
    cData = String(32, Space(1)) & ";" & String(10, Space(1)) & ";" & String(10, Space(1)) & ";" & String(14, Space(1)) & ";" & String(14, Space(1))
    DD_PEDDETALLE.Rows.Add sgFormatCharSeparatedValue, cData, ";"
    '////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
    
    rsHeader.MoveNext
Loop

'cData = "Comprado: " & Format(nCompras, "CURRENCY") & " - Pagado: " & Format(nPagado, "CURRENCY")
cData = "Compras a Crédito: " & Format((nCompras - nPagado), "CURRENCY")
DD_PEDDETALLE.Rows.Add sgFormatCharSeparatedValue, cData, ";"
cData = "Compras  Efectivo : " & Format(nPagado, "CURRENCY")
DD_PEDDETALLE.Rows.Add sgFormatCharSeparatedValue, cData, ";"
'INFO: 06NOV2010. agregar por suguerencia de La Scala
cData = "Total de Compras del Periodo: " & Format(nCompras, "CURRENCY")
DD_PEDDETALLE.Rows.Add sgFormatCharSeparatedValue, cData, ";"
rsHeader.Close

Set rsHeader = Nothing
Set rsDetalle = Nothing

DD_PEDDETALLE.Columns(1).Width = 4800: DD_PEDDETALLE.Columns(1).Caption = "Proovedor/Usuario"
DD_PEDDETALLE.Columns(2).Width = 1600: DD_PEDDETALLE.Columns(2).Caption = "Fecha"
DD_PEDDETALLE.Columns(3).Width = 1600: DD_PEDDETALLE.Columns(3).Caption = "# Orden"
'DD_PEDDETALLE.Columns(4).Width = 2100: DD_PEDDETALLE.Columns(4).Caption = "Monto Total"
'DD_PEDDETALLE.Columns(5).Width = 2100: DD_PEDDETALLE.Columns(5).Caption = "Monto Pagado"
DD_PEDDETALLE.Columns(4).Width = 2100: DD_PEDDETALLE.Columns(4).Caption = "Crédito"
DD_PEDDETALLE.Columns(5).Width = 2100: DD_PEDDETALLE.Columns(5).Caption = "Contado"

DD_PEDDETALLE.EvenOddStyle = sgEvenOddRows
DD_PEDDETALLE.AllowAddNew = False
DD_PEDDETALLE.ColumnClickSort = False
DD_PEDDETALLE.RedrawEnabled = True

Me.MousePointer = vbDefault

Exit Sub

ErrAdm:
    ShowMsg Err.Number & " - " & Err.Description
    On Error Resume Next
    If rsHeader.State = adStateOpen Then
        rsHeader.Close
        Set rsHeader = Nothing
    End If
    On Error GoTo 0
    Me.MousePointer = vbDefault
    Resume
End Sub

Private Sub Menu_VS_Inventario()
Dim rsEventos As ADODB.Recordset
Dim cSQL As String
Dim iLoop As Integer
Dim dF1 As String
Dim dF2 As String

'DD_PEDDETALLE.Rows.RemoveAll True
DD_PEDDETALLE.Columns.RemoveAll True
dF1 = Format(txtFecIni, "YYYYMMDD")
dF2 = Format(txtFecFin, "YYYYMMDD")

On Error GoTo ErrAdm:

cSQL = "SELECT B.DESCRIP AS DEPTO, C.DESCRIP AS PRODUCTO_VENTA, "
cSQL = cSQL & "  D.DESCRIP AS DEPT_INV, E.NOMBRE AS ARTICULO_INVENT, "
cSQL = cSQL & "  CANT, F.DESCRIP AS CONSUMO"
cSQL = cSQL & "  FROM PLU_INVENT AS A, DEPTO AS B, PLU AS C, "
        cSQL = cSQL & "  DEP_INV AS D, INVENT AS E, UNID_CONSUMO AS F"
cSQL = cSQL & "  WHERE A.ID_DEPT = B.CODIGO AND A.ID_PLU=C.CODIGO "
cSQL = cSQL & "  AND A.ID_DEPT_INV=D.CODIGO AND A.ID_PROD_INV=E.ID "
cSQL = cSQL & "  AND A.ID_UNID_CONSUMO = F.ID"

Set rsEventos = New ADODB.Recordset
rsEventos.Open cSQL, msConn, adOpenStatic, adLockOptimistic

'rsEventos.Fields.Append "Envase", adChar, 20
If rsEventos.EOF Then
    'DO NOTHING
    ShowMsg "NO HAY REGISTROS PARA REPORTE MENU vs INVENTARIO"
Else

    With DD_PEDDETALLE
       .LoadArray rsEventos.GetRows()
       ' define each column from the recordsets' fields collection
       For iLoop = 1 To rsEventos.Fields.Count
          .Columns(iLoop).Caption = rsEventos.Fields(iLoop - 1).Name
          .Columns(iLoop).DBField = rsEventos.Fields(iLoop - 1).Name
          .Columns(iLoop).Key = rsEventos.Fields(iLoop - 1).Name
       Next iLoop

        .ColumnClickSort = True
        .EvenOddStyle = sgEvenOddRows
        .ColorEven = vbWhite
        .ColorOdd = &HE0E0E0
    
        .Columns(1).Width = 2600: .Columns(2).Width = 3300: .Columns(3).Width = 2500:
        .Columns(4).Width = 2000: .Columns(5).Width = 800: .Columns(6).Width = 1500
        '.Columns(7).Width = 1200:
        .Columns(1).Style.TextAlignment = sgAlignLeftCenter
        .Columns(2).Style.TextAlignment = sgAlignLeftCenter
        .Columns(3).Style.TextAlignment = sgAlignLeftCenter
        .Columns(4).Style.TextAlignment = sgAlignLeftCenter
        .Columns(5).Style.TextAlignment = sgAlignRightCenter
        .Columns(6).Style.TextAlignment = sgAlignRightCenter
        '.Columns(7).Style.TextAlignment = sgAlignRightCenter
        '.Redraw sgRedrawAll
    End With
End If

Set rsEventos = Nothing

Exit Sub

ErrAdm:
    ShowMsg Err.Number & " - " & Err.Description
    On Error Resume Next
    If msPED.State = adStateOpen Then MesasPED "CLOSE"
    If rsEventos.State = adStateOpen Then
        rsEventos.Close
        Set rsEventos = Nothing
    End If
    On Error GoTo 0
End Sub
Private Sub EmpaqueConsumo()
Dim rsEventos As ADODB.Recordset
Dim cSQL As String
Dim iLoop As Integer
Dim dF1 As String
Dim dF2 As String

'DD_PEDDETALLE.Rows.RemoveAll True
DD_PEDDETALLE.Columns.RemoveAll True
dF1 = Format(txtFecIni, "YYYYMMDD")
dF2 = Format(txtFecFin, "YYYYMMDD")

On Error GoTo ErrAdm:

cSQL = "SELECT B.DESCRIP AS DEPTO, A.NOMBRE, C.DESCRIP AS UNID_COMPRA, "
cSQL = cSQL & " A.CANTIDAD AS CANT_COMPRA, D.DESCRIP AS UNID_CONSUMO, "
cSQL = cSQL & " A.CANTIDAD2 AS CANT_CONSUMO "
cSQL = cSQL & " FROM INVENT AS A, DEP_INV AS B, UNIDADES AS C, UNID_CONSUMO AS D "
cSQL = cSQL & " WHERE A.COD_DEPT = B.CODIGO And A.UNID_MEDIDA = C.ID And A.UNID_CONSUMO = D.ID "
cSQL = cSQL & "  ORDER BY 1, 2 "


Set rsEventos = New ADODB.Recordset
rsEventos.Open cSQL, msConn, adOpenStatic, adLockOptimistic

If rsEventos.EOF Then
    'DO NOTHING
    ShowMsg "NO HAY REGISTROS PARA REPORTE EMPAQUE vs CONSUMO"
Else

    With DD_PEDDETALLE
       .LoadArray rsEventos.GetRows()
       ' define each column from the recordsets' fields collection
       For iLoop = 1 To rsEventos.Fields.Count
          .Columns(iLoop).Caption = rsEventos.Fields(iLoop - 1).Name
          .Columns(iLoop).DBField = rsEventos.Fields(iLoop - 1).Name
          .Columns(iLoop).Key = rsEventos.Fields(iLoop - 1).Name
       Next iLoop

        .ColumnClickSort = True
        .EvenOddStyle = sgEvenOddRows
        .ColorEven = vbWhite
        .ColorOdd = &HE0E0E0
    
        .Columns(1).Width = 3300: .Columns(2).Width = 3300: .Columns(3).Width = 1500:
        .Columns(4).Width = 1600: .Columns(5).Width = 1700: .Columns(6).Width = 1700:
        '.Columns(7).Width = 1200:
        .Columns(1).Style.TextAlignment = sgAlignLeftCenter
        .Columns(2).Style.TextAlignment = sgAlignLeftCenter
        .Columns(3).Style.TextAlignment = sgAlignRightCenter
        .Columns(4).Style.TextAlignment = sgAlignRightCenter
        .Columns(5).Style.TextAlignment = sgAlignRightCenter
        .Columns(6).Style.TextAlignment = sgAlignRightCenter
        .Columns(6).SortType = sgSortTypeNumber
        '.Columns(7).Style.TextAlignment = sgAlignRightCenter
        '.Redraw sgRedrawAll
    End With
End If

Set rsEventos = Nothing

Exit Sub

ErrAdm:
    ShowMsg Err.Number & " - " & Err.Description
    On Error Resume Next
    If msPED.State = adStateOpen Then MesasPED "CLOSE"
    If rsEventos.State = adStateOpen Then
        rsEventos.Close
        Set rsEventos = Nothing
    End If
    On Error GoTo 0
End Sub
Private Sub InventarioCosto()
Dim rsEventos As ADODB.Recordset
Dim cSQL As String
Dim iLoop As Integer
Dim dF1 As String
Dim dF2 As String

'DD_PEDDETALLE.Rows.RemoveAll True
DD_PEDDETALLE.Columns.RemoveAll True
dF1 = Format(txtFecIni, "YYYYMMDD")
dF2 = Format(txtFecFin, "YYYYMMDD")

On Error GoTo ErrAdm:

cSQL = "SELECT C.DESCRIP AS DEPTO, a.Nombre, b.Descrip AS Envase, "
cSQL = cSQL & " a.Cantidad AS Cant_Env, a.Exist1 as Exist_1, "
cSQL = cSQL & " FORMAT(a.Exist2,'#,###.00') AS Exist_2, "
cSQL = cSQL & " FORMAT(a.COSTO_EMPAQUE,'STANDARD') as Costo, a.ITBM"
cSQL = cSQL & " FROM invent AS a, unidades AS b, DEP_INV AS C"
cSQL = cSQL & " Where A.UNID_MEDIDA = b.id And A.COD_DEPT = C.CODIGO"
cSQL = cSQL & " ORDER BY C.DESCRIP, b.descrip, a.nombre"


Set rsEventos = New ADODB.Recordset
rsEventos.Open cSQL, msConn, adOpenStatic, adLockOptimistic

If rsEventos.EOF Then
    'DO NOTHING
    ShowMsg "NO HAY REGISTROS PARA REPORTE DE INVENTARIO vs COSTO"
Else

    With DD_PEDDETALLE
       .LoadArray rsEventos.GetRows()
       ' define each column from the recordsets' fields collection
       For iLoop = 1 To rsEventos.Fields.Count
          .Columns(iLoop).Caption = rsEventos.Fields(iLoop - 1).Name
          .Columns(iLoop).DBField = rsEventos.Fields(iLoop - 1).Name
          .Columns(iLoop).Key = rsEventos.Fields(iLoop - 1).Name
       Next iLoop

        .ColumnClickSort = True
        .EvenOddStyle = sgEvenOddRows
        .ColorEven = vbWhite
        .ColorOdd = &HE0E0E0
    
        .Columns(1).Width = 2200: .Columns(2).Width = 3300: .Columns(3).Width = 1200:
        .Columns(4).Width = 1300: .Columns(5).Width = 900: .Columns(6).Width = 1200:
        .Columns(7).Width = 1200:
        .Columns(1).Style.TextAlignment = sgAlignLeftCenter
        .Columns(2).Style.TextAlignment = sgAlignLeftCenter
        .Columns(3).Style.TextAlignment = sgAlignRightCenter
        .Columns(4).Style.TextAlignment = sgAlignRightCenter
        .Columns(5).Style.TextAlignment = sgAlignRightCenter
        .Columns(6).Style.TextAlignment = sgAlignRightCenter
        .Columns(7).Style.TextAlignment = sgAlignRightCenter
        .Columns(8).Style.TextAlignment = sgAlignRightCenter
        
        .Columns(4).SortType = sgSortTypeNumber
        .Columns(5).SortType = sgSortTypeNumber
        .Columns(6).SortType = sgSortTypeNumber
        .Columns(7).SortType = sgSortTypeNumber
        .Columns(8).SortType = sgSortTypeNumber
        '.Redraw sgRedrawAll
    End With
End If

Set rsEventos = Nothing

Exit Sub

ErrAdm:
    ShowMsg Err.Number & " - " & Err.Description
    On Error Resume Next
    If msPED.State = adStateOpen Then MesasPED "CLOSE"
    If rsEventos.State = adStateOpen Then
        rsEventos.Close
        Set rsEventos = Nothing
    End If
    On Error GoTo 0
End Sub
Private Sub ReodenInventario()
Dim rsEventos As ADODB.Recordset
Dim cSQL As String
Dim iLoop As Integer
Dim dF1 As String
Dim dF2 As String
Dim i As Long
Dim nST As Single

'DD_PEDDETALLE.Rows.RemoveAll True
DD_PEDDETALLE.Columns.RemoveAll True
dF1 = Format(txtFecIni, "YYYYMMDD")
dF2 = Format(txtFecFin, "YYYYMMDD")

On Error GoTo ErrAdm:

cSQL = "SELECT LEFT(B.DESCRIP,13) AS DEPTO, A.NOMBRE, FORMAT(A.EXIST2,'STANDARD') AS EN_STOCK, "
cSQL = cSQL & " A.NIV_REORDEN & ' - ' & C.DESCRIP AS REORDEN, "
cSQL = cSQL & " FORMAT(A.COSTO_EMPAQUE,'STANDARD')  AS COSTO, "
cSQL = cSQL & " IIF(A.EXIST2 < 0,  IIF(A.NIV_REORDEN=0,1,A.NIV_REORDEN) & ' - ' & C.DESCRIP, IIF(A.NIV_REORDEN=0,1,A.NIV_REORDEN) & ' - ' & C.DESCRIP) AS TO_ORDER, "
'cSQL = cSQL & " FORMAT((A.NIV_REORDEN - A.EXIST2) * 1.15,'#,###.00') AS TO_ORDER, "
'cSQL = cSQL & " FORMAT((A.COSTO * (A.NIV_REORDEN - A.EXIST2) * 1.15),'CURRENCY') AS MONTO "
cSQL = cSQL & " FORMAT((A.COSTO_EMPAQUE * IIF(A.NIV_REORDEN=0,1,A.NIV_REORDEN)),'CURRENCY') AS MONTO "
cSQL = cSQL & " FROM INVENT AS A, DEP_INV AS B, UNIDADES AS C "
'cSQL = cSQL & " WHERE A.NIV_REORDEN > 0 "
cSQL = cSQL & " WHERE A.UNID_MEDIDA = C.ID "
cSQL = cSQL & " AND A.NIV_REORDEN  * (A.CANTIDAD2) > A.EXIST2 "
cSQL = cSQL & " AND A.cod_dept = B.CODIGO "
cSQL = cSQL & " ORDER BY B.DESCRIP, A.NOMBRE"

Set rsEventos = New ADODB.Recordset
rsEventos.Open cSQL, msConn, adOpenStatic, adLockOptimistic

If rsEventos.EOF Then
    'DO NOTHING
    ShowMsg "NO HAY REGISTROS PARA REPORTE DE RE-ORDEN DE PRODUCTOS"
Else

    With DD_PEDDETALLE
       .LoadArray rsEventos.GetRows()
       ' define each column from the recordsets' fields collection
       For iLoop = 1 To rsEventos.Fields.Count
          .Columns(iLoop).Caption = rsEventos.Fields(iLoop - 1).Name
          .Columns(iLoop).DBField = rsEventos.Fields(iLoop - 1).Name
          .Columns(iLoop).Key = rsEventos.Fields(iLoop - 1).Name
       Next iLoop
        
        .ColumnClickSort = True
        .EvenOddStyle = sgEvenOddRows
        .ColorEven = vbWhite
        .ColorOdd = &HE0E0E0
    
        .Columns(1).Width = 1800: .Columns(2).Width = 3300: .Columns(3).Width = 1200:
        .Columns(4).Width = 1500: .Columns(5).Width = 900: .Columns(6).Width = 1200:
        .Columns(7).Width = 1400: '.Columns(8).Width = 1000:
        .Columns(1).Style.TextAlignment = sgAlignLeftCenter
        .Columns(2).Style.TextAlignment = sgAlignLeftCenter
        .Columns(3).Style.TextAlignment = sgAlignRightCenter
        .Columns(4).Style.TextAlignment = sgAlignRightCenter
        .Columns(5).Style.TextAlignment = sgAlignRightCenter
        .Columns(6).Style.TextAlignment = sgAlignRightCenter
        .Columns(7).Style.TextAlignment = sgAlignRightCenter
        
        .Columns(3).SortType = sgSortTypeNumber
        .Columns(4).SortType = sgSortTypeNumber
        .Columns(5).SortType = sgSortTypeNumber
        .Columns(6).SortType = sgSortTypeNumber
        .Columns(7).SortType = sgSortTypeNumber

    End With
End If

rsEventos.MoveFirst
Do While Not rsEventos.EOF
    nST = nST + rsEventos!MONTO
    rsEventos.MoveNext
Loop

cData = ";" & Space(15) & "TOTAL GENERAL " & ";;;;;" & Format(nST, "CURRENCY")
DD_PEDDETALLE.Rows.InsertAt DD_PEDDETALLE.Rows.Count + 1, sgFormatCharSeparatedValue, cData, ";"

DD_PEDDETALLE.TopRow = 0

Set rsEventos = Nothing

Exit Sub

ErrAdm:
    ShowMsg Err.Number & " - " & Err.Description
    On Error Resume Next
    If msPED.State = adStateOpen Then MesasPED "CLOSE"
    If rsEventos.State = adStateOpen Then
        rsEventos.Close
        Set rsEventos = Nothing
    End If
    On Error GoTo 0
End Sub
Private Sub EventosAutorizacion()
Dim rsEventos As ADODB.Recordset
Dim cSQL As String
Dim iLoop As Integer
Dim dF1 As String
Dim dF2 As String

'DD_PEDDETALLE.Rows.RemoveAll True
DD_PEDDETALLE.Columns.RemoveAll True
dF1 = Format(txtFecIni, "YYYYMMDD")
dF2 = Format(txtFecFin, "YYYYMMDD")

On Error GoTo ErrAdm:

MesasPED "OPEN"
Set rsEventos = New ADODB.Recordset

cSQL = "SELECT  RIGHT(FECHA,2) + '/' + MID(FECHA,5,2) + '/' + LEFT(FECHA,4)  AS FECHA_1, "
cSQL = cSQL & " HORA, DESCRIPCION "
cSQL = cSQL & " FROM LOG "
cSQL = cSQL & " WHERE LEFT(DESCRIPCION,4) "
'INFO. SE AGREGA CORRECCION DEL SISTEMA FAST
cSQL = cSQL & " IN ('ANUL','DESC','CORT','EXON','CIER','ABON','PAG0','ABRI','CORR')"
cSQL = cSQL & " AND FECHA Between '" & dF1 & "'"
cSQL = cSQL & " AND '" & dF2 & "'"
cSQL = cSQL & " ORDER BY FECHA DESC, HORA DESC"

rsEventos.Open cSQL, msPED, adOpenStatic, adLockOptimistic

With DD_PEDDETALLE
    
    .LoadArray rsEventos.GetRows()
       ' define each column from the recordsets' fields collection
        For iLoop = 1 To rsEventos.Fields.Count
           .Columns(iLoop).Caption = rsEventos.Fields(iLoop - 1).Name
           .Columns(iLoop).DBField = rsEventos.Fields(iLoop - 1).Name
           .Columns(iLoop).Key = rsEventos.Fields(iLoop - 1).Name
        Next iLoop
    
    .ColumnClickSort = True
    .EvenOddStyle = sgEvenOddRows
    .ColorEven = vbWhite
        
    .ColorOdd = &HE0E0E0
    .Columns(1).Width = 1400: .Columns(2).Width = 1200: .Columns(3).Width = 8500:
    .Columns(1).Style.TextAlignment = sgAlignLeftCenter
    .Columns(2).Style.TextAlignment = sgAlignLeftCente
    .Columns(3).Style.TextAlignment = sgAlignLeftCente
    
End With

MesasPED "CLOSE"
Set rsEventos = Nothing

Exit Sub

ErrAdm:
    ShowMsg Err.Number & " - " & Err.Description
    On Error Resume Next
    If msPED.State = adStateOpen Then MesasPED "CLOSE"
    If rsEventos.State = adStateOpen Then
        rsEventos.Close
        Set rsEventos = Nothing
    End If
    On Error GoTo 0
End Sub

Private Sub EventosCaja()
Dim rsEventos As ADODB.Recordset
Dim cSQL As String
Dim iLoop As Integer
Dim dF1 As String
Dim dF2 As String

'DD_PEDDETALLE.Rows.RemoveAll True
DD_PEDDETALLE.Columns.RemoveAll True
dF1 = Format(txtFecIni, "YYYYMMDD")
dF2 = Format(txtFecFin, "YYYYMMDD")

On Error GoTo ErrAdm:

MesasPED "OPEN"
Set rsEventos = New ADODB.Recordset

cSQL = "SELECT  RIGHT(FECHA,2) + '/' + MID(FECHA,5,2) + '/' + LEFT(FECHA,4)  AS FECHA, "
cSQL = cSQL & " HORA, DESCRIPCION "
cSQL = cSQL & " FROM LOG "
cSQL = cSQL & " WHERE FECHA Between '" & dF1 & "'"
cSQL = cSQL & " AND '" & dF2 & "'"
'INFO: 18FEB2011
cSQL = cSQL & " AND LEFT(DESCRIPCION,5) <> 'Admin'"
cSQL = cSQL & " ORDER BY FECHA DESC, HORA DESC"

rsEventos.Open cSQL, msPED, adOpenStatic, adLockOptimistic
'rsEventos.Filter = " DESCRIPCION NOT LIKE 'Admin*'"
With DD_PEDDETALLE
    
    .LoadArray rsEventos.GetRows()
       ' define each column from the recordsets' fields collection
        For iLoop = 1 To rsEventos.Fields.Count
           .Columns(iLoop).Caption = rsEventos.Fields(iLoop - 1).Name
           .Columns(iLoop).DBField = rsEventos.Fields(iLoop - 1).Name
           .Columns(iLoop).Key = rsEventos.Fields(iLoop - 1).Name
        Next iLoop
    
    .ColumnClickSort = True
    .EvenOddStyle = sgEvenOddRows
    .ColorEven = vbWhite
        
    .ColorOdd = &HE0E0E0
    .Columns(1).Width = 1700: .Columns(2).Width = 1200: .Columns(3).Width = 7500:
    .Columns(1).Style.TextAlignment = sgAlignLeftCenter
    .Columns(2).Style.TextAlignment = sgAlignLeftCente
    .Columns(3).Style.TextAlignment = sgAlignLeftCente
    
End With
MesasPED "CLOSE"
Set rsEventos = Nothing

Exit Sub

ErrAdm:
    ShowMsg Err.Number & " - " & Err.Description
    On Error Resume Next
    If msPED.State = adStateOpen Then MesasPED "CLOSE"
    If rsEventos.State = adStateOpen Then
        rsEventos.Close
        Set rsEventos = Nothing
    End If
    On Error GoTo 0
End Sub
Private Sub EventosAdministracion()
Dim rsEventos As ADODB.Recordset
Dim cSQL As String
Dim iLoop As Integer
Dim dF1 As String
Dim dF2 As String

'INFO: 18FEB2011
DD_PEDDETALLE.Columns.RemoveAll True
dF1 = Format(txtFecIni, "YYYYMMDD")
dF2 = Format(txtFecFin, "YYYYMMDD")

On Error GoTo ErrAdm:

MesasPED "OPEN"
Set rsEventos = New ADODB.Recordset

cSQL = "SELECT  RIGHT(FECHA,2) + '/' + MID(FECHA,5,2) + '/' + LEFT(FECHA,4)  AS FECHA, "
cSQL = cSQL & " HORA, DESCRIPCION "
cSQL = cSQL & " FROM LOG "
cSQL = cSQL & " WHERE FECHA Between '" & dF1 & "'"
cSQL = cSQL & " AND '" & dF2 & "'"
cSQL = cSQL & " AND LEFT(DESCRIPCION,5) = 'Admin' "
cSQL = cSQL & " ORDER BY FECHA DESC, HORA DESC"

rsEventos.Open cSQL, msPED, adOpenStatic, adLockOptimistic

With DD_PEDDETALLE
    
    .LoadArray rsEventos.GetRows()
       ' define each column from the recordsets' fields collection
        For iLoop = 1 To rsEventos.Fields.Count
           .Columns(iLoop).Caption = rsEventos.Fields(iLoop - 1).Name
           .Columns(iLoop).DBField = rsEventos.Fields(iLoop - 1).Name
           .Columns(iLoop).Key = rsEventos.Fields(iLoop - 1).Name
        Next iLoop
    
    .ColumnClickSort = True
    .EvenOddStyle = sgEvenOddRows
    .ColorEven = vbWhite
        
    .ColorOdd = &HE0E0E0
    .Columns(1).Width = 1700: .Columns(2).Width = 1200: .Columns(3).Width = 7500:
    .Columns(1).Style.TextAlignment = sgAlignLeftCenter
    .Columns(2).Style.TextAlignment = sgAlignLeftCente
    .Columns(3).Style.TextAlignment = sgAlignLeftCente
    
End With
MesasPED "CLOSE"
Set rsEventos = Nothing

Exit Sub

ErrAdm:
    ShowMsg Err.Number & " - " & Err.Description
    On Error Resume Next
    If msPED.State = adStateOpen Then MesasPED "CLOSE"
    If rsEventos.State = adStateOpen Then
        rsEventos.Close
        Set rsEventos = Nothing
    End If
    On Error GoTo 0
End Sub

Private Sub Old_EventosAdministracion()
Dim rsEventos As ADODB.Recordset
Dim cSQL As String
Dim iLoop As Integer
Dim dF1 As String
Dim dF2 As String

'DD_PEDDETALLE.Rows.RemoveAll True
DD_PEDDETALLE.Columns.RemoveAll True
dF1 = Format(txtFecIni, "YYYYMMDD")
dF2 = Format(txtFecFin, "YYYYMMDD")

On Error GoTo ErrAdm:

With DD_PEDDETALLE
    .ImportData "C:\WINDOWS\ADMLOG.SOL", sgFormatCharSeparatedValue, vbTab
    .ColumnClickSort = True
    .EvenOddStyle = sgEvenOddRows
    .ColorEven = vbWhite

    .Columns(1).Caption = "FECHA"
    .Columns(2).Caption = "USUARIO"
    .Columns(3).Caption = "EVENTO"

    .ColorOdd = &HE0E0E0
    .Columns(1).Width = 2600: .Columns(2).Width = 1700: .Columns(3).Width = 6500:
    .Columns(1).Style.TextAlignment = sgAlignLeftCenter
    .Columns(2).Style.TextAlignment = sgAlignLeftCente
    .Columns(3).Style.TextAlignment = sgAlignLeftCente
End With

Exit Sub


ErrAdm:
    
    If Err.Number = -2147024894 Then
        ShowMsg Err.Number & " - " & Err.Description & " (ADMLOG.SOL)"
        Exit Sub
    Else
        ShowMsg Err.Number & " - " & Err.Description & " (ADMLOG.SOL)"
    End If
    
    On Error Resume Next
    If msPED.State = adStateOpen Then MesasPED "CLOSE"
    If rsEventos.State = adStateOpen Then
        rsEventos.Close
        Set rsEventos = Nothing
    End If
    On Error GoTo 0
End Sub
'---------------------------------------------------------------------------------------
' Procedure : ListadoDescuentos
' Author    : hsequeira
' Date      : 20/02/2016
' Purpose   : LISTA LOS DESCUENTOS MARCADOS
'---------------------------------------------------------------------------------------
'
Private Sub ListadoDescuentos()
Dim rsDescuentos As ADODB.Recordset
Dim cSQL As String
Dim iLoop As Integer
Dim dF1 As String
Dim dF2 As String
Dim cData As String
Dim nTotCant As Long
Dim nTotDesc As Double

'DD_PEDDETALLE.Rows.RemoveAll True
'DD_PEDDETALLE.Columns.RemoveAll True

Me.MousePointer = vbHourglass
DD_PEDDETALLE.Columns.RemoveAll True
dF1 = Format(txtFecIni, "YYYYMMDD")
dF2 = Format(txtFecFin, "YYYYMMDD")

On Error GoTo ErrAdm:

Set rsDescuentos = New ADODB.Recordset

cSQL = "SELECT LEFT(FECHA,6) AS YEAR_MES, DESCRIP AS TIPO_DESCUENTO, COUNT(PRECIO) AS CANT_DESCUENTOS, "
cSQL = cSQL & " ABS(SUM(PRECIO)) AS TOTAL_DESCUENTOS"
cSQL = cSQL & " FROM HIST_TR"
cSQL = cSQL & " WHERE FECHA BETWEEN '" & dF1 & "' AND '" & dF2 & "'"
'cSQL = cSQL & " AND DESCRIP LIKE '%DESCUENTO%'"
cSQL = cSQL & " AND ID_DESCUENTO > 0 "
cSQL = cSQL & " AND VALID AND PRECIO_UNIT < 0"
cSQL = cSQL & " GROUP BY LEFT(FECHA,6), DESCRIP"

DD_PEDDETALLE.DataMode = sgUnbound

rsDescuentos.Open cSQL, msConn, adOpenStatic, adLockOptimistic
'rsDescuentos.Filter = " DESCRIPCION NOT LIKE 'Admin*'"
With DD_PEDDETALLE
    
    .LoadArray rsDescuentos.GetRows()
       ' define each column from the recordsets' fields collection
        For iLoop = 1 To rsDescuentos.Fields.Count
           .Columns(iLoop).Caption = rsDescuentos.Fields(iLoop - 1).Name
           .Columns(iLoop).DBField = rsDescuentos.Fields(iLoop - 1).Name
           .Columns(iLoop).Key = rsDescuentos.Fields(iLoop - 1).Name
        Next iLoop
    
    .ColumnClickSort = True
    .EvenOddStyle = sgEvenOddRows
    .ColorEven = vbWhite
        
    .ColorOdd = &HE0E0E0
    .Columns(1).Width = 1200: .Columns(2).Width = 5000: .Columns(3).Width = 2000:
    .Columns(4).Width = 2100
    .Columns(1).Style.TextAlignment = sgAlignLeftCenter
    .Columns(2).Style.TextAlignment = sgAlignLeftCente
    .Columns(3).Style.TextAlignment = sgAlignCenterCenter
    .Columns(4).Style.TextAlignment = sgAlignRightCenter
    .Columns(4).Style.Format = "Standard"

End With

rsDescuentos.Close

cSQL = "SELECT DESCRIP AS TIPO_DESCUENTO, COUNT(PRECIO) AS CANT_DESCUENTOS, "
cSQL = cSQL & " ABS(SUM(PRECIO)) AS TOTAL_DESCUENTOS"
cSQL = cSQL & " FROM HIST_TR"
cSQL = cSQL & " WHERE FECHA BETWEEN '" & dF1 & "' AND '" & dF2 & "'"
'cSQL = cSQL & " AND DESCRIP LIKE '%DESCUENTO%'"
'INFO: 30JUN2018
cSQL = cSQL & " AND ID_DESCUENTO > 0 "
cSQL = cSQL & " AND VALID AND PRECIO_UNIT < 0"
cSQL = cSQL & " GROUP BY DESCRIP"
cSQL = cSQL & " ORDER BY DESCRIP"

rsDescuentos.Open cSQL, msConn, adOpenStatic, adLockOptimistic

rsDescuentos.MoveFirst

cData = "======="
cData = cData & ";================"
cData = cData & ";============="
cData = cData & ";============="
DD_PEDDETALLE.Rows.Add sgFormatCharSeparatedValue, cData, ";"

Do While Not rsDescuentos.EOF
    cData = "   Totales"
    cData = cData & ";<<" & rsDescuentos!TIPO_DESCUENTO & ">>"
    cData = cData & ";" & Format(rsDescuentos!CANT_DESCUENTOS, "#,###")
    cData = cData & ";" & Format(rsDescuentos!TOTAL_DESCUENTOS, "STANDARD")
    DD_PEDDETALLE.Rows.Add sgFormatCharSeparatedValue, cData, ";"

    nTotCant = nTotCant + rsDescuentos!CANT_DESCUENTOS
    nTotDesc = nTotDesc + rsDescuentos!TOTAL_DESCUENTOS

    rsDescuentos.MoveNext
Loop

cData = "Total Global"
cData = cData & ";" & txtFecIni.value & " - " & txtFecFin.value & ";" & Format(nTotCant, "#,###") & ";" & Format(nTotDesc, "STANDARD")
DD_PEDDETALLE.Rows.Add sgFormatCharSeparatedValue, cData, ";"

rsDescuentos.Close
Set rsDescuentos = Nothing
Me.MousePointer = vbDefault

Exit Sub

ErrAdm:
    ShowMsg Err.Number & " - " & Err.Description, vbRed, vbYellow
    On Error Resume Next
    'If rsDescuentos.State = adStateOpen Then
    '    rsDescuentos.Close
    '    Set rsDescuentos = Nothing
    'End If
    On Error GoTo 0
    Me.MousePointer = vbDefault
End Sub

'---------------------------------------------------------------------------------------
' Procedure : ListadoProductosDescuentos
' Author    : hsequeira
' Date      : 20/02/2016
' Date      : 30/06/2018
' Purpose   : MUESTRA PRODUCTOS CON MAS DESCUENTOS SEGUN EL PERIODO
' UPDATE: MUESTRA LOS PRODUCTOS MARCADOS EN CADA TIPO DE DESCUENTO.
'---------------------------------------------------------------------------------------
'
Private Sub ListadoProductosDescuentos()
Dim rsDescuentos As ADODB.Recordset
Dim cSQL As String
Dim iLoop As Integer
Dim dF1 As String
Dim dF2 As String
Dim cData As String
Dim nTotCant As Long
Dim nTotDesc As Double
Dim nTotDescGEN As Double
Dim oTemp As String
Dim rsSUBTotales As ADODB.Recordset

'DD_PEDDETALLE.Rows.RemoveAll True
'DD_PEDDETALLE.Columns.RemoveAll True

Me.MousePointer = vbHourglass
DD_PEDDETALLE.Columns.RemoveAll True
dF1 = Format(txtFecIni, "YYYYMMDD")
dF2 = Format(txtFecFin, "YYYYMMDD")

On Error GoTo ErrAdm:

Set rsDescuentos = New ADODB.Recordset

'cSQL = "SELECT B.DESCRIP AS PRODUCTO,"
'cSQL = cSQL & " COUNT(A.PRECIO) AS CANT_DESCUENTOS,"
'cSQL = cSQL & " ABS(SUM(A.PRECIO)) AS TOTAL_DESCUENTOS"
'cSQL = cSQL & " FROM HIST_TR AS A, PLU AS B"
'cSQL = cSQL & " WHERE A.PLU = B.CODIGO"
'cSQL = cSQL & " AND A.FECHA BETWEEN '" & dF1 & "' AND '" & dF2 & "'"
'cSQL = cSQL & " AND A.DESCRIP LIKE '%DESCUENTO%' AND A.VALID AND A.PRECIO_UNIT < 0"
'cSQL = cSQL & " GROUP BY B.DESCRIP"
'cSQL = cSQL & " ORDER BY 3 DESC"

cSQL = "SELECT VAL(D.FISCAL) AS FISCAL, FORMAT(A.FECHA_TRANS,'####-##-##') AS FECHA, "
cSQL = cSQL & "FORMAT(A.HORA_TRANS,'SHORT TIME') AS HORA, A.ID_DESCUENTO, "
cSQL = cSQL & "C.DESCRIP & ' (' & C.PORCENTAJE & ' %)' AS DESCUENTO_DESCRIP, "
cSQL = cSQL & "B.DESCRIP AS PRODUCTO, A.PRECIO, A.NUM_TRANS "
cSQL = cSQL & "FROM HIST_TR AS A, PLU AS B, DESCUENTO  AS C, TRANSAC_FISCAL AS D "
cSQL = cSQL & "WHERE A.PLU = B.CODIGO AND A.ID_DESCUENTO = C.TIPO "
cSQL = cSQL & "AND A.FECHA BETWEEN '" & dF1 & "' AND '" & dF2 & "'"
cSQL = cSQL & " AND A.DESCUENTO <> 0 AND A.VALID "
cSQL = cSQL & "AND A.NUM_TRANS = D.DOC_SOLO "
cSQL = cSQL & "ORDER BY C.DESCRIP, A.NUM_TRANS"

DD_PEDDETALLE.DataMode = sgUnbound

rsDescuentos.Open cSQL, msConn, adOpenStatic, adLockOptimistic
'rsDescuentos.Filter = " DESCRIPCION NOT LIKE 'Admin*'"
With DD_PEDDETALLE
    
    .LoadArray rsDescuentos.GetRows()
       ' define each column from the recordsets' fields collection
        For iLoop = 1 To rsDescuentos.Fields.Count
           .Columns(iLoop).Caption = rsDescuentos.Fields(iLoop - 1).Name
           .Columns(iLoop).DBField = rsDescuentos.Fields(iLoop - 1).Name
           .Columns(iLoop).Key = rsDescuentos.Fields(iLoop - 1).Name
        Next iLoop
    
    .ColumnClickSort = False
    .EvenOddStyle = sgEvenOddRows
    .ColorEven = vbWhite
        
    .ColorOdd = &HE0E0E0
    .Columns(1).Width = 1000: .Columns(2).Width = 1400: .Columns(3).Width = 1000:
    .Columns(4).Width = 0:
    .Columns(5).Width = 3200: .Columns(6).Width = 3200: .Columns(7).Width = 1000:
    .Columns(1).Style.TextAlignment = sgAlignLeftCenter
    .Columns(2).Style.TextAlignment = sgAlignLeftCenter
    .Columns(3).Style.TextAlignment = sgAlignLeftCenter
    '.Columns(2).Style.Format = "#,###"
    '.Columns(3).Style.TextAlignment = sgAlignLeftCenter
    '.Columns(3).Style.Format = "Standard"
    .Columns(5).Style.TextAlignment = sgAlignLeftCenter
    .Columns(6).Style.TextAlignment = sgAlignLeftCenter
    .Columns(7).Style.TextAlignment = sgAlignRightCenter
    .Columns(7).Style.Format = "Standard"
    .Columns(8).Width = 0:
End With

cData = ";;;;" & "  Periodo Seleccionado: " & ";"
cData = cData & txtFecIni.value & " - " & txtFecFin.value & ";"
'DD_PEDDETALLE.Rows.Add sgFormatCharSeparatedValue, cData, ";"
DD_PEDDETALLE.Rows.InsertAt 0, sgFormatCharSeparatedValue, cData, ";"


DD_PEDDETALLE.Row = 2
DD_PEDDETALLE.Col = 4
oTemp = DD_PEDDETALLE.CurrentCell.value
i = 2
'For i = 2 To DD_PEDDETALLE.RowCount - 1
Do While Not DD_PEDDETALLE.EOF
    DD_PEDDETALLE.Col = 4
    DD_PEDDETALLE.Row = i
    
    If i > DD_PEDDETALLE.Row Then
        Exit Do
    End If
  
    If oTemp = DD_PEDDETALLE.CurrentCell.value Then
        DD_PEDDETALLE.Col = 6
        nTotDesc = nTotDesc + DD_PEDDETALLE.CurrentCell.value
    Else
        oTemp = DD_PEDDETALLE.CurrentCell.value
        cData = ";;;;;" & Space(30) & "SUB TOTAL " & ";" & Format(nTotDesc, "STANDARD")

        DD_PEDDETALLE.Rows.InsertAt DD_PEDDETALLE.Row, sgFormatCharSeparatedValue, cData, ";"
        'DD_PEDDETALLE.Col = 5
        'DD_PEDDETALLE.CurrentCell.Style.Font.Bold = True
        'DD_PEDDETALLE.CurrentCell.Style.Font.Bold = False
        'DD_PEDDETALLE.Col = 4
        nTotDescGEN = nTotDescGEN + nTotDesc
        nTotDesc = 0
    End If
    i = i + 1
Loop
'Next

cData = ";;;;;" & Space(30) & "SUB TOTAL " & ";" & Format(nTotDesc, "STANDARD")
DD_PEDDETALLE.Rows.InsertAt i, sgFormatCharSeparatedValue, cData, ";"
i = i + 1
cData = ";;;;;" & Space(25) & "TOTAL GENERAL " & ";" & Format(nTotDescGEN, "STANDARD")
DD_PEDDETALLE.Rows.InsertAt i, sgFormatCharSeparatedValue, cData, ";"

cSQL = "SELECT C.PORCENTAJE, SUM(A.PRECIO) AS MONTO_DESCUENTO "
cSQL = cSQL & "FROM HIST_TR AS A, DESCUENTO  AS C "
cSQL = cSQL & "WHERE A.ID_DESCUENTO = C.TIPO "
cSQL = cSQL & "AND A.FECHA BETWEEN '" & dF1 & "' AND '" & dF2 & "'"
cSQL = cSQL & " AND A.DESCUENTO <> 0 AND A.VALID "
cSQL = cSQL & "GROUP BY C.PORCENTAJE"

Set rsSUBTotales = New ADODB.Recordset
rsSUBTotales.Open cSQL, msConn, adOpenStatic, adLockOptimistic

i = i + 1
Do While Not rsSUBTotales.EOF
    i = i + 1
    cData = ";;;;;" & Space(5) & "TOTAL DESCUENTO (" & rsSUBTotales!PORCENTAJE & " %)" & ";" & Format(rsSUBTotales!MONTO_DESCUENTO, "STANDARD")
    DD_PEDDETALLE.Rows.InsertAt i, sgFormatCharSeparatedValue, cData, ";"
    rsSUBTotales.MoveNext
Loop

DD_PEDDETALLE.TopRow = 0

rsDescuentos.Close
Set rsDescuentos = Nothing

rsSUBTotales.Close
Set rsSUBTotales = Nothing

Me.MousePointer = vbDefault

Exit Sub

ErrAdm:
    ShowMsg Err.Number & " - " & Err.Description, vbRed, vbYellow
    On Error Resume Next
    'If rsDescuentos.State = adStateOpen Then
    '    rsDescuentos.Close
    '    Set rsDescuentos = Nothing
    'End If
    On Error GoTo 0
    Me.MousePointer = vbDefault
End Sub

'---------------------------------------------------------------------------------------
' Procedure : ClientesEnMesa
' Author    : hsequeira
' Date      : 25/02/2022
' Purpose   : REPORTE DE LOS CLIENTES QUE ESTAN SIENDO ATENDIDOS
'---------------------------------------------------------------------------------------
'
Private Sub ClientesEnMesa()
Dim rsEventos As ADODB.Recordset
Dim rsEventos2 As ADODB.Recordset
Dim cSQL As String
Dim iLoop As Integer
Dim dF1 As String
Dim dF2 As String

'INFO: 12DIC2021
DD_PEDDETALLE.Columns.RemoveAll True
dF1 = Format(txtFecIni, "YYYYMMDD")
dF2 = Format(txtFecFin, "YYYYMMDD")

On Error GoTo ErrAdm:

'MesasPED "OPEN"
Set rsEventos = New ADODB.Recordset

cSQL = "SELECT  RIGHT(FECHA,2) + '/' + MID(FECHA,5,2) + '/' + LEFT(FECHA,4)  AS FECHA, "
cSQL = cSQL & " MESA, CLIENTES "
cSQL = cSQL & " FROM CLIENTES_COUNTER "
cSQL = cSQL & " WHERE FECHA Between '" & dF1 & "'"
cSQL = cSQL & " AND '" & dF2 & "'"
cSQL = cSQL & " ORDER BY FECHA ASC, MESA ASC"

rsEventos.Open cSQL, msConn, adOpenStatic, adLockOptimistic

With DD_PEDDETALLE
    
    .LoadArray rsEventos.GetRows()
       ' define each column from the recordsets' fields collection
        For iLoop = 1 To rsEventos.Fields.Count
           .Columns(iLoop).Caption = rsEventos.Fields(iLoop - 1).Name
           .Columns(iLoop).DBField = rsEventos.Fields(iLoop - 1).Name
           .Columns(iLoop).Key = rsEventos.Fields(iLoop - 1).Name
        Next iLoop
    
    .ColumnClickSort = True
    .EvenOddStyle = sgEvenOddRows
    .ColorEven = vbWhite
        
    .ColorOdd = &HE0E0E0
    .Columns(1).Width = 1700: .Columns(2).Width = 3000: .Columns(3).Width = 1900:
    .Columns(1).Style.TextAlignment = sgAlignLeftCenter
    .Columns(2).Style.TextAlignment = sgAlignCenterCenter
    .Columns(3).Style.TextAlignment = sgAlignCenterCenter
    
End With
'MesasPED "CLOSE"

Set rsEventos2 = New ADODB.Recordset

cSQL = "SELECT MESA, "
cSQL = cSQL & " ABS(SUM(CLIENTES)) AS TOTAL_CLIENTES"
cSQL = cSQL & " FROM CLIENTES_COUNTER"
cSQL = cSQL & " WHERE FECHA BETWEEN '" & dF1 & "' AND '" & dF2 & "'"
cSQL = cSQL & " GROUP BY MESA"
cSQL = cSQL & " ORDER BY MESA"

rsEventos2.Open cSQL, msConn, adOpenStatic, adLockOptimistic

rsEventos2.MoveFirst

cData = "======="
cData = cData & ";================"
cData = cData & ";============="
cData = cData & ";============="
DD_PEDDETALLE.Rows.Add sgFormatCharSeparatedValue, cData, ";"

Do While Not rsEventos2.EOF
    cData = "   Totales"
    cData = cData & ";<< MESA " & rsEventos2!MESA & " >>"
    cData = cData & ";" & Format(rsEventos2!TOTAL_CLIENTES, "###,###")
    DD_PEDDETALLE.Rows.Add sgFormatCharSeparatedValue, cData, ";"

    nTotCant = nTotCant + rsEventos2!TOTAL_CLIENTES
    'nTotDesc = nTotDesc + rsDescuentos!TOTAL_DESCUENTOS

    rsEventos2.MoveNext
Loop

cData = "Total Global"
cData = cData & ";" & txtFecIni.value & " - " & txtFecFin.value & ";" & Format(nTotCant, "#,###")
DD_PEDDETALLE.Rows.Add sgFormatCharSeparatedValue, cData, ";"

'cData = "Total Global"
'cData = cData & ";" & txtFecIni.value & " - " & txtFecFin.value & ";" & Format(nTotCant, "#,###") & ";" & Format(nTotDesc, "###,###")
'DD_PEDDETALLE.Rows.Add sgFormatCharSeparatedValue, cData, ";"



rsEventos.Close
rsEventos2.Close

Set rsEventos = Nothing
Set rsEventos2 = Nothing
Me.MousePointer = vbDefault


Exit Sub

ErrAdm:
    ShowMsg Err.Number & " - " & Err.Description
    On Error Resume Next
    'If msPED.State = adStateOpen Then MesasPED "CLOSE"
    If rsEventos.State = adStateOpen Then
        rsEventos.Close
        Set rsEventos = Nothing
    End If
    On Error GoTo 0

End Sub

'---------------------------------------------------------------------------------------
' Procedure : ListadoAcompanates
' Author    : hsequeira
' Date      : 11/10/2022
' Purpose   : LISTA LOS ACOMPANANTES
'---------------------------------------------------------------------------------------
'
Private Sub ListadoAcompanantes()
Dim rsAcompanates As ADODB.Recordset
Dim cSQL As String
Dim iLoop As Integer
Dim dF1 As String
Dim dF2 As String
Dim cData As String
Dim nTotCant As Long
Dim nTotDesc As Double

'DD_PEDDETALLE.Rows.RemoveAll True
'DD_PEDDETALLE.Columns.RemoveAll True

Me.MousePointer = vbHourglass
DD_PEDDETALLE.Columns.RemoveAll True
dF1 = Format(txtFecIni, "YYYYMMDD")
dF2 = Format(txtFecFin, "YYYYMMDD")

On Error GoTo ErrAdm:

Set rsAcompanates = New ADODB.Recordset



cSQL = "SELECT DESCRIP as ACOMPAÑANTE, COUNT(*) AS CANTIDAD "
cSQL = cSQL & " FROM HIST_TR "
cSQL = cSQL & " WHERE FECHA BETWEEN '" & dF1 & "' AND '" & dF2 & "'"
cSQL = cSQL & " AND LEFT(DESCRIP,3) = ' @@' "
cSQL = cSQL & " AND VALID "
cSQL = cSQL & " GROUP BY DESCRIP ORDER BY 2 DESC"

DD_PEDDETALLE.DataMode = sgUnbound

rsAcompanates.Open cSQL, msConn, adOpenStatic, adLockOptimistic
'rsAcompanates.Filter = " DESCRIPCION NOT LIKE 'Admin*'"
With DD_PEDDETALLE
    
    .LoadArray rsAcompanates.GetRows()
       ' define each column from the recordsets' fields collection
        For iLoop = 1 To rsAcompanates.Fields.Count
           .Columns(iLoop).Caption = rsAcompanates.Fields(iLoop - 1).Name
           .Columns(iLoop).DBField = rsAcompanates.Fields(iLoop - 1).Name
           .Columns(iLoop).Key = rsAcompanates.Fields(iLoop - 1).Name
        Next iLoop
    
    .ColumnClickSort = True
    .EvenOddStyle = sgEvenOddRows
    .ColorEven = vbWhite
        
    .ColorOdd = &HE0E0E0
    .Columns(1).Width = 5000: .Columns(2).Width = 1200:
    .Columns(1).Style.TextAlignment = sgAlignLeftCenter
    .Columns(2).Style.TextAlignment = sgAlignRightCenter
'    .Columns(3).Style.TextAlignment = sgAlignCenterCenter
'    .Columns(4).Style.TextAlignment = sgAlignRightCenter
'    .Columns(4).Style.Format = "Standard"

End With

rsAcompanates.Close
Set rsAcompanates = Nothing
Me.MousePointer = vbDefault

Exit Sub

ErrAdm:
    ShowMsg Err.Number & " - " & Err.Description, vbRed, vbYellow
    On Error Resume Next
    'If rsAcompanates.State = adStateOpen Then
    '    rsAcompanates.Close
    '    Set rsAcompanates = Nothing
    'End If
    On Error GoTo 0
    Me.MousePointer = vbDefault
End Sub

