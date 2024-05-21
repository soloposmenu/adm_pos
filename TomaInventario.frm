VERSION 5.00
Object = "{A45D986F-3AAF-4A3B-A003-A6C53E8715A2}#1.0#0"; "ARVIEW2.OCX"
Object = "{43135020-B751-4DDD-AB4A-B6D8A342216E}#1.0#0"; "sg20o.ocx"
Begin VB.Form TomaInventario 
   BackColor       =   &H00B39665&
   Caption         =   "Toma de Inventario"
   ClientHeight    =   8685
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   15120
   Icon            =   "TomaInventario.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   8685
   ScaleWidth      =   15120
   StartUpPosition =   2  'CenterScreen
   Begin VB.ListBox LV_DEPTO 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1260
      Left            =   2640
      TabIndex        =   5
      ToolTipText     =   "Seleccione el Departamento Especifico o TODOS LOS DEPARTAMENTOS"
      Top             =   240
      Width           =   3255
   End
   Begin VB.CommandButton cmdPrintInvent 
      BackColor       =   &H0000FFFF&
      Height          =   495
      Left            =   12600
      Picture         =   "TomaInventario.frx":030A
      Style           =   1  'Graphical
      TabIndex        =   3
      ToolTipText     =   "Envia Seleccion a la Impresora"
      Top             =   240
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.CommandButton cmdUpdateInvent 
      Caption         =   "Actualiza CAMBIOS al Inventario"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   7440
      TabIndex        =   2
      Top             =   240
      Visible         =   0   'False
      Width           =   4575
   End
   Begin VB.CommandButton cmdBuscaInventario 
      Caption         =   "Lista Inventario"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   240
      TabIndex        =   1
      Top             =   240
      Width           =   2175
   End
   Begin DDSharpGridOLEDB2.SGGrid DD_PEDDETALLE 
      Height          =   6855
      Left            =   120
      TabIndex        =   0
      Top             =   1680
      Width           =   14895
      _cx             =   26273
      _cy             =   12091
      DataMember      =   ""
      DataMode        =   0
      AutoFields      =   -1  'True
      Enabled         =   -1  'True
      GridBorderStyle =   1
      ScrollBars      =   3
      FlatScrollBars  =   0
      ScrollBarTrack  =   0   'False
      DataRowCount    =   2
      BeginProperty HeadingFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
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
         Name            =   "Tahoma"
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
      AllowEdit       =   -1  'True
      ScrollBarTips   =   0
      CellTips        =   0
      CellTipsDelay   =   1000
      SpecialMode     =   0
      OutlineLines    =   1
      CacheAllRecords =   0   'False
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
      AutoGroup       =   0   'False
      GroupByBoxVisible=   0   'False
      GroupByBoxText  =   "Drag a column header here to group by that column"
      AlphaBlendEnabled=   0   'False
      DragAlphaLevel  =   206
      AutoSearch      =   0
      AutoSearchDelay =   2000
      OutlineIcons    =   1
      CellTipsDisplayLength=   3000
      StylesCollection=   $"TomaInventario.frx":0614
      ColumnsCollection=   $"TomaInventario.frx":23E5
      ValueItems      =   $"TomaInventario.frx":31A8
   End
   Begin DDActiveReportsViewer2Ctl.ARViewer2 arv 
      Height          =   495
      Left            =   13800
      TabIndex        =   4
      Top             =   240
      Width           =   495
      _ExtentX        =   873
      _ExtentY        =   873
      SectionData     =   "TomaInventario.frx":358D
   End
   Begin VB.Label Label1 
      BackColor       =   &H00B39665&
      Caption         =   "Ahora Puede Ordenar el listado haciendo click en la columna deseada"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFF00&
      Height          =   375
      Left            =   6240
      TabIndex        =   7
      Top             =   1200
      Width           =   8295
   End
   Begin VB.Label lbFisicoReal 
      BackColor       =   &H00B39665&
      Caption         =   "** El fisico Real es el valor del producto en cantidades fisicas"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   375
      Left            =   6240
      TabIndex        =   6
      Top             =   840
      Width           =   7455
   End
End
Attribute VB_Name = "TomaInventario"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private cFile As String
Private nOpc As Integer
Private nDeptoSeleccionado As Integer
Private Sub ImprimirGrid()
On Error GoTo ErrAdm:
    
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
        
    End With
    
   With DD_PEDDETALLE.PrintSettings
'      .MarginBottom = 300
'      .MarginLeft = 900
'      .MarginRight = 900
'      .MarginTop = 900
    
      .HeaderHeight = 750
      .HeaderStyle.Font.Name = "Tahoma"
      .HeaderStyle.Font.Bold = True
      .HeaderStyle.Font.Size = 10
      .HeaderStyle.TextAlignment = sgAlignCenterCenter
      .HeaderText = rs00!DESCRIP & vbCrLf & "Reporte de " & Me.Caption
      
      .FooterHeight = 750
'      .FooterStyle.ForeColor = vbRed
      .FooterStyle.Font.Name = "Tahoma"
      .FooterStyle.Font.Bold = False
      .FooterStyle.Font.Size = 10
      .FooterText = "Fecha: " & Format(Date, "LONG DATE") & "           Hora: " & Format(Time, "LONG TIME")
      '.FooterText = DD_PEDDETALLE.DataRowCount & " files"

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
      Else
        '.EvenOddStyle = sgEvenOddRows
        '.GridLines = sgGridLineFlat
        '.HeadingBackColor = nHeadingBackColor
        '.CellBkgStyle = nCellBkgStyle

         .Visible = True
         Me.Controls("arv").Visible = False
      End If
   End With
On Error GoTo 0

Exit Sub

ErrAdm:
ShowMsg Err.Number & " - " & Err.Description

End Sub
Private Sub cmdBuscaInventario_Click()
    Call BuscaInventario("OK")
End Sub
'---------------------------------------------------------------------------------------
' Procedure : BuscaInventario
' Author    : hsequeira
' Date      : 02/09/2015
' Purpose   :
' - ACTUALIZANDO EL DISPLAY PARA QUE NO INCLUYA EXISTENCIA1, QUE NO SE USA EN EL SISTEMA
' - PERMITE AHORA ORDENAR POR COLUMNAS
' - EL VALOR QUE SE DEBE INTRODUCIR ES EL VALOR DE LA CANTIDAD EN EXISTENCIA ENTRE EL VALOR DE CANTIDAD2 DEFINIDA EN INVENTARIO
'---------------------------------------------------------------------------------------
'
Private Function BuscaInventario(cTipo As String)

Dim cSQL As String
Dim rsTomaInv As ADODB.Recordset
Dim iLoop As Integer

If arv.Visible = True Then
    arv.Visible = False
    DD_PEDDETALLE.Visible = True
Else
End If

DD_PEDDETALLE.Columns.RemoveAll True

Set rsTomaInv = New ADODB.Recordset

'cSQL = "SELECT A.ID, A.COD_DEPT, B.DESCRIP AS DEPARTAMENTO, A.NOMBRE, "
cSQL = "SELECT A.ID, A.COD_DEPT, B.DESCRIP & ' - ' & A.NOMBRE AS DEPTO_PRODUCTO, "
'cSQL = cSQL & " C.DESCRIP AS COMPRA, D.DESCRIP AS CONSUMO, "
cSQL = cSQL & " C.DESCRIP AS COMPRA, A.CANTIDAD2 & ' ' & D.DESCRIP AS CONSUMO, "
If cTipo = "OK" Then
    'cSQL = cSQL & " A.EXIST1, A.EXIST1 AS FISICO1, A.EXIST2 , A.EXIST2 AS FISICO2, SPACE(50) AS COMENTARIO "
    'cSQL = cSQL & " A.EXIST2/A.CANTIDAD2 AS EXIST2, A.EXIST2 AS FISICO2, SPACE(50) AS COMENTARIO "
    'cSQL = cSQL & " A.EXIST2/A.CANTIDAD2 AS EXIST2, A.EXIST2 AS FISICO2, SPACE(50) AS COMENTARIO "
    cSQL = cSQL & " A.EXIST2 AS Exist_Consumo, "
    cSQL = cSQL & " IIF(D.DESCRIP = 'UNIDAD' OR D.DESCRIP='Unidad',FORMAT(A.EXIST2,'#0.000'), FORMAT(A.EXIST2/A.CANTIDAD2,'#0.000')) AS Exist_Real, "
    cSQL = cSQL & " IIF(D.DESCRIP = 'UNIDAD' OR D.DESCRIP='Unidad',FORMAT(A.EXIST2,'#0.000'), FORMAT(A.EXIST2/A.CANTIDAD2,'#0.000')) AS Fisico_Real, "
    cSQL = cSQL & " SPACE(50) AS COMENTARIO, A.CANTIDAD2, D.DESCRIP AS DESCRIP_CONSUMO "
Else
    'cSQL = cSQL & " A.EXIST1, '__________' AS FISICO1, A.EXIST2, '__________' AS FISICO2, '______________________________' AS COMENTARIO "
    'cSQL = cSQL & " A.EXIST2/A.CANTIDAD2 AS EXIST2, '__________' AS FISICO2, '______________________________' AS COMENTARIO "
    cSQL = cSQL & " A.EXIST2 AS Exist_Consumo, A.EXIST2/A.CANTIDAD2 AS Exist_Real, "
    cSQL = cSQL & " '__________' AS Fisico_Real, '______________________________' AS COMENTARIO, A.CANTIDAD2, D.DESCRIP AS DESCRIP_CONSUMO "
End If
cSQL = cSQL & " FROM INVENT AS A, DEP_INV AS B, UNIDADES AS C, UNID_CONSUMO AS D"
cSQL = cSQL & " WHERE A.COD_DEPT = B.CODIGO AND A.UNID_MEDIDA = C.ID "
cSQL = cSQL & " AND A.UNID_CONSUMO = D.ID "
If nDeptoSeleccionado = -1 Then
Else
    cSQL = cSQL & " AND A.COD_DEPT = " & nDeptoSeleccionado
End If
cSQL = cSQL & " ORDER BY 3,4"

rsTomaInv.Open cSQL, msConn, adOpenDynamic, adLockReadOnly

If rsTomaInv.EOF Then
    ShowMsg "NO HAY REGISTROS EN LA TABLA DE INVENTARIO", vbYellow, vbRed
    rsTomaInv.Close
    Set rsTomaInv = Nothing
    Exit Function
End If

DD_PEDDETALLE.DataMode = sgUnbound
DD_PEDDETALLE.ColumnClickSort = True
'Set DD_PEDDETALLE.DataSource = rsTomaInv
With DD_PEDDETALLE
    
    .LoadArray rsTomaInv.GetRows()
       ' define each column from the recordsets' fields collection
       For iLoop = 1 To rsTomaInv.Fields.Count
          .Columns(iLoop).Caption = rsTomaInv.Fields(iLoop - 1).Name
          .Columns(iLoop).DBField = rsTomaInv.Fields(iLoop - 1).Name
          .Columns(iLoop).Key = rsTomaInv.Fields(iLoop - 1).Name
       Next iLoop
        
        .AllowEdit = True
        .Columns(1).Hidden = True   'ID
        .Columns(2).Hidden = True   'COD_DEPT
        .Columns(10).Hidden = True   'CANTIDAD2
        .Columns(11).Hidden = True   'DESCRIP_CONSUMO
        '.Columns(7).Style.BackColor = vbYellow
        '.Columns(9).Style.BackColor = vbYellow
        .Columns(8).Style.BackColor = vbYellow
        '.Columns(10).Style.BackColor = vbGreen
        .Columns(9).Style.BackColor = vbGreen
        
        .Columns(6).Style.TextAlignment = sgAlignRightCenter
        .Columns(8).Style.TextAlignment = sgAlignRightCenter
        
        '.Columns(3).Width = 1800: .Columns(4).Width = 3700: .Columns(5).Width = 1500: .Columns(6).Width = 1100:
        .Columns(3).Width = 5000    'DEPTO_PRODUCTO
        .Columns(4).Width = 1500    'COMPRA
        .Columns(5).Width = 1100    'CONSUMO
        .Columns(6).Width = 1300    'EXIST_CONSUMO
        .Columns(7).Width = 1000    'EXIST_REAL
        .Columns(8).Width = 1200    'FISICO_REAL
        .Columns(9).Width = 3100    'COMENTARIO
        
        .Columns(2).ReadOnly = True
        .Columns(3).ReadOnly = True
        .Columns(4).ReadOnly = True
        .Columns(5).ReadOnly = True
        .Columns(6).ReadOnly = True
        .Columns(7).ReadOnly = True
        
        .Columns(8).ReadOnly = False
        .Columns(9).ReadOnly = False
         
         '.Columns("EXIST1").Style.Format = "#0.000"
         .Columns("Exist_Consumo").Style.Format = "#0.000"
         .Columns("Exist_Real").Style.Format = "#0.000"

        If cTipo = "OK" Then
            '.Columns("FISICO1").Style.Format = "#0.000"
            .Columns("Fisico_Real").Style.Format = "#0.000"
            .Columns(6).Style.TextAlignment = sgAlignRightCenter
            .Columns(7).Style.TextAlignment = sgAlignRightCenter
            .Columns(8).Style.TextAlignment = sgAlignRightCenter
            .Columns(9).Style.TextAlignment = sgAlignGeneral
        Else
            .Columns(6).Style.TextAlignment = sgAlignGeneral
            .Columns(7).Style.TextAlignment = sgAlignGeneral
            .Columns(8).Style.TextAlignment = sgAlignGeneral
            .Columns(9).Style.TextAlignment = sgAlignGeneral
        End If
        .Redraw sgRedrawAll
        
'        .Columns(1).Style.TextAlignment = sgAlignLeftCenter
'        .Columns(8).SortType = sgSortTypeNumber
End With
Set rsTomaInv = Nothing
cmdBuscaInventario.Enabled = False
LV_DEPTO.Enabled = False
cmdPrintInvent.Visible = True
'cmdUpdateInvent.Visible = True

End Function

Private Sub cmdPrintInvent_Click()
If cFile = "" Then
    Call BuscaInventario("TOMA")
    Call ImprimirGrid
    cmdPrintInvent.Visible = False
Else
    Call PrintFile(cFile, Me.Caption)
End If
End Sub

Private Sub cmdUpdateInvent_Click()

If MsgBox("¿ Desea realizar la actualizacion del Inventario ?", vbQuestion + vbYesNo, "Actualizacion del Invetario Fisico") = vbYes Then
    If UpdateInventario Then
        cmdUpdateInvent.Visible = False
        'cmdUpdateInvent.Refresh
        Call BuscaInventario("OK")
    End If
End If

End Sub
'---------------------------------------------------------------------------------------
' Procedure : UpdateInventario
' Author    : hsequeira
' Date      : 02/09/2015
' Purpose   : ACTUALIZA LOS VALORES DE Exist_Real en el inventario
'---------------------------------------------------------------------------------------
'
Private Function UpdateInventario() As Boolean
'LEE EL GRID, GUARDA DATOS CAMBIADOS e IMPRIME LOS CAMBIOS REALIZADOS
Dim iFila As Long
Dim cData As String
Dim nFileNumber As Integer
Dim bTag As Boolean

On Error GoTo ErrAdm:

bTag = False
nFileNumber = FreeFile()

iFila = 1

cFile = "\INVENTARIO_FISICO_" & Format(Date, "DD_MM_YYYY") & "_" & Format(Time, "HH_MM") & ".txt"
Open App.Path & cFile For Output As #nFileNumber

Print #nFileNumber, Space(1)
Print #nFileNumber, Format(Date, "long date") & "  " & Time
Print #nFileNumber, Space(1)
Print #nFileNumber, rs00!DESCRIP
Print #nFileNumber, Space(1)
Print #nFileNumber, "Actualización de Inventario Fisico"
Print #nFileNumber, Space(1)
Print #nFileNumber, "Realizado por : " & NOM_ADMINISTRADOR
Print #nFileNumber, String(40, "=")

For iFila = 1 To DD_PEDDETALLE.Rows.Count - 1
    DD_PEDDETALLE.Row = iFila
    If (DD_PEDDETALLE.Rows.At(iFila).Cells(6).value <> DD_PEDDETALLE.Rows.At(iFila).Cells(7).value) Then
        'Or (DD_PEDDETALLE.Rows.At(iFila).Cells(7).value <> DD_PEDDETALLE.Rows.At(iFila).Cells(8).value) Then
        'DIFF EXIST1 or DIFF EXIST2
        cData = DD_PEDDETALLE.Rows.At(iFila).Cells(2).value & vbCrLf
        'cData = cData & String(5, Str(126)) & " Bodega 1: " & DD_PEDDETALLE.Rows.At(iFila).Cells(5).value & " ==> Nuevo Valor: " & DD_PEDDETALLE.Rows.At(iFila).Cells(6).value & vbCrLf
        cData = cData & String(5, Str(126)) & " Bodega 2: " & DD_PEDDETALLE.Rows.At(iFila).Cells(6).value & " ==> Nuevo Valor: " & DD_PEDDETALLE.Rows.At(iFila).Cells(7).value & vbCrLf
        cData = cData & String(5, Str(126)) & " Razón del Cambio: " & DD_PEDDETALLE.Rows.At(iFila).Cells(8).value
        
        'Call UpdateInvent2(DD_PEDDETALLE.Rows.At(iFila).Cells(0).value, _
            DD_PEDDETALLE.Rows.At(iFila).Cells(6).value, _
            DD_PEDDETALLE.Rows.At(iFila).Cells(8).value, _
            DD_PEDDETALLE.Rows.At(iFila).Cells(2).value, _
            DD_PEDDETALLE.Rows.At(iFila).Cells(9).value)
            
        Call UpdateInvent2(DD_PEDDETALLE.Rows.At(iFila).Cells(0).value, _
            DD_PEDDETALLE.Rows.At(iFila).Cells(6).value, _
            DD_PEDDETALLE.Rows.At(iFila).Cells(7).value, _
            DD_PEDDETALLE.Rows.At(iFila).Cells(2).value, _
            DD_PEDDETALLE.Rows.At(iFila).Cells(8).value, _
            DD_PEDDETALLE.Rows.At(iFila).Cells(9).value, _
            DD_PEDDETALLE.Rows.At(iFila).Cells(10).value)
            
        Print #nFileNumber, cData
        bTag = True
        cData = ""
    End If
    
'    If DD_PEDDETALLE.Rows.At(iFila).Cells(9).value <> DD_PEDDETALLE.Rows.At(iFila).Cells(10).value Then
'        'DIFF EXIST2
'        Print #nFileNumber, cData
'    End If
    
Next
Close #nFileNumber
If bTag = False Then
    'Kill cFile
End If
UpdateInventario = bTag
DD_PEDDETALLE.Row = 1
Exit Function

ErrAdm:
If Err.Number = 13 Then
    ShowMsg "Articulo: " & DD_PEDDETALLE.Rows.At(iFila).Cells(2).value & vbCrLf & _
            "DEBE INTRODUCIR UN VALOR NUMERICO EN LA BODEGA QUE ESTA ACTUALIZANDO", vbBlue, vbYellow
Else
    ShowMsg "Articulo: " & DD_PEDDETALLE.Rows.At(iFila).Cells(2).value & vbCrLf & _
            Err.Number & " - " & Err.Description, vbYellow, vbRed
End If
Resume Next
End Function
Private Function UpdateInvent2(nIDInvent As Long, Exist1 As Single, Exist2 As Single, cProducto As String, _
                                                cRazon As String, nCantidad2 As Single, cDesripConsumo As String) As Boolean
Dim cSQL As String
Dim cLog As String
Dim nCalculaValorDesempacado As Single


If LTrim(RTrim(cDesripConsumo)) = "Unidad" Or LTrim(RTrim(cDesripConsumo)) = "UNIDAD" Then
    'SI EL VALOR DEL PRODUCTO ES UNIDAD, PONE LA CANTIDAD MARCADA
    nCalculaValorDesempacado = Exist2
Else
    'DE LO CONTRARIO MULTIPLICA EL VALOR MARCADO POR SU FACTOR DEFINIDO
    nCalculaValorDesempacado = Exist2 * nCantidad2
End If
msConn.BeginTrans
'cSQL = "UPDATE INVENT SET EXIST1 = " & E1 & ", EXIST2 = " & E2 & " WHERE ID = " & nID
cSQL = "UPDATE INVENT SET EXIST2 = " & nCalculaValorDesempacado & " WHERE ID = " & nIDInvent
msConn.Execute cSQL
msConn.CommitTrans

'cLog = "Admin.Inventario Fisico Cambio: " & cInvent & ", Bod1: " & E1 & ", Bod2: " & E2 & " ==> Razon: " & cRazon
cLog = "Admin.Inventario Fisico Cambio: " & cProducto & ", Bodega2: " & Exist2 & " ==> Razon: " & cRazon
EscribeLog cLog
End Function

Private Sub DD_PEDDETALLE_KeyPressEdit(ByVal RowKey As Long, ByVal ColIndex As Long, KeyAscii As Integer)
If cmdUpdateInvent.Visible Then
Else
    cmdUpdateInvent.Visible = True
    cmdPrintInvent.Visible = False
End If
End Sub

'Private Sub cmdPrint_Click()
'Dim nLinea As Integer
'Dim RSREPORT As New ADODB.Recordset
'Dim nPage As Integer
'Dim nAcumProducto As Single
'Dim nAcumDepto As Single
'Dim cSQL As String
'Dim bErrFlag As Boolean
'
'On Error GoTo ErrAdm:
'
'cSQL = "SELECT C.DESCRIP AS DEPTO, a.Nombre, b.Descrip AS Envase, "
'cSQL = cSQL & " a.Cantidad AS Cant_Envase, a.Exist1 as Existencia_1, a.Exist2 AS Existencia_2, "
'cSQL = cSQL & " a.COSTO, a.ITBM"
'cSQL = cSQL & " FROM invent AS a, unidades AS b, DEP_INV AS C"
'cSQL = cSQL & " Where A.UNID_MEDIDA = b.id And A.COD_DEPT = C.CODIGO"
'cSQL = cSQL & " AND C.DESCRIP = '" & Combo1.Text & "' "
'cSQL = cSQL & " ORDER BY C.DESCRIP, b.descrip, a.nombre"
'
'RSREPORT.Open cSQL, msConn, adOpenStatic, adLockOptimistic
'
''" a.COSTO_EMPAQUE as Costo, a.ITBM"
'MainMant.spDoc.DocBegin
'MainMant.spDoc.WindowTitle = "TOMA DE INVENTARIO"
'MainMant.spDoc.FirstPage = 1
'MainMant.spDoc.PageOrientation = SPOR_PORTRAIT
'MainMant.spDoc.Units = SPUN_LOMETRIC
'nPage = nPage + 1
'MainMant.spDoc.Page = nPage
'
'MainMant.spDoc.TextOut 300, 200, Format(Date, "long date") & "  " & Time
'MainMant.spDoc.TextOut 300, 250, rs00!DESCRIP
'MainMant.spDoc.TextOut 300, 300, "Toma de Inventario"
'MainMant.spDoc.TextOut 300, 350, "DEPARTAMENTO : " & Combo1.Text
'MainMant.spDoc.TextOut 300, 400, "NOMBRE                                ENVASE  Cant_Envase      Exist 1         Exist 2          Costo     ITBM             TOTAL"
'MainMant.spDoc.TextOut 300, 450, "----------------------------------------------------------------------------------------------------------------------------------------------------"
'nLinea = 500
'Do Until RSREPORT.EOF
'    MainMant.spDoc.TextOut 300, nLinea, Mid(RSREPORT!NOMBRE, 1, 15)
'    MainMant.spDoc.TextOut 800, nLinea, Mid(RSREPORT!envase, 1, 10)
'    MainMant.spDoc.TextAlign = SPTA_RIGHT
'    MainMant.spDoc.TextOut 1100, nLinea, Format(RSREPORT!cant_envase, "####0")
'    MainMant.spDoc.TextOut 1300, nLinea, Format(RSREPORT!Existencia_1, "####.00")
'    MainMant.spDoc.TextOut 1500, nLinea, Format(RSREPORT!Existencia_2, "####.00")
'    MainMant.spDoc.TextOut 1700, nLinea, Format(RSREPORT!COSTO, "###0.00")
'    MainMant.spDoc.TextOut 1830, nLinea, Format(RSREPORT!ITBM, "##.00")
'    '' nAcumProducto = (((RSREPORT!Existencia_1 + RSREPORT!Existencia_2) / RSREPORT!cant_envase) * RSREPORT!COSTO)
'    nAcumProducto = (((RSREPORT!Existencia_1 + RSREPORT!Existencia_2)) * RSREPORT!COSTO)
'    MainMant.spDoc.TextOut 2070, nLinea, Format(nAcumProducto, "####0.00")
'    nAcumProducto = 0#
'
'    'INFO: REV ENE2010
'    If RSREPORT!cant_envase = 0 Then
'        nAcumDepto = nAcumDepto + 0
'        bErrFlag = True
'    Else
'        nAcumDepto = nAcumDepto + (((RSREPORT!Existencia_1 + RSREPORT!Existencia_2) / RSREPORT!cant_envase) * RSREPORT!COSTO)
'    End If
'
'    MainMant.spDoc.TextAlign = SPTA_LEFT
'    nLinea = nLinea + 50
'    If nLinea > 2400 Then
''''        MainMant.spDoc.DocBegin
''''        MainMant.spDoc.WindowTitle = "TOMA DE INVENTARIO"
''''        MainMant.spDoc.FirstPage = 1
''''        MainMant.spDoc.PageOrientation = SPOR_PORTRAIT
''''        MainMant.spDoc.Units = SPUN_LOMETRIC
'        nPage = nPage + 1
'        MainMant.spDoc.Page = nPage
'
'        MainMant.spDoc.TextOut 300, 200, Format(Date, "long date") & "  " & Time
'        MainMant.spDoc.TextOut 300, 250, rs00!DESCRIP
'        MainMant.spDoc.TextOut 300, 300, "Toma de Inventario"
'        MainMant.spDoc.TextOut 300, 350, "DEPARTAMENTO : " & Combo1.Text
'        MainMant.spDoc.TextOut 300, 400, "NOMBRE                                ENVASE  Cant_Envase      Exist 1         Exist 2          Costo     ITBM             TOTAL"
'        MainMant.spDoc.TextOut 300, 450, "----------------------------------------------------------------------------------------------------------------------------------------------------"
'        nLinea = 500
'    End If
'    RSREPORT.MoveNext
'Loop
'MainMant.spDoc.TextOut 300, nLinea + 100, "Total Departamental : " & Format(nAcumDepto, "CURRENCY")
'If bErrFlag Then
'    MainMant.spDoc.TextOut 300, nLinea + 200, " (******) ALGUNOS PRODUCTOS TIENEN CERO (0) EN LA CANTIDAD DEL ENVASE. FAVOR REVISAR. "
'End If
'RSREPORT.Close
'MainMant.spDoc.DoPrintPreview
'Set RSREPORT = Nothing
'On Error GoTo 0
'
'Call Seguridad
'
'Exit Sub
'
'ErrAdm:
'On Error Resume Next
'MsgBox "Error : " & Err.Number & " - " & Err.Description & vbCrLf & "Producto : " & RSREPORT!NOMBRE, vbCritical, "Error en Producto"
'On Error GoTo 0
'Resume Next
'End Sub

Private Sub Form_Load()

''''Dim rsLoc As New ADODB.Recordset
''''Dim cSQL As String
''''rsLoc.Open "SELECT CODIGO,DESCRIP FROM DEP_INV ORDER BY DESCRIP", msConn, adOpenStatic, adLockOptimistic
''''If Not rsLoc.EOF Then nLocDep = rsLoc!CODIGO
''''Do While Not rsLoc.EOF
''''    Combo1.AddItem rsLoc!DESCRIP
''''    rsLoc.MoveNext
''''Loop
''''rsLoc.Close
''''Set rsLoc = Nothing
''''

nDeptoSeleccionado = -1
Call CargarDepartamentos

Me.Controls("arv").Visible = False
Call Seguridad
End Sub
Private Sub CargarDepartamentos()
Dim rsDepto As ADODB.Recordset
Dim cSQL As String

Set rsDepto = New ADODB.Recordset

'cSQL = "SELECT A.ID, A.COD_DEPT, B.DESCRIP AS DEPARTAMENTO, A.NOMBRE, "
cSQL = "SELECT CODIGO, DESCRIP FROM DEP_INV ORDER BY DESCRIP"

rsDepto.Open cSQL, msConn, adOpenDynamic, adLockReadOnly

If rsDepto.EOF Then
    ShowMsg "NO HAY REGISTROS EN LA TABLA DE DEPARTAMENTOS", vbYellow, vbRed
    rsDepto.Close
    Set rsDepto = Nothing
    Exit Sub
End If

LV_DEPTO.AddItem "TODOS LOS DEPARTAMENTOS" & Space(130) & "~-1"
Do While Not rsDepto.EOF
    LV_DEPTO.AddItem rsDepto!DESCRIP & Space(130) & "~" & rsDepto!CODIGO
    rsDepto.MoveNext
Loop
rsDepto.Close
Set rsDepto = Nothing
LV_DEPTO.ListIndex = 0
'cmdUpdateInvent.Visible = True

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
        cmdPrint.Enabled = False
    Case "N"        'SIN DERECHOS
        Combo1.Enabled = False
        cmdPrint.Enabled = False
End Select
End Function

Private Sub Form_Resize()
arv.Move DD_PEDDETALLE.Left, DD_PEDDETALLE.Top, DD_PEDDETALLE.Width, DD_PEDDETALLE.Height
End Sub

Private Sub LV_DEPTO_Click()
Dim aCodigo() As String
aCodigo = Split(LV_DEPTO.Text, Chr(126))
nDeptoSeleccionado = Val(aCodigo(1))
End Sub
