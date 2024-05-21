VERSION 5.00
Object = "{43135020-B751-4DDD-AB4A-B6D8A342216E}#1.0#0"; "sg20o.ocx"
Begin VB.Form TransferenciaRECETAS 
   BackColor       =   &H00B39665&
   Caption         =   "TRANSFERENCIAS A RECETAS"
   ClientHeight    =   6090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   8340
   Icon            =   "TransferenciaRECETAS.frx":0000
   LinkTopic       =   "Form2"
   ScaleHeight     =   6090
   ScaleWidth      =   8340
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdSalir 
      Caption         =   "SALIR"
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
      Left            =   120
      TabIndex        =   2
      Top             =   5280
      Width           =   2055
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Aplicar Transferecias"
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
      Left            =   6120
      TabIndex        =   1
      Top             =   5280
      Width           =   2055
   End
   Begin DDSharpGridOLEDB2.SGGrid DD_RECETAS 
      Height          =   4575
      Left            =   120
      TabIndex        =   0
      Top             =   480
      Width           =   8055
      _cx             =   14208
      _cy             =   8070
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
      StylesCollection=   $"TransferenciaRECETAS.frx":0442
      ColumnsCollection=   $"TransferenciaRECETAS.frx":2251
      ValueItems      =   $"TransferenciaRECETAS.frx":2766
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00B39665&
      Caption         =   "Al Terminar haga click en el boton de Aplicar Transferencias ==>"
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
      Height          =   495
      Left            =   2880
      TabIndex        =   4
      Top             =   5400
      Width           =   3135
   End
   Begin VB.Label Label2 
      BackColor       =   &H00B39665&
      Caption         =   "Escriba cuanto se esta creando de la Receta en la Columna de ENTRADA"
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
      Left            =   120
      TabIndex        =   3
      Top             =   120
      Width           =   6615
   End
End
Attribute VB_Name = "TransferenciaRECETAS"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private rsInvent As ADODB.Recordset

Private Sub cmdSalir_Click()
Set rsInvent = Nothing
Unload Me
End Sub

Private Sub Command1_Click()
If vbYes = MsgBox("DESEA APLICAR LA TRANSFERENCIA DE ARTICULOS ?", vbQuestion + vbYesNo, "AGREGAR EXISTENCIA EN RECETAS") Then
    Call ProcessTransferencias
End If

Call Seguridad

End Sub

Private Function ProcessTransferencias() As Boolean
Dim iRow As Long
Dim varID As Variant
Dim varUpdate As Variant
Dim cSQL As String
Dim bFlag As Boolean
Dim nSecuencialRecetaTransfer As Long
Dim cReceta As Variant
Dim nCostoTT As Single

For iRow = 0 To DD_RECETAS.DataRowCount - 1

     Call DD_RECETAS.Array.Get(iRow, 0, varID)
     Call DD_RECETAS.Array.Get(iRow, 1, cReceta)
     Call DD_RECETAS.Array.Get(iRow, 4, varUpdate)

    If varUpdate = "" Or varUpdate = Null Or varUpdate = 0 Or varUpdate = Empty Then
        'DO NOTHING
    Else
        bFlag = True

        'cSQL = "UPDATE RECETAS SET EXIST2 = EXIST2 + " & CLng(varUpdate) & " WHERE ID = " & CLng(varID)
        'INFO: CAMBIANDO A DECIMAL EN VEZ DE LONG
        cSQL = "UPDATE RECETAS SET EXIST2 = EXIST2 + " & CSng(varUpdate) & " WHERE ID = " & CLng(varID)

        msConn.BeginTrans
        msConn.Execute cSQL
        msConn.CommitTrans
       
        nSecuencialRecetaTransfer = GetNEWIndice("RECETA_TRANSFERS.ID")
        
        cSQL = "INSERT INTO RECETAS_TRANSFERS VALUES (" & nSecuencialRecetaTransfer & "," & CLng(varID) & ",'"
        cSQL = cSQL & Format(Date, "YYYYMMDD") & "','" & Format(Time, "HH:MM") & "',"
        cSQL = cSQL & npNumCaj & "," & CSng(varUpdate) & ")"
        
        msConn.BeginTrans
        msConn.Execute cSQL
        msConn.CommitTrans
        
        'HACE LAS REBAJAS EN INVENTARIO SEGUN LA DEFINICION DE LA RECETA
        nCostoTT = UpdateInvent(CLng(varID), CSng(varUpdate), nSecuencialRecetaTransfer, cReceta)
        
        Call EscribeLog("TRANSFERENCIAS A RECETA: " & cReceta & ", CANTIDAD: " & varUpdate & _
                                       ". COSTO DE TRANSFERENCIA: " & Format(nCostoTT, "#,###.###"))
        
    End If
Next

If bFlag Then
    ShowMsg "TRANSFERENCIA(S) RECETAS REALIZADAS CON EXITO", vbYellow, vbBlue
    Call LoadData
    Call DD_RECETAS_OnInit
End If

End Function
'Private Function UpdateInvent(nIDReceta As Long, nCant As Long, idTransfer As Long) As Single
Private Function UpdateInvent(nIDReceta As Long, nCant As Single, idTransfer As Long, cTexto As Variant) As Single 'UPDATE 21DIC2010
'HACE LAS REBAJAS EN INVENTARIO SEGUN LA DEFINICION DE LA RECETA
Dim rsRecetasInvent As ADODB.Recordset
Dim cSQL As String
Dim nCostoTotal As Single

   On Error GoTo UpdateInvent_Error

Set rsRecetasInvent = New ADODB.Recordset
cSQL = "SELECT A.ID_INVENT, A.CANTIDAD_CONSUME, B.COSTO "
cSQL = cSQL & " FROM RECETAS_INVENT AS A, INVENT AS B "
cSQL = cSQL & " WHERE A.ID = " & nIDReceta
cSQL = cSQL & " AND A.ID_INVENT = B.ID "

rsRecetasInvent.Open cSQL, msConn, adOpenStatic, adLockOptimistic
rsRecetasInvent.MoveFirst
Do While Not rsRecetasInvent.EOF

    cSQL = "UPDATE INVENT SET EXIST2 = EXIST2 - " & (rsRecetasInvent!CANTIDAD_CONSUME * nCant)
    cSQL = cSQL & " WHERE ID = " & rsRecetasInvent!ID_INVENT
    
    nCostoTotal = nCostoTotal + _
                                (rsRecetasInvent!CANTIDAD_CONSUME * nCant) * rsRecetasInvent!COSTO
    
    msConn.BeginTrans
    msConn.Execute cSQL
    msConn.CommitTrans
    
    cSQL = "INSERT INTO RECETAS_TR_INV VALUES (" & idTransfer & "," & rsRecetasInvent!ID_INVENT & ","
    cSQL = cSQL & (rsRecetasInvent!CANTIDAD_CONSUME * nCant) & "," & rsRecetasInvent!COSTO & ")"
    
    msConn.BeginTrans
    msConn.Execute cSQL
    msConn.CommitTrans
    
    rsRecetasInvent.MoveNext
Loop

rsRecetasInvent.Close
Set rsRecetasInvent = Nothing

UpdateInvent = Format(nCostoTotal, "CURRENCY")

   On Error GoTo 0
   Exit Function

UpdateInvent_Error:

    If Err.Number = 3021 Then
        ShowMsg "Esta Receta " & cTexto & ", no tiena Materia Prima en Inventario, no se puede Actualizar", vbYellow, vbRed
    Else
        ShowMsg "Error " & Err.Number & " (" & Err.Description & ") EN Transferencia RECETAS", vbYellow, vbRed
    End If

End Function

Private Sub DD_RECETAS_OnInit()

DD_RECETAS.ColumnClickSort = True
DD_RECETAS.EvenOddStyle = sgEvenOddRows
DD_RECETAS.ColorOdd = &HE0E0E0

DD_RECETAS.DataColCount = 5

DD_RECETAS.Columns(1).Hidden = True
DD_RECETAS.Columns(2).Caption = "Descripción"
DD_RECETAS.Columns(2).Width = 3000
DD_RECETAS.Columns(2).ReadOnly = True

DD_RECETAS.Columns(3).Caption = "Se Crea Por"
DD_RECETAS.Columns(3).Width = 1300
DD_RECETAS.Columns(3).ReadOnly = True

DD_RECETAS.Columns(4).Caption = "On Hand"
DD_RECETAS.Columns(4).Style.TextAlignment = sgAlignRightCenter
DD_RECETAS.Columns(4).Style.Format = "#,#0.00"
DD_RECETAS.Columns(4).Width = 900
DD_RECETAS.Columns(4).ReadOnly = True

DD_RECETAS.Columns(5).Caption = "Entrada"
DD_RECETAS.Columns(5).Style.TextAlignment = sgAlignRightCenter
DD_RECETAS.Columns(5).Style.Format = "#,##0.00"
DD_RECETAS.Columns(5).ReadOnly = False
DD_RECETAS.Columns(5).Width = 900

End Sub

Private Sub Form_Load()
If Not LoadData() Then
    ShowMsg "NO HAY RECETAS DEFINIDAS PARA TRANSFERIR", vbRed
End If

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
        Command1.Enabled = False
    Case "N"        'SIN DERECHOS
        DD_RECETAS.Enabled = False
        Command1.Enabled = False
End Select
End Function

Private Function LoadData() As Boolean
Dim rsTABLA As ADODB.Recordset
Dim cSQL As String

On Error GoTo ErrAdm:

Set rsTABLA = New ADODB.Recordset

cSQL = "SELECT A.ID, A.NOMBRE, B.DESCRIP , A.EXIST2 AS CANT "
cSQL = cSQL & " FROM RECETAS AS A, UNIDADES AS B"
cSQL = cSQL & " WHERE A.UNID_MEDIDA = B.ID "
cSQL = cSQL & " ORDER BY A.NOMBRE"

rsTABLA.Open cSQL, msConn, adOpenStatic, adLockReadOnly

With DD_RECETAS
    rsTABLA.MoveFirst
   .DataMode = sgUnbound
   .LoadArray rsTABLA.GetRows()
End With

'INFO: ARCHIVO QUE SE ESTA EXPORTANDO

rsTABLA.Close
Set rsTABLA = Nothing
LoadData = True
On Error GoTo 0
Exit Function

ErrAdm:
End Function

