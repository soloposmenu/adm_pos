VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form lstPagos 
   BackColor       =   &H00B39665&
   Caption         =   "LISTADO DE PAGOS A PROVEEDORES"
   ClientHeight    =   6405
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8070
   Icon            =   "lstPagos.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   6405
   ScaleWidth      =   8070
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdPrint 
      Caption         =   "Imprimir"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   120
      TabIndex        =   3
      Top             =   5640
      Width           =   1215
   End
   Begin MSComctlLib.TreeView TV2 
      Height          =   2535
      Left            =   120
      TabIndex        =   2
      Top             =   3000
      Width           =   7695
      _ExtentX        =   13573
      _ExtentY        =   4471
      _Version        =   393217
      LabelEdit       =   1
      Style           =   7
      BorderStyle     =   1
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSComctlLib.ListView LV 
      Height          =   2535
      Left            =   120
      TabIndex        =   0
      Top             =   360
      Width           =   6975
      _ExtentX        =   12303
      _ExtentY        =   4471
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   0
   End
   Begin VB.Shape ErrorShape 
      BackColor       =   &H00FFFFFF&
      BorderColor     =   &H0000FFFF&
      BorderWidth     =   2
      FillColor       =   &H000000FF&
      FillStyle       =   0  'Solid
      Height          =   495
      Left            =   7200
      Shape           =   3  'Circle
      Top             =   2400
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Label Label1 
      BackColor       =   &H00B39665&
      Caption         =   "SELECCIONE EL PROVEEDOR"
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
      TabIndex        =   1
      Top             =   120
      Width           =   3735
   End
End
Attribute VB_Name = "lstPagos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdPrint_Click()
Dim tvwPrint As CPrintTvw
Set tvwPrint = New CPrintTvw

On Error GoTo ErrAdm:
With tvwPrint

  .FontBold = True
  .FontSize = 10
  .PrintFooter = True
  .PrintStyle = ePrintAll
  .SecondTitle = ""
  .SecondTitleFontBold = False
  .SecondTitleFontSize = 10
  .Title = Me.Caption
  .TitleFontBold = True
  .TitleFontSize = 14
  .ConnectorStyle = econnectlines
End With

Set tvwPrint.tvwToPrint = Me.TV2
Call tvwPrint.PrintTvw
On Error GoTo 0

Exit Sub
ErrAdm:
    ShowMsg Err.Number & " - " & Err.Description, vbYellow, vbRed
End Sub

Private Sub Form_Load()
Dim rsProvedor As ADODB.Recordset
Dim cSQL As String
Dim nFila  As Integer

Set rsProvedor = New ADODB.Recordset

cSQL = "SELECT CODIGO, NOMBRE + SPACE(1) + APELLIDO AS PROVEED, "
cSQL = cSQL & " EMPRESA , TELEFONO1, "
cSQL = cSQL & " TELEFONO2, SALDO "
cSQL = cSQL & " FROM PROVEEDORES "
cSQL = cSQL & " ORDER BY EMPRESA "

rsProvedor.Open cSQL, msConn, adOpenStatic, adLockOptimistic

nFila = 1
Call PutHeaders(1)
Do While Not rsProvedor.EOF
    LV.ListItems.Add , , rsProvedor!CODIGO
    LV.ListItems.Item(nFila).ListSubItems.Add , , rsProvedor!PROVEED
    LV.ListItems.Item(nFila).ListSubItems.Add , , rsProvedor!EMPRESA
    LV.ListItems.Item(nFila).ListSubItems.Add , , rsProvedor!TELEFONO1
    LV.ListItems.Item(nFila).ListSubItems.Add , , rsProvedor!telefono2
    LV.ListItems.Item(nFila).ListSubItems.Add , , Format(rsProvedor!SALDO, "STANDARD")
    nFila = nFila + 1
    rsProvedor.MoveNext
Loop
rsProvedor.Close
Set rsProvedor = Nothing

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
        cmdPrint.Enabled = False
    Case "N"        'SIN DERECHOS
        LV.Enabled = False
        TV2.Enabled = False
        cmdPrint.Enabled = False
End Select
End Function

Private Sub LV_Click()
Dim cMsgBox As String
Dim cSQL As String
Dim rsResult As ADODB.Recordset
Dim nFila As Integer
Dim cLVProv  As String
Dim cErrorLog As String

nFila = 1
ErrorShape.Visible = False
cErrorLog = ""
TV2.Nodes.Clear
cLVProv = LV.SelectedItem.ListSubItems(2).Text
cMsgBox = "¿ Desea ver los pagos de : " & cLVProv & " ?"
If vbYes = MsgBox(cMsgBox, vbQuestion + vbYesNo, "IMPRESION DE PAGOS") Then
    
    Set rsResult = New ADODB.Recordset
    
    cSQL = "SELECT A.TIPO_DOC, A.NUM_DOC,  B.LIN, A.CODIGO_PROV, C.EMPRESA, "
    cSQL = cSQL & " A.FECHA_DOC, A.VALOR_DOC, B.REF_DOC_NUM, B.REF_VAL_PAG, "
    cSQL = cSQL & " A.COMMENT, B.REF_TDOC, D.MONTO "
    cSQL = cSQL & " FROM CXP_REC AS A , PAGOS_PROV AS B, PROVEEDORES AS C,  COMPRAS_HEAD AS D"
    cSQL = cSQL & " WHERE A.NUM_DOC = B.NUM_DOC And A.CODIGO_PROV = C.CODIGO"
    cSQL = cSQL & " AND A.CODIGO_PROV = " & Val(LV.SelectedItem.Text)
    cSQL = cSQL & " AND B.REF_DOC_NUM = D.INDICE "
    cSQL = cSQL & " ORDER BY A.NUM_DOC, B.LIN"
    
    rsResult.Open cSQL, msConn, adOpenStatic, adLockOptimistic
    If rsResult.EOF Then
        rsResult.Close
        Set rsResult = Nothing
        ShowMsg "NO HAY DATOS DE PAGOS REGISTRADOS ", vbBlack
        Exit Sub
    End If
    
    Call PutHeaders(2)
    
    
    Do While Not rsResult.EOF
        On Error Resume Next
        'Debug.Print cLVProv & ". # " & rsResult!NUM_DOC & ". VALOR. : " & Format(rsResult!VALOR_DOC, "CURRENCY") & ". Fecha: " & GetFecha(rsResult!FECHA_DOC)
        TV2.Nodes.Add , , "MAIN" & cLVProv & ". # " & rsResult!NUM_DOC, _
                    cLVProv & ". DOCUMENTO # " & rsResult!NUM_DOC & ". VALOR. : " & Format(rsResult!VALOR_DOC, "CURRENCY") & ".  del: " & GetFecha(rsResult!FECHA_DOC)
        On Error GoTo 0
        rsResult.MoveNext
    Loop
    rsResult.MoveFirst
    
    Do While Not rsResult.EOF
        On Error Resume Next
        TV2.Nodes.Add "MAIN" & cLVProv & ". # " & rsResult!NUM_DOC, tvwChild, _
                     "C" & rsResult!NUM_DOC, "Ref: " & RemoveNull(rsResult!COMMENT)
        On Error GoTo 0
        nFila = nFila + 1
        rsResult.MoveNext
    Loop
    rsResult.MoveFirst
    
    Do While Not rsResult.EOF
        'Debug.Print "KEY: " & "L" & rsResult!TIPO_DOC & rsResult!NUM_DOC & rsResult!REF_DOC_NUM
        On Error GoTo ErrAdm:
        TV2.Nodes.Add "MAIN" & cLVProv & ". # " & rsResult!NUM_DOC, tvwChild, _
            "L" & rsResult!TIPO_DOC & rsResult!NUM_DOC & rsResult!REF_DOC_NUM, " (" & rsResult!REF_DOC_NUM & ")" & "  Tipo: " & rsResult!REF_TDOC & ". Original: " & Format(rsResult!MONTO, "STANDARD") & ".  Pagado: " & Format(rsResult!REF_VAL_PAG, "STANDARD")
        On Error GoTo 0
        rsResult.MoveNext
    Loop
    rsResult.Close
    Set rsResult = Nothing
    Call ExpandAllNodes
End If
If cErrorLog <> "" Then MsgBox cErrorLog
On Error GoTo 0
Exit Sub

ErrAdm:
    Call EscribeLog("Error mostrando detalle de Pagos de Compras: " & cLVProv & ". # " & rsResult!NUM_DOC)
    If Not ErrorShape.Visible Then ErrorShape.Visible = True
    cErrorLog = cErrorLog & "Error mostrando detalle de Pagos de Compras: " & cLVProv & ". # " & rsResult!NUM_DOC & vbCrLf
    Resume Next
End Sub
Private Sub ExpandAllNodes()
Dim i As Integer
For i = 1 To TV2.Nodes.Count
    TV2.Nodes(i).Expanded = True
Next i
End Sub
Private Sub PutHeaders(nOpc As Integer)
If nOpc = 1 Then
    LV.ListItems.Clear
    LV.ColumnHeaders.Clear
    LV.ColumnHeaders.Add , , "CODIGO"
    LV.ColumnHeaders.Add , , "CONTACTO"
    LV.ColumnHeaders.Add , , "EMPRESA"
    LV.ColumnHeaders.Add , , "TELEFONO"
    LV.ColumnHeaders.Add , , "TELEFONO2"
    LV.ColumnHeaders.Add , , "SALDO"
    
    LV.ColumnHeaders.Item(6).Alignment = lvwColumnRight
    LV.ColumnHeaders.Item(1).Width = 500
    LV.ColumnHeaders.Item(2).Width = 2200
    LV.ColumnHeaders.Item(3).Width = 2700
    LV.ColumnHeaders.Item(4).Width = 1300
    LV.ColumnHeaders.Item(5).Width = 1300
    LV.ColumnHeaders.Item(6).Width = 1700
Else
''''    LV2.ListItems.Clear
''''    LV2.ColumnHeaders.Clear
''''    LV2.ColumnHeaders.Add , , "TIPO_DOC"
''''    LV2.ColumnHeaders.Add , , "NUM_DOC"
''''    LV2.ColumnHeaders.Add , , "LIN"
''''    LV2.ColumnHeaders.Add , , "PROVEEDOR"
''''    LV2.ColumnHeaders.Add , , "FECHA"
''''    LV2.ColumnHeaders.Add , , "VALOR"
''''    LV2.ColumnHeaders.Add , , "REFERENCIA"
''''    LV2.ColumnHeaders.Add , , "PAGADO"
    
''    LV.ColumnHeaders.Item(6).Alignment = lvwColumnRight
''    LV.ColumnHeaders.Item(1).Width = 500
''    LV.ColumnHeaders.Item(2).Width = 2200
''    LV.ColumnHeaders.Item(3).Width = 2700
''    LV.ColumnHeaders.Item(4).Width = 1300
''    LV.ColumnHeaders.Item(5).Width = 1300
''    LV.ColumnHeaders.Item(6).Width = 1700
End If
End Sub
