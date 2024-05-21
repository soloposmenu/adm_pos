VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomct2.ocx"
Begin VB.Form NCProveedor 
   BackColor       =   &H00404000&
   Caption         =   "Nota de Credito (Proveedores)"
   ClientHeight    =   5910
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9000
   Icon            =   "NCProveedor.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   5910
   ScaleWidth      =   9000
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox txtObservacion 
      Height          =   735
      Left            =   2040
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   27
      Top             =   5040
      Width           =   3735
   End
   Begin MSComCtl2.DTPicker txtFecIni 
      Height          =   330
      Left            =   5280
      TabIndex        =   2
      Top             =   465
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   582
      _Version        =   393216
      Format          =   131334145
      CurrentDate     =   37474
   End
   Begin VB.CommandButton cmdAplicaNC 
      Caption         =   "&Aplicar NC"
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
      Left            =   120
      TabIndex        =   8
      Top             =   4800
      Width           =   1335
   End
   Begin VB.ComboBox cmbInvent 
      Height          =   315
      Left            =   2520
      Style           =   2  'Dropdown List
      TabIndex        =   4
      Top             =   1200
      Width           =   3375
   End
   Begin VB.TextBox txtDocumento 
      Alignment       =   1  'Right Justify
      Enabled         =   0   'False
      Height          =   330
      Left            =   120
      MaxLength       =   15
      TabIndex        =   0
      Top             =   465
      Width           =   1215
   End
   Begin VB.CommandButton cmdGO 
      BackColor       =   &H00FFC0C0&
      Caption         =   "&GO !"
      Height          =   375
      Left            =   8400
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   1170
      Width           =   495
   End
   Begin VB.TextBox txtCosto 
      Alignment       =   1  'Right Justify
      Height          =   330
      Left            =   7200
      TabIndex        =   6
      Top             =   1170
      Width           =   1095
   End
   Begin VB.TextBox txtCant 
      Alignment       =   1  'Right Justify
      Height          =   330
      Left            =   6000
      TabIndex        =   5
      Top             =   1170
      Width           =   615
   End
   Begin VB.ComboBox cmbDeptInv 
      Height          =   315
      Left            =   120
      Style           =   2  'Dropdown List
      TabIndex        =   3
      Top             =   1200
      Width           =   2295
   End
   Begin VB.ComboBox cmbProveedor 
      Height          =   315
      Left            =   1440
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   480
      Width           =   3375
   End
   Begin MSComctlLib.ListView LV 
      Height          =   2655
      Left            =   120
      TabIndex        =   9
      Top             =   2040
      Width           =   8655
      _ExtentX        =   15266
      _ExtentY        =   4683
      View            =   3
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      SmallIcons      =   "ImageList1"
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   0
   End
   Begin VB.Label Label1 
      BackColor       =   &H00404000&
      Caption         =   "Costo Actual (Sin ITBM)"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   14
      Left            =   5160
      TabIndex        =   30
      Top             =   1680
      Width           =   1815
   End
   Begin VB.Label lbCostoEmpaque 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   7080
      TabIndex        =   29
      Top             =   1680
      Width           =   1215
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H000000C0&
      BorderWidth     =   2
      Height          =   900
      Left            =   60
      Top             =   50
      Width           =   6855
   End
   Begin VB.Label Label1 
      BackColor       =   &H00404000&
      Caption         =   "Observaciones"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   13
      Left            =   2040
      TabIndex        =   28
      Top             =   4800
      Width           =   1215
   End
   Begin VB.Label Label1 
      BackColor       =   &H00404000&
      Caption         =   "Fecha del Documento"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   9
      Left            =   5040
      TabIndex        =   26
      Top             =   240
      Width           =   1575
   End
   Begin VB.Label lbUnidades 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   4320
      TabIndex        =   25
      Top             =   1680
      Width           =   615
   End
   Begin VB.Label Label1 
      BackColor       =   &H00404000&
      Caption         =   "Cant x Unidad"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   12
      Left            =   3120
      TabIndex        =   24
      Top             =   1680
      Width           =   1215
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H00000000&
      Height          =   255
      Index           =   11
      Left            =   1560
      TabIndex        =   23
      Top             =   1680
      Width           =   1455
   End
   Begin VB.Label Label1 
      BackColor       =   &H00404000&
      Caption         =   "Unidad de Medida"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   10
      Left            =   120
      TabIndex        =   22
      Top             =   1680
      Width           =   1455
   End
   Begin VB.Label txtTOTAL 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "0.00"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   7440
      TabIndex        =   21
      Top             =   5325
      Width           =   1290
   End
   Begin VB.Label txtITBM 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "0.00"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   7440
      TabIndex        =   20
      Top             =   4995
      Width           =   1290
   End
   Begin VB.Label txtSUBTOT 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "0.00"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   7440
      TabIndex        =   19
      Top             =   4680
      Width           =   1290
   End
   Begin VB.Label Label1 
      BackColor       =   &H00404000&
      Caption         =   "Costo Lineal sin ITBM"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   8
      Left            =   6840
      TabIndex        =   18
      Top             =   960
      Width           =   1695
   End
   Begin VB.Label Label1 
      BackColor       =   &H00404000&
      Caption         =   "Cantidad"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   7
      Left            =   6000
      TabIndex        =   17
      Top             =   960
      Width           =   735
   End
   Begin VB.Label Label1 
      BackColor       =   &H00404000&
      Caption         =   "Seleccione Articulo de Inventario"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   6
      Left            =   2520
      TabIndex        =   16
      Top             =   960
      Width           =   2775
   End
   Begin VB.Label Label1 
      BackColor       =   &H00404000&
      Caption         =   "Departamento"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   5
      Left            =   120
      TabIndex        =   15
      Top             =   960
      Width           =   1215
   End
   Begin VB.Label Label1 
      BackColor       =   &H00404000&
      Caption         =   "Proveedor"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   4
      Left            =   1440
      TabIndex        =   14
      Top             =   240
      Width           =   1215
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00404000&
      Caption         =   "Número de Documento"
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   3
      Left            =   240
      TabIndex        =   13
      Top             =   60
      Width           =   1095
   End
   Begin VB.Label Label1 
      BackColor       =   &H00404000&
      Caption         =   "TOTAL"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   2
      Left            =   6240
      TabIndex        =   12
      Top             =   5400
      Width           =   1215
   End
   Begin VB.Label Label1 
      BackColor       =   &H00404000&
      Caption         =   "ITBM"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   1
      Left            =   6240
      TabIndex        =   11
      Top             =   5100
      Width           =   1215
   End
   Begin VB.Label Label1 
      BackColor       =   &H00404000&
      Caption         =   "Sub Total"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   0
      Left            =   6240
      TabIndex        =   10
      Top             =   4800
      Width           =   1215
   End
End
Attribute VB_Name = "NCProveedor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim nIdInv As Long
Dim nIdProv As Long
Dim nITBM As Single
Dim iLinea As Byte
Dim rsDetalle As New ADODB.Recordset
Dim nConsecutivo As Long
Dim GET_BODEGA As Integer

Private Sub AbreInv()
Dim rsUnid As New ADODB.Recordset
Dim cSQL As String

cSQL = "SELECT A.ITBM, A.CANTIDAD, B.ID, B.DESCRIP AS UNIDAD,"
cSQL = cSQL & " B.UNIDAD AS MEDIDA, A.COSTO_EMPAQUE "
'INFO: 24FEB2014
cSQL = cSQL & ", A.CANTIDAD2 "
cSQL = cSQL & " FROM INVENT AS A, UNIDADES AS B "
cSQL = cSQL & " WHERE A.ID = " & nIdInv
cSQL = cSQL & " AND A.UNID_MEDIDA = B.ID"

rsUnid.Open cSQL, msConn, adOpenDynamic, adLockOptimistic

If Not rsUnid.EOF Then
    Label1(11) = rsUnid!UNIDAD
    Label1(11).Tag = rsUnid!ID
    nITBM = rsUnid!ITBM
    lbUnidades = IIf(IsNull(rsUnid!Cantidad), 0, rsUnid!Cantidad)
    'INFO: 24FEB2014
    'CAMBIO AL PRESENTAR COSTO LINEAL y LA CANTIDAD2 QUE ES LA DEL PRODUCTO DESEMPACADO
    'lbCostoEmpaque = Format(rsUnid!COSTO_EMPAQUE / (1 + (rsUnid!ITBM / 100)), "#,##0.0000")
    lbUnidades.Tag = rsUnid!CANTIDAD2
    lbCostoEmpaque = Format(rsUnid!COSTO_EMPAQUE, "#,##0.0000")
    'Label3 = Format(rsUnid!COSTO_EMPAQUE, "####0.00")
End If
rsUnid.Close
Set rsUnid = Nothing
End Sub
Private Sub cmbDeptInv_Click()
Dim POSIC As Integer
On Error Resume Next
POSIC = Len(cmbDeptInv.Text)
cmbDeptInv.Tag = Val(Mid(cmbDeptInv.Text, POSIC - 2, 3))
On Error GoTo 0
Call MuestraInventario
End Sub
Private Sub MuestraInventario()
Dim rsInvent As New ADODB.Recordset
cmbInvent.Clear
rsInvent.Open "SELECT ID,NOMBRE FROM INVENT " & _
    " WHERE COD_DEPT = " & Val(cmbDeptInv.Tag) & _
    " ORDER BY NOMBRE", msConn, adOpenStatic, adLockOptimistic
Do While Not rsInvent.EOF
    cmbInvent.AddItem rsInvent.Fields(1).value & Space(60) & rsInvent.Fields(0).value
    rsInvent.MoveNext
Loop
rsInvent.Close
Set rsInvent = Nothing
End Sub

Private Sub cmbDeptInv_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then cmbInvent.SetFocus
End Sub

Private Sub cmbInvent_Click()
Dim POSIC As Integer
On Error Resume Next
POSIC = Len(cmbInvent.Text)
nIdInv = Val(Mid(cmbInvent.Text, POSIC - 2, 3))
On Error GoTo 0
Call AbreInv
End Sub
Private Sub cmbInvent_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then txtCant.SetFocus
End Sub
Private Sub cmbProveedor_Click()
Dim POSIC As Integer
On Error Resume Next
POSIC = Len(cmbProveedor.Text)
nIdProv = Val(Mid(cmbProveedor.Text, POSIC - 2, 3))
On Error GoTo 0
End Sub

Private Sub cmbProveedor_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then txtFecIni.SetFocus
End Sub

Private Sub cmdAplicaNC_Click()
Dim vResp
vResp = MsgBox("¿ DESEA GRABAR ESTA NOTA DE CREDITO ?", vbQuestion + vbYesNo, "NOTA DE CREDITO # " & txtDocumento)
If vResp = vbYes Then
    Me.MousePointer = vbHourglass
    If ActualizaInventario Then
        Call ActualizaProveedores(txtTOTAL)
        'Call ActualizaDevoluciones
        Call ActualizaNotaCredito
        Call ImprimeDocumento
        Call ClearScreen
    Else
        MsgBox "Imposible Aplicar Nota de Crédito", vbCritical, "Revise e Intente nuevamente"
    End If
    Me.MousePointer = vbDefault
End If

Call Seguridad

End Sub
Private Sub ActualizaNotaCredito()
Dim cSQL As String
Dim cFecDoc As String

cFecDoc = Format(txtFecIni, "YYYYMMDD")
cfecsys = Format(Date, "YYYYMMDD")

rsDetalle.MoveFirst
Do While Not rsDetalle.EOF
    cSQL = "INSERT INTO NC "
    cSQL = cSQL & " VALUES ('"
    cSQL = cSQL & txtDocumento & "',"
    cSQL = cSQL & rsDetalle!LINEA & ","
    cSQL = cSQL & rsDetalle!dept_inv & ","
    cSQL = cSQL & rsDetalle!COD_INV & ","
    cSQL = cSQL & rsDetalle!CANT & ","
    cSQL = cSQL & rsDetalle!COSTO_UNIT & ","
    cSQL = cSQL & rsDetalle!ITBM & ","
    cSQL = cSQL & nIdProv & ","
    cSQL = cSQL & rs!NUMERO & ",'"
    cSQL = cSQL & cFecDoc & "','"
    cSQL = cSQL & cfecsys & "','"
    cSQL = cSQL & Format(Time, "HH:MM") & "','"
    cSQL = cSQL & txtObservacion & "')"
    
    msConn.BeginTrans
    msConn.Execute cSQL
    msConn.CommitTrans
    
    Debug.Print rsDetalle!ITEM_DESCRIPTOR
    EscribeLog ("Admin.Nota Crédito # " & txtDocumento.Text & " " & LTrim(RTrim(rsDetalle!ITEM_DESCRIPTOR)) & " (" & rsDetalle!CANT & ") " & rsDetalle!UNID_MEDIDA)
    rsDetalle.MoveNext
Loop

cSQL = "UPDATE NOTA_CREDITO SET CONTADOR = CONTADOR + 1"
msConn.BeginTrans
msConn.Execute cSQL
msConn.CommitTrans

End Sub
Private Sub ImprimeDocumento()
Dim nLinea As Integer
Dim nPage As Integer

MainMant.spDoc.DocBegin
MainMant.spDoc.WindowTitle = "Impresión de NOTA DE CREDITO"
MainMant.spDoc.FirstPage = 1
MainMant.spDoc.PageOrientation = SPOR_PORTRAIT
MainMant.spDoc.Units = SPUN_LOMETRIC
nPage = 1
MainMant.spDoc.Page = nPage

MainMant.spDoc.TextOut 300, 200, Format(Date, "long date") & "  " & Time
MainMant.spDoc.TextOut 300, 250, rs00!DESCRIP
MainMant.spDoc.TextOut 300, 300, "NOTA DE CREDITO"
MainMant.spDoc.TextOut 300, 400, "Documento : " & txtDocumento
MainMant.spDoc.TextOut 300, 450, "Fecha del Documento : " & txtFecIni
MainMant.spDoc.TextOut 300, 500, "Proveedor : " & Left(cmbProveedor.Text, 25)

MainMant.spDoc.TextOut 300, 600, "Linea    Producto                                          Cantidad                   Costo"
MainMant.spDoc.TextOut 300, 650, "------------------------------------------------------------------------------------------------"

nLinea = 700
rsDetalle.MoveFirst
Do While Not rsDetalle.EOF
    MainMant.spDoc.TextAlign = SPTA_LEFT
    MainMant.spDoc.TextOut 300, nLinea, rsDetalle!LINEA
    MainMant.spDoc.TextOut 430, nLinea, rsDetalle!ITEM_DESCRIPTOR
    MainMant.spDoc.TextAlign = SPTA_RIGHT
    MainMant.spDoc.TextOut 1120, nLinea, rsDetalle!CANT
    MainMant.spDoc.TextOut 1400, nLinea, Format(rsDetalle!COSTO_UNIT, "#,##0.00")
    nLinea = nLinea + 50
    rsDetalle.MoveNext
Loop
nLinea = nLinea + 50
MainMant.spDoc.TextAlign = SPTA_LEFT
MainMant.spDoc.TextOut 1120, nLinea, "Sub Total :"
MainMant.spDoc.TextAlign = SPTA_RIGHT
MainMant.spDoc.TextOut 1400, nLinea, txtSUBTOT
nLinea = nLinea + 50
MainMant.spDoc.TextAlign = SPTA_LEFT
MainMant.spDoc.TextOut 1120, nLinea, "ITBM : "
MainMant.spDoc.TextAlign = SPTA_RIGHT
MainMant.spDoc.TextOut 1400, nLinea, txtITBM
nLinea = nLinea + 50
MainMant.spDoc.TextAlign = SPTA_LEFT
MainMant.spDoc.TextOut 1120, nLinea, "TOTAL : "
MainMant.spDoc.TextAlign = SPTA_RIGHT
MainMant.spDoc.TextOut 1400, nLinea, txtTOTAL
MainMant.spDoc.TextAlign = SPTA_LEFT
MainMant.spDoc.DoPrintPreview
End Sub

Private Sub ClearScreen()
txtCosto = ""
txtCant = ""
txtSUBTOT = 0#
txtITBM = 0#
txtTOTAL = 0#
txtObservacion = ""
Label1(11) = ""
lbUnidades = 0
iLinea = 0
rsDetalle.MoveFirst
Do While Not rsDetalle.EOF
    rsDetalle.Delete
    rsDetalle.Update
    rsDetalle.MoveNext
Loop
cmbDeptInv.ListIndex = -1
cmbInvent.ListIndex = -1
cmbProveedor.ListIndex = -1

LV.ListItems.Clear
LV.Refresh

txtDocumento.Enabled = True
txtDocumento.Text = GetConsecutivo
txtDocumento.Enabled = False

End Sub

Private Function ActualizaProveedores(nUpdate As Single)
Dim cSQL As String

cSQL = "UPDATE PROVEEDORES SET SALDO = SALDO - " & nUpdate & " WHERE CODIGO = " & nIdProv
msConn.BeginTrans
msConn.Execute cSQL
msConn.CommitTrans
'EscribeLog ("Admin." & "Nota Crédito : " & cmbInvent)
EscribeLog ("Admin." & "Nota Crédito # " & txtDocumento.Text & " (" & Format(nUpdate, "CURRENCY") & ") para " & LTrim(RTrim(Left(cmbProveedor, 30))))

End Function
Private Function ActualizaInventario() As Boolean
Dim cSQL As String

On Error GoTo ErrAdm:
ActualizaInventario = False
rsDetalle.MoveFirst
Do While Not rsDetalle.EOF
    
    If GET_BODEGA = 1 Then
        cSQL = "UPDATE INVENT SET EXIST1 = EXIST1 - " & rsDetalle!CANT
        cSQL = cSQL & " WHERE ID = " & rsDetalle!COD_INV
    Else
        cSQL = "UPDATE INVENT SET EXIST2 = EXIST2 - " & rsDetalle!CANT * rsDetalle!CANTIDAD2
        cSQL = cSQL & " WHERE ID = " & rsDetalle!COD_INV
    End If
    
    msConn.BeginTrans
    msConn.Execute cSQL
    msConn.CommitTrans
    rsDetalle.MoveNext
Loop
ActualizaInventario = True
On Error GoTo 0
Exit Function

ErrAdm:
End Function
Private Sub cmdGO_Click()
Dim lITBM As Single
Dim nLocSB As Double

On Error GoTo ErrAdm:

If cmbInvent.Text = "" Or txtCosto = "" Or Val(txtCosto) < 0# Or _
        Val(txtCant) = 0 Or Val(txtCant) < 0 Or cmbDeptInv.Text = "" Then
    
    MsgBox "Por favor revise los datos. Falta Informacion o estan equivocados", vbExclamation, Me.Caption
    txtCosto.SelStart = 0
    txtCosto.SelLength = Len(txtCosto.Text)
    txtCosto.SetFocus
    Exit Sub

End If
BoxResp = MsgBox("¿ Desea Añadir a la Nota de Crédito " & Trim(Mid(cmbInvent.Text, 1, 40)) & " ?", vbQuestion + vbYesNo, Me.Caption)
If BoxResp = vbYes Then
    'txtCosto ES EL COSTO x CAJA
    lITBM = nITBM / 100
    lITBM = Format((Val(txtCosto) * lITBM), "standard")
    'ID.DEPT-ID.INV-DESCR.MEDIDA-CANT-DESCRP.PROD-COSTO.IN-CANTxCOSTO.IN-
    nLocSB = Format(Val(txtCosto) + lITBM, "currency")
    LV.ListItems.Add , , iLinea + 1
    LV.ListItems.Item(iLinea + 1).ListSubItems.Add , , Left(cmbInvent.Text, 25)
    LV.ListItems.Item(iLinea + 1).ListSubItems.Add , , Format(txtCant / lbUnidades, "#0.00")
    LV.ListItems.Item(iLinea + 1).ListSubItems.Add , , txtCant
    LV.ListItems.Item(iLinea + 1).ListSubItems.Add , , txtCosto
    iLinea = iLinea + 1
    
    With rsDetalle
        .AddNew
        !LINEA = iLinea
        !dept_inv = cmbDeptInv.Tag
        !COD_INV = nIdInv
        !CANT = txtCant
        !COSTO_UNIT = txtCosto  'Costo de caja sin impuesto
        !ITBM = lITBM           'impuesto
        !ITEM_DESCRIPTOR = Left(cmbInvent.Text, 25)
        !CANTIDAD2 = Val(lbUnidades.Tag)
        !UNID_MEDIDA = LTrim(RTrim(Left(Label1(11).Caption, 25)))
        .Update
    End With
    txtSUBTOT = Format(Val(txtSUBTOT) + Val(txtCosto), "#,##0.00")
    txtITBM = Format((Val(txtITBM) + lITBM), "#,##0.00")
    txtTOTAL = Format((Val(txtTOTAL) + nLocSB), "#,##0.00")

    txtCosto = 0#
    cmbInvent.SetFocus
End If
On Error GoTo 0

Exit Sub

ErrAdm:
    MsgBox Err.Description, vbCritical, "OCURRIO UN ERROR. CORRIJA SUS DATOS o REVISE SU INVENTARIO"
    MsgBox "EL DATO NO SERA GUARDADO HASTA QUE SE CORRIJA EL ERROR", vbCritical, BoxTit
End Sub
Private Function GetConsecutivo() As Long
Dim rsNotaCredito As New ADODB.Recordset

rsNotaCredito.Open "SELECT CONTADOR FROM NOTA_CREDITO", msConn, adOpenStatic, adLockOptimistic
GetConsecutivo = rsNotaCredito.Fields(0).value + 1
rsNotaCredito.Close
Set rsNotaCredito = Nothing
End Function
Private Sub Form_Load()
Dim rsTemp As New ADODB.Recordset

GET_BODEGA = GetFromINI("Administracion", "ActualizaBodega", App.Path & "\soloini.ini")

On Error GoTo ErrAdm:
rsTemp.Open "SELECT CODIGO,EMPRESA FROM PROVEEDORES ORDER BY EMPRESA", msConn, adOpenStatic, adLockOptimistic
Do While Not rsTemp.EOF
    cmbProveedor.AddItem rsTemp.Fields(1).value & Space(60) & rsTemp.Fields(0).value
    rsTemp.MoveNext
Loop
rsTemp.Close
rsTemp.Open "SELECT CODIGO,DESCRIP FROM DEP_INV ORDER BY DESCRIP", msConn, adOpenStatic, adLockOptimistic
Do While Not rsTemp.EOF
    cmbDeptInv.AddItem rsTemp.Fields(1).value & Space(60) & rsTemp.Fields(0).value
    rsTemp.MoveNext
Loop
rsTemp.Close
Set rsTemp = Nothing
nITBM = 0
iLinea = 0
txtFecIni = Format(Date, "SHORT DATE")
LV.ColumnHeaders.Clear

LV.ColumnHeaders.Add , , "Linea"
LV.ColumnHeaders.Add , , "Producto"
LV.ColumnHeaders.Add , , "Cantidad"
LV.ColumnHeaders.Add , , "Cantidad (Extendida)"
LV.ColumnHeaders.Add , , "Costo"
LV.ColumnHeaders.Item(3).Alignment = lvwColumnRight
LV.ColumnHeaders.Item(4).Alignment = lvwColumnRight
LV.ColumnHeaders.Item(5).Alignment = lvwColumnRight

LV.ColumnHeaders.Item(1).Width = 700
LV.ColumnHeaders.Item(2).Width = 3000
LV.ColumnHeaders.Item(3).Width = 900
LV.ColumnHeaders.Item(4).Width = 1700
LV.ColumnHeaders.Item(5).Width = 900
Call CreaTablaDetalle
txtDocumento.Enabled = True
txtDocumento.Text = GetConsecutivo
txtDocumento.Enabled = False
On Error GoTo 0

Call Seguridad

Exit Sub
ErrAdm:
    MsgBox Err.Number & " - " & Err.Description, vbCritical, BoxTit
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
        txtDocumento.Enabled = False: cmbProveedor.Enabled = False: cmdGO.Enabled = False
        LV.Enabled = False: cmdAplicaNC.Enabled = False: txtObservacion.Enabled = False
    Case "N"        'SIN DERECHOS
        txtDocumento.Enabled = False: cmbProveedor.Enabled = False: cmdGO.Enabled = False
        LV.Enabled = False: cmdAplicaNC.Enabled = False: txtObservacion.Enabled = False
End Select
End Function

Private Sub CreaTablaDetalle()

With rsDetalle
    .Fields.Append "LINEA", adInteger, , adFldUpdatable
    .Fields.Append "DEPT_INV", adInteger, , adFldUpdatable
    .Fields.Append "COD_INV", adInteger, , adFldUpdatable
    .Fields.Append "CANT", adSingle, , adFldUpdatable
    .Fields.Append "COSTO_UNIT", adCurrency, , adFldUpdatable
    .Fields.Append "ITBM", adSingle, , adFldUpdatable
    .Fields.Append "ITEM_DESCRIPTOR", adChar, 25, adFldUpdatable
    'INFO: 24FEB2014
    .Fields.Append "CANTIDAD2", adSingle, , adFldUpdatable
    .Fields.Append "UNID_MEDIDA", adChar, 25, adFldUpdatable
    .CursorType = adOpenKeyset
    .LockType = adLockOptimistic
    .Open
End With
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
On Error GoTo ErrAdm:
    rsDetalle.Close
    Set rsDetalle = Nothing
    Unload Me
On Error GoTo 0
Exit Sub

ErrAdm:
    EscribeLog ("Admin." & "NCProveedor.QueryUnload. (" & Err.Number & ") - " & Err.Description)
End Sub
Private Sub LV_DblClick()
Dim vResp As Variant
Dim nLineaEnRs As Integer

On Error Resume Next
vResp = MsgBox("¿ Desea Eliminar Linea # " & LV.SelectedItem.Text & " ?", vbYesNo + vbQuestion, "Elimina Articulo")
If vResp = vbYes Then
    nLineaEnRs = Val(LV.SelectedItem.Text)
    LV.ListItems.Remove (LV.SelectedItem.Index)
    LV.Refresh
    rsDetalle.MoveFirst
    rsDetalle.Find "LINEA = " & nLineaEnRs
    rsDetalle.Delete
End If
On Error GoTo 0
End Sub
Private Sub txtCant_KeyPress(KeyAscii As Integer)
On Error GoTo ErrAdm:
If KeyAscii = 13 Then
    txtCosto = Format(txtCant * (Val(lbCostoEmpaque) / Val(lbUnidades)), "#0.000")
    txtCosto.SetFocus
End If
On Error GoTo 0
Exit Sub

ErrAdm:
    MsgBox Err.Number & " - " & Err.Description, vbCritical, "CANTIDAD INVALIDA"
End Sub

Private Sub txtCosto_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    If Not IsNumeric(txtCosto) Then
        MsgBox "Escriba un valor valido", vbExclamation, BoxTit
        txtCosto.SetFocus
        Exit Sub
    End If
    txtCosto = Format(txtCosto, "#####0.000")
    cmdGO_Click
End If
End Sub
Private Sub txtFecIni_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then cmbDeptInv.SetFocus
End Sub
