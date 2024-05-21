VERSION 5.00
Begin VB.Form conCompras 
   BackColor       =   &H00B39665&
   Caption         =   "Consulta de Factura de Compra"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   Icon            =   "conCompras.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdBusca 
      Caption         =   "Buscar"
      Height          =   495
      Left            =   3840
      TabIndex        =   9
      Top             =   1320
      Width           =   735
   End
   Begin VB.ComboBox Combo4 
      Height          =   315
      Left            =   120
      Style           =   2  'Dropdown List
      TabIndex        =   7
      Top             =   1440
      Width           =   3375
   End
   Begin VB.ComboBox Combo3 
      Height          =   315
      ItemData        =   "conCompras.frx":030A
      Left            =   2040
      List            =   "conCompras.frx":030C
      TabIndex        =   3
      Text            =   "Combo3"
      ToolTipText     =   "Se muestran los años con compras"
      Top             =   600
      Width           =   1215
   End
   Begin VB.ComboBox Combo2 
      Height          =   315
      Left            =   120
      Style           =   2  'Dropdown List
      TabIndex        =   2
      Top             =   600
      Width           =   1575
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      Left            =   120
      TabIndex        =   1
      Top             =   2040
      Width           =   2295
   End
   Begin VB.CommandButton cmdPrint 
      Caption         =   "&Imprimir"
      Height          =   495
      Left            =   120
      TabIndex        =   0
      Top             =   2520
      Width           =   1215
   End
   Begin VB.Shape Shape1 
      BorderWidth     =   3
      Height          =   855
      Left            =   60
      Top             =   240
      Width           =   3375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00B39665&
      Caption         =   "Proveedor"
      Height          =   255
      Index           =   3
      Left            =   120
      TabIndex        =   8
      Top             =   1200
      Width           =   1215
   End
   Begin VB.Label Label1 
      BackColor       =   &H00B39665&
      Caption         =   "Número de Documento"
      Height          =   255
      Index           =   2
      Left            =   120
      TabIndex        =   6
      Top             =   1800
      Width           =   1815
   End
   Begin VB.Label Label1 
      BackColor       =   &H00B39665&
      Caption         =   "Año"
      Height          =   255
      Index           =   1
      Left            =   2040
      TabIndex        =   5
      Top             =   360
      Width           =   1215
   End
   Begin VB.Label Label1 
      BackColor       =   &H00B39665&
      Caption         =   "Mes"
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   4
      Top             =   360
      Width           =   1215
   End
End
Attribute VB_Name = "conCompras"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private iLin As Long
Private nPagina As Long
Private Sub cmdBusca_Click()
Dim rsCompras As New ADODB.Recordset
Dim rsTemporal As New ADODB.Recordset
Dim cSQL As String
Dim c2Seek As String

c2Seek = Trim(Combo3.Text) & Format(Combo2.ListIndex + 1, "00")
If Combo4.ListIndex = 0 Then
    'Todos los Proveedores
    cSQL = "SELECT NUMERO, FECHA FROM COMPRAS_HEAD "
    cSQL = cSQL & " WHERE LEFT(FECHA,6) = '" & c2Seek & "'"
    cSQL = cSQL & " ORDER BY NUMERO "
Else
    'Se selecciono un provedor
    rsTemporal.Open "SELECT CODIGO FROM PROVEEDORES WHERE EMPRESA = '" & Combo4.Text & "'", msConn, adOpenStatic, adLockOptimistic
    cSQL = "SELECT NUMERO, FECHA FROM COMPRAS_HEAD "
    cSQL = cSQL & " WHERE COD_PROV = " & rsTemporal!CODIGO
    cSQL = cSQL & " AND LEFT(FECHA,6) = '" & c2Seek & "'"
    cSQL = cSQL & " ORDER BY NUMERO "
    rsTemporal.Close
    Set rsTemporal = Nothing
End If
rsCompras.Open cSQL, msConn, adOpenStatic, adLockOptimistic
Combo1.Clear
Combo1.ListIndex = -1
If rsCompras.EOF Then
    MsgBox "No existen compras para los datos seleccionados", vbInformation, BoxTit
    rsCompras.Close
    Set rsCompras = Nothing
    Exit Sub
End If

Do While Not rsCompras.EOF
    Combo1.AddItem rsCompras!NUMERO
    rsCompras.MoveNext
Loop

Combo1.ListIndex = 0
rsCompras.Close
Set rsCompras = Nothing

Call Seguridad

End Sub

Private Sub cmdPrint_Click()
Dim rsTemporal As New ADODB.Recordset
Dim cSQL As String
Dim nTotalSinITBM As Single
Dim nITBM As Single
Dim nNeto As Single
Dim nTotal As Single
Dim nAcumITBM_5 As Single
Dim nAcumITBM_10 As Single
Dim nAcumITBM_Else As Single
Dim cFecDoc As String
Dim cEmpresa  As String

Me.MousePointer = vbHourglass

cSQL = "SELECT A.TIPO,A.FECHA, A.MONTO,A.USUARIO,A.PAGADO, "
cSQL = cSQL & " B.LINEA,B.CANT,B.UNIDADES,B.COSTO_UNIT,B.COSTO_IN, "
cSQL = cSQL & " C.NOMBRE, C.ITBM, D.DESCRIP, E.EMPRESA "
cSQL = cSQL & " FROM COMPRAS_HEAD AS A, COMPRAS_DETA AS B, "
cSQL = cSQL & " INVENT AS C, UNIDADES AS D, PROVEEDORES AS E "
cSQL = cSQL & " WHERE A.NUMERO = " & Val(Combo1.Text)
cSQL = cSQL & " AND A.COD_PROV = E.CODIGO "
cSQL = cSQL & " AND A.INDICE = B.NUM_COMPRA "
cSQL = cSQL & " AND B.CODI_INV = C.ID "
cSQL = cSQL & " AND B.UNID_ID = D.ID "
cSQL = cSQL & " ORDER BY B.LINEA "
rsTemporal.Open cSQL, msConn, adOpenStatic, adLockOptimistic

If rsTemporal.EOF Then
    Me.MousePointer = vbDefault
    rsTemporal.Close
    Set rsTemporal = Nothing
    Exit Sub
End If

cFecDoc = Right(rsTemporal!FECHA, 2) & "/" & Mid(rsTemporal!FECHA, 5, 2) & "/" & Left(rsTemporal!FECHA, 4)
cEmpresa = rsTemporal!EMPRESA
Do While Not rsTemporal.EOF
    nNeto = nNeto + (rsTemporal!COSTO_IN / (1 + (rsTemporal!ITBM / 100)))   'SIN ITBM
    nTotal = nTotal + rsTemporal!COSTO_IN
    Select Case rsTemporal!ITBM
        Case 5
            nAcumITBM_5 = nAcumITBM_5 + rsTemporal!COSTO_IN - (rsTemporal!COSTO_IN / (1 + (rsTemporal!ITBM / 100)))
        Case 10
            nAcumITBM_10 = nAcumITBM_10 + rsTemporal!COSTO_IN - (rsTemporal!COSTO_IN / (1 + (rsTemporal!ITBM / 100)))
        Case Else
            nAcumITBM_Else = nAcumITBM_Else + rsTemporal!COSTO_IN - (rsTemporal!COSTO_IN / (1 + (rsTemporal!ITBM / 100)))
    End Select
    rsTemporal.MoveNext
Loop

MainMant.spDoc.DocBegin
Call PrintTit(nNeto, nTotal, nAcumITBM_5, nAcumITBM_10, nAcumITBM_Else, cFecDoc, cEmpresa)
EscribeLog ("Admin." & "Revisión de Compra # " & Combo1.Text)

nTotalSinITBM = 0#
rsTemporal.MoveFirst
Do While Not rsTemporal.EOF
    MainMant.spDoc.TextAlign = SPTA_LEFT
    MainMant.spDoc.TextOut 300, iLin, rsTemporal!LINEA + 1
    MainMant.spDoc.TextOut 400, iLin, rsTemporal!CANT
    MainMant.spDoc.TextOut 500, iLin, rsTemporal!NOMBRE
    MainMant.spDoc.TextAlign = SPTA_RIGHT
    nTotalSinITBM = (rsTemporal!COSTO_IN / (1 + (rsTemporal!ITBM / 100)))
    nITBM = rsTemporal!COSTO_IN - (rsTemporal!COSTO_IN / (1 + (rsTemporal!ITBM / 100)))
    MainMant.spDoc.TextOut 1100, iLin, Format(Format((nTotalSinITBM / rsTemporal!CANT), "###0.00"), "@@@@@@@@")
    MainMant.spDoc.TextOut 1280, iLin, Format(Format(nTotalSinITBM, "###0.00"), "@@@@@@@@")
    MainMant.spDoc.TextOut 1480, iLin, Format(Format(nITBM, "####0.00"), "@@@@@@@@@")
    MainMant.spDoc.TextOut 1680, iLin, Format(Format(rsTemporal!COSTO_IN, "####0.00"), "@@@@@@@@@")
    MainMant.spDoc.TextOut 1920, iLin, Format(Format((rsTemporal!COSTO_IN / rsTemporal!UNIDADES), "####0.0000"), "@@@@@@@@@@")
    MainMant.spDoc.TextAlign = SPTA_LEFT
    iLin = iLin + 50
    rsTemporal.MoveNext
    If iLin > 2400 Then
        Call PrintTit(nNeto, nTotal, nAcumITBM_5, nAcumITBM_10, nAcumITBM_Else, cFecDoc, cEmpresa)
    End If
Loop
Me.MousePointer = vbDefault
MainMant.spDoc.DoPrintPreview
rsTemporal.Close
Set rsTemporal = Nothing
nPagina = 0
iLin = 0

Call Seguridad

End Sub

Private Sub Combo1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then cmdPrint_Click
End Sub

Private Sub Combo2_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then Combo2.SetFocus
End Sub
Private Sub Combo3_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then Combo4.SetFocus
End Sub

Private Sub Form_Load()
Dim rsTemporal As New ADODB.Recordset
On Error Resume Next
Combo2.AddItem "ENERO"
Combo2.AddItem "FEBRERO"
Combo2.AddItem "MARZO"
Combo2.AddItem "ABRIL"
Combo2.AddItem "MAYO"
Combo2.AddItem "JUNIO"
Combo2.AddItem "JULIO"
Combo2.AddItem "AGOSTO"
Combo2.AddItem "SEPTIEMBRE"
Combo2.AddItem "OCTUBRE"
Combo2.AddItem "NOVIEMBRE"
Combo2.AddItem "DICIEMBRE"
Combo2.ListIndex = 0
rsTemporal.Open "SELECT LEFT(FECHA,4) AS ANNO FROM COMPRAS_HEAD GROUP BY LEFT(FECHA,4) ORDER BY 1", msConn, adOpenStatic, adLockOptimistic
Do While Not rsTemporal.EOF
    Combo3.AddItem rsTemporal!ANNO
    rsTemporal.MoveNext
Loop
Combo3.ListIndex = 0
rsTemporal.Close

rsTemporal.Open "SELECT EMPRESA FROM PROVEEDORES ORDER BY 1"
Combo4.AddItem "Todos los Proveedores"
Do While Not rsTemporal.EOF
    Combo4.AddItem rsTemporal!EMPRESA
    rsTemporal.MoveNext
Loop
Combo4.ListIndex = 0
rsTemporal.Close
Set rsTemporal = Nothing
On Error GoTo 0

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
        'INFO: NO HAY RESTRICCIONES
    Case "N"        'SIN DERECHOS
        Combo2.Enabled = False: Combo3.Enabled = False
        Combo4.Enabled = False: cmdBusca.Enabled = False
        Combo1.Enabled = False
        cmdPrint.Enabled = False
End Select
End Function

Private Sub PrintTit(nNeto As Single, nTotal As Single, nITBM5 As Single, nITBM10 As Single, nITBMElse As Single, cFecha As String, cEmpresa As String)

If nPagina = 0 Then
    MainMant.spDoc.WindowTitle = "Impresión de " & Me.Caption
    MainMant.spDoc.FirstPage = 1
    MainMant.spDoc.PageOrientation = SPOR_PORTRAIT
    MainMant.spDoc.Units = SPUN_LOMETRIC
End If
MainMant.spDoc.Page = nPagina + 1

MainMant.spDoc.TextOut 300, 200, Format(Date, "long date") & "  " & Time
MainMant.spDoc.TextOut 300, 250, "Página : " & nPagina + 1
MainMant.spDoc.TextOut 300, 350, rs00!DESCRIP
MainMant.spDoc.TextOut 300, 450, "(Copia) REGISTRO DE COMPRA # " & Combo1.Text
   
MainMant.spDoc.TextOut 300, 550, "Proveedor : " & cEmpresa
MainMant.spDoc.TextOut 300, 600, "Numero de Documento : " & Combo1.Text
MainMant.spDoc.TextOut 300, 650, "Fecha del Documento : " & cFecha
MainMant.spDoc.TextOut 300, 700, "Monto Neto  :" & Format(nNeto, "currency")
MainMant.spDoc.TextOut 300, 750, "ITBM  5% :   " & Format(nITBM5, "#,##0.00")
MainMant.spDoc.TextOut 300, 800, "ITBM 10% :   " & Format(nITBM10, "#,##0.00")
MainMant.spDoc.TextOut 300, 850, "Monto Total : " & Format(nTotal, "currency")
MainMant.spDoc.TextOut 300, 900, "LIN  CANT  DESCRIPCION                        PRECIO       TOTAL          ITBM          TOTAL     PRECIO UNIT"
MainMant.spDoc.TextOut 300, 950, "-----------------------------------------------------------------------------------------------------------------------------------------------"
iLin = 1000
nPagina = nPagina + 1
End Sub
