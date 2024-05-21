VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form Desbloquear 
   BackColor       =   &H00B39665&
   Caption         =   "Desbloqueo de Mesas"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8040
   Icon            =   "Desbloquear.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   3195
   ScaleWidth      =   8040
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdSalir 
      Caption         =   "Salir"
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
      Left            =   6720
      TabIndex        =   1
      Top             =   2520
      Width           =   1215
   End
   Begin MSComctlLib.ListView LV 
      Height          =   2775
      Left            =   120
      TabIndex        =   0
      ToolTipText     =   "Haga Doble Click para Desbloquear"
      Top             =   240
      Width           =   6375
      _ExtentX        =   11245
      _ExtentY        =   4895
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483624
      BorderStyle     =   1
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   0
   End
End
Attribute VB_Name = "Desbloquear"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmdSalir_Click()
Unload Me
End Sub

Private Sub Form_Load()
'Que hace desbloquar:
'2002-07-08
'Quita el bloqueo de mesas y la marca de ocupada
'Deja el Cajero que la tenia bloqueada
SHOWDATA

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
        LV.Enabled = False
    Case "N"        'SIN DERECHOS
        LV.Enabled = False
End Select
End Function

Private Sub LV_DblClick()
Dim nNewValue As Variant
Dim vResp As Variant
On Error Resume Next
vResp = MsgBox("¿ Desea Desbloquear Mesa " & LV.SelectedItem.Text & " ?", vbYesNo + vbQuestion, "UnLock Tables")
If vResp = vbYes Then
    msConn.BeginTrans
    msConn.Execute "UPDATE MESAS SET LOCK = 0 WHERE NUMERO = " & LV.SelectedItem.Text
    msConn.CommitTrans
    EscribeLog ("Admin." & "Desbloqueo de Mesas: " & LV.SelectedItem.Text & ", Cajero : " & LV.SelectedItem.ListSubItems.Item(2).Text)
    SHOWDATA
End If
On Error GoTo 0
End Sub
Private Sub SHOWDATA()
Dim nFila As Integer
Dim cSQL As String
Dim rsLocal As New ADODB.Recordset
nFila = 1
LV.ListItems.Clear
LV.ColumnHeaders.Clear
LV.ColumnHeaders.Add , , "Mesa"
LV.ColumnHeaders.Add , , "En Uso?"
LV.ColumnHeaders.Add , , "Mesero"
LV.ColumnHeaders.Add , , "Status (1)"
'LV.ColumnHeaders.Item(2).Alignment = lvwColumnRight

cSQL = "SELECT A.NUMERO,OCUPADA,(TRIM(B.NOMBRE) & ' ' & TRIM(B.APELLIDO)) AS CAJERO, A.LOCK "
cSQL = cSQL & " FROM MESAS AS A, MESEROS AS B "
cSQL = cSQL & " WHERE A.LOCK = 1 AND A.MESERO_ACTUAL = B.NUMERO order by A.NUMERO"

rsLocal.Open cSQL, msConn, adOpenStatic, adLockOptimistic
Do While Not rsLocal.EOF
    LV.ListItems.Add , , rsLocal!NUMERO
    LV.ListItems.Item(nFila).ListSubItems.Add , , IIf(rsLocal!OCUPADA = 0, "SI", "NO")
    LV.ListItems.Item(nFila).ListSubItems.Add , , rsLocal!CAJERO
    LV.ListItems.Item(nFila).ListSubItems.Add , , rsLocal!LOCK
    nFila = nFila + 1
    rsLocal.MoveNext
Loop
rsLocal.Close
'*******************************************
cSQL = "SELECT A.NUMERO,OCUPADA,A.LOCK "
cSQL = cSQL & " FROM MESAS AS A "
cSQL = cSQL & " WHERE A.LOCK = 1 order by A.NUMERO"

rsLocal.Open cSQL, msConn, adOpenStatic, adLockOptimistic
Do While Not rsLocal.EOF
    LV.ListItems.Add , , rsLocal!NUMERO
    'INFO: AHORA DICE S/N EN VEZ DE 0 y 1
    LV.ListItems.Item(nFila).ListSubItems.Add , , IIf(rsLocal!OCUPADA = 0, "SI", "NO")
    LV.ListItems.Item(nFila).ListSubItems.Add , , "Sin Mesero"
    LV.ListItems.Item(nFila).ListSubItems.Add , , rsLocal!LOCK
    nFila = nFila + 1
    rsLocal.MoveNext
Loop
rsLocal.Close
Set rsLocal = Nothing
End Sub
