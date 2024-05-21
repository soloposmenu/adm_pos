VERSION 5.00
Begin VB.Form ClearData 
   Caption         =   "Clear Database data"
   ClientHeight    =   7515
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8955
   Icon            =   "ClearData.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   7515
   ScaleWidth      =   8955
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command2 
      Caption         =   "Marca Todos"
      Height          =   375
      Left            =   7560
      TabIndex        =   6
      Top             =   4440
      Width           =   1095
   End
   Begin VB.ListBox List2 
      BackColor       =   &H00C0FFFF&
      Height          =   1860
      Left            =   120
      Style           =   1  'Checkbox
      TabIndex        =   5
      Top             =   4920
      Width           =   8535
   End
   Begin VB.CommandButton cmdMarkAll 
      Caption         =   "Marca Todos"
      Height          =   375
      Left            =   7595
      TabIndex        =   4
      Top             =   60
      Width           =   1095
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   120
      TabIndex        =   3
      Top             =   120
      Width           =   7455
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00C0E0FF&
      Caption         =   "Limpieza de Base de Datos"
      Height          =   495
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   6960
      Width           =   1215
   End
   Begin VB.ListBox List1 
      Height          =   3660
      Left            =   120
      Style           =   1  'Checkbox
      TabIndex        =   0
      Top             =   480
      Width           =   8535
   End
   Begin VB.Label Label1 
      BackColor       =   &H0080FFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   $"ClearData.frx":0442
      ForeColor       =   &H000000C0&
      Height          =   495
      Left            =   1800
      TabIndex        =   2
      Top             =   6960
      Width           =   6855
   End
End
Attribute VB_Name = "ClearData"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
Private Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpstring As Any, ByVal lpFileName As String) As Long

Private msConn As New ADODB.Connection

Private Sub cmdMarkAll_Click()
Dim i As Byte
For i = 0 To List1.ListCount - 1
    If List1.Selected(i) = True Then List1.Selected(i) = False Else List1.Selected(i) = True
Next
List1.ListIndex = 0
End Sub

Private Sub Command1_Click()
Dim i As Integer
Dim ccLista As String

On Error GoTo ErrAdm:
ccLista = "Lista1"
For i = 0 To List1.ListCount - 1
    If List1.Selected(i) Then
        Text1 = List1.Text
        Text1.Refresh
        msConn.BeginTrans
        msConn.Execute List1.List(i)
        msConn.CommitTrans
    End If
Next

ccLista = "Lista2"
For i = 0 To List2.ListCount - 1
    If List2.Selected(i) Then
        Text1 = List2.Text
        Text1.Refresh
        msConn.BeginTrans
        msConn.Execute List2.List(i)
        msConn.CommitTrans
    End If
Next

On Error GoTo 0
MsgBox "LIMPIEZA FINALIZADA"

Exit Sub
ErrAdm:
    If ccLista = "Lista1" Then
        MsgBox "Apunte donde Ocurrio el Error" & vbCrLf & _
            List1.List(i) & vbCrLf & _
            Err.Number & " - " & Err.Description, vbCritical, "Error en Aplicacion de la Operacion"
    Else
        MsgBox "Apunte donde Ocurrio el Error" & vbCrLf & _
            List2.List(i) & vbCrLf & _
            Err.Number & " - " & Err.Description, vbCritical, "Error en Aplicacion de la Operacion"
    End If
    Resume Next
End Sub

Private Sub Command2_Click()
Dim i As Byte
For i = 0 To List2.ListCount - 1
    If List2.Selected(i) = True Then List2.Selected(i) = False Else List2.Selected(i) = True
Next
List2.ListIndex = 0

End Sub

Private Sub Form_Load()
Dim cDataPath As String
Me.MousePointer = vbHourglass
On Error Resume Next
If Dir("C:\ACCESS\CLEARSOLO.MDB") = "" Then
Else
    Kill "C:\ACCESS\CLEARSOLO.MDB"
End If
'FileCopy "C:\ACCESS\SOLO.MDB", "C:\ACCESS\CLEARSOLO.MDB"
'MsgBox "SE HA CREADO UN RESPALDO LLAMADO : CLEARSOLO.MDB"
On Error GoTo 0
Me.MousePointer = vbDefault
'msConn.Open "Provider=Microsoft.Jet.OLEDB.4.0;Jet OLEDB:Database Password=master24;Data Source=C:\ACCESS\SOLO.mdb;"
cDataPath = GetFromINI("General", "DirectorioDatos", App.Path & "\soloini.ini")
'msConn.Open "Provider=Microsoft.Jet.OLEDB.4.0;Jet OLEDB:Database Password=master24;Data Source=C:\SOLO SOFTWARE\ACCESS\_CLEAR DATABASE_\SOLO.mdb;"
msConn.Open "Provider=Microsoft.Jet.OLEDB.4.0;Jet OLEDB:Database Password=master24;Data Source=" & cDataPath & "\SOLO.MDB" & ";"
List1.AddItem "UPDATE plu SET x_count = 0, z_count = 0, valor = 0.00, X_PERIOD_CNT = 0, Z_PERIOD_CNT = 0, PERIOD_VAL = 0.00;"
List1.AddItem "UPDATE depto SET x_count = 0, z_count = 0, valor = 0.00, X_PERIOD_CNT = 0, Z_PERIOD_CNT = 0, PERIOD_VAL = 0.00;"
List1.AddItem "UPDATE PAGOS SET x_count = 0, z_count = 0, valor = 0.00, X_PERIOD_CNT = 0, Z_PERIOD_CNT = 0, PERIOD_VAL = 0.00;"
List1.AddItem "UPDATE CAJEROS SET X_COUNT = 0,Z_COUNT = 0, VALOR = 0.00, X_C=0, Z_C=0, VTA_TOT= 0.00"
List1.AddItem "UPDATE MESAS SET OCUPADA = FALSE, X_COUNT = 0, Z_COUNT = 0, VALOR = 0, MESERO_ACTUAL = 0, LOCK=0"
List1.AddItem "UPDATE MESAS SET GUEST_COUNTER =0, GUEST_TOTAL =0"
List1.AddItem "UPDATE MESEROS SET X_COUNT=0, Z_COUNT = 0, VALOR = 0.00"
List1.AddItem "UPDATE CONTEND_02 SET x_count = 0, z_count = 0, valor = 0.00, X_PERIOD_CNT = 0, Z_PERIOD_CNT = 0, PERIOD_VAL = 0.00"
List1.AddItem "UPDATE INVENT SET EXIST1 = 0, EXIST2 = 0"
List1.AddItem "UPDATE ISC SET MES01 = 0, MES02 = 0, MES03 = 0, MES04 = 0, MES05 = 0, MES06 = 0, MES07 = 0, MES08 = 0, MES09 = 0, MES10 = 0, MES11 = 0, MES12 = 0, DIARIO = 0"
List1.AddItem "UPDATE NOTA_CREDITO SET CONTADOR = 0"
List1.AddItem "UPDATE PROVEEDORES SET SALDO = 0, LIMITE = 0"
'08:51 PM 20/02/2007
List1.AddItem "UPDATE INVENT SET NUM_ULT_COMPRA = ''"
List1.AddItem "UPDATE INVENT SET FEC_ULT_COMPRA = ''"
List1.AddItem "DELETE * FROM HIST_TR"
List1.AddItem "DELETE * FROM HIST_TR_CLI"
List1.AddItem "DELETE * FROM HIST_TR_PAGO"
List1.AddItem "DELETE * FROM HIST_TR_PROP"
List1.AddItem "UPDATE ORGANIZACION SET TRANS = 0, VTA_TOT = 0, X_CDEP = 0, Z_CDEP = 0, X_CMESAS = 0, Z_CMESAS = 0, X_CMESEROS = 0, Z_CMESEROS = 0, X_CPLU = 0, Z_CPLU = 0, CONT_RO = 0, CONT_ND = 0, CONT_NC = 0, TOT_HASH = 0,DIA_COUNT =0"
List1.AddItem "DELETE * FROM TRANSAC"
List1.AddItem "DELETE * FROM TRANSAC_CLI"
List1.AddItem "DELETE * FROM TRANSAC_PAGO"
List1.AddItem "DELETE * FROM TRANSAC_PROP"
List1.AddItem "DELETE * FROM TMP_CLI"
List1.AddItem "DELETE * FROM TMP_CUENTAS"
List1.AddItem "DELETE * FROM TMP_PAR_PAGO"
List1.AddItem "DELETE * FROM TMP_PAR_PROP"
List1.AddItem "DELETE * FROM TMP_TRANS"
List1.AddItem "DELETE * FROM TMP_VTA_TRANS"
'List1.AddItem "DELETE * FROM DEP_INV"
List1.AddItem "DELETE * FROM COMPRAS_HEAD"
List1.AddItem "DELETE * FROM CONTADORES"
List1.AddItem "DELETE * FROM CXC_REC"
List1.AddItem "DELETE * FROM CXP_REC"
List1.AddItem "DELETE * FROM DEV_HISTORY"
List1.AddItem "DELETE * FROM DEVOLUCION_HEAD"
List1.AddItem "DELETE * FROM DEVOLUCION_DETA"
List1.AddItem "DELETE * FROM HIST_INVENT"
List1.AddItem "DELETE * FROM NC"
List1.AddItem "DELETE * FROM PAGOS_CLI"
List1.AddItem "DELETE * FROM PAGOS_PROV"
List1.AddItem "DELETE * FROM TRANSFERENCIA"
List1.AddItem "DELETE * FROM VTA_TRANS"
List1.AddItem "DELETE * FROM Z_COUNTER"
'~~~~~~~~~~ 10ABRIL2013 ~~~~~~~~~~~~~~~~~~~
List2.AddItem "DELETE * FROM ACOMPA"
List2.AddItem "DELETE * FROM CONTENED"
List2.AddItem "DELETE * FROM CONTEND_01"
List2.AddItem "DELETE * FROM CONTEND_02"
List2.AddItem "DELETE * FROM DEPTO"
List2.AddItem "DELETE * FROM PLU_ACOMP"
List2.AddItem "DELETE * FROM PLU_INVENT"
List2.AddItem "DELETE * FROM PLU_RECETAS"
List2.AddItem "DELETE * FROM RECETAS"
List2.AddItem "DELETE * FROM RECETAS_TR_INV"
List2.AddItem "DELETE * FROM RECETAS_TRANSFERS"
List2.AddItem "DELETE * FROM RECETAS_INVENT"
List2.AddItem "DELETE * FROM SEGURIDAD"
List2.AddItem "DELETE * FROM MESEROS WHERE NUMERO <> 500"
List2.AddItem "DELETE * FROM CAJEROS WHERE NUMERO NOT IN (100,200) "
List2.AddItem "DELETE * FROM TRANSAC_FISCAL"
List2.AddItem "DELETE * FROM DEP_INV"
List2.AddItem "DELETE * FROM INVENT"
List2.AddItem "DELETE * FROM INVENT_02"
List2.AddItem "DELETE * FROM Z_FISCAL"
List2.AddItem "DELETE * FROM USUARIOS WHERE NUMERO <> 1967"
List2.AddItem "DELETE * FROM CLIENTES"
List2.AddItem "UPDATE INDICES SET INDICE = 1"
List2.AddItem "DELETE * FROM ISC"
List2.AddItem "DELETE * FROM MESES"
List2.AddItem "DELETE * FROM SUPER_DET"
List2.AddItem "DELETE * FROM SUPER_GRP"
List2.AddItem "DELETE * FROM PROVEEDORES"



End Sub

Public Function GetFromINI(Section As String, Key As String, Directory As String) As String
Dim strBuffer As String

On Error GoTo FileError:
    strBuffer = String(750, Chr(0))
    Key$ = LCase$(Key$)
    GetFromINI$ = Left(strBuffer, GetPrivateProfileString(Section$, ByVal Key$, "", strBuffer, Len(strBuffer), Directory$))
On Error GoTo 0
Exit Function

FileError:
    MsgBox Err.Number & ": NO SE ENCUENTRA ARCHIVO DE INICIALIZACION", vbCritical, "ERROR AL INICIAR"
    Resume Next
End Function

