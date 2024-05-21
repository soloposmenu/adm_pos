VERSION 5.00
Begin VB.Form MisClientes 
   BackColor       =   &H00808080&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "SELECCION DE CLIENTE"
   ClientHeight    =   7155
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10005
   ControlBox      =   0   'False
   FillStyle       =   0  'Solid
   Icon            =   "MisClientes.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7155
   ScaleWidth      =   10005
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdHotel 
      Caption         =   "HUESPEDES"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   5640
      TabIndex        =   1
      Top             =   720
      Width           =   1935
   End
   Begin VB.CommandButton cmdCli 
      Caption         =   "CLIENTES"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   1440
      TabIndex        =   0
      Top             =   720
      Width           =   1935
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Regresar"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   5520
      TabIndex        =   6
      Top             =   6360
      Width           =   1935
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "Aceptar"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   1440
      TabIndex        =   5
      Top             =   6360
      Width           =   1935
   End
   Begin VB.ListBox lstCli 
      BackColor       =   &H00008080&
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4515
      Left            =   120
      TabIndex        =   3
      Top             =   1560
      Width           =   8655
   End
   Begin VB.VScrollBar VSOcup 
      Height          =   4575
      Left            =   9000
      TabIndex        =   2
      Top             =   1560
      Width           =   855
   End
   Begin VB.Label Label1 
      BackColor       =   &H00808080&
      Caption         =   "Clientes "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   375
      Left            =   120
      TabIndex        =   4
      Top             =   120
      Width           =   9615
   End
End
Attribute VB_Name = "MisClientes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private rsCli As New ADODB.Recordset
Private POSIC As Integer
Private Sub GetClientes(Opc As Integer)
Dim MiOpcion As Boolean
Set rsCli = New Recordset
If Opc = 0 Then MiOpcion = False Else MiOpcion = True
txt = "SELECT Apellido,Nombre,Empresa,Codigo " & _
    " FROM CLIENTES " & _
    " WHERE HUESPED = " & MiOpcion & _
    " ORDER BY 1,2,3"
rsCli.Open txt, msConn, adOpenDynamic, adLockOptimistic

Do Until rsCli.EOF
    lstCli.AddItem FormatTexto((IIf(rsCli!nombre = "", "", rsCli!nombre) & "," & IIf(rsCli!apellido = "", "", rsCli!apellido)), 26) & Chr(9) & FormatTexto(rsCli!empresa, 20) & Space(60) & rsCli!codigo
    rsCli.MoveNext
Loop

VSOcup.Min = 0: VSOcup.Max = (lstCli.ListCount - 1)
nCliNum = 0    'CLIENTE 0 ES EFECTIVO
cmdCli.Enabled = False
cmdHotel.Enabled = False
End Sub
Private Sub cmdCancel_Click()
If IsObject(rsCli) Then
If rsCli.State = adStateOpen Then
    rsCli.Close
    Set rsCli = Nothing
End If
End If
cNombreCliente = ""
nCliNum = 0
Unload Me
End Sub
Private Sub cmdCli_Click()
'SELECCION DE CLIENTES REGULARES DEL RESTAURANTE
GetClientes 0
End Sub

Private Sub cmdHotel_Click()
'SELECCION DE CLIENTES HUESPEDES DEL HOTEL/MOTEL
GetClientes 1
End Sub

Private Sub cmdOK_Click()
If rsCli.State <> adStateOpen Then
    MsgBox "DEBE SELECCIONAR UN CLIENTE O UN HUESPED", vbExclamation, BoxTit
    Exit Sub
End If
On Error Resume Next
rsCli.MoveFirst
On Error GoTo 0
rsCli.Find "CODIGO =" & nCliNum
If Not rsCli.EOF Then cNombreCliente = rsCli!nombre & " " & rsCli!apellido Else cNombreCliente = ""
rsCli.Close
Set rsCli = Nothing
Unload Me
End Sub
Private Sub Form_Load()
Dim txt As String

Label1 = Label1 + rs00!descrip
End Sub

Private Sub lstCli_Click()
POSIC = Len(lstCli.Text)
nCliNum = Val(Mid(lstCli.Text, POSIC - 4, 5))
End Sub
Private Sub VSOcup_Change()
lstCli.ListIndex = VSOcup.Value
End Sub
