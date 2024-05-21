VERSION 5.00
Begin VB.Form CambiarMesa 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "CAMBIO DE MESA"
   ClientHeight    =   6135
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7455
   ControlBox      =   0   'False
   Icon            =   "CambiarMesa.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   6135
   ScaleWidth      =   7455
   StartUpPosition =   2  'CenterScreen
   Begin VB.VScrollBar VSDisp 
      Height          =   4455
      Left            =   6360
      TabIndex        =   7
      Top             =   600
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.VScrollBar VSOcup 
      Height          =   4455
      Left            =   2760
      TabIndex        =   6
      Top             =   600
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.CommandButton cmdSalir 
      Cancel          =   -1  'True
      Caption         =   "Regresar"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   5760
      TabIndex        =   4
      Top             =   5400
      Width           =   1455
   End
   Begin VB.ListBox lstDisponible 
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   21.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4365
      Left            =   3840
      TabIndex        =   1
      Top             =   600
      Width           =   2295
   End
   Begin VB.ListBox lstOcupada 
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   21.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4365
      Left            =   240
      TabIndex        =   0
      Top             =   600
      Width           =   2295
   End
   Begin VB.Label Label1 
      Caption         =   "Indicaciones : Toque primero la Mesa Ocupada, y despues la Disponible"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   615
      Index           =   2
      Left            =   240
      TabIndex        =   5
      Top             =   5280
      Width           =   4095
   End
   Begin VB.Label Label1 
      Caption         =   "Mesas Disponibles"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   255
      Index           =   1
      Left            =   3840
      TabIndex        =   3
      Top             =   240
      Width           =   2175
   End
   Begin VB.Label Label1 
      Caption         =   "Mesas Ocupadas"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   255
      Index           =   0
      Left            =   240
      TabIndex        =   2
      Top             =   240
      Width           =   1935
   End
End
Attribute VB_Name = "CambiarMesa"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rsOcup As Recordset
Dim rsDisp As Recordset
Dim nOcup As Integer, nDisp As Integer
Private Sub cmdSalir_Click()
Unload Me
End Sub

Private Sub Form_Load()
Dim nErrCount As Integer

On Error GoTo AdmErr:
nDisp = 0: nOcup = 0

Set rsOcup = New Recordset
Set rsDisp = New Recordset

rsOcup.Open "SELECT numero FROM MESAS WHERE ocupada = TRUE ORDER BY NUMERO", msConn, adOpenStatic, adLockOptimistic
rsDisp.Open "SELECT numero FROM MESAS WHERE ocupada = FALSE AND NUMERO <> -99 ORDER BY NUMERO", msConn, adOpenStatic, adLockOptimistic

Do Until rsOcup.EOF
    lstOcupada.AddItem rsOcup!numero
    rsOcup.MoveNext
Loop

Do Until rsDisp.EOF
    lstDisponible.AddItem rsDisp!numero
    rsDisp.MoveNext
Loop

VSOcup.Min = 0: VSOcup.Max = (lstOcupada.ListCount - 1)
VSDisp.Min = 0: VSDisp.Max = (lstDisponible.ListCount - 1)

On Error GoTo 0
Exit Sub

AdmErr:
nErrCount = nErrCount + 1
If nErrCount < 3 Then
    Resume
ElseIf nErrCount > 2 And nErrCount < 6 Then
    Resume Next
Else
    MsgBox "IMPOSIBLE MOSTRAR LAS MESAS EN ESTE MOMENTO, INTENTE MAS TARDE", vbCritical, BoxTit
    Exit Sub
End If
End Sub

Private Sub lstDisponible_Click()
On Error GoTo ErrAdm:
nDisp = Val(lstDisponible.Text)
If nOcup = 0 Then
    MsgBox "NO HA SELECCIONADO UNA MESA OCUPADA", vbExclamation, BoxTit
    lstOcupada.SetFocus
Else
    msConn.BeginTrans
    msConn.Execute "UPDATE MESAS SET OCUPADA = FALSE WHERE NUMERO = " & nOcup
    msConn.Execute "UPDATE TMP_TRANS SET MESA = " & nDisp & " WHERE MESA = " & nOcup
    msConn.Execute "UPDATE MESAS SET OCUPADA = TRUE WHERE NUMERO = " & nDisp
    msConn.Execute "UPDATE TMP_PAR_PAGO SET MESA = " & nDisp & " WHERE MESA = " & nOcup
    msConn.Execute "UPDATE TMP_PAR_PROP SET MESA = " & nDisp & " WHERE MESA = " & nOcup
    msConn.Execute "UPDATE TMP_CUENTAS SET MESA = " & nDisp & " WHERE MESA = " & nOcup
    msConn.CommitTrans
    StatMesa nOcup, 0
    Unload Me
End If
Exit Sub

ErrAdm:
If Err.Number = -2147467259 Then
    MsgBox "NO ES POSIBLE MOVERSE A ESTA MESA EN ESTE MOMENTO", vbInformation, "INTENTELO MAS TARDE"
Else
    MsgBox Err.Description, vbCritical, BoxTit
End If
msConn.RollbackTrans
Exit Sub
End Sub

Private Sub lstOcupada_Click()
nOcup = Val(lstOcupada.Text)
End Sub

Private Sub VSDisp_Change()
lstDisponible.ListIndex = VSDisp.Value
End Sub

Private Sub VSOcup_Change()
lstOcupada.ListIndex = VSOcup.Value
End Sub
