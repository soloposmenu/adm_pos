VERSION 5.00
Begin VB.Form Opc01 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "SELECCIONE OPCIONES DE LA PRE-CUENTA"
   ClientHeight    =   6105
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9885
   ControlBox      =   0   'False
   Icon            =   "Opc01.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   6105
   ScaleWidth      =   9885
   StartUpPosition =   2  'CenterScreen
   Begin VB.CheckBox Check3 
      Caption         =   "Cargo por Servicio Habitación"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   240
      TabIndex        =   42
      Top             =   1440
      Width           =   3615
   End
   Begin VB.Frame Frame4 
      BackColor       =   &H00800000&
      Caption         =   "Billetes mas Frecuentes para anotar Descuento"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   2175
      Left            =   5160
      TabIndex        =   26
      Top             =   600
      Width           =   4455
      Begin VB.CommandButton cdmBill 
         Height          =   855
         Index           =   2
         Left            =   120
         Picture         =   "Opc01.frx":030A
         Style           =   1  'Graphical
         TabIndex        =   30
         Tag             =   "15.00"
         Top             =   1200
         Width           =   2055
      End
      Begin VB.CommandButton cdmBill 
         Height          =   855
         Index           =   3
         Left            =   2280
         Picture         =   "Opc01.frx":6FF4
         Style           =   1  'Graphical
         TabIndex        =   29
         Tag             =   "20.00"
         Top             =   1200
         Width           =   2055
      End
      Begin VB.CommandButton cdmBill 
         Height          =   855
         Index           =   0
         Left            =   120
         Picture         =   "Opc01.frx":DEA6
         Style           =   1  'Graphical
         TabIndex        =   28
         Tag             =   "5.00"
         Top             =   240
         Width           =   2055
      End
      Begin VB.CommandButton cdmBill 
         Height          =   855
         Index           =   1
         Left            =   2280
         Picture         =   "Opc01.frx":14B90
         Style           =   1  'Graphical
         TabIndex        =   27
         Tag             =   "10.00"
         Top             =   240
         Width           =   2055
      End
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00008000&
      Caption         =   "DESCUENTO"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   2895
      Left            =   5400
      TabIndex        =   20
      Top             =   3000
      Width           =   3975
      Begin VB.CommandButton Command4 
         Caption         =   "&Imprimir Pre Cuenta"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   630
         Left            =   120
         TabIndex        =   24
         Top             =   2130
         Width           =   1380
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         IMEMode         =   3  'DISABLE
         Left            =   2640
         TabIndex        =   23
         Text            =   "0.00"
         Top             =   330
         Width           =   1125
      End
      Begin VB.Frame Frame2 
         BackColor       =   &H00008000&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1335
         Index           =   0
         Left            =   120
         TabIndex        =   22
         Top             =   690
         Width           =   3735
         Begin VB.CommandButton Command6 
            Caption         =   "0"
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
            Index           =   0
            Left            =   2940
            TabIndex        =   32
            Top             =   720
            Width           =   735
         End
         Begin VB.CommandButton Command6 
            Caption         =   "9"
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
            Index           =   9
            Left            =   2220
            TabIndex        =   41
            Top             =   720
            Width           =   735
         End
         Begin VB.CommandButton Command6 
            Caption         =   "8"
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
            Index           =   8
            Left            =   1500
            TabIndex        =   40
            Top             =   720
            Width           =   735
         End
         Begin VB.CommandButton Command6 
            Caption         =   "7"
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
            Index           =   7
            Left            =   780
            TabIndex        =   39
            Top             =   720
            Width           =   735
         End
         Begin VB.CommandButton Command6 
            Caption         =   "6"
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
            Index           =   6
            Left            =   60
            TabIndex        =   38
            Top             =   720
            Width           =   735
         End
         Begin VB.CommandButton Command6 
            Caption         =   "5"
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
            Index           =   5
            Left            =   2940
            TabIndex        =   37
            Top             =   240
            Width           =   735
         End
         Begin VB.CommandButton Command6 
            Caption         =   "4"
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
            Index           =   4
            Left            =   2220
            TabIndex        =   36
            Top             =   240
            Width           =   735
         End
         Begin VB.CommandButton Command6 
            Caption         =   "3"
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
            Index           =   3
            Left            =   1500
            TabIndex        =   35
            Top             =   240
            Width           =   735
         End
         Begin VB.CommandButton Command6 
            Caption         =   "2"
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
            Index           =   2
            Left            =   780
            TabIndex        =   34
            Top             =   240
            Width           =   735
         End
         Begin VB.CommandButton Command6 
            Caption         =   "1"
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
            Index           =   1
            Left            =   60
            TabIndex        =   33
            Top             =   240
            Width           =   735
         End
      End
      Begin VB.CommandButton Command3 
         Height          =   615
         Left            =   2760
         Picture         =   "Opc01.frx":1B87A
         Style           =   1  'Graphical
         TabIndex        =   21
         Top             =   2130
         Width           =   1095
      End
      Begin VB.Label Label1 
         BackColor       =   &H00008000&
         Caption         =   "Escribir el Monto para Anotar Descuento"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   615
         Left            =   120
         TabIndex        =   25
         Top             =   240
         Width           =   2385
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "CONTRASEÑA"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2895
      Left            =   240
      TabIndex        =   4
      Top             =   3000
      Width           =   4935
      Begin VB.PictureBox PictOK 
         BorderStyle     =   0  'None
         Height          =   495
         Left            =   3840
         Picture         =   "Opc01.frx":1BCBC
         ScaleHeight     =   495
         ScaleWidth      =   495
         TabIndex        =   31
         Top             =   240
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.CommandButton Borrar 
         Height          =   495
         Left            =   3000
         Picture         =   "Opc01.frx":1C0FE
         Style           =   1  'Graphical
         TabIndex        =   18
         Top             =   2280
         Width           =   1095
      End
      Begin VB.Frame Frame2 
         BackColor       =   &H00C0C0C0&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1335
         Index           =   3
         Left            =   240
         TabIndex        =   7
         Top             =   810
         Width           =   3855
         Begin VB.CommandButton Command8 
            Caption         =   "0"
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
            Index           =   0
            Left            =   3000
            TabIndex        =   17
            Top             =   720
            Width           =   735
         End
         Begin VB.CommandButton Command8 
            Caption         =   "1"
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
            Index           =   1
            Left            =   120
            TabIndex        =   16
            Top             =   180
            Width           =   735
         End
         Begin VB.CommandButton Command8 
            Caption         =   "2"
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
            Index           =   2
            Left            =   840
            TabIndex        =   15
            Top             =   180
            Width           =   735
         End
         Begin VB.CommandButton Command8 
            Caption         =   "3"
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
            Index           =   3
            Left            =   1560
            TabIndex        =   14
            Top             =   180
            Width           =   735
         End
         Begin VB.CommandButton Command8 
            Caption         =   "4"
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
            Index           =   4
            Left            =   2280
            TabIndex        =   13
            Top             =   180
            Width           =   735
         End
         Begin VB.CommandButton Command8 
            Caption         =   "5"
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
            Index           =   5
            Left            =   3000
            TabIndex        =   12
            Top             =   180
            Width           =   735
         End
         Begin VB.CommandButton Command8 
            Caption         =   "6"
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
            Index           =   6
            Left            =   120
            TabIndex        =   11
            Top             =   720
            Width           =   735
         End
         Begin VB.CommandButton Command8 
            Caption         =   "7"
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
            Index           =   7
            Left            =   840
            TabIndex        =   10
            Top             =   720
            Width           =   735
         End
         Begin VB.CommandButton Command8 
            Caption         =   "8"
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
            Index           =   8
            Left            =   1560
            TabIndex        =   9
            Top             =   720
            Width           =   735
         End
         Begin VB.CommandButton Command8 
            Caption         =   "9"
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
            Index           =   9
            Left            =   2280
            TabIndex        =   8
            Top             =   720
            Width           =   735
         End
      End
      Begin VB.TextBox txtLin 
         Alignment       =   1  'Right Justify
         Height          =   345
         IMEMode         =   3  'DISABLE
         Left            =   2760
         PasswordChar    =   "@"
         TabIndex        =   6
         Top             =   330
         Width           =   1125
      End
      Begin VB.CommandButton cmdOK 
         Caption         =   "&Aceptar"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   510
         Left            =   240
         TabIndex        =   5
         Top             =   2250
         Width           =   1260
      End
      Begin VB.Image PictApunta2 
         Height          =   675
         Left            =   4200
         Picture         =   "Opc01.frx":1C540
         Stretch         =   -1  'True
         Top             =   1080
         Visible         =   0   'False
         Width           =   720
      End
      Begin VB.Label lblLabels 
         Caption         =   "Anote Contraseña para Descuento"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   495
         Left            =   240
         TabIndex        =   19
         Top             =   240
         Width           =   2400
      End
   End
   Begin VB.CommandButton Cancelar 
      Cancel          =   -1  'True
      Caption         =   "&Cancelar"
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
      Left            =   2880
      TabIndex        =   3
      Top             =   2160
      Width           =   1575
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Imprimir Pre Cuenta"
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
      Left            =   240
      TabIndex        =   2
      Top             =   2160
      Width           =   1575
   End
   Begin VB.CheckBox Check2 
      Caption         =   "¿ Desea Anotar un Descuento Global ?"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   240
      TabIndex        =   1
      Top             =   840
      Width           =   4335
   End
   Begin VB.CheckBox Check1 
      Caption         =   "¿ Desea Marcar el 10 % de Propina ?"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   4215
   End
   Begin VB.Image PictApunta 
      Height          =   915
      Left            =   2040
      Picture         =   "Opc01.frx":1C982
      Stretch         =   -1  'True
      Top             =   1920
      Visible         =   0   'False
      Width           =   600
   End
End
Attribute VB_Name = "Opc01"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private nClavPase As Integer
Private nDescPase As Integer
Private nMntOculto As String

Private Sub Borrar_Click()
nClavPase = 0
txtLin = ""
End Sub

Private Sub cdmBill_Click(Index As Integer)
Text1 = cdmBill(Index).Tag
End Sub

Private Sub Check1_Click()
If Check1.Value = 1 Then OKProp = 1 Else OKProp = 0
End Sub

Private Sub Check2_Click()
If Check2.Value = 1 Then        'SI SELECCIONO DESCUENTO
    PictApunta.Visible = True
    Frame1.Enabled = True
    lblLabels.Visible = True
    txtLin.SetFocus
ElseIf Check2.Value = 0 Then    'NO HAY DESCUENTO
    Frame1.Enabled = False: Frame3.Enabled = False
    Frame4.Enabled = False
    PictApunta.Visible = False
End If
End Sub

Private Sub cmdOK_Click()
Dim rsUsr As Recordset
If txtLin = "" Or Not IsNumeric(txtLin) Then
    MsgBox "!! CLAVE NO ES VALIDA !!", vbExclamation, BoxTit
Else
    Set rsUsr = New Recordset
    rsUsr.Open "SELECT numero,nombre FROM USUARIOS " & _
        " WHERE CLAVE = " & "'" & LTrim(txtLin) & "'", msConn, adOpenForwardOnly, adLockReadOnly
    If Not rsUsr.EOF Then
        OKDesc = 1
        PictOK.Visible = True: PictApunta2.Visible = True
        Frame3.Enabled = True
        Frame4.Enabled = True
        Text1.SetFocus
        rsUsr.Close
    Else
        MsgBox "CLAVE PARA DESCUENTO ES INVALIDA", vbCritical, BoxTit
        txtLin = ""
        txtLin.SetFocus
    End If
End If
End Sub
Private Sub Cancelar_Click()
OKDesc = 0
OKProp = 0
OKCancelar = 1
Unload Me
End Sub

Private Sub Command1_Click()
Command4_Click
End Sub

Private Sub Command3_Click()
nDescPase = 0
Text1 = 0#
End Sub

Private Sub Command4_Click()
DescPreCta = Format(Text1, "standard")
Unload Me
End Sub

Private Sub Command6_Click(Index As Integer)
Dim cCant As String

If nDescPase = 0 Then
    nMntOculto = Command6(Index).Index
Else
    nMntOculto = nMntOculto & Command6(Index).Index
End If
Text1 = Format(Val(nMntOculto) / 100, "standard")
nDescPase = nDescPase + 1

End Sub

Private Sub Command8_Click(Index As Integer)
Dim cCant As String

If nClavPase = 0 Then
    txtLin = Command8(Index).Index
Else
    cCant = CStr(txtLin)
    cCant = cCant & Command8(Index).Index
    txtLin = cCant
End If
nClavPase = nClavPase + 1
End Sub

Private Sub Form_Load()
nClavPase = 0
nDescPase = 0
lblLabels.Visible = False
nMntOculto = ""
DescPreCta = 0#
If HABITACION_OK = False Then
    Check3.Visible = False
End If
End Sub
Private Sub txtLin_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then cmdOK.SetFocus
End Sub
