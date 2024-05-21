VERSION 5.00
Begin VB.Form AskDesc 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "DESCUENTO PRE-CUENTA"
   ClientHeight    =   2865
   ClientLeft      =   2835
   ClientTop       =   3480
   ClientWidth     =   4095
   Icon            =   "AskDesc.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   1692.736
   ScaleMode       =   0  'User
   ScaleWidth      =   3844.984
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton Borrar 
      Height          =   495
      Left            =   1680
      Picture         =   "AskDesc.frx":000C
      Style           =   1  'Graphical
      TabIndex        =   15
      Top             =   2160
      Width           =   615
   End
   Begin VB.Frame Frame2 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Index           =   3
      Left            =   120
      TabIndex        =   4
      Top             =   720
      Width           =   3735
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
         Height          =   375
         Index           =   0
         Left            =   3000
         TabIndex        =   14
         Top             =   720
         Width           =   615
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
         Height          =   375
         Index           =   1
         Left            =   120
         TabIndex        =   13
         Top             =   180
         Width           =   615
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
         Height          =   375
         Index           =   2
         Left            =   840
         TabIndex        =   12
         Top             =   180
         Width           =   615
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
         Height          =   375
         Index           =   3
         Left            =   1560
         TabIndex        =   11
         Top             =   180
         Width           =   615
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
         Height          =   375
         Index           =   4
         Left            =   2280
         TabIndex        =   10
         Top             =   180
         Width           =   615
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
         Height          =   375
         Index           =   5
         Left            =   3000
         TabIndex        =   9
         Top             =   180
         Width           =   615
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
         Height          =   375
         Index           =   6
         Left            =   120
         TabIndex        =   8
         Top             =   720
         Width           =   615
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
         Height          =   375
         Index           =   7
         Left            =   840
         TabIndex        =   7
         Top             =   720
         Width           =   615
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
         Height          =   375
         Index           =   8
         Left            =   1560
         TabIndex        =   6
         Top             =   720
         Width           =   615
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
         Height          =   375
         Index           =   9
         Left            =   2280
         TabIndex        =   5
         Top             =   720
         Width           =   615
      End
   End
   Begin VB.TextBox txtLin 
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
      TabIndex        =   1
      Text            =   "0.00"
      Top             =   240
      Width           =   1125
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "&Aceptar"
      Default         =   -1  'True
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
      Left            =   120
      TabIndex        =   2
      Top             =   2160
      Width           =   1260
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "&Cancelar"
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
      Left            =   2640
      TabIndex        =   3
      Top             =   2160
      Width           =   1260
   End
   Begin VB.Label lblLabels 
      Caption         =   "Escribir el Monto para Aplicar Descuento"
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
      Left            =   120
      TabIndex        =   0
      Top             =   150
      Width           =   2385
   End
End
Attribute VB_Name = "AskDesc"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim nlPase As Integer
Dim nMntOculto As String

Private Sub Borrar_Click()
nlPase = 0
txtLin = Format(0#, "standard")
nMntOculto = ""
End Sub

Private Sub cmdCancel_Click()
Unload Me
End Sub

Private Sub cmdOK_Click()

If txtLin = "" Or Not IsNumeric(txtLin) Then
    BoxPreg = "Descuento NO es Valido!!"
    BoxResp = MsgBox(BoxPreg, vbOKOnly, BoxTit)
Else
    DescPreCta = txtLin
End If
Unload Me
End Sub

Private Sub Command8_Click(Index As Integer)
Dim cCant As String

If nlPase = 0 Then
    nMntOculto = Command8(Index).Index
Else
    nMntOculto = nMntOculto & Command8(Index).Caption
End If
txtLin = Format(Val(nMntOculto) / 100, "standard")
nlPase = nlPase + 1
End Sub

Private Sub Form_Load()
nlPase = 0
nMntOculto = ""
End Sub

Private Sub txtLin_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    cmdOK.SetFocus
End If
End Sub
