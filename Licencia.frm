VERSION 5.00
Begin VB.Form Licencia 
   BackColor       =   &H00B39665&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   4125
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   7050
   LinkTopic       =   "Form1"
   ScaleHeight     =   4125
   ScaleWidth      =   7050
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture 
      BackColor       =   &H00FFFFFF&
      Height          =   1140
      Left            =   120
      Picture         =   "Licencia.frx":0000
      ScaleHeight     =   1080
      ScaleWidth      =   1080
      TabIndex        =   6
      Top             =   120
      Width           =   1140
   End
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "ACEPTAR"
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
      Left            =   2880
      TabIndex        =   4
      Top             =   3360
      Width           =   1455
   End
   Begin VB.CommandButton cmdCancelar 
      Caption         =   "CANCELAR"
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
      Left            =   5535
      TabIndex        =   3
      Top             =   3360
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.TextBox txtLic2 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      IMEMode         =   3  'DISABLE
      Left            =   5520
      PasswordChar    =   "*"
      TabIndex        =   1
      Text            =   "txtLic1"
      Top             =   2040
      Width           =   1335
   End
   Begin VB.TextBox txtLic1 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      IMEMode         =   3  'DISABLE
      Left            =   3840
      PasswordChar    =   "*"
      TabIndex        =   0
      Text            =   "txtLic1"
      Top             =   2040
      Width           =   1335
   End
   Begin VB.Label lbMensaje 
      BackColor       =   &H00FF0000&
      Caption         =   "Gracias por utilizar nuestro Sistema, por favor introduzca su Código de  Activación para proceder"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   855
      Left            =   1560
      TabIndex        =   5
      Top             =   240
      Width           =   4575
   End
   Begin VB.Label Label 
      BackColor       =   &H00B39665&
      Caption         =   "Código de Activación"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   120
      TabIndex        =   2
      Top             =   2160
      Width           =   3375
   End
   Begin VB.Shape Shape 
      BackColor       =   &H00FF0000&
      BackStyle       =   1  'Opaque
      Height          =   1455
      Left            =   0
      Top             =   0
      Width           =   7095
   End
End
Attribute VB_Name = "Licencia"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
