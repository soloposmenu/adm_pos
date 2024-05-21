VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3825
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5655
   LinkTopic       =   "Form1"
   ScaleHeight     =   3825
   ScaleWidth      =   5655
   StartUpPosition =   3  'Windows Default
   Begin Project1.uctClassClientesDataGrid uctClassClientesDataGrid1 
      Height          =   1215
      Left            =   120
      TabIndex        =   2
      Top             =   1440
      Width           =   4215
      _ExtentX        =   7435
      _ExtentY        =   2143
      ManualInitialize=   0   'False
      GridEditable    =   0   'False
      SaveMode        =   0
   End
   Begin Project1.uctclsUnidMedidaComboBox uctclsUnidMedidaComboBox1 
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Top             =   840
      Width           =   2535
      _ExtentX        =   4471
      _ExtentY        =   661
      ManualInitialize=   0   'False
      NoneFirst       =   0   'False
   End
   Begin Project1.uctClassDepInvComboBox uctClassDepInvComboBox1 
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   2055
      _ExtentX        =   3625
      _ExtentY        =   661
      ManualInitialize=   0   'False
      NoneFirst       =   0   'False
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
