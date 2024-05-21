VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form PLU 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "MODULO DE VENTAS"
   ClientHeight    =   8625
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11910
   ControlBox      =   0   'False
   ForeColor       =   &H00000000&
   Icon            =   "PLU.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   8625
   ScaleWidth      =   11910
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin VB.CommandButton Command13 
      BackColor       =   &H000000FF&
      Caption         =   "CORTESIA"
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
      Left            =   2040
      Style           =   1  'Graphical
      TabIndex        =   56
      Top             =   2640
      Width           =   1245
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Enabled         =   0   'False
      Height          =   495
      Left            =   11280
      TabIndex        =   55
      Top             =   5760
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00808080&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   3800
      Index           =   1
      Left            =   1920
      TabIndex        =   5
      Top             =   3120
      Width           =   7335
      Begin VB.CommandButton cmdPlus 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   530
         Index           =   0
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   240
         Width           =   2295
      End
   End
   Begin VB.CommandButton Command13 
      Caption         =   "Descuento/Plato"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   650
      Index           =   0
      Left            =   2040
      TabIndex        =   54
      Top             =   7080
      Width           =   1215
   End
   Begin VB.CommandButton Command13 
      Caption         =   "Anulación"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   650
      Index           =   1
      Left            =   3390
      Picture         =   "PLU.frx":0442
      Style           =   1  'Graphical
      TabIndex        =   53
      Top             =   7080
      Width           =   1215
   End
   Begin VB.CommandButton Command13 
      Caption         =   "Pre-Cuenta"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   650
      Index           =   3
      Left            =   2040
      Picture         =   "PLU.frx":0884
      Style           =   1  'Graphical
      TabIndex        =   52
      Top             =   7845
      Width           =   1215
   End
   Begin VB.CommandButton Command13 
      Appearance      =   0  'Flat
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
      Height          =   650
      Index           =   4
      Left            =   4710
      Picture         =   "PLU.frx":0B8E
      Style           =   1  'Graphical
      TabIndex        =   51
      Top             =   7845
      Width           =   1455
   End
   Begin VB.CommandButton Command13 
      Caption         =   "Reporte X"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   650
      Index           =   2
      Left            =   4710
      Picture         =   "PLU.frx":102D
      Style           =   1  'Graphical
      TabIndex        =   50
      Top             =   7080
      Width           =   1455
   End
   Begin VB.CommandButton Command13 
      BackColor       =   &H80000018&
      Caption         =   "Pago Parcial"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   650
      Index           =   5
      Left            =   3390
      Style           =   1  'Graphical
      TabIndex        =   49
      Top             =   7845
      Width           =   1215
   End
   Begin VB.CommandButton Command8 
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   0
      Left            =   9480
      TabIndex        =   48
      Top             =   7935
      Width           =   735
   End
   Begin VB.CommandButton Command8 
      Caption         =   "1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   1
      Left            =   9480
      TabIndex        =   47
      Top             =   6480
      Width           =   735
   End
   Begin VB.CommandButton Command8 
      Caption         =   "2"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   2
      Left            =   10200
      TabIndex        =   46
      Top             =   6480
      Width           =   735
   End
   Begin VB.CommandButton Command8 
      Caption         =   "3"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   3
      Left            =   10920
      TabIndex        =   45
      Top             =   6480
      Width           =   735
   End
   Begin VB.CommandButton Command8 
      Caption         =   "4"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   4
      Left            =   9480
      TabIndex        =   44
      Top             =   6960
      Width           =   735
   End
   Begin VB.CommandButton Command8 
      Caption         =   "5"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   5
      Left            =   10200
      TabIndex        =   43
      Top             =   6960
      Width           =   735
   End
   Begin VB.CommandButton Command8 
      Caption         =   "6"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   6
      Left            =   10920
      TabIndex        =   42
      Top             =   6960
      Width           =   735
   End
   Begin VB.CommandButton Command8 
      Caption         =   "7"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   7
      Left            =   9480
      TabIndex        =   41
      Top             =   7440
      Width           =   735
   End
   Begin VB.CommandButton Command8 
      Caption         =   "8"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   8
      Left            =   10200
      TabIndex        =   40
      Top             =   7440
      Width           =   735
   End
   Begin VB.CommandButton Command8 
      Caption         =   "9"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   9
      Left            =   10920
      TabIndex        =   39
      Top             =   7440
      Width           =   735
   End
   Begin VB.CommandButton Correccion 
      Caption         =   "Correción"
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
      Left            =   10320
      Picture         =   "PLU.frx":1337
      Style           =   1  'Graphical
      TabIndex        =   38
      Top             =   7935
      Width           =   1335
   End
   Begin VB.CommandButton cmdRestoAco 
      DisabledPicture =   "PLU.frx":1779
      Enabled         =   0   'False
      Height          =   495
      Left            =   9900
      Picture         =   "PLU.frx":1BBB
      Style           =   1  'Graphical
      TabIndex        =   37
      Top             =   5800
      Width           =   1215
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H0080C0FF&
      Caption         =   "GENERAL"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3135
      Index           =   6
      Left            =   9840
      TabIndex        =   19
      Top             =   0
      Width           =   2055
      Begin VB.CommandButton Clear 
         Caption         =   "CLEAR"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   60
         TabIndex        =   30
         Top             =   1245
         Width           =   850
      End
      Begin VB.TextBox Text1 
         Alignment       =   2  'Center
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Index           =   2
         Left            =   960
         TabIndex        =   27
         Top             =   1485
         Width           =   975
      End
      Begin VB.TextBox Text1 
         Alignment       =   2  'Center
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
         Height          =   285
         Index           =   0
         Left            =   840
         TabIndex        =   25
         Top             =   195
         Width           =   495
      End
      Begin VB.TextBox Text1 
         Alignment       =   2  'Center
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
         Height          =   285
         Index           =   3
         Left            =   840
         TabIndex        =   23
         Top             =   840
         Width           =   1095
      End
      Begin VB.TextBox Text1 
         Alignment       =   2  'Center
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
         Height          =   285
         Index           =   1
         Left            =   840
         TabIndex        =   20
         Top             =   520
         Width           =   1095
      End
      Begin VB.Label Label1 
         BackColor       =   &H0080C0FF&
         Caption         =   "Cuenta Actual"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   5
         Left            =   60
         TabIndex        =   33
         Top             =   1995
         Width           =   1335
      End
      Begin VB.Label lbCuenta 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   1440
         TabIndex        =   32
         Top             =   1900
         Width           =   495
      End
      Begin VB.Label lbMensaje 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
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
         Height          =   465
         Left            =   60
         TabIndex        =   29
         Top             =   2280
         Width           =   1885
         WordWrap        =   -1  'True
      End
      Begin VB.Label Label1 
         BackColor       =   &H0080C0FF&
         Caption         =   "CANT."
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   3
         Left            =   1200
         TabIndex        =   28
         Top             =   1250
         Width           =   615
      End
      Begin VB.Label Label1 
         BackColor       =   &H0080C0FF&
         Caption         =   "Mesa #"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   1
         Left            =   60
         TabIndex        =   26
         Top             =   240
         Width           =   735
      End
      Begin VB.Label Label1 
         BackColor       =   &H0080C0FF&
         Caption         =   "Cajer@"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   4
         Left            =   60
         TabIndex        =   24
         Top             =   900
         Width           =   855
      End
      Begin VB.Label Hora 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00000000&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "*"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFF00&
         Height          =   300
         Left            =   840
         TabIndex        =   22
         Top             =   2760
         Width           =   1095
      End
      Begin VB.Label Label1 
         BackColor       =   &H0080C0FF&
         Caption         =   "Meser@"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   2
         Left            =   60
         TabIndex        =   21
         Top             =   550
         Width           =   855
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00C00000&
      Caption         =   "Acompañantes"
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
      Height          =   2620
      Index           =   2
      Left            =   9240
      TabIndex        =   9
      Top             =   3120
      Width           =   2520
      Begin VB.CommandButton cmdAcomp 
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
         Left            =   120
         TabIndex        =   36
         Top             =   240
         Width           =   2295
      End
   End
   Begin VB.CommandButton cmdCtas 
      BackColor       =   &H0080FFFF&
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
      Left            =   4680
      Picture         =   "PLU.frx":1FFD
      Style           =   1  'Graphical
      TabIndex        =   31
      Top             =   2520
      Width           =   1095
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid PlatosMesa 
      Height          =   2175
      Left            =   3480
      TabIndex        =   16
      TabStop         =   0   'False
      Top             =   240
      Width           =   6375
      _ExtentX        =   11245
      _ExtentY        =   3836
      _Version        =   393216
      BackColor       =   0
      ForeColor       =   65280
      Rows            =   1
      Cols            =   5
      FixedRows       =   0
      FixedCols       =   0
      GridColor       =   16777215
      AllowBigSelection=   0   'False
      SelectionMode   =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _NumberOfBands  =   1
      _Band(0).Cols   =   5
      _Band(0).GridLinesBand=   1
      _Band(0).TextStyleBand=   0
      _Band(0).TextStyleHeader=   0
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00C0C0C0&
      Caption         =   "TAMAÑO"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   2600
      Index           =   4
      Left            =   1920
      TabIndex        =   10
      Top             =   0
      Width           =   1455
      Begin VB.CommandButton cmdEnvases 
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
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   15
         Top             =   2040
         Width           =   1215
      End
      Begin VB.CommandButton cmdEnvases 
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
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   14
         Top             =   1440
         Width           =   1215
      End
      Begin VB.CommandButton cmdEnvases 
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
         Style           =   1  'Graphical
         TabIndex        =   13
         Top             =   840
         Width           =   1215
      End
      Begin VB.CommandButton cmdEnvases 
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
         Height          =   480
         Index           =   0
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   11
         Top             =   240
         Width           =   1200
      End
   End
   Begin VB.CommandButton cmdSelMesa 
      Caption         =   "MESAS"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   3480
      TabIndex        =   17
      Top             =   2520
      Width           =   1095
   End
   Begin VB.CommandButton cmdSalir 
      Caption         =   "Salir"
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
      Left            =   6720
      TabIndex        =   12
      Top             =   7920
      Width           =   2535
   End
   Begin VB.CommandButton Command6 
      BackColor       =   &H0000C0C0&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   8040
      Picture         =   "PLU.frx":243F
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   6960
      Width           =   1215
   End
   Begin VB.CommandButton Command5 
      BackColor       =   &H00C0C000&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   6720
      Picture         =   "PLU.frx":3741
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   6960
      Width           =   1215
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00008080&
      Caption         =   "Departamentos"
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
      Height          =   8415
      Index           =   0
      Left            =   0
      TabIndex        =   1
      Top             =   120
      Width           =   1815
      Begin VB.CommandButton Command3 
         BackColor       =   &H00C0C000&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   800
         Left            =   50
         Picture         =   "PLU.frx":4A43
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   7560
         Width           =   855
      End
      Begin VB.CommandButton Command2 
         BackColor       =   &H0000C0C0&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   800
         Left            =   920
         Picture         =   "PLU.frx":5D45
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   7560
         Width           =   855
      End
      Begin VB.CommandButton cmdDepto 
         Appearance      =   0  'Flat
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
         Height          =   635
         Index           =   0
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   240
         Width           =   1575
      End
   End
   Begin VB.CommandButton cmdSlip 
      Caption         =   "CHEF"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   5880
      TabIndex        =   35
      Top             =   2520
      Width           =   975
   End
   Begin VB.Shape Shape2 
      BackColor       =   &H00008000&
      BackStyle       =   1  'Opaque
      Height          =   1575
      Left            =   1920
      Top             =   6990
      Width           =   4335
   End
   Begin VB.Shape Shape1 
      FillColor       =   &H00800080&
      FillStyle       =   0  'Solid
      Height          =   2295
      Left            =   9360
      Top             =   6360
      Width           =   2415
   End
   Begin VB.Label SubTot 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00000000&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   495
      Left            =   7920
      TabIndex        =   34
      Top             =   2520
      Width           =   1905
   End
   Begin VB.Label Label2 
      Caption         =   "Lin    Platos en la Mesa                            Cant    P.Unit       Precio"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   3480
      TabIndex        =   18
      Top             =   0
      Width           =   8055
   End
   Begin VB.Label Label1 
      Caption         =   "Sub-Tot"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   0
      Left            =   6960
      TabIndex        =   0
      Top             =   2730
      Width           =   975
   End
End
Attribute VB_Name = "PLU"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private num As Integer
Private numplu As Integer
Private nNLinSel As Integer
Private Arreg_Deptos(10) As Long
Private Arreg_Plu(17) As Integer
Private nPase As Integer 'Cantidad de Clicks a Cantidad
Private ElDepto As Long 'Es el Departamento Seleccionado
Private nGlobEnv As Long    'El envase seleccionado
Private TextEnv As String
Private rsTmpAco As New ADODB.Recordset
Private nAcoBookMark As Variant
Private lGo As Boolean
Private nCortesia As Integer
'REEMPLAZO DE ####.## por #0.00 CUANDO Sea Necesario
Private Sub AddOpenDeptItem()
'Agrega registros a TMP_TRANS desde un Departamento Abierto
Dim SOLO_FECHA As String

ValOpenDept = 0
TXT_OPEN_DEPT = ""

ActHost.Show 1
If TXT_OPEN_DEPT = "" Then Exit Sub

CajLin = CajLin + 1
SOLO_FECHA = Format(Date, "YYYYMMDD")
    
CadenaSql = "INSERT INTO TMP_TRANS " & _
    "(CAJA,CAJERO,MESA,MESERO,VALID,LIN,DESCRIP,CANT,DEPTO,PLU,ENVASE,PRECIO_UNIT,PRECIO,FECHA,HORA,TIPO,DESCUENTO,CUENTA,IMPRESO,IMPRESORA) VALUES (" & _
    cCaja & "," & npNumCaj & "," & nMesa & "," & nMesero & "," & -1 & "," & CajLin & ",'" & _
    rs02!corto & TXT_OPEN_DEPT & "'," & nMulti & "," & rs02!CODIGO & "," & 0 & "," & 0 & "," & _
    Format((ValOpenDept / nMulti) / 100, "#0.00") & "," & Format(ValOpenDept / 100, "#0.00") & ",'" & _
    SOLO_FECHA & "','" & Time & "','  '," & 0# & "," & nCta & ",FALSE," & BARRA_01 & ")"

msConn.BeginTrans
msConn.Execute CadenaSql
msConn.CommitTrans

If CajLin = 1 Then msConn.Execute "UPDATE Mesas SET ocupada = TRUE WHERE numero = " & nMesa

rs07.Open "SELECT a.lin,a.descrip,a.cant," & _
    " format(precio_unit,'##0.00') as mPrecio_unit," & _
    " format(precio,'##0.00') as mPrecio," & _
    " a.cajero,a.mesero,a.depto,a.plu,a.mesa,a.valid, " & _
    " a.envase,a.fecha,a.hora,a.tipo,a.descuento,a.cuenta, " & _
    " a.caja " & _
    " FROM tmp_trans as a " & _
    " WHERE a.mesa = " & nMesa & _
    " ORDER BY a.lin", msConn, adOpenStatic, adLockOptimistic

Set PlatosMesa.DataSource = rs07
SetupPantalla

nLineas = PlatosMesa.Rows - 1

Set rsParciales = New Recordset
rsParciales.Open "SELECT MESA,SUM(MONTO) AS VALOR FROM TMP_PAR_PAGO " & _
    " WHERE MESA = " & nMesa & _
    " GROUP BY MESA", msConn, adOpenDynamic, adLockOptimistic

If rsParciales.EOF Then lParc = 0 Else lParc = 1

rs07.Close
rs07.Open "SELECT sum(a.precio) as precio FROM tmp_trans as a " & _
    " WHERE a.mesa = " & nMesa, msConn, adOpenStatic, adLockReadOnly
SubTot = FormatCurrency(IIf(IsNull(rs07!precio), 0#, rs07!precio), 2)
On Error Resume Next
SubTot = FormatCurrency((SubTot + (rs07!precio * iISC)), 2)
iISCTransaccion = rs07!precio * iISC
SBTot = Format(SubTot, "standard")
On Error GoTo 0
rs07.Close
If (PlatosMesa.Rows - 1) >= 1 Then
    PlatosMesa.TopRow = (PlatosMesa.Rows - 1)
End If
nCantidad = 1: nPase = 0
nNLinSel = 0
Text1(2) = nCantidad

If lParc = 1 Then
    PlatosMesa.AddItem (CajLin + 1) & Chr(9) & _
        "PAGO.PARCIAL" & Chr(9) & 1 & Chr(9) & _
        Format(rsParciales!VALOR, "STANDARD") & Chr(9) & Format(rsParciales!VALOR, "STANDARD")
    SubTot = Format(SubTot - rsParciales!VALOR, "STANDARD")
    lbMensaje.BackColor = &HFFFF00
    lbMensaje = "MESA CON PAGOS PARCIALES"
Else
    lbMensaje.BackColor = &HFFFFFF
    lbMensaje = ""
End If

End Sub
Private Sub ImprPreCta()
Dim sqltext As String
Dim LinTx As String
Dim rsPreCta As Recordset
Dim MiMatriz(0, 3) As String
Dim MiLen1, Milen2 As Integer
Dim MiPropina As Single, nProp As Single
Dim rsParciales As Recordset
Dim lParc As Integer
Dim nLinCta As Integer
Dim nTotCta As Single
Dim nTotProp As Single
Dim nErrCount As Integer
Dim i As Integer
Dim txtString As String
Dim nTempISC As Single
Dim nTempSubTot As Single
Dim nTempSubFinal As Single

Set rsPreCta = New Recordset
Set rsParciales = New Recordset

OKCancelar = 0
If OPEN_PROPINA = True Then
    Opc01.Show 1
End If

nTotProp = 0#

On Error GoTo AdmErr:
'PROPINA_DESCRIP
If OKCancelar = 1 Then OKCancelar = 0: Exit Sub

rsParciales.Open "SELECT MESA,SUM(MONTO) AS VALOR " & _
        " FROM TMP_PAR_PAGO " & _
        " WHERE MESA = " & nMesa & _
        " GROUP BY MESA", msConn, adOpenDynamic, adLockOptimistic

'VERIFICA SI HAY PAGOS PARCIALES
If rsParciales.EOF Then lParc = 0 Else lParc = 1

sqltext = "SELECT MESA,CUENTA,LIN,DESCRIP,CANT,PRECIO " & _
        " FROM TMP_TRANS " & _
        " WHERE MESA = " & nMesa & _
        " AND CUENTA = " & nCta & _
        " ORDER BY CUENTA,LIN "

rsPreCta.Open sqltext, msConn, adOpenStatic, adLockReadOnly
If rsPreCta.EOF Then
    MsgBox "NO HAY PLATOS EN LA MESA. PRE-CUENTA NO SE IMPRIMIRA", vbInformation, BoxTit
    Exit Sub
End If
rsPreCta.MoveFirst
nLinCta = rsPreCta!CUENTA

Sys_Pos.Coptr1.PrintNormal PTR_S_RECEIPT, rs00!DESCRIP & Chr(&HD) & Chr(&HA)
Sys_Pos.Coptr1.PrintNormal PTR_S_RECEIPT, rs00!RAZ_SOC & Chr(&HD) & Chr(&HA)
Sys_Pos.Coptr1.PrintNormal PTR_S_RECEIPT, "RUC:" & rs00!RUC & Chr(&HD) & Chr(&HA)
Sys_Pos.Coptr1.PrintNormal PTR_S_RECEIPT, "SERIAL:" & rs00!SERIAL & Chr(&HD) & Chr(&HA)
Sys_Pos.Coptr1.PrintNormal PTR_S_RECEIPT, Space(2) & Chr(&HD) & Chr(&HA)
Sys_Pos.Coptr1.PrintNormal PTR_S_RECEIPT, "         PRE-CUENTA" & Chr(&HD) & Chr(&HA)
Sys_Pos.Coptr1.PrintNormal PTR_S_RECEIPT, Format(Date, "short date") & Chr(&HD) & Chr(&HA)
Sys_Pos.Coptr1.PrintNormal PTR_S_RECEIPT, "Mesero : " & cNomMesero & Chr(&HD) & Chr(&HA)
Sys_Pos.Coptr1.PrintNormal PTR_S_RECEIPT, Space(2) & Chr(&HD) & Chr(&HA)
Sys_Pos.Coptr1.PrintNormal PTR_S_RECEIPT, "Mesa : " & nMesa & "      " & Time & Chr(&HD) & Chr(&HA)
Sys_Pos.Coptr1.PrintNormal PTR_S_RECEIPT, Space(2) & Chr(&HD) & Chr(&HA)
If nLinCta <> 0 Then Sys_Pos.Coptr1.PrintNormal PTR_S_RECEIPT, "Cuenta # : " & nLinCta & Chr(&HD) & Chr(&HA)

Do Until rsPreCta.EOF
    Do Until rsPreCta.EOF
        MiMatriz(0, 0) = FormatTexto(rsPreCta!DESCRIP, 15)
        MiMatriz(0, 1) = Format(rsPreCta!CANT, "general number")
        MiMatriz(0, 2) = Format(rsPreCta!precio, "#,###.00")
        MiLen1 = Len(MiMatriz(0, 1))
        Milen2 = Len(MiMatriz(0, 2))
        LinTx = MiMatriz(0, 0) & Space(5 - MiLen1) & MiMatriz(0, 1) & _
            Space(10 - Milen2) & MiMatriz(0, 2)
        Sys_Pos.Coptr1.PrintNormal PTR_S_RECEIPT, LinTx & Chr(&HD) & Chr(&HA)
        nTotCta = nTotCta + rsPreCta!precio
        rsPreCta.MoveNext
        If rsPreCta.EOF Then Exit Do
        If nLinCta <> rsPreCta!CUENTA Then
            nLinCta = rsPreCta!CUENTA
            Exit Do
        End If
    Loop
    
    If nLinCta <> 0 Then
        MiLen1 = Len(Format(nTotCta, "standard"))
        nProp = 0#
        If OKProp = 1 Or OPEN_PROPINA = False Then
            If nMesa <> rs00!mesa_barra Then
                MiPropina = RoundToNearest(SBTot * 0.1, 0.05, 1)
                nProp = MiPropina * 100
                nProp = nProp / 100
                Sys_Pos.Coptr1.PrintNormal PTR_S_RECEIPT, PROPINA_DESCRIP & " : " & Format(nProp, "##0.00") & Chr(&HD) & Chr(&HA)
            End If
            nTotProp = nTotProp + nProp
            OKProp = 1
        End If
        nTotCta = nTotCta + nProp
        Sys_Pos.Coptr1.PrintNormal PTR_S_RECEIPT, "SUB-TOTAL Cuenta " & Space(13 - MiLen1) & Format(nTotCta, "STANDARD") & Chr(&HD) & Chr(&HA)
        Sys_Pos.Coptr1.PrintNormal PTR_S_RECEIPT, Space(1) & Chr(&HD) & Chr(&HA)
        If Not rsPreCta.EOF Then
            For i = 1 To 10
                Sys_Pos.Coptr1.PrintNormal PTR_S_RECEIPT, Space(2) & Chr(&HD) & Chr(&HA)
            Next
            Sys_Pos.Coptr1.CutPaper 100
            Sys_Pos.Coptr1.PrintNormal PTR_S_RECEIPT, Format(Date, "short date") & Chr(&HD) & Chr(&HA)
            Sys_Pos.Coptr1.PrintNormal PTR_S_RECEIPT, "Mesero : " & cNomMesero & Chr(&HD) & Chr(&HA)
            Sys_Pos.Coptr1.PrintNormal PTR_S_RECEIPT, Space(1) & Chr(&HD) & Chr(&HA)
            Sys_Pos.Coptr1.PrintNormal PTR_S_RECEIPT, "Mesa : " & nMesa & "      " & Time & Chr(&HD) & Chr(&HA)
            Sys_Pos.Coptr1.PrintNormal PTR_S_RECEIPT, Space(1) & Chr(&HD) & Chr(&HA)
            Sys_Pos.Coptr1.PrintNormal PTR_S_RECEIPT, "Cuenta # : " & nLinCta & Chr(&HD) & Chr(&HA)
        End If

    End If
    nTotCta = 0#
Loop

If lParc = 1 Then
    MiLen1 = -1
    Milen2 = Len(Format(rsParciales!VALOR * (-1), "standard"))
    Sys_Pos.Coptr1.PrintNormal PTR_S_RECEIPT, "PAGO.PARCIAL " & Space(4 - MiLen1) & MiLen1 & _
            Space(10 - Milen2) & Format(rsParciales!VALOR * (-1), "standard") & Chr(&HD) & Chr(&HA)
End If

'*******************************************
'*******************************************
On Error Resume Next
Sys_Pos.Coptr1.PrintNormal PTR_S_RECEIPT, Space(1) & Chr(&HD) & Chr(&HA)
Milen2 = Len(Format(SubTot, "standard"))
Sys_Pos.Coptr1.PrintNormal PTR_S_RECEIPT, "   * Sub-Total :" & Space(14 - Milen2) & Format((SubTot + nTotProp) - FormatCurrency(iISCTransaccion, 2), "standard") & Chr(&HD) & Chr(&HA)
'Sys_Pos.Coptr1.PrintNormal PTR_S_RECEIPT, Space(1) & Chr(&HD) & Chr(&HA)

If OKDesc = 1 And DescPreCta > 0.01 Then
    Sys_Pos.Coptr1.PrintNormal PTR_S_RECEIPT, "  Su DESCUENTO :" & Space(9) & Format((DescPreCta * -1), "standard") & Chr(&HD) & Chr(&HA)
'    Sys_Pos.Coptr1.PrintNormal PTR_S_RECEIPT, Space(1) & Chr(&HD) & Chr(&HA)
    
    Milen2 = Len(Format(SubTot, "standard"))
    Sys_Pos.Coptr1.PrintNormal PTR_S_RECEIPT, "     Sub-Total :" & Space(14 - Milen2) & Format((SubTot - iISCTransaccion) + (DescPreCta * -1), "standard") & Chr(&HD) & Chr(&HA)
'    Sys_Pos.Coptr1.PrintNormal PTR_S_RECEIPT, Space(1) & Chr(&HD) & Chr(&HA)
    nTempSubTot = (SubTot - iISCTransaccion) + (DescPreCta * -1)
    
    Milen2 = Len(Format(iISCTransaccion, "STANDARD"))
    txtString = "     ITBMS (5%):" & Space(14 - Milen2) & Format(((SubTot - iISCTransaccion) + (DescPreCta * -1)) * iISC, "STANDARD")
    Sys_Pos.Coptr1.PrintNormal PTR_S_RECEIPT, txtString & Chr(&HD) & Chr(&HA)
    nTempISC = FormatCurrency(((SubTot - iISCTransaccion) + (DescPreCta * -1)) * iISC, 2)

    Milen2 = Len(Format(SubTot, "standard"))
    Sys_Pos.Coptr1.PrintNormal PTR_S_RECEIPT, "     Sub-Total :" & Space(14 - Milen2) & Format((SubTot - iISCTransaccion) + (DescPreCta * -1) + nTempISC, "standard") & Chr(&HD) & Chr(&HA)
    Sys_Pos.Coptr1.PrintNormal PTR_S_RECEIPT, Space(1) & Chr(&HD) & Chr(&HA)
    nTempSubFinal = (SubTot - iISCTransaccion) + (DescPreCta * -1) + nTempISC
    
    If (OKProp = 1 And nLinCta = 0) Or (OPEN_PROPINA = False And mlincta = 0) Then
        'SI HAY PROPINA Y ES UNA SOLA CUENTA
        'MiPropina = Format(Round(SubTot * 0.1, 1) * 1#, "standard")
        If nMesa <> rs00!mesa_barra Then
            MiPropina = RoundToNearest(nTempSubTot * 0.1, 0.05, 1)
            nProp = MiPropina * 100
            nProp = nProp / 100
            nTotProp = nProp
            NLEN = Len(PROPINA_DESCRIP)
            Sys_Pos.Coptr1.PrintNormal PTR_S_RECEIPT, PROPINA_DESCRIP & " : " & Space(23 - NLEN) & Format(nProp, "##0.00") & Chr(&HD) & Chr(&HA)
            Sys_Pos.Coptr1.PrintNormal PTR_S_RECEIPT, Space(1) & Chr(&HD) & Chr(&HA)
        End If
        OKProp = 0
    End If

    Milen2 = Len(Format((SubTot + nTotProp - DescPreCta), "currency"))
    Sys_Pos.Coptr1.PrintNormal PTR_S_RECEIPT, "     SUMA      :" & Space(14 - Milen2) & Format((nTempSubFinal + nTotProp), "currency") & Chr(&HD) & Chr(&HA)

    OKDesc = 0
Else
    Milen2 = Len(Format(iISCTransaccion, "STANDARD"))
    txtString = "     ITBMS (5%):" & Space(14 - Milen2) & Format(iISCTransaccion, "STANDARD")
    Sys_Pos.Coptr1.PrintNormal PTR_S_RECEIPT, txtString & Chr(&HD) & Chr(&HA)
    
    Milen2 = Len(Format(SubTot, "standard"))
    Sys_Pos.Coptr1.PrintNormal PTR_S_RECEIPT, "     Sub-Total :" & Space(14 - Milen2) & Format((SubTot + nTotProp), "standard") & Chr(&HD) & Chr(&HA)
    Sys_Pos.Coptr1.PrintNormal PTR_S_RECEIPT, Space(1) & Chr(&HD) & Chr(&HA)
    
    If (OKProp = 1 And nLinCta = 0) Or (OPEN_PROPINA = False And mlincta = 0) Then
        'SI HAY PROPINA Y ES UNA SOLA CUENTA
        'MiPropina = Format(Round(SubTot * 0.1, 1) * 1#, "standard")
        If nMesa <> rs00!mesa_barra Then
            MiPropina = RoundToNearest((SBTot - iISCTransaccion) * 0.1, 0.05, 1)
            nProp = MiPropina * 100
            nProp = nProp / 100
            nTotProp = nProp
            NLEN = Len(PROPINA_DESCRIP)
            Sys_Pos.Coptr1.PrintNormal PTR_S_RECEIPT, PROPINA_DESCRIP & " : " & Space(23 - NLEN) & Format(nProp, "##0.00") & Chr(&HD) & Chr(&HA)
            Sys_Pos.Coptr1.PrintNormal PTR_S_RECEIPT, Space(1) & Chr(&HD) & Chr(&HA)
        End If
        OKProp = 0
    End If
    
    Milen2 = Len(Format((SubTot + nTotProp - DescPreCta), "currency"))
    Sys_Pos.Coptr1.PrintNormal PTR_S_RECEIPT, "     SUMA      :" & Space(14 - Milen2) & Format((SubTot + nTotProp - DescPreCta), "currency") & Chr(&HD) & Chr(&HA)
    
End If

Sys_Pos.Coptr1.PrintNormal PTR_S_RECEIPT, Space(1) & Chr(&HD) & Chr(&HA)
Sys_Pos.Coptr1.PrintNormal PTR_S_RECEIPT, Space(1) & Chr(&HD) & Chr(&HA)
Sys_Pos.Coptr1.PrintNormal PTR_S_RECEIPT, Space(1) & Chr(&HD) & Chr(&HA)
For i = 1 To 10
    Sys_Pos.Coptr1.PrintNormal PTR_S_RECEIPT, Space(2) & Chr(&HD) & Chr(&HA)
Next
Sys_Pos.Coptr1.CutPaper 100

On Error GoTo 0
'*******************************************
'*******************************************

If nErrCount >= 3 Then
    MsgBox "POR FAVOR REVISE LA PRE-CUENTA", vbInformation, BoxTit
End If
On Error GoTo 0
Exit Sub

AdmErr:
nErrCount = nErrCount + 1
Milen2 = 10
If nErrCount < 3 Then
    Resume
Else
    Resume Next
End If
'msConn.BeginTrans
'msConn.Execute "UPDATE ORGANIZACION SET TRANS = TRANS + 1"
'msConn.CommitTrans
End Sub
Private Sub DescProducto()
'PROCEDIMIENTO PARA DAR DESCUENTO A UN PRODUCTO

Dim MiDesc As Single
Dim nDescImpre As Single
Dim rsParciales As Recordset
Dim lParc As Integer
Dim sqltext As String

If PlatosMesa.Rows = 0 Then
    MsgBox " No hay nada Marcado ", vbExclamation, BoxTit
    Exit Sub
End If

If nCantidad > MAX_DESCUENTO Then
    MsgBox "ES IMPOSIBLE DAR ESE DESCUENTO. INTENTE DAR UN PORCENTAJE MAS BAJO", vbInformation, BoxTit
    Clear_Click
    Exit Sub
End If

'nCantidad es el valor del Cuadro de Numeros de la Derecha Abajo
If nCantidad > 1 Then
    MiDesc = Format(nCantidad / 100, "standard")
Else
    'nDesc01 es el Descuento Marcado
    MiDesc = Format(nDesc01 / 100, "standard")
End If

MiDesc = Format(MiDesc, "standard")

Dim rsFixTmpTrans As New Recordset
Dim rsGetMaxLin As New Recordset
Dim nMaxLin As Integer
Dim txto As String

If nNLinSel <> 0 Then   'PREGUNTA SI HIZO CLICK A PLATOSMESAS
    txto = "SELECT * FROM tmp_trans " & _
        " WHERE mesa = " & nMesa & " AND lin = " & nNLinSel
Else
    rsGetMaxLin.Open "SELECT MAX(LIN) AS MAX_LIN " & _
        "FROM TMP_TRANS WHERE MESA = " & nMesa, msConn, adOpenStatic, adLockOptimistic
    
    If (PlatosMesa.Rows - 1) >= 1 Then
        PlatosMesa.TopRow = (PlatosMesa.Rows - 1)
    End If
    PlatosMesa.Col = 0
    PlatosMesa.Row = (PlatosMesa.Rows - 1)

    txto = "SELECT * FROM tmp_trans " & _
        " WHERE MESA = " & nMesa & " AND LIN = " & Val(PlatosMesa.Text)
    nMaxLin = rsGetMaxLin!max_lin
    rsGetMaxLin.Close
    Set rsGetMaxLin = Nothing
    'txto = "SELECT * FROM tmp_trans " & _
        " WHERE MESA = " & nMesa & " AND LIN = " & Val(PlatosMesa.Text)
End If

rsFixTmpTrans.Open txto, msConn, adOpenStatic, adLockReadOnly

If rsFixTmpTrans.EOF = True Then
    rsFixTmpTrans.Close
    MsgBox "Por Favor SELECCIONE un Producto", vbInformation, BoxTit
    nCantidad = 1
    Exit Sub
End If

If rsFixTmpTrans!CANT < 0 Then
    'Si la Cantidad es 0 entonces...
    MsgBox "NO puede dar DESCUENTO a este Producto", vbInformation, BoxTit
    rsFixTmpTrans.Close
    nCantidad = 1
    Exit Sub
End If
    
If Mid(rsFixTmpTrans!DESCRIP, 1, 9) = "DESCUENTO" Then
    MsgBox "NO puede dar DESCUENTO a un Descuento", vbExclamation, BoxTit
    rsFixTmpTrans.Close
    nCantidad = 1
    Exit Sub
End If
    
If Mid(rsFixTmpTrans!TIPO, 1, 1) = "B" Then
    MsgBox "PRODUCTO YA FUE ANULADO/CORREGIDO/SE DIO DESCUENTO EN LA LINEA " & Val(Mid(rsFixTmpTrans!TIPO, 5, 2)), vbInformation, BoxTit
    rsFixTmpTrans.Close
    nCantidad = 1
    Exit Sub
End If
    
nCtaLinAnul = rsFixTmpTrans!CUENTA
CajLin = CajLin + 1

'------------REVISION DE PAGOS PARCIALES-------------------
Set rsParciales = New Recordset
rsParciales.Open "SELECT MESA,SUM(MONTO) AS VALOR " & _
            " FROM TMP_PAR_PAGO " & _
            " WHERE MESA = " & nMesa & _
            " GROUP BY MESA", msConn, adOpenDynamic, adLockOptimistic

If rsParciales.EOF Then lParc = 0 Else lParc = 1
'--------------------------------

Dim nTestDesc As Integer
nDescImpre = Format(MiDesc * rsFixTmpTrans!precio * (-1), "standard")
nTestDesc = Val(Mid(nDescImpre, Len(nDescImpre) + 1, 1))
    
'Proceso que quita los centavos del Descuento y los redondea al mas bajo
'y Asigna su valor a nDescImpre
If nTestDesc = 0 Or nTestDesc = 5 Then
ElseIf nTestDesc < 5 Then
    nDescImpre = nDescImpre + (nTestDesc / 100)
ElseIf nTestDesc > 5 And nTestDesc <= 9 Then
    nDescImpre = nDescImpre + ((nTestDesc - 5) / 100)
End If
    
Dim SOLO_FECHA As String
SOLO_FECHA = Format(Date, "YYYYMMDD")

If nNLinSel <> 0 Then
  CadenaSql = "INSERT INTO TMP_TRANS " & _
    "(CAJA,CAJERO,MESA,MESERO,VALID,LIN," & _
    "DESCRIP,CANT,DEPTO," & _
    "PLU,ENVASE,PRECIO_UNIT,PRECIO,FECHA," & _
    "HORA,TIPO,DESCUENTO,CUENTA) VALUES (" & _
    "" & cCaja & "," & npNumCaj & "," & nMesa & "," & nMesero & "," & -1 & "," & CajLin & "," & _
    "'DESCUENTO : " & Format(MiDesc, "#.00") & "%'" & "," & 1 & "," & rsFixTmpTrans!depto & "," & _
    rsFixTmpTrans!PLU & "," & rsFixTmpTrans!envase & "," & nDescImpre & "," & nDescImpre & "," & _
    "'" & SOLO_FECHA & "'" & "," & "'" & Time & "'" & _
    ",'DC-" & nNLinSel & "'," & MiDesc & "," & nCtaLinAnul & ")"

    sqltext = "UPDATE TMP_TRANS SET TIPO = 'BDC" & Str((CajLin)) & _
            "' WHERE MESA = " & nMesa & _
            " AND LIN = " & nNLinSel
Else
  CadenaSql = "INSERT INTO TMP_TRANS " & _
    "(CAJA,CAJERO,MESA,MESERO,VALID,LIN," & _
    "DESCRIP,CANT,DEPTO,PLU,ENVASE,PRECIO_UNIT," & _
    "PRECIO,FECHA,HORA,TIPO,DESCUENTO,CUENTA) " & _
    " VALUES (" & _
    "" & cCaja & "," & npNumCaj & "," & nMesa & "," & nMesero & "," & -1 & "," & nMaxLin + 1 & "," & _
    "'DESCUENTO : " & Format(MiDesc, "#.00") & "%'" & "," & 1 & "," & rsFixTmpTrans!depto & "," & rsFixTmpTrans!PLU & "," & _
    rsFixTmpTrans!envase & "," & nDescImpre & "," & nDescImpre & "," & "'" & SOLO_FECHA & "'" & "," & "'" & Time & "'" & _
    ",'DC-" & Val(PlatosMesa.Text) & "'," & MiDesc & "," & nCtaLinAnul & ")"
        
    sqltext = "UPDATE TMP_TRANS SET TIPO = 'BDC" & Str(Val(PlatosMesa.Text)) & _
        "' WHERE MESA = " & nMesa & _
        "  AND LIN = " & nMaxLin
    CajLin = (nMaxLin + 1)
End If
    
msConn.BeginTrans
msConn.Execute CadenaSql
msConn.Execute sqltext
msConn.CommitTrans

'''''''''''msConnLoc.BeginTrans
'''''''''''msConnLoc.Execute CadenaSql
'''''''''''msConnLoc.Execute sqltext
'''''''''''msConnLoc.CommitTrans

rs07.Open "SELECT a.lin,a.descrip,a.cant," & _
    " format(precio_unit,'##0.00') as mPrecio_unit," & _
    " format(precio,'##0.00') as mPrecio," & _
    " a.cajero,a.mesero,a.depto,a.plu,a.mesa,a.valid, " & _
    " a.envase,a.fecha,a.hora,a.tipo,a.descuento,a.cuenta, " & _
    " a.caja " & _
    " FROM tmp_trans as a " & _
    " WHERE a.mesa = " & nMesa & _
    " AND A.CUENTA = " & nCta & _
    " ORDER BY a.lin ", msConn, adOpenStatic, adLockOptimistic

Set PlatosMesa.DataSource = rs07
SetupPantalla
    
nLineas = PlatosMesa.Rows - 1

If (PlatosMesa.Rows - 1) >= 1 Then
    PlatosMesa.TopRow = (PlatosMesa.Rows - 1)
End If
    
rs07.Close
rs07.Open "SELECT sum(a.precio) as precio FROM tmp_trans as a " & _
    " WHERE a.mesa = " & nMesa & _
    " AND A.CUENTA = " & nCta, msConn, adOpenStatic, adLockReadOnly

SubTot = FormatCurrency(IIf(IsNull(rs07!precio), 0#, rs07!precio), 2)
On Error Resume Next
SubTot = FormatCurrency((SubTot + (rs07!precio * iISC)), 2)
iISCTransaccion = rs07!precio * iISC
SBTot = Format(SubTot, "standard")
On Error GoTo 0
rs07.Close
rsFixTmpTrans.Close

nCantidad = 1: nPase = 0
Text1(2) = nCantidad
    
If lParc = 1 Then
    PlatosMesa.AddItem (CajLin + 1) & Chr(9) & _
    "PAGO.PARCIAL" & Chr(9) & 1 & Chr(9) & _
    Format(rsParciales!VALOR, "STANDARD") & Chr(9) & Format(rsParciales!VALOR, "STANDARD")
    SubTot = Format(SubTot - rsParciales!VALOR, "STANDARD")
    lbMensaje.BackColor = &HFFFF00
    lbMensaje = "MESA CON PAGOS PARCIALES"
Else
    lbMensaje.BackColor = &HFFFFFF
    lbMensaje = ""
End If
End Sub
Private Sub MuestraPLU_del_Envase(nElEnvase As Long)
Dim MiTop As Integer, MiLeft As Integer, StayLeft As Integer
Dim iTam As Integer
Dim sqltext As String

'Busca PLUS x Enase Seleccionado
nGlobEnv = nElEnvase
sqltext = "SELECT a.depto,a.contenedor,b.codigo,b.descrip,b.corto,c.precio " & _
    " FROM CONTEND_01 as a,PLU as b, CONTEND_02 as c " & _
    " WHERE a.DEPTO = " & ElDepto & _
    " AND a.contenedor = " & nElEnvase & _
    " AND a.depto = b.depto " & _
    " AND b.codigo = c.codigo " & _
    " AND c.contenedor = " & nElEnvase & _
    " ORDER BY b.CORTO "
rs08.Open sqltext, msConn, adOpenStatic, adLockOptimistic
iTam = 0

MiTop = 240: StayLeft = 120
MiLeft = 0: numplu = 0

Do Until rs08.EOF
    If numplu < 1 Then
        cmdPlus(numplu).Caption = rs08!DESCRIP
        cmdPlus(numplu).Tag = rs08!CODIGO
        'Muestra los PLUs del primer departamento
    Else
        If Not IsObject(cmdPlus(numplu)) Then
           Load cmdPlus(numplu)
        End If
        cmdPlus(numplu).Visible = True
        cmdPlus(numplu).Caption = rs08!DESCRIP
        cmdPlus(numplu).Tag = rs08!CODIGO
        cmdPlus(numplu).Left = MiLeft + StayLeft
        cmdPlus(numplu).Top = MiTop
        StayLeft = 120
    End If
    numplu = numplu + 1
    MiLeft = MiLeft + 2400
    If numplu = 3 Or numplu = 6 Or numplu = 9 Or numplu = 12 Or numplu = 15 Then
        MiTop = MiTop + 600
        MiLeft = 0
    End If
    If numplu = 18 Then Exit Do
    rs08.MoveNext
Loop
rs08.Close
End Sub
Private Sub SetupPantalla()
Dim i As Integer
'Formato de la Pantalla de Facturacion

With PlatosMesa
    For i = 0 To 17
        .ColWidth(i) = 0
    Next
    .ColWidth(0) = 400: .ColWidth(1) = 3200: .ColWidth(2) = 500
    .ColWidth(3) = 800: .ColWidth(4) = 1000:
    .ColAlignmentFixed(3) = flexAlignRightCenter
    .ColAlignmentFixed(4) = flexAlignRightCenter
    .ColAlignment(3) = flexAlignRightCenter
    .ColAlignment(4) = flexAlignRightCenter
End With

End Sub
Private Sub QuitarPLUS()
Dim nNum As Integer, lNum As Integer

nNum = rs03.RecordCount

For lNum = 1 To 17
    cmdPlus(1).Caption = ""
    If Not IsObject(cmdPlus(lNum)) Then
        Load cmdPlus(lNum)
    End If
    cmdPlus(lNum).Visible = False
Next
cmdPlus(0).Visible = True
End Sub
Private Sub MuestraPLU(ElDepto As Integer)
'VIENE DE HACER CLICK A LOS DEPARTAMENTOS
Dim MiTop As Integer, MiLeft As Integer, StayLeft As Integer
Dim iTam As Integer

'Muestra los productos a Vender
Set rs03 = New Recordset
'Set rs04 = New Recordset
rs04.Close
'Busca PLUS del Depto
rs03.Open "SELECT codigo,depto,descrip,corto,precio1,envases,IMPRESORA " & _
        " FROM PLU " & _
        " WHERE depto = " & ElDepto & _
        " ORDER BY CORTO", msConn, adOpenStatic, adLockReadOnly
'Busca Envases del Departamento
rs04.Open "SELECT a.depto,a.contenedor,b.descrip " & _
        " FROM contend_01 as a,contened as b " & _
        " WHERE a.DEPTO = " & ElDepto & " AND " & _
        " a.contenedor = b.contenedor " & _
        " ORDER BY a.depto,a.contenedor", msConn, adOpenStatic, adLockOptimistic
iTam = 0

'Prepara los Envases del Departamento
For iTam = 0 To 3
    cmdEnvases(iTam).Enabled = True
    cmdEnvases(iTam).BackColor = &HC0C0C0
    cmdEnvases(iTam).Caption = ""
Next

iTam = 0

'If Not rs04.EOF Then FlashControl Frame2(4)
If rs04.EOF Then Frame2(4).BackColor = &H8000000F Else Frame2(4).BackColor = &HFFFF&

Do Until rs04.EOF
    cmdEnvases(iTam).Caption = rs04!DESCRIP
    cmdEnvases(iTam).Tag = rs04!contenedor
    iTam = iTam + 1
    rs04.MoveNext
Loop

For iTam = 0 To 3
    If cmdEnvases(iTam).Caption = "" Then
        cmdEnvases(iTam).Enabled = False
    End If
Next

MiTop = 240: StayLeft = 120
MiLeft = 0: numplu = 0

'Si No hay productos, quitar los que estan visibles
If rs03.EOF Then
    Dim lNum As Integer
    cmdPlus(0).Tag = ""
    For lNum = 0 To 17
        cmdPlus(0).Caption = ""
        If Not IsObject(cmdPlus(lNum)) Then
            Load cmdPlus(lNum)
        End If
        cmdPlus(lNum).Visible = False
    Next
    cmdPlus(0).Visible = True
    rs02.MoveFirst
    rs02.Find "CODIGO = " & ElDepto
    If Not rs02.EOF Then
        If rs02!ABIERTO = True Then
            'MsgBox "DEPARTAMENTO ABIERTO", vbCritical, BoxTit
            If nMesa = 0 Or nMesero = 0 Then
                MsgBox "Antes Debe Seleccionar una Mesa y su Mesero", vbCritical, BoxTit
                Exit Sub
            End If
            AddOpenDeptItem
        End If
    End If
    Exit Sub
End If

'SI HAY PRODUCTOS EN EL DEPARTAMENTO, LOS MUESTRO
Do Until rs03.EOF
    If numplu < 1 Then
        cmdPlus(numplu).Caption = rs03!DESCRIP
        cmdPlus(numplu).Tag = rs03!CODIGO
        'Muestra los PLUs del primer departamento
    Else
        If Not IsObject(cmdPlus(numplu)) Then
           Load cmdPlus(numplu)
        End If
        cmdPlus(numplu).Visible = True
        cmdPlus(numplu).Caption = rs03!DESCRIP
        cmdPlus(numplu).Tag = rs03!CODIGO
        cmdPlus(numplu).Left = MiLeft + StayLeft
        cmdPlus(numplu).Top = MiTop
        StayLeft = 120
    End If
    numplu = numplu + 1
    MiLeft = MiLeft + 2400
    If numplu = 3 Or numplu = 6 Or numplu = 9 Or numplu = 12 Or numplu = 15 Then
        MiTop = MiTop + 600
        MiLeft = 0
    End If
    If numplu = 18 Then Exit Do
    rs03.MoveNext
Loop

End Sub
Private Sub Quita_Subrallado(var As Integer)
Dim i As Integer

i = 0

For i = 0 To cmdDepto.Count - 1
    cmdDepto(i).BackColor = &HC0C0C0
Next
If var <> 67 Then
    cmdDepto(var).BackColor = &HFFFF80
End If
End Sub

Private Sub QuitarDeptos()
Dim nNum As Integer

For nNum = 1 To cmdDepto.Count - 1
    cmdDepto(1).Caption = ""
    cmdDepto(nNum).Visible = False
Next
End Sub

Private Sub Clear_Click()
nPase = 0
nCantidad = 1
Text1(2) = nCantidad
End Sub

Private Sub cmdAcomp_Click(Index As Integer)
Dim SOLO_FECHA As String

If cmdAcomp(Index).Caption = "" Then Exit Sub

CajLin = CajLin + 1
SOLO_FECHA = Format(Date, "YYYYMMDD")

CadenaSql = "INSERT INTO TMP_TRANS " & _
    "(CAJA,CAJERO,MESA,MESERO,VALID,LIN,DESCRIP,CANT,DEPTO,PLU,ENVASE,PRECIO_UNIT,PRECIO,FECHA,HORA,TIPO,DESCUENTO,CUENTA,IMPRESO,IMPRESORA) VALUES (" & _
    cCaja & "," & _
    npNumCaj & "," & _
    nMesa & "," & _
    nMesero & "," & _
    -1 & "," & _
    CajLin & ",' @@ " & _
    cmdAcomp(Index).Caption & "'," & _
    1 & "," & _
    cmdAcomp(Index).Tag & "," & _
    0 & "," & _
    0 & "," & _
    0 & "," & _
    0 & ",'" & _
    SOLO_FECHA & "','" & _
    Time & "','  '," & _
    0# & "," & _
    nCta & _
    ",FALSE," & _
    nSeleccionCocina & ")"
    
'''    COCINA_01 & ")"

msConn.BeginTrans
msConn.Execute CadenaSql
msConn.CommitTrans

If CajLin = 1 Then msConn.Execute "UPDATE Mesas SET ocupada = TRUE WHERE numero = " & nMesa

rs07.Open "SELECT a.lin,a.descrip,a.cant," & _
    " format(precio_unit,'##0.00') as mPrecio_unit," & _
    " format(precio,'##0.00') as mPrecio," & _
    " a.cajero,a.mesero,a.depto,a.plu,a.mesa,a.valid, " & _
    " a.envase,a.fecha,a.hora,a.tipo,a.descuento,a.cuenta, " & _
    " a.caja " & _
    " FROM tmp_trans as a " & _
    " WHERE a.mesa = " & nMesa & _
    " ORDER BY a.lin", msConn, adOpenStatic, adLockOptimistic

Set PlatosMesa.DataSource = rs07
SetupPantalla

nLineas = PlatosMesa.Rows - 1

Set rsParciales = New Recordset
rsParciales.Open "SELECT MESA,SUM(MONTO) AS VALOR FROM TMP_PAR_PAGO " & _
    " WHERE MESA = " & nMesa & _
    " GROUP BY MESA", msConn, adOpenDynamic, adLockOptimistic

If rsParciales.EOF Then lParc = 0 Else lParc = 1

rs07.Close
rs07.Open "SELECT sum(a.precio) as precio FROM tmp_trans as a " & _
    " WHERE a.mesa = " & nMesa, msConn, adOpenStatic, adLockReadOnly
SubTot = FormatCurrency(IIf(IsNull(rs07!precio), 0#, rs07!precio), 2)
On Error Resume Next
SubTot = FormatCurrency((SubTot + (rs07!precio * iISC)), 2)
iISCTransaccion = rs07!precio * iISC
SBTot = Format(SubTot, "standard")
On Error GoTo 0
rs07.Close
If (PlatosMesa.Rows - 1) >= 1 Then
    PlatosMesa.TopRow = (PlatosMesa.Rows - 1)
End If
nCantidad = 1: nPase = 0
nNLinSel = 0
Text1(2) = nCantidad
'PUEDE SELECCIONAR MAS DE UN ACOMPAÑANTE
'POR ESO FRAME2(2)=ENABLED
'Frame2(2).Enabled = False

If lParc = 1 Then
    PlatosMesa.AddItem (CajLin + 1) & Chr(9) & _
        "PAGO.PARCIAL" & Chr(9) & 1 & Chr(9) & _
        Format(rsParciales!VALOR, "STANDARD") & Chr(9) & Format(rsParciales!VALOR, "STANDARD")
    SubTot = Format(SubTot - rsParciales!VALOR, "STANDARD")
    lbMensaje.BackColor = &HFFFF00
    lbMensaje = "MESA CON PAGOS PARCIALES"
Else
    lbMensaje.BackColor = &HFFFFFF
    lbMensaje = ""
End If

End Sub

Private Sub cmdCtas_Click()
Dim rsParciales As Recordset
Dim rsMaxLin As New ADODB.Recordset
Dim cSQL As String

On Error GoTo ErrAdm:

If nCortesia = 0 Then nCortesia = 1: Command13(6).BackColor = &HFF&

Set rsParciales = New Recordset
rsParciales.Open "SELECT CAJERO,MESA,MESERO,TIPO_PAGO,LIN,MONTO " & _
        " FROM TMP_PAR_PAGO " & _
        " WHERE MESA = " & nMesa, msConn, adOpenDynamic, adLockOptimistic

If rsParciales.EOF Then
    rsParciales.Close
    cSQL = "SELECT DISTINCT TIPO FROM TMP_TRANS "
    cSQL = cSQL & " WHERE MESA = " & nMesa
    cSQL = cSQL & " AND CUENTA = 0 "
    rsParciales.Open cSQL, msConn, adOpenStatic, adLockOptimistic
    If rsParciales.RecordCount > 1 Then
        MsgBox "NO ES POSIBLE ASIGNAR CUENTAS. ESTA MESA YA TIENE CORRECCIONES, ANULACIONES o DESCUENTOS " & vbCrLf & _
            "SE LE SUGIERE ABRIR UNA NUEVA MESA", vbExclamation, "NO ES POSIBLE ASIGNAR CUENTAS"
        rsParciales.Close
        Set rsParciales = Nothing
        Exit Sub
    End If
    Set rsParciales = Nothing
    FacCtaPlato.Show 1
    lbCuenta = nCta
    rs07.Open "SELECT a.lin,a.descrip,a.cant," & _
        " format(precio_unit,'##0.00') as mPrecio_unit," & _
        " format(precio,'##0.00') as mPrecio," & _
        " a.cajero,a.mesero,a.depto,a.plu,a.mesa,a.valid, " & _
        " a.envase,a.fecha,a.hora,a.tipo,a.descuento,a.cuenta, " & _
        " a.caja " & _
        " FROM tmp_trans as a " & _
        " WHERE a.mesa = " & nMesa & _
        " AND A.CUENTA = " & nCta & _
        " ORDER BY a.lin", msConn, adOpenStatic, adLockOptimistic

    Set PlatosMesa.Recordset = rs07
    SetupPantalla
    
''    rsMaxLin.Open "SELECT MAX(LIN) AS MAX_LIN " & _
''        " FROM TMP_TRANS " & _
''        " WHERE MESA = " & nMesa & _
''        " AND CUENTA = " & nCta, msConn, adOpenStatic, adLockOptimistic
''    CajLin = rsMaxLin!MAX_LIN
''    rsMaxLin.Close
''    Set rsMaxLin = Nothing
    
    rs07.Close
    rs07.Open "SELECT sum(a.precio) as precio FROM tmp_trans as a " & _
          " WHERE a.mesa = " & nMesa & _
          " AND A.CUENTA = " & nCta, msConn, adOpenStatic, adLockReadOnly
    SubTot = FormatCurrency(IIf(IsNull(rs07!precio), 0#, rs07!precio), 2)
    On Error Resume Next
    SubTot = FormatCurrency((SubTot + (rs07!precio * iISC)), 2)
    iISCTransaccion = rs07!precio * iISC
    SBTot = Format(SubTot, "standard")
    On Error GoTo 0
    rs07.Close
    If (PlatosMesa.Rows - 1) >= 1 Then
        PlatosMesa.TopRow = (PlatosMesa.Rows - 1)
    End If
    nCantidad = 1: nPase = 0
    nNLinSel = 0
Else
    MsgBox "ESTA MESA NO SE PUEDE TRABAJAR POR CUENTAS, YA QUE TIENE PAGOS PARCIALES", vbExclamation, BoxTit
    Exit Sub
End If
On Error GoTo 0
Exit Sub

ErrAdm:
MsgBox Err.Number & " ->>>> " & Err.Description, vbCritical, BoxTit
Resume Next
End Sub

Private Sub cmdDepto_Click(Index As Integer)
Dim iAcoCnt As Integer
Frame2(1).Caption = "MENU " & cmdDepto(Index).Caption
Quita_Subrallado (Index)
'TRATAMIENTO DE ACOMPAÑANTES
iAcoCnt = 0
If Frame2(2).Enabled = True Then
    For iAcoCnt = 0 To 3
        cmdAcomp(iAcoCnt).Caption = ""
        cmdAcomp(iAcoCnt).Tag = 0
    Next
    Frame2(2).Enabled = False
End If
'---------------------------------
TextEnv = ""
QuitarPLUS
ElDepto = Arreg_Deptos(Index)
MuestraPLU (Arreg_Deptos(Index))
nGlobEnv = 0
nNLinSel = 0
End Sub

Private Sub cmdEnvases_Click(Index As Integer)
'cmdEnvases(Index).Tag
TextEnv = "-" + cmdEnvases(Index).Caption
QuitarPLUS
For i = 0 To 3
    cmdEnvases(i).BackColor = &HC0C0C0
Next
If cmdEnvases(Index).BackColor = &HC0C0C0 Then
    cmdEnvases(Index).BackColor = &HFFFF00
Else
    cmdEnvases(Index).BackColor = &HC0C0C0
End If
MuestraPLU_del_Envase (cmdEnvases(Index).Tag)
End Sub
Private Sub cmdPlus_GotFocus(Index As Integer)
cmdPlus(Index).BackColor = &HFFFF00
End Sub
Private Sub cmdPlus_LostFocus(Index As Integer)
cmdPlus(Index).BackColor = &HC0C0C0
End Sub

Private Sub cmdRestoAco_Click()
Dim iLoc As Integer
Dim iAcom As Integer
Dim nAcoTop As Integer

iLoc = 0: iAcom = 0: nAcoTop = 240

For iLoc = 0 To cmdAcomp.Count - 1
    cmdAcomp(iLoc).Caption = ""
    cmdAcomp(iLoc).Tag = 0
    If iLoc > 0 Then
        cmdAcomp(iLoc).Visible = False
    End If
Next
rsTmpAco.Bookmark = nAcoBookMark
rsTmpAco.MoveNext
Do Until rsTmpAco.EOF
    If iAcom = 0 Then
        cmdAcomp(iAcom).Visible = True
        cmdAcomp(iAcom).Caption = rsTmpAco!DESCRIP
        cmdAcomp(iAcom).Tag = rs03!depto
        iAcom = iAcom + 1
        rsTmpAco.MoveNext
    Else
        cmdAcomp(iAcom).Visible = True
        cmdAcomp(iAcom).Top = nAcoTop + 600
        cmdAcomp(iAcom).Caption = rsTmpAco!DESCRIP
        cmdAcomp(iAcom).Tag = rs03!depto
        iAcom = iAcom + 1
        nAcoTop = nAcoTop + 600
        rsTmpAco.MoveNext
    End If
Loop
cmdRestoAco.Enabled = False
End Sub
Private Sub cmdSalir_Click()
If rsTmpAco.State = adStateOpen Then rsTmpAco.Close
Set rsTmpAco = Nothing
StatMesa nMesa, 0
Unload Me
End Sub

Private Sub cmdSlip_Click()
Dim rsCocina As New ADODB.Recordset
Dim nFlag As Boolean
Dim nFile As Byte

If nCortesia = 0 Then nCortesia = 1: Command13(6).BackColor = &HFF&

nFlag = False

'=======================
'=======================
    'COPIA DE IMPRESION DE LA COMANDA QUE SE ENVIA A LA COCINA
    'SOLO QUE ESTA COPIA APARECE EN LA IMPRESORA 950 DE FACTURACION
If GetFromINI("Facturacion", "RepetirCocinaEnFacturacion", App.Path & "\soloini.ini") = "Pereza" Then
    Call CocinaOtraVez
End If
'=======================
'=======================

'PRIMERO SELECCIONO COCINA
rsCocina.Open "SELECT A.MESERO,A.MESA,A.LIN,A.DESCRIP,A.CANT,A.IMPRESO,A.IMPRESORA, " & _
        " (B.NOMBRE + ' ' + B.APELLIDO) AS NOMBRE" & _
        " FROM TMP_TRANS AS A, MESEROS AS B " & _
        " WHERE A.MESA = " & nMesa & _
        " AND A.MESERO = B.NUMERO " & _
        " AND A.IMPRESO = FALSE " & _
        " AND A.IMPRESORA = " & COCINA_01 & _
        " ORDER BY A.LIN,A.IMPRESORA ", msConn, adOpenStatic, adLockOptimistic
If rsCocina.EOF Then
    'MsgBox "NO HAY ARTICULOS PENDIENTES PARA ENVIAR A LA COCINA o LA BARRA", vbInformation, BoxTit
    nFlag = True
    rsCocina.Close
    GoTo ChequeaBarra:
End If

'Seleccion_KitchenPrinter
Sys_Pos.ImprCocina.PrintNormal PTR_S_RECEIPT, Space(2) & Chr(&HD) & Chr(&HA)
Sys_Pos.ImprCocina.PrintNormal PTR_S_RECEIPT, "Impresora:" & Sys_Pos.ImprCocina.DeviceName & "-" & Chr(&HD) & Chr(&HA)
Sys_Pos.ImprCocina.PrintNormal PTR_S_RECEIPT, Format(Date, "LONG DATE") & Chr(&HD) & Chr(&HA)
Sys_Pos.ImprCocina.PrintNormal PTR_S_RECEIPT, Time & Chr(&HD) & Chr(&HA)
Sys_Pos.ImprCocina.PrintNormal PTR_S_RECEIPT, Space(2) & Chr(&HD) & Chr(&HA)
Sys_Pos.ImprCocina.PrintNormal PTR_S_RECEIPT, "SOLICITUD DE PLATOS" & Chr(&HD) & Chr(&HA)
Sys_Pos.ImprCocina.PrintNormal PTR_S_RECEIPT, "Mesero : " & rsCocina!nombre & Chr(&HD) & Chr(&HA)
Sys_Pos.ImprCocina.PrintNormal PTR_S_RECEIPT, "Mesa # : " & rsCocina!mesa & Chr(&HD) & Chr(&HA)
If rs00!mesa_barra = nMesa Then Sys_Pos.ImprCocina.PrintNormal PTR_S_RECEIPT, "P A R A   L L E V A R" & Chr(&HD) & Chr(&HA)
Sys_Pos.ImprCocina.PrintNormal PTR_S_RECEIPT, "---------------------------------" & Chr(&HD) & Chr(&HA)
Do Until rsCocina.EOF
    If Mid(LTrim(rsCocina!DESCRIP), 1, 2) = "@@" Then
        Sys_Pos.ImprCocina.PrintNormal PTR_S_RECEIPT, Space(3) & Mid(rsCocina!DESCRIP, 1, 26) & Chr(&HD) & Chr(&HA)
    Else
        Sys_Pos.ImprCocina.PrintNormal PTR_S_RECEIPT, Format(rsCocina!CANT, "##") & Space(2) & Mid(rsCocina!DESCRIP, 1, 26) & Chr(&HD) & Chr(&HA)
    End If
    rsCocina.MoveNext
Loop
Sys_Pos.ImprCocina.PrintNormal PTR_S_RECEIPT, Space(2) & Chr(&HD) & Chr(&HA)
Sys_Pos.ImprCocina.PrintNormal PTR_S_RECEIPT, Space(2) & Chr(&HD) & Chr(&HA)
rsCocina.MovePrevious
Sys_Pos.ImprCocina.PrintNormal PTR_S_RECEIPT, Space(8) & "FIN PEDIDO MESA #: " & rsCocina!mesa & Chr(&HD) & Chr(&HA)
Sys_Pos.ImprCocina.PrintNormal PTR_S_RECEIPT, Space(2) & Chr(&HD) & Chr(&HA)
Sys_Pos.ImprCocina.PrintNormal PTR_S_RECEIPT, Space(2) & Chr(&HD) & Chr(&HA)
Sys_Pos.ImprCocina.PrintNormal PTR_S_RECEIPT, Space(2) & Chr(&HD) & Chr(&HA)
Sys_Pos.ImprCocina.PrintNormal PTR_S_RECEIPT, Space(2) & Chr(&HD) & Chr(&HA)
Sys_Pos.ImprCocina.PrintNormal PTR_S_RECEIPT, Space(2) & Chr(&HD) & Chr(&HA)
Sys_Pos.ImprCocina.PrintNormal PTR_S_RECEIPT, Space(2) & Chr(&HD) & Chr(&HA)
rc = Sys_Pos.ImprCocina.CutPaper(100)
rsCocina.Close

ChequeaBarra:
'DESPUES SELECCIONO LA BARRA
rsCocina.Open "SELECT A.MESERO,A.MESA,A.LIN,A.DESCRIP,A.CANT,A.IMPRESO,A.IMPRESORA, " & _
        " (B.NOMBRE + ',' + B.APELLIDO) AS NOMBRE" & _
        " FROM TMP_TRANS AS A, MESEROS AS B " & _
        " WHERE A.MESA = " & nMesa & _
        " AND A.MESERO = B.NUMERO " & _
        " AND A.IMPRESO = FALSE " & _
        " AND A.IMPRESORA = " & BARRA_01 & _
        " ORDER BY A.LIN,A.IMPRESORA ", msConn, adOpenStatic, adLockOptimistic

If rsCocina.EOF Then
    rsCocina.Close
    Set rsCocina = Nothing
    'Seleccion_Impresora_Default
    If nFlag = True Then
        MsgBox "NO HAY ARTICULOS PENDIENTES PARA ENVIAR A LA COCINA o LA BARRA", vbInformation, BoxTit
        Exit Sub
    End If
    GoTo MarcaImpresos:
End If

'Seleccion_Impresora_Default
'Seleccion_BarraPrinter
Sys_Pos.ImpresoraBarra.PrintNormal PTR_S_RECEIPT, Format(Date, "SHORT DATE") & Space(4) & Time & Chr(&HD) & Chr(&HA)
Sys_Pos.ImpresoraBarra.PrintNormal PTR_S_RECEIPT, Space(2) & Chr(&HD) & Chr(&HA)
Sys_Pos.ImpresoraBarra.PrintNormal PTR_S_RECEIPT, "SOLICITUD DE BEBIDAS" & Chr(&HD) & Chr(&HA)
Sys_Pos.ImpresoraBarra.PrintNormal PTR_S_RECEIPT, "Mesero : " & rsCocina!nombre & Chr(&HD) & Chr(&HA)
Sys_Pos.ImpresoraBarra.PrintNormal PTR_S_RECEIPT, "Mesa # : " & rsCocina!mesa & Chr(&HD) & Chr(&HA)
If rs00!mesa_barra = nMesa Then Sys_Pos.ImpresoraBarra.PrintNormal PTR_S_RECEIPT, "P A R A   L L E V A R" & Chr(&HD) & Chr(&HA)
Sys_Pos.ImpresoraBarra.PrintNormal PTR_S_RECEIPT, "------------------------------" & Chr(&HD) & Chr(&HA)
Do Until rsCocina.EOF
    If Mid(LTrim(rsCocina!DESCRIP), 1, 2) = "@@" Then
        Sys_Pos.ImpresoraBarra.PrintNormal PTR_S_RECEIPT, Space(3) & rsCocina!DESCRIP & Chr(&HD) & Chr(&HA)
    Else
        Sys_Pos.ImpresoraBarra.PrintNormal PTR_S_RECEIPT, Format(rsCocina!CANT, "##") & Space(2) & rsCocina!DESCRIP & Chr(&HD) & Chr(&HA)
    End If
    rsCocina.MoveNext
Loop

Sys_Pos.ImpresoraBarra.PrintNormal PTR_S_RECEIPT, Space(3) & Chr(&HD) & Chr(&HA)
Sys_Pos.ImpresoraBarra.PrintNormal PTR_S_RECEIPT, Space(3) & Chr(&HD) & Chr(&HA)
Sys_Pos.ImpresoraBarra.PrintNormal PTR_S_RECEIPT, Space(3) & Chr(&HD) & Chr(&HA)
Sys_Pos.ImpresoraBarra.PrintNormal PTR_S_RECEIPT, Space(3) & Chr(&HD) & Chr(&HA)
Sys_Pos.ImpresoraBarra.PrintNormal PTR_S_RECEIPT, Space(3) & Chr(&HD) & Chr(&HA)

Sys_Pos.ImpresoraBarra.CutPaper 100
rsCocina.Close

Set rsCocina = Nothing
'Seleccion_Impresora_Default

MarcaImpresos:
msConn.BeginTrans
msConn.Execute "UPDATE TMP_TRANS SET IMPRESO = TRUE WHERE MESA = " & nMesa
msConn.CommitTrans
Call BuscaMesa(False)
End Sub
Private Sub cmdSelMesa_Click()
If nCortesia = 0 Then nCortesia = 1: Command13(6).BackColor = &HFF&
Call BuscaMesa(False)
End Sub
Private Sub BuscaMesa(iOpc As Boolean)
Dim rsParciales As Recordset
Dim rsCuentas As Recordset
Dim lParc As Integer

lGo = False
StatMesa nMesa, 0

'Llama a la Pantalla que muestra las mesas (Ocupadas/Disponibles)
Mesas.Show 1
rs07.Open "SELECT a.lin,a.descrip,a.cant," & _
        " format(precio_unit,'##0.00') as mPrecio_unit," & _
        " format(precio,'##0.00') as mPrecio," & _
        " a.cajero,a.mesero,a.depto,a.plu,a.mesa,a.valid, " & _
        " a.envase,a.fecha,a.hora,a.tipo,a.descuento,a.cuenta, " & _
        " a.caja " & _
        " FROM tmp_trans as a " & _
        " WHERE a.mesa = " & nMesa & _
        " ORDER BY a.lin", msConn, adOpenStatic, adLockOptimistic
CajLin = rs07.RecordCount

Set rsCuentas = New Recordset
Set rsParciales = New Recordset

rsParciales.Open "SELECT MESA,SUM(MONTO) AS VALOR FROM TMP_PAR_PAGO " & _
            " WHERE MESA = " & nMesa & _
            " GROUP BY MESA", msConn, adOpenDynamic, adLockOptimistic

If rsParciales.EOF Then lParc = 0 Else lParc = 1

'Si la linea de detalle esta en 0, llama al mesero
If CajLin = 0 Then
    If nFlag = 0 Then
        Meseros.Show 1
    Else
        nMesero = 1
        nFlag = 0
        rs05.MoveFirst
        rs05.Find "numero = " & nMesero
        If Not rs05.EOF Then cNomMesero = rs05!nombre
    End If
Else
    nMesero = rs07!MESERO
    rs05.MoveFirst
    rs05.Find "numero = " & nMesero
    If Not rs05.EOF Then
        cNomMesero = rs05!nombre
    End If
End If

On Error Resume Next
    Set PlatosMesa.DataSource = rs07
On Error GoTo 0
SetupPantalla

If PlatosMesa.Rows <> 0 Then
    PlatosMesa.TopRow = (PlatosMesa.Rows - 1)
End If

If rs07.RecordCount > 0 Then
    If npNumCaj <> rs07!CAJERO Then
        '--- PARA MODULO DE MESEROS
        'MsgBox "USTED ES UN CAJERO DIFERENTE, SE LE PASARA ESTA MESA A SU NOMBRE", vbInformation, BoxTit
        msConn.BeginTrans
        msConn.Execute "UPDATE TMP_TRANS SET CAJERO = " & npNumCaj & _
                " WHERE MESA = " & nMesa
        msConn.CommitTrans
    End If
End If
rs07.Close
rs07.Open "SELECT sum(a.precio) as precio FROM tmp_trans as a WHERE a.mesa = " & nMesa, msConn, adOpenStatic, adLockReadOnly
SubTot = FormatCurrency(IIf(IsNull(rs07!precio), 0#, rs07!precio), 2)
On Error Resume Next
SubTot = FormatCurrency((SubTot + (rs07!precio * iISC)), 2)
iISCTransaccion = rs07!precio * iISC
SBTot = Format(SubTot, "standard")
On Error GoTo 0
rs07.Close

Text1(1) = cNomMesero
Text1(0) = nMesa

If nMesa = 0 Then
    Frame2(1).Enabled = False
    lbMensaje.BackColor = &HFFFF&
    lbMensaje = "¡¡ DEBE SELECCIONAR UNA MESA !!"
Else
    Frame2(1).Enabled = True
    lbMensaje.BackColor = &HFFFFFF
    lbMensaje = ""
End If

If lParc = 1 Then
    PlatosMesa.AddItem (CajLin + 1) & Chr(9) & _
    "PAGO.PARCIAL" & Chr(9) & 1 & Chr(9) & _
    Format(rsParciales!VALOR, "STANDARD") & Chr(9) & Format(rsParciales!VALOR, "STANDARD")
    SubTot = Format(SubTot - rsParciales!VALOR, "STANDARD")
    lbMensaje.BackColor = &HFFFF00
    lbMensaje = "MESA CON PAGOS PARCIALES"
Else
    lbMensaje.BackColor = &HFFFFFF
    lbMensaje = ""
End If

sqltxt = "SELECT MESA,CUENTA FROM TMP_CUENTAS " & _
        " WHERE MESA = " & nMesa & _
        " ORDER BY MESA,CUENTA"
rsCuentas.Open sqltxt, msConn, adOpenKeyset, adLockOptimistic

If rsCuentas.RecordCount > 0 Then lGo = True

If Not rsCuentas.EOF Then
    rsCuentas.MoveFirst
    nCta = rsCuentas!CUENTA
Else
    nCta = 0
End If

rsCuentas.Close
lbCuenta = nCta
If lGo = True Then
    If iOpc = False Then cmdCtas_Click
End If
End Sub

Private Sub Command1_Click()
Dim i As Integer
DoEvents
For i = 1 To 200
rs07.Open "SELECT a.lin,a.descrip,a.cant," & _
        " format(precio_unit,'##0.00') as mPrecio_unit," & _
        " format(precio,'##0.00') as mPrecio," & _
        " a.cajero,a.mesero,a.depto,a.plu,a.mesa,a.valid, " & _
        " a.envase,a.fecha,a.hora,a.tipo,a.descuento,a.cuenta, " & _
        " a.caja " & _
        " FROM tmp_trans as a " & _
        " WHERE a.mesa = " & nMesa & _
        " ORDER BY a.lin", msConn, adOpenStatic, adLockOptimistic
On Error GoTo ErrAdm:
    Set PlatosMesa.DataSource = rs07
    SetupPantalla
    rs07.Close
On Error GoTo 0
Next
MsgBox "200 ok"
Exit Sub

ErrAdm:
MsgBox Err.Source & "---" & Err.Description, vbCritical, BoxTit
End Sub

Private Sub Command13_Click(Index As Integer)
Dim DescResp As Variant

Select Case Index
Case 0  'DESCUENTO
    'Descuento al ultimo producto de la lista
    If nCortesia = 0 Then nCortesia = 1: Command13(6).BackColor = &HFF&
    txtInfo = "Escriba Clave para DESCUENTO"
    AskClave.Show 1
    If OkAnul = 1 Then
        Call DescProducto
        OkAnul = 0
        Call SetupPantalla
    End If
Case 1  'ANULACION DE LINEA
    If nCortesia = 0 Then nCortesia = 1: Command13(6).BackColor = &HFF&
    txtInfo = "Escriba Clave para ANULAR Linea"
    AskClave.Show 1
    If OkAnul = 1 Then
        BorraLin.Show 1
        OkAnul = 0
        Call SetupPantalla
    Else
        MsgBox "NO Tiene AUTORIZACION para ANULAR esta linea", vbExclamation, BoxTit
    End If
Case 2  'Reporte de X
    If nCortesia = 0 Then nCortesia = 1: Command13(6).BackColor = &HFF&
    If REPCAJAX_OK = True Then
        RptCajas.RepCajX
        MsgBox "GENERACION DEL REPORTE EN (X) HA FINALIZADO", vbInformation, BoxTit
        cmdSalir_Click
    Else
        MsgBox "ESTA OPCION NO ESTA DISPONIBLE. CONTACTE A SU ADMINISTRADOR", vbInformation, BoxTit
    End If
    'HacerRep_X
Case 3  'Impresion de Precuenta
    If nCortesia = 0 Then nCortesia = 1: Command13(6).BackColor = &HFF&
    Call ImprPreCta
    Call cmdSelMesa_Click
Case 4
    Dim rsTempo As Recordset
    
    If nCortesia = 0 Then nCortesia = 1: Command13(6).BackColor = &HFF&
    
    Set rsTempo = New Recordset
    rsTempo.Open "SELECT DISTINCT CUENTA FROM TMP_TRANS" & _
            " WHERE MESA = " & nMesa & " ORDER BY CUENTA ", msConn, adOpenStatic, adLockOptimistic
    If rsTempo.EOF Then
        Pagos.Show 1
        Call BuscaMesa(False)
        Exit Sub
    End If
    If rsTempo!CUENTA <> 0 Then
    'If rsTempo.RecordCount > 1 Then
        'PAGOS POR CUENTA
        rsTempo.Close
        sqltxt = "SELECT CUENTA,SUM(PRECIO) AS VALOR "
        sqltxt = sqltxt & " FROM TMP_TRANS "
        sqltxt = sqltxt & " WHERE MESA = " & nMesa
        sqltxt = sqltxt & " AND CUENTA = " & nCta
        sqltxt = sqltxt & " GROUP BY CUENTA "
        rsTempo.Open sqltxt, msConn, adOpenStatic, adLockOptimistic
        If rsTempo.RecordCount <= 0 Then
            MsgBox "NO HAY PLATOS MARCADOS PARA ESTA CUENTA", vbInformation, "SELECCIONE OTRA CUENTA POR FAVOR       "
            Exit Sub
        End If
        Set rsTempo = Nothing
        CtaPago.Show 1
    Else
        Pagos.Show 1
    End If
    Call BuscaMesa(False)
Case 5
    'PAGOS PARCIALES
    Dim rsTempo01 As Recordset
    
    If nCortesia = 0 Then nCortesia = 1: Command13(6).BackColor = &HFF&
    
    Set rsTempo01 = New Recordset
    rsTempo01.Open "SELECT DISTINCT CUENTA FROM TMP_TRANS" & _
            " WHERE MESA = " & nMesa, msConn, adOpenStatic, adLockOptimistic
    If rsTempo01.EOF Then
        PagParcial.Show 1
        Call BuscaMesa(False)
        SetupPantalla
        Exit Sub
    End If
    If rsTempo01!CUENTA <> 0 Then
        MsgBox "NO PUEDE HACER PAGOS PARCIALES A UNA MESA QUE TIENE CUENTAS ABIERTAS", vbInformation, BoxTit
    Else
        PagParcial.Show 1
        Call BuscaMesa(False)
        SetupPantalla
    End If
Case 6
    'CORTESIA DE LA CASA. EL PLATO DEBE APARECER CON PRECIO 0.00
    'COLOR ORIGINAL = &HFF&
    nCortesia = 0
    Command13(6).BackColor = &HFF00&
End Select
nNLinSel = 0
End Sub

Private Sub Command2_Click()
'AVANZA HACIA ABAJO
Dim num As Integer
num = 0
nNLinSel = 0

'Llamar proc. de limpiar deptos anteriores
Quita_Subrallado (67)
QuitarDeptos

If rs02.EOF = True Then
    rs02.MovePrevious
    cmdDepto(num).Caption = rs02!corto
End If

Do Until rs02.EOF
    If num < 1 Then
        cmdDepto(num).Caption = rs02!corto
        Arreg_Deptos(num) = rs02!CODIGO
    Else
        If Not IsObject(cmdDepto(num)) Then
           Load cmdDepto(num)
        End If
        cmdDepto(num).Caption = rs02!corto
        Arreg_Deptos(num) = rs02!CODIGO
        cmdDepto(num).Left = 120
        cmdDepto(num).Top = cmdDepto(num - 1).Top + 660
        cmdDepto(num).Visible = True
    End If
    num = num + 1
    If num = 11 Then Exit Do
    rs02.MoveNext
Loop

End Sub

Private Sub Command3_Click()
'AVANZA HACIA ARRIBA
Dim nNum As Integer
num = 0
nNLinSel = 0

rs02.MoveFirst
rs02.Find "codigo = " & Arreg_Deptos(0)

If rs02.EOF Then
    'El PG-DOWN llego a la ultima pantalla
    rs02.MovePrevious
    'cmdDepto(num).Caption = rs02!CORTO
End If

rs02.Move -11
If rs02.BOF Then rs02.MoveFirst
'If (nNum - 11) <= 0 Then
'    ' Desde el principio
'    rs02.MoveFirst
'Else
'    'rs02.Move (-12)
'    rs02.Move (-11)
'End If
'Llamar proc. de limpiar deptos anteriores
Quita_Subrallado (67)
QuitarDeptos

'CARGANDO EL CODIGO DE LOS DEPARTAMENTO EN LOS BOTONES DISPONIBLES
Do Until rs02.EOF
    If num < 1 Then
        cmdDepto(num).Caption = rs02!corto
        Arreg_Deptos(num) = rs02!CODIGO
    Else
        If Not IsObject(cmdDepto(num)) Then
           Load cmdDepto(num)
        End If
        cmdDepto(num).Caption = rs02!corto
        Arreg_Deptos(num) = rs02!CODIGO
        cmdDepto(num).Left = 120
        cmdDepto(num).Top = cmdDepto(num - 1).Top + 660
        cmdDepto(num).Visible = True
    End If
    num = num + 1
    If num = 11 Then Exit Do
    rs02.MoveNext
Loop

End Sub

Private Sub cmdPlus_Click(Index As Integer)
Dim CadenaSql As String
Dim nLineas As Long
Dim i As Integer
Dim rsParciales As Recordset
Dim lParc As Integer
Dim SOLO_FECHA As String
Dim nAcoTop As Integer

If cmdPlus(Index).Tag = "" Then Beep: Exit Sub

i = 0
'Si quiere marcar PRODUCTOS y no hay Mesero, EXIGIRLO!!!!
Do Until nMesero > 0
    Meseros.Show 1
Loop

On Error Resume Next
    rs03.MoveFirst
On Error GoTo 0

rs03.Find "codigo = " + cmdPlus(Index).Tag

nAcoTop = 240

If rs03!ENVASES = True Then    'El producto tiene Envase(s)
    If nGlobEnv < 1 Then
        MsgBox "Por Favor Seleccione ENVASE", vbInformation, BoxTit
        nCortesia = 1: Command13(6).BackColor = &HFF&
        Exit Sub
    End If

    CadenaSql = "SELECT a.contenedor,a.codigo,a.precio, " & _
            " b.depto,b.descrip,b.corto,b.IMPRESORA " & _
            " FROM CONTEND_02 as a, PLU as b " & _
            " WHERE a.CODIGO = " & rs03!CODIGO & _
            " AND a.CONTENEDOR = " & nGlobEnv & _
            " AND a.codigo = b.codigo "
    rs09.Open CadenaSql, msConn, adOpenStatic, adLockReadOnly
    
    If rs09.EOF Then
        MsgBox "Por Favor Seleccione ENVASE", vbInformation, BoxTit
        rs09.Close
        nCortesia = 1: Command13(6).BackColor = &HFF&
        Exit Sub
    End If

    CajLin = CajLin + 1
    
    SOLO_FECHA = Format(Date, "YYYYMMDD")
    
    CadenaSql = "INSERT INTO TMP_TRANS " & _
        "(CAJA,CAJERO,MESA,MESERO,VALID,LIN,DESCRIP,CANT,DEPTO,PLU,ENVASE,PRECIO_UNIT,PRECIO,FECHA,HORA,TIPO,DESCUENTO,CUENTA,IMPRESO,IMPRESORA) VALUES (" & _
        "" & cCaja & "," & npNumCaj & "," & nMesa & "," & nMesero & "," & -1 & "," & CajLin & "," & "'" & _
        rs09!corto + TextEnv & "'" & "," & nCantidad & "," & rs09!depto & "," & rs09!CODIGO & "," & _
        nGlobEnv & "," & (rs09!precio * nCortesia) & "," & (rs09!precio * nCortesia * nCantidad) & "," & "'" & SOLO_FECHA & "'" & "," & "'" & Time & "'" & _
        ",'  '," & 0# & "," & nCta & ",FALSE," & rs09!IMPRESORA & ")"
    nSeleccionCocina = rs09!IMPRESORA
Else
    CajLin = CajLin + 1

    SOLO_FECHA = Format(Date, "YYYYMMDD")
    
    CadenaSql = "INSERT INTO TMP_TRANS " & _
        "(CAJA,CAJERO,MESA,MESERO,VALID,LIN,DESCRIP,CANT,DEPTO,PLU,ENVASE,PRECIO_UNIT,PRECIO,FECHA,HORA,TIPO,DESCUENTO,CUENTA,IMPRESO,IMPRESORA) VALUES (" & _
        "" & cCaja & "," & npNumCaj & "," & nMesa & "," & nMesero & "," & -1 & "," & CajLin & "," & "'" & _
        rs03!DESCRIP + TextEnv & "'" & "," & nCantidad & "," & rs03!depto & "," & rs03!CODIGO & "," & _
        0 & "," & (rs03!precio1 * nCortesia) & "," & (rs03!precio1 * nCortesia * nCantidad) & "," & "'" & SOLO_FECHA & "'" & "," & "'" & Time & "'" & _
        ",'  '," & 0# & "," & nCta & ",FALSE," & rs03!IMPRESORA & ")"
    nSeleccionCocina = rs03!IMPRESORA
End If

msConn.BeginTrans
msConn.Execute CadenaSql
msConn.CommitTrans

'---------  CORTESIA ----------------
nCortesia = 1: Command13(6).BackColor = &HFF&
'------------------------------------
' ------ TRATAMIENTO DE ACOMPAÑANTES
If rsTmpAco.State = adStateOpen Then rsTmpAco.Close
rsTmpAco.Open "SELECT A.PLU_ID,A.ACOMP_ID,B.DESCRIP " & _
        " FROM PLU_ACOMP AS A, ACOMPA AS B " & _
        " WHERE A.PLU_ID = " & rs03!CODIGO & _
        " AND A.ACOMP_ID = B.CODIGO " & _
        " ORDER BY B.DESCRIP ", msConn, adOpenStatic, adLockOptimistic

If rsTmpAco.EOF Then
    If rsTmpAco.State = adStateOpen Then rsTmpAco.Close
    cmdRestoAco.Enabled = False
    cmdAcomp(0).Caption = ""
    cmdAcomp(0).Tag = 0
    For iLocal = 1 To cmdAcomp.Count - 1
        cmdAcomp(iLocal).Visible = False
    Next
    Frame2(2).Enabled = False
Else
    Frame2(2).Enabled = True
    iAcom = 0: iLocal = 0
    For iLocal = 1 To 3
        On Error Resume Next
            Load cmdAcomp(iLocal)
        On Error GoTo 0
    Next
    On Error Resume Next
    Do Until rsTmpAco.EOF
        If iAcom = 0 Then
            cmdAcomp(iAcom).Visible = True
            cmdAcomp(iAcom).Caption = rsTmpAco!DESCRIP
            cmdAcomp(iAcom).Tag = rs03!depto
            iAcom = iAcom + 1
            rsTmpAco.MoveNext
            If rsTmpAco.EOF Then Exit Do
        End If
        
        cmdAcomp(iAcom).Visible = True
        cmdAcomp(iAcom).Top = nAcoTop + 600
        cmdAcomp(iAcom).Caption = rsTmpAco!DESCRIP
        cmdAcomp(iAcom).Tag = rs03!depto
        iAcom = iAcom + 1
        nAcoTop = nAcoTop + 600
        nAcoBookMark = rsTmpAco.Bookmark
        rsTmpAco.MoveNext
        If iAcom = 4 Then cmdRestoAco.Enabled = True: Exit Do
    Loop
    On Error GoTo 0
'    Do Until rsTmpAco.EOF
'        cmdAcomp(iAcoCnt).Caption = rsTmpAco!descrip
'        cmdAcomp(iAcoCnt).Tag = rs03!depto
'        iAcoCnt = iAcoCnt + 1
'        rsTmpAco.MoveNext
'    Loop
'    cmdAcomp(iAcoCnt).Caption = "SIN ACOMPAÑANTE"
'    cmdAcomp(iAcoCnt).Tag = rs03!depto
End If

' ------------------FIN TRATAMIENTO DE ACOMPAÑANTES---------------------------------
If rs03!ENVASES = True Then rs09.Close

If CajLin = 1 Then
    msConn.Execute "UPDATE Mesas SET ocupada = TRUE " & _
           " WHERE numero = " & nMesa
End If

rs07.Open "SELECT a.lin,a.descrip,a.cant," & _
        " format(precio_unit,'##0.00') as mPrecio_unit," & _
        " format(precio,'##0.00') as mPrecio," & _
        " a.cajero,a.mesero,a.depto,a.plu,a.mesa,a.valid, " & _
        " a.envase,a.fecha,a.hora,a.tipo,a.descuento,a.cuenta, " & _
        " a.caja " & _
        " FROM tmp_trans as a " & _
        " WHERE a.mesa = " & nMesa & _
        " AND A.CUENTA = " & nCta & _
        " ORDER BY a.lin", msConn, adOpenStatic, adLockOptimistic

On Error GoTo ErrAdm:
Set PlatosMesa.DataSource = rs07
On Error GoTo 0
SetupPantalla

nLineas = PlatosMesa.Rows - 1

Set rsParciales = New Recordset
rsParciales.Open "SELECT MESA,SUM(MONTO) AS VALOR FROM TMP_PAR_PAGO " & _
            " WHERE MESA = " & nMesa & _
            " GROUP BY MESA", msConn, adOpenDynamic, adLockOptimistic

If rsParciales.EOF Then lParc = 0 Else lParc = 1

rs07.Close
rs07.Open "SELECT sum(a.precio) as precio FROM tmp_trans as a " & _
          " WHERE a.mesa = " & nMesa & _
          " AND A.CUENTA = " & nCta, msConn, adOpenStatic, adLockReadOnly
SubTot = FormatCurrency(IIf(IsNull(rs07!precio), 0#, rs07!precio), 2)
On Error Resume Next
SubTot = FormatCurrency((SubTot + (rs07!precio * iISC)), 2)
iISCTransaccion = rs07!precio * iISC
SBTot = Format(SubTot, "standard")
On Error GoTo 0
rs07.Close
If (PlatosMesa.Rows - 1) >= 1 Then
        PlatosMesa.TopRow = (PlatosMesa.Rows - 1)
End If
nCantidad = 1: nPase = 0
nNLinSel = 0
Text1(2) = nCantidad

If lParc = 1 Then
    PlatosMesa.AddItem (CajLin + 1) & Chr(9) & _
        "PAGO.PARCIAL" & Chr(9) & 1 & Chr(9) & _
        Format(rsParciales!VALOR, "STANDARD") & Chr(9) & Format(rsParciales!VALOR, "STANDARD")
    SubTot = Format(SubTot - rsParciales!VALOR, "STANDARD")
    lbMensaje.BackColor = &HFFFF00
    lbMensaje = "MESA CON PAGOS PARCIALES"
Else
    lbMensaje.BackColor = &HFFFFFF
    lbMensaje = ""
End If
Exit Sub

ErrAdm:
    EscribeLog (Err.Number & " - " & Err.Description)
    Resume Next
End Sub

Private Sub Command5_Click()
Dim nNum As String
num = 0
'Tengo que saber quien se ve de primero para mostrar
'los 11 anteriores

If rs03.EOF Then
    If rs03.BOF Then Exit Sub
    rs03.MovePrevious
    cmdPlus(num).Caption = rs03!DESCRIP
    cmdPlus(numplu).Tag = rs03!CODIGO
End If
nNum = rs03!CODIGO

If (numplu - 17) <= 0 Then
    ' Desde el principio
    rs03.MoveFirst
Else
    rs03.Move (-17)
    If rs03.BOF Then
        rs03.MoveFirst
    End If
End If
'Llamar proc. de limpiar deptos anteriores
'Quita_Subrallado (67)
'QuitarDeptos
'QuitarPLUS
MiTop = 240: StayLeft = 120
MiLeft = 0: numplu = 0

Do Until rs03.EOF
    If numplu < 1 Then
        cmdPlus(numplu).Caption = rs03!DESCRIP
        cmdPlus(numplu).Tag = rs03!CODIGO
        'Muestra los PLUs del primer departamento
    Else
        If Not IsObject(cmdPlus(numplu)) Then
           Load cmdPlus(numplu)
        End If
        cmdPlus(numplu).Visible = True
        cmdPlus(numplu).Caption = rs03!DESCRIP
        cmdPlus(numplu).Tag = rs03!CODIGO
        cmdPlus(numplu).Left = MiLeft + StayLeft
        cmdPlus(numplu).Top = MiTop
        StayLeft = 120
    End If
    numplu = numplu + 1
    MiLeft = MiLeft + 2400
    If numplu = 3 Or numplu = 6 Or numplu = 9 Or numplu = 12 Or numplu = 15 Then
        MiTop = MiTop + 600
        MiLeft = 0
    End If
    If numplu = 18 Then Exit Do
    rs03.MoveNext
Loop
nNLinSel = 0
End Sub

Private Sub Command6_Click()

If cmdPlus(17).Visible = False Then
    Exit Sub
End If

numplu = 0
'Llamar proc. de limpiar PLUS anteriores
QuitarPLUS
If rs03.EOF Then
    rs03.MovePrevious
    cmdPlus(numplu).Caption = rs03!DESCRIP
End If

MiTop = 240: StayLeft = 120
MiLeft = 0: numplu = 0

Do Until rs03.EOF
    If numplu < 1 Then
        cmdPlus(numplu).Caption = rs03!DESCRIP
        cmdPlus(numplu).Tag = rs03!CODIGO
        'Muestra los PLUs del primer departamento
    Else
        If Not IsObject(cmdPlus(numplu)) Then
           Load cmdPlus(numplu)
        End If
        cmdPlus(numplu).Visible = True
        cmdPlus(numplu).Caption = rs03!DESCRIP
        cmdPlus(numplu).Tag = rs03!CODIGO
        cmdPlus(numplu).Left = MiLeft + StayLeft
        cmdPlus(numplu).Top = MiTop
        StayLeft = 120
    End If
    numplu = numplu + 1
    MiLeft = MiLeft + 2400
    If numplu = 3 Or numplu = 6 Or numplu = 9 Or numplu = 12 Or numplu = 15 Then
        MiTop = MiTop + 600
        MiLeft = 0
    End If
    If numplu = 18 Then Exit Do
    rs03.MoveNext
Loop
nNLinSel = 0
End Sub

Private Sub Command8_Click(Index As Integer)
Dim cCant As String

On Error GoTo FixError:

If nPase = 0 Then
    nCantidad = Command8(Index).Index
Else
    cCant = Str(nCantidad)
    cCant = cCant & Command8(Index).Index
    nCantidad = Val(cCant)
    If Len(cCant) = 6 Then
        MsgBox "CANTIDAD/MONTO NO ES VALIDO, ESTABLECIENDO UNO (1)", vbExclamation, BoxTit
        nPase = 0
        nCantidad = 1
    End If
End If

On Error GoTo 0

Text1(2) = nCantidad
nPase = nPase + 1
Exit Sub

FixError:
    nCantidad = 1
    Resume Next
End Sub
Private Sub Correccion_Click()
'------------------- CORRECCION / ERROR CORRECT ----------------
Dim rsFixTmpTrans As Recordset
Dim txto As String
Dim rsParciales As Recordset
Dim lParc As Integer
Dim sqltext As String
Dim SSD As Single
Dim nTp  As Integer
Dim nn, i As Integer
Dim zz As Integer
Dim SOLO_FECHA As String
Dim nVeriCant As Integer
'---------------------
Dim nLocLin As Integer
Dim nLocCan As Integer

nNLinSel = 0: nTp = 0
OkAnul = 0
If GetFromINI("Meseros", "PermiteCorreccion", App.Path & "\soloini.ini") = "Pereza" Then
    txtInfo = "Escriba Clave para ANULAR Linea"
    AskClave.Show 1
    If OkAnul = 1 Then
    Else
        MsgBox "NO Tiene AUTORIZACION para CORREGIR esta linea", vbExclamation, BoxTit
        Exit Sub
    End If
End If

On Error Resume Next
    PlatosMesa.Row = PlatosMesa.Rows - 1
    PlatosMesa.Col = 0
    nLocLin = Val(PlatosMesa.Text)
    PlatosMesa.Col = 2
    nLocCan = PlatosMesa.Text
On Error GoTo 0

On Error GoTo ErrAdm:
Set rsFixTmpTrans = New Recordset
txto = "SELECT * FROM tmp_trans " & _
    " WHERE mesa = " & nMesa & " AND lin = " & nLocLin
rsFixTmpTrans.Open txto, msConn, adOpenStatic, adLockReadOnly

If rsFixTmpTrans.EOF = True Then
    rsFixTmpTrans.Close
    Exit Sub
End If

If rsFixTmpTrans!CANT < 0 Then
    MsgBox "NO puede CORREGIR este Producto", vbExclamation, BoxTit
    rsFixTmpTrans.Close
    Exit Sub
End If

If Mid(rsFixTmpTrans!TIPO, 1, 1) = "B" Then
    MsgBox "PRODUCTO YA FUE ANULADO/CORREGIDO/SE DIO DESCUENTO EN LA LINEA " & Val(Mid(rsFixTmpTrans!TIPO, 5, 2)), vbExclamation, BoxTit
    rsFixTmpTrans.Close
    Exit Sub
End If

'---------------------------------------

nn = 0: i = 1: zz = 0
'Pregunta si hay un Numero en TIPO, si hay significa que tiene Desc
For i = i To 9
    nn = InStr(1, rsFixTmpTrans!TIPO, i)
    If nn <> 0 Then Exit For
Next
If nn <> 0 Then zz = Val(Mid(rsFixTmpTrans!TIPO, nn, 2))
'----------------------------------------

SSD = rsFixTmpTrans!precio * (-1)

SOLO_FECHA = Format(Date, "YYYYMMDD")
CajLin = CajLin + 1

'INSERTA LA LINEA DE CORRECCION
CadenaSql = "INSERT INTO TMP_TRANS " & _
    "(CAJA,CAJERO,MESA,MESERO,VALID,LIN,DESCRIP,CANT,DEPTO,PLU,ENVASE,PRECIO_UNIT,PRECIO,FECHA,HORA,TIPO,DESCUENTO,CUENTA,IMPRESO,IMPRESORA) VALUES (" & _
    "" & cCaja & "," & npNumCaj & "," & nMesa & "," & nMesero & "," & -1 & "," & CajLin & "," & "'EC-" & _
    rsFixTmpTrans!DESCRIP & "'" & "," & rsFixTmpTrans!CANT * (-1) & "," & rsFixTmpTrans!depto & "," & rsFixTmpTrans!PLU & "," & _
    rsFixTmpTrans!envase & "," & rsFixTmpTrans!precio_unit * (-1) & "," & SSD & "," & "'" & SOLO_FECHA & "'" & "," & "'" & Time & "'" & _
    ",'EC-" & CajLin - 1 & "'," & 0# & "," & rsFixTmpTrans!CUENTA & ",FALSE," & rsFixTmpTrans!IMPRESORA & ")"

'MARCA LA LINEA QUE SE ESTA CORRIGIENDO PARA QUE NO PUEDA SER CORREGIDA/ANULADA
'DE NUEVO
sqltext = "UPDATE TMP_TRANS SET VALID = 0,TIPO = 'BEC" & Str(CajLin) & _
          "' WHERE MESA = " & nMesa & _
          " AND LIN = " & (nLocLin)

msConn.BeginTrans
''''''''''''''' msConnLoc.BeginTrans
If zz > 0 Then
    sqltxt = "UPDATE TMP_TRANS SET TIPO = ' ' WHERE MESA = " & nMesa & _
        " AND LIN = " & zz
End If
msConn.Execute CadenaSql
msConn.Execute sqltext
If zz > 0 Then
    msConn.Execute sqltxt
End If
msConn.CommitTrans

rs07.Open "SELECT a.lin,a.descrip,a.cant," & _
    " format(precio_unit,'##0.00') as mPrecio_unit," & _
    " format(precio,'##0.00') as mPrecio," & _
    " a.cajero,a.mesero,a.depto,a.plu,a.mesa,a.valid, " & _
    " a.envase,a.fecha,a.hora,a.tipo,a.descuento,a.cuenta, " & _
    " a.caja " & _
    " FROM tmp_trans as a " & _
    " WHERE a.mesa = " & nMesa & _
    " AND A.CUENTA = " & nCta & _
    " ORDER BY a.lin ", msConn, adOpenStatic, adLockOptimistic

Set PlatosMesa.DataSource = rs07
SetupPantalla

nLineas = PlatosMesa.Rows - 1

If (PlatosMesa.Rows - 1) >= 1 Then
    PlatosMesa.TopRow = (PlatosMesa.Rows - 1)
End If

Set rsParciales = New Recordset
rsParciales.Open "SELECT MESA,SUM(MONTO) AS VALOR" & _
        " FROM TMP_PAR_PAGO " & _
        " WHERE MESA = " & nMesa & _
        " GROUP BY MESA", msConn, adOpenDynamic, adLockOptimistic

If rsParciales.EOF Then lParc = 0 Else lParc = 1

rs07.Close
rs07.Open "SELECT sum(a.precio) as precio FROM tmp_trans as a " & _
      " WHERE a.mesa = " & nMesa & _
      " AND A.CUENTA = " & nCta, msConn, adOpenStatic, adLockReadOnly
SubTot = FormatCurrency(IIf(IsNull(rs07!precio), 0#, rs07!precio), 2)
On Error Resume Next
SubTot = FormatCurrency((SubTot + (rs07!precio * iISC)), 2)
iISCTransaccion = rs07!precio * iISC
SBTot = Format(SubTot, "standard")
On Error GoTo 0
rs07.Close
rsFixTmpTrans.Close

nCantidad = 1: nPase = 0
Text1(2) = nCantidad

If lParc = 1 Then
    PlatosMesa.AddItem (CajLin + 1) & Chr(9) & _
        "PAGO.PARCIAL" & Chr(9) & 1 & Chr(9) & _
        Format(rsParciales!VALOR, "STANDARD") & Chr(9) & Format(rsParciales!VALOR, "STANDARD")
    SubTot = Format(SubTot - rsParciales!VALOR, "STANDARD")
    lbMensaje.BackColor = &HFFFF00
    lbMensaje = "MESA CON PAGOS PARCIALES"
Else
    lbMensaje.BackColor = &HFFFFFF
    lbMensaje = ""
End If
On Error GoTo 0
Exit Sub

ErrAdm:
MsgBox Err.Number & " " & Err.Description, vbCritical, BoxTit
Resume Next
End Sub
Private Sub Form_Load()
Dim MiTop As Integer, MiLeft As Integer, StayLeft As Integer
Dim iTam As Integer
Dim rsCuentas As Recordset

Set rs01 = New Recordset
Set rs02 = New Recordset
Set rs03 = New Recordset
Set rs04 = New Recordset
Set rs05 = New Recordset
Set rs06 = New Recordset
Set rs07 = New Recordset
Set rs08 = New Recordset 'Para Precios de PLU con Envases
Set rs09 = New Recordset 'Para Precios de PLU con Envases
Set rsParciales = New Recordset
Set rsCuentas = New Recordset

nCantidad = 1: cCaja = 1: Text1(2) = nCantidad
num = 0: iTam = 0: nMesa = 0: CajLin = 0: nPase = 0:
nGlobEnv = 0: nCta = 0: nCliNum = 0
nCortesia = 1

nFlag = 0
OkAnul = 0
nNLinSel = 0    'Linea Seleccionada
OKProp = 0
OKDesc = 0
OKCancelar = 0

SetupPantalla

Show    'Muestra la Pantalla de Facturacion

'Mesas
rs01.Open "SELECT numero, iif(ocupada=TRUE,'Ocupada','Libre') AS status FROM mesas", msConn, adOpenDynamic, adLockOptimistic
'Departamentos
rs02.Open "SELECT codigo, corto, abierto FROM depto ORDER BY ORDEN", msConn, adOpenDynamic, adLockOptimistic
If rs02.EOF Then
    MsgBox "NO EXISTEN DEPARTAMENTOS. ES NECESARIO CREAR DEPARTAMENTOS DE VENTAS. EL PROGRAMA TERMINARA AHORA", vbCritical, BoxTit
    Unload Me
    End
End If
'PLUS del Primer Departamento
rs03.Open "SELECT codigo,depto,descrip,corto,precio1,envases,IMPRESORA " & _
        " FROM PLU " & _
        " WHERE depto = " & rs02!CODIGO & " ORDER BY CORTO", msConn, adOpenStatic, adLockOptimistic
'Contened_01
rs04.Open "SELECT a.depto,a.contenedor,b.descrip FROM contend_01 as a,contened as b WHERE a.DEPTO = " & rs02!CODIGO & " and a.contenedor = b.contenedor ORDER BY a.depto,a.contenedor", msConn, adOpenStatic, adLockOptimistic
'Meseros
rs05.Open "SELECT numero,nombre,apellido FROM meseros WHERE NUMERO <> 999 ORDER BY NUMERO", msConn, adOpenStatic, adLockReadOnly
If rs05.EOF Then
    MsgBox "NO EXISTEN MESEROS/SALONEROS. ES NECESARIO CREARLOS. EL PROGRAMA TERMINARA AHORA", vbCritical, BoxTit
    Unload Me
    End
End If
'Cajas
rs06.Open "SELECT caja_cod, descrip FROM cajas WHERE caja_cod = " & cCaja, msConn, adOpenStatic, adLockReadOnly

If TipoApplicacion <> "" Then
    Command13(1).Enabled = False
    Correccion.Enabled = False
End If

If ON_LINE Then
    PLU.Caption = PLU.Caption + "." + rs06!DESCRIP + "." + rs00!DESCRIP + ".ON-LINE" + TipoApplicacion
Else
    PLU.Caption = PLU.Caption + "." + rs06!DESCRIP + "." + rs00!DESCRIP + ".OFF-LINE" + TipoApplicacion
End If

BuscaMesa (True)   'Pantalla de Seleccion de Mesa

' Lo maximo que puede caber en el frame de Departamentos son 11
' botones. Indice es entonces 10

'CARGA LOS DATOS EN LOS BOTONES DEPARTAMENTALES DISPONIBLE
'Y SU VALOR EN EL ARREGLO ARREG_DEPTOS
Do Until rs02.EOF
    If num < 1 Then
        cmdDepto(num).Caption = rs02!corto
        Arreg_Deptos(num) = rs02!CODIGO
        ElDepto = rs02!CODIGO
    Else
        Load cmdDepto(num)
        cmdDepto(num).Caption = rs02!corto
        Arreg_Deptos(num) = rs02!CODIGO
        cmdDepto(num).Left = 120
        cmdDepto(num).Top = cmdDepto(num - 1).Top + 660
        cmdDepto(num).Visible = True
    End If
    num = num + 1
    If num = 11 Then Exit Do
    rs02.MoveNext
Loop

MiTop = 240: StayLeft = 120
MiLeft = 0: numplu = 0

'Muestra los PLUs(Botones) del primer departamento
For i = 1 To 18
    Load cmdPlus(i)
Next

Do Until rs03.EOF
    If numplu < 1 Then
        cmdPlus(numplu).Caption = rs03!DESCRIP
        cmdPlus(numplu).Tag = rs03!CODIGO
        'Arreg_Plu(numplu) = numplu
        'Muestra los PLUs del primer departamento
    Else
        'Load cmdPlus(numplu)
        cmdPlus(numplu).Visible = True
        cmdPlus(numplu).Caption = rs03!DESCRIP
        cmdPlus(numplu).Tag = rs03!CODIGO
        cmdPlus(numplu).Left = MiLeft + StayLeft
        cmdPlus(numplu).Top = MiTop
    End If
    numplu = numplu + 1
    MiLeft = MiLeft + 2400
    If numplu = 3 Or numplu = 6 Or numplu = 9 Or numplu = 12 Or numplu = 15 Then
        MiTop = MiTop + 600
        MiLeft = 0
    End If
    If numplu = 18 Then Exit Do
    rs03.MoveNext
Loop

Do Until rs04.EOF
    cmdEnvases(iTam).Caption = rs04!DESCRIP
    cmdEnvases(iTam).Tag = rs04!contenedor
    iTam = iTam + 1
    rs04.MoveNext
Loop

For iTam = 0 To 3
    If cmdEnvases(iTam).Caption = "" Then
        cmdEnvases(iTam).Enabled = False
    End If
Next

Text1(3) = cNomCaj
Text1(0) = nMesa
If nMesa = 0 Then
    Frame2(1).Enabled = False
    'lbMensaje.BackColor = &HFFFF&
    lbMensaje.BackColor = &HFFFF00
    lbMensaje = "! Debe Seleccionar una Mesa !"
Else
    Frame2(1).Enabled = True
    lbMensaje.BackColor = &HFFFFFF
    lbMensaje = ""
End If

sqltxt = "SELECT MESA,CUENTA " & _
        " FROM TMP_CUENTAS " & _
        " WHERE MESA = " & nMesa & _
        " ORDER BY MESA,CUENTA"
rsCuentas.Open sqltxt, msConn, adOpenDynamic, adLockOptimistic

If Not rsCuentas.EOF Then
    rsCuentas.MoveFirst
    nCta = rsCuentas!CUENTA
Else
    nCta = 0
End If

rsCuentas.Close
lbCuenta = nCta
If lGo = True Then cmdCtas_Click
End Sub
Private Sub PlatosMesa_Click()
If PlatosMesa.Rows = 0 Then
    MsgBox "DEBE MARCAR UN PLATO", vbExclamation, BoxTit
    Exit Sub
End If
nNLinSel = Val(PlatosMesa.Text)
PlatosMesa.Col = 16
nCta = Val(PlatosMesa.Text)
lbCuenta = nCta
PlatosMesa.Col = 0
End Sub

Private Sub CocinaOtraVez()
'COPIA DE IMPRESION DE LA COMANDA QUE SE ENVIA A LA COCINA
'SOLO QUE ESTA COPIA APARECE EN LA IMPRESORA 950 DE FACTURACION
Dim rsCocina As New ADODB.Recordset
Dim nFlag As Boolean
Dim i As Integer

If nCortesia = 0 Then nCortesia = 1: Command13(6).BackColor = &HFF&

nFlag = False
'SELECCIONO TODOS LOS PRODUCTOS, COMO EN LA COMANDA
rsCocina.Open "SELECT A.MESERO,A.MESA,A.LIN,A.DESCRIP,A.CANT,A.IMPRESO,A.IMPRESORA, " & _
        " (B.NOMBRE + ' ' + B.APELLIDO) AS NOMBRE" & _
        " FROM TMP_TRANS AS A, MESEROS AS B " & _
        " WHERE A.MESA = " & nMesa & _
        " AND A.MESERO = B.NUMERO " & _
        " AND A.IMPRESO = FALSE " & _
        " ORDER BY A.LIN,A.IMPRESORA ", msConn, adOpenStatic, adLockOptimistic
If rsCocina.EOF Then
    'MsgBox "NO HAY ARTICULOS PENDIENTES PARA ENVIAR A LA COCINA o LA BARRA", vbInformation, BoxTit
    nFlag = True
    rsCocina.Close
    Set rsCocina = Nothing
    Exit Sub
End If

'
'Sys_Pos.Coptr1.PrintNormal PTR_S_RECEIPT, "Cocina OtraVez" & Chr(&HD) & Chr(&HA)
Sys_Pos.Coptr1.PrintNormal PTR_S_RECEIPT, Space(2) & Chr(&HD) & Chr(&HA)
Sys_Pos.Coptr1.PrintNormal PTR_S_RECEIPT, "Impresora:" & Sys_Pos.Coptr1.DeviceName & "-" & Chr(&HD) & Chr(&HA)
Sys_Pos.Coptr1.PrintNormal PTR_S_RECEIPT, Format(Date, "LONG DATE") & Chr(&HD) & Chr(&HA)
Sys_Pos.Coptr1.PrintNormal PTR_S_RECEIPT, Time & Chr(&HD) & Chr(&HA)
Sys_Pos.Coptr1.PrintNormal PTR_S_RECEIPT, Space(2) & Chr(&HD) & Chr(&HA)
Sys_Pos.Coptr1.PrintNormal PTR_S_RECEIPT, "SOLICITUD DE ARTICULOS" & Chr(&HD) & Chr(&HA)
Sys_Pos.Coptr1.PrintNormal PTR_S_RECEIPT, "Mesero : " & rsCocina!nombre & Chr(&HD) & Chr(&HA)
Sys_Pos.Coptr1.PrintNormal PTR_S_RECEIPT, "Mesa # : " & rsCocina!mesa & Chr(&HD) & Chr(&HA)
If rs00!mesa_barra = nMesa Then Sys_Pos.ImprCocina.PrintNormal PTR_S_RECEIPT, "P A R A   L L E V A R" & Chr(&HD) & Chr(&HA)
Sys_Pos.Coptr1.PrintNormal PTR_S_RECEIPT, "---------------------------------" & Chr(&HD) & Chr(&HA)
Do Until rsCocina.EOF
    If Mid(LTrim(rsCocina!DESCRIP), 1, 2) = "@@" Then
        Sys_Pos.Coptr1.PrintNormal PTR_S_RECEIPT, Space(3) & Mid(rsCocina!DESCRIP, 1, 26) & Chr(&HD) & Chr(&HA)
    Else
        Sys_Pos.Coptr1.PrintNormal PTR_S_RECEIPT, Format(rsCocina!CANT, "##") & Space(2) & Mid(rsCocina!DESCRIP, 1, 26) & Chr(&HD) & Chr(&HA)
    End If
    rsCocina.MoveNext
Loop
Sys_Pos.Coptr1.PrintNormal PTR_S_RECEIPT, Space(2) & Chr(&HD) & Chr(&HA)
Sys_Pos.Coptr1.PrintNormal PTR_S_RECEIPT, Space(2) & Chr(&HD) & Chr(&HA)
rsCocina.MovePrevious
Sys_Pos.Coptr1.PrintNormal PTR_S_RECEIPT, Space(8) & "FIN PEDIDO MESA #: " & rsCocina!mesa & Chr(&HD) & Chr(&HA)
Sys_Pos.Coptr1.PrintNormal PTR_S_RECEIPT, Space(2) & Chr(&HD) & Chr(&HA)
Sys_Pos.Coptr1.PrintNormal PTR_S_RECEIPT, Space(2) & Chr(&HD) & Chr(&HA)
Sys_Pos.Coptr1.PrintNormal PTR_S_RECEIPT, Space(2) & Chr(&HD) & Chr(&HA)
Sys_Pos.Coptr1.PrintNormal PTR_S_RECEIPT, Space(2) & Chr(&HD) & Chr(&HA)
Sys_Pos.Coptr1.PrintNormal PTR_S_RECEIPT, Space(2) & Chr(&HD) & Chr(&HA)
Sys_Pos.Coptr1.PrintNormal PTR_S_RECEIPT, Space(2) & Chr(&HD) & Chr(&HA)
'Sys_Pos.Coptr1.AsyncMode = False
For i = 1 To 10
    Sys_Pos.Coptr1.PrintNormal PTR_S_RECEIPT, Space(3) & Chr(&HD) & Chr(&HA)
Next
Sys_Pos.Coptr1.CutPaper (100)
rsCocina.Close

Set rsCocina = Nothing
'
End Sub



