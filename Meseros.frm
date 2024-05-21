VERSION 5.00
Begin VB.Form Meseros 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "SELECCION DE MESERO"
   ClientHeight    =   7995
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8460
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   9.75
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "Meseros.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   7995
   ScaleWidth      =   8460
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command1 
      Caption         =   "Aceptar Selección"
      Height          =   615
      Index           =   0
      Left            =   840
      TabIndex        =   1
      Top             =   7200
      Width           =   2415
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Regresar"
      Height          =   615
      Index           =   1
      Left            =   5040
      TabIndex        =   2
      Top             =   7200
      Width           =   2415
   End
   Begin VB.ListBox Meseros 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   6780
      IntegralHeight  =   0   'False
      ItemData        =   "Meseros.frx":000C
      Left            =   240
      List            =   "Meseros.frx":000E
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   240
      Width           =   7935
   End
End
Attribute VB_Name = "Meseros"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click(Index As Integer)
Dim iPos As Integer
If Index = 1 Then
    nMesero = 0
    Unload Me
End If
For iPos = 1 To 6
    If Mid(Meseros.List(Meseros.ListIndex), iPos, 1) = " " Then
        Exit For
    End If
Next
nMesero = Val(Mid(Meseros.List(Meseros.ListIndex), 1, iPos))
On Error Resume Next
    rs05.MoveFirst
On Error GoTo 0
rs05.Find "numero = " & nMesero
If Not rs05.EOF Then
    cNomMesero = rs05!nombre
    PLU.Text1(1) = cNomMesero
End If
Unload Me
End Sub
Private Sub Form_Load()
rs05.MoveFirst
Do Until rs05.EOF
    Meseros.AddItem Format(rs05!numero) + "   " + rs05!nombre + " " + Mid(rs05!apellido, 1, 1) & "."
    rs05.MoveNext
Loop
End Sub
