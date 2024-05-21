VERSION 5.00
Begin VB.Form ScreenReport 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "REPORTES POR PANTALLA"
   ClientHeight    =   6855
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11610
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6855
   ScaleWidth      =   11610
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Salir 
      Caption         =   "Sa&lir"
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
      Left            =   10200
      TabIndex        =   1
      Top             =   6240
      Width           =   1215
   End
   Begin VB.ListBox MSHFText 
      BeginProperty Font 
         Name            =   "Fixedsys"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5685
      Left            =   120
      TabIndex        =   0
      Top             =   360
      Width           =   11295
   End
End
Attribute VB_Name = "ScreenReport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
Dim MyString
Dim cDirecc As String

On Error Resume Next
cDirecc = App.Path

Open cDirecc & "SOLOFILE.TXT" For Input As #1
Do While Not EOF(1)   ' Loop until end of file.
   Input #1, MyString
   MSHFText.AddItem MyString
Loop
Close #1
On Error GoTo 0
End Sub

Private Sub Salir_Click()
Unload Me
End Sub
