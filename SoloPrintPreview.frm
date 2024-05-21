VERSION 5.00
Object = "{A45D986F-3AAF-4A3B-A003-A6C53E8715A2}#1.0#0"; "ARVIEW2.OCX"
Begin VB.Form SoloPrintPreview 
   ClientHeight    =   5985
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7905
   LinkTopic       =   "Form2"
   ScaleHeight     =   5985
   ScaleWidth      =   7905
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin DDActiveReportsViewer2Ctl.ARViewer2 ARViewer21 
      Height          =   1335
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   2655
      _ExtentX        =   4683
      _ExtentY        =   2355
      SectionData     =   "SoloPrintPreview.frx":0000
   End
End
Attribute VB_Name = "SoloPrintPreview"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private mActiveControl As Control
Public Function SoloPrintPreviewFunction(SGSolo As SGGrid, cEncabezado As String) As Boolean
Set mActiveControl = ARViewer21
'mActiveControl.Top = SGSolo.Top
      
Set SGSolo.PrintSettings.Viewer = ARViewer21
SGSolo.PrintSettings.PrintGrid

With SGSolo.PrintSettings
    .RepeatColumnHeaders = True
End With

Me.Show
'Me.Caption = cEncabezado
End Function

Private Sub Form_Resize()
   On Error Resume Next
   
   With mActiveControl
      .Move 0, .Top, Me.ScaleWidth, Me.ScaleHeight - .Top
   End With
   
End Sub
