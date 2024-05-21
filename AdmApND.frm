VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form AdmApND 
   BackColor       =   &H00B39665&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "APLICAR NOTAS DE DEBITO"
   ClientHeight    =   3615
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9510
   Icon            =   "AdmApND.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3615
   ScaleWidth      =   9510
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdSalir 
      Caption         =   "&Salir"
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
      Left            =   7920
      TabIndex        =   7
      Top             =   2760
      Width           =   1455
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Aplicar ND"
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
      Left            =   4320
      Picture         =   "AdmApND.frx":0442
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   1920
      Width           =   1215
   End
   Begin VB.CommandButton cmdBorra 
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
      Height          =   615
      Left            =   4320
      Picture         =   "AdmApND.frx":0884
      Style           =   1  'Graphical
      TabIndex        =   6
      ToolTipText     =   "Borra Toda la Información"
      Top             =   2760
      Width           =   1215
   End
   Begin VB.TextBox cComment 
      Height          =   1455
      Left            =   120
      MultiLine       =   -1  'True
      TabIndex        =   4
      Top             =   1920
      Width           =   3615
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
      Height          =   375
      Index           =   0
      Left            =   6360
      TabIndex        =   2
      Top             =   525
      Width           =   1335
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  'Right Justify
      Height          =   375
      Index           =   1
      Left            =   8040
      TabIndex        =   3
      Text            =   "0.00"
      Top             =   525
      Width           =   1335
   End
   Begin VB.ComboBox cmbCli 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   120
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   525
      Width           =   4575
   End
   Begin MSComCtl2.DTPicker txtFecIni 
      Height          =   345
      Left            =   4800
      TabIndex        =   1
      Top             =   525
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   609
      _Version        =   393216
      Format          =   63766529
      CurrentDate     =   36431
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      BackColor       =   &H0080FFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "0.00"
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
      Left            =   1680
      TabIndex        =   14
      Top             =   1000
      Width           =   1335
   End
   Begin VB.Label Label2 
      BackColor       =   &H00B39665&
      Caption         =   "Saldo del Cliente"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   3
      Left            =   120
      TabIndex        =   13
      Top             =   1200
      Width           =   1575
   End
   Begin VB.Label Label2 
      BackColor       =   &H00B39665&
      Caption         =   "Concepto/Comentario"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   2
      Left            =   120
      TabIndex        =   12
      Top             =   1680
      Width           =   2055
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00B39665&
      Caption         =   "Numero de Documento"
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
      Left            =   6360
      TabIndex        =   11
      Top             =   120
      Width           =   1335
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00B39665&
      Caption         =   "Monto de Documento"
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
      Left            =   8160
      TabIndex        =   10
      Top             =   120
      Width           =   1215
   End
   Begin VB.Label Label2 
      BackColor       =   &H00B39665&
      Caption         =   "Fecha"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   6
      Left            =   4800
      TabIndex        =   9
      Top             =   285
      Width           =   1335
   End
   Begin VB.Label Label2 
      BackColor       =   &H00B39665&
      Caption         =   "Seleccione Cliente"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   7
      Left            =   120
      TabIndex        =   8
      Top             =   285
      Width           =   1695
   End
End
Attribute VB_Name = "AdmApND"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private rsCli As Recordset
Private nCodTP As Integer
Private nCli As Long
Private rsNDs As New ADODB.Recordset
Private nPagina As Integer
Private iLin As Integer
Private Sub ClearScreen()
Text1(0) = "": Text1(1) = 0#: Label2(4) = "": cComment = ""
End Sub

Private Sub cmbCli_Click()
nCli = Int(Val(RTrim(Mid(cmbCli.Text, 1, 6))))
rsCli.MoveFirst
rsCli.Find "CODIGO = " & nCli
If Not rsCli.EOF Then
    Label2(4) = Format(rsCli!SALDO, "STANDARD")
Else
    Label2(4) = Format(0#, "STANDARD")
End If
End Sub

Private Sub cmbCli_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then txtFecIni.SetFocus
End Sub

Private Sub cmdBorra_Click()
ClearScreen
Call Seguridad
End Sub

Private Sub cmdSalir_Click()
Unload Me
End Sub
Private Sub PrintTit()
If nPagina = 0 Then
    MainMant.spDoc.WindowTitle = "Impresión de " & Me.Caption
    MainMant.spDoc.FirstPage = 1
    MainMant.spDoc.PageOrientation = SPOR_PORTRAIT
    MainMant.spDoc.Units = SPUN_LOMETRIC
End If
MainMant.spDoc.Page = nPagina + 1

MainMant.spDoc.TextOut 300, 200, Format(Date, "long date") & "  " & Time
MainMant.spDoc.TextOut 300, 250, "Página : " & nPagina + 1
MainMant.spDoc.TextOut 300, 300, Me.Caption
MainMant.spDoc.TextOut 300, 350, rs00!DESCRIP

MainMant.spDoc.TextOut 300, 450, "CLIENTE"
MainMant.spDoc.TextOut 950, 450, "FECHA"
MainMant.spDoc.TextOut 1100, 450, "NUM_DOC"
MainMant.spDoc.TextOut 1300, 450, "MONTO"
MainMant.spDoc.TextOut 1500, 450, "COMENTARIO"
MainMant.spDoc.TextOut 300, 500, "--------------------------------------------------------------------------------------------------------------------------------"

iLin = 550
nPagina = nPagina + 1
End Sub
Private Sub ImprimirDatos()

nPagina = 0
MainMant.spDoc.DocBegin
PrintTit    'Rutina de Titulos

rsNDs.MoveFirst
Do While Not rsNDs.EOF
    If iLin > 2400 Then PrintTit

    MainMant.spDoc.TextAlign = SPTA_LEFT
    MainMant.spDoc.TextOut 300, iLin, rsNDs!CLIENTE
    MainMant.spDoc.TextOut 900, iLin, rsNDs!FECHA
    MainMant.spDoc.TextAlign = SPTA_RIGHT
    MainMant.spDoc.TextOut 1260, iLin, rsNDs!NUM_DOC
    MainMant.spDoc.TextOut 1440, iLin, Format(rsNDs!MONTO, "#,###.00")
    MainMant.spDoc.TextAlign = SPTA_LEFT
    MainMant.spDoc.TextOut 1500, iLin, rsNDs!COMENTARIO
    iLin = iLin + 50
    rsNDs.MoveNext
Loop
MainMant.spDoc.DoPrintPreview
End Sub
Private Sub Command1_Click()
Dim DocFec As String
Dim cTxt As String

DocFec = Format(txtFecIni, "YYYYMMDD")
If Text1(0) = "" Or Val(Text1(0)) = 0 Or Text1(1) = "" Or Val(Text1(1)) = 0 Then
    MsgBox "FALTA INFORMACION DEL DOCUMENTO y/o MONTO", vbExclamation, BoxTit
    Exit Sub
End If
BoxResp = MsgBox("¿ Desea Aplicar esta Nota de Debito ?", vbQuestion + vbYesNo, BoxTit)
If BoxResp = vbYes Then
    
    On Error GoTo ErrAplica:
    msConn.BeginTrans

    cTxt = "INSERT INTO HIST_TR_CLI " & _
        "(CODIGO_TP,CODIGO_CLI,NUM_TRANS,FECHA,STATUS,MONTO,RECIBIDO,SALDO,TIPO_TRANS,COMMENT,USUARIO) " & _
        " VALUES (" & _
        nCodTP & "," & _
        nCli & "," & _
        Val(Text1(0)) & ",'" & _
        DocFec & "'," & _
        0 & "," & _
        Format(Text1(1), "######.00") & "," & _
        0# & "," & _
        Format(Text1(1), "######.00") & _
        ",'ND','" & _
        cComment & "'," & _
        npNumCaj & ")"

    With rsNDs
        .AddNew
        !CLIENTE = Left(cmbCli.Text, 30)
        !FECHA = Format(txtFecIni, "DD/MM/YYYY")
        !NUM_DOC = Text1(0)
        !MONTO = Text1(1)
        !COMENTARIO = Left(cComment, 40)
        .Update
    End With

    msConn.Execute cTxt
    msConn.Execute "UPDATE CLIENTES SET SALDO = SALDO + " & Format(Text1(1), "######.00") & _
            " WHERE CODIGO = " & nCli
    msConn.Execute "UPDATE ORGANIZACION SET CONT_ND = CONT_ND + 1 "
    msConn.CommitTrans
    
    EscribeLog ("Admin." & "Aplicación ND al Cliente " & cmbCli.Text & " por " & Text1(1))
    MsgBox "EL DOCUMENTO HA SIDO APLICADO CON EXITO", vbInformation, BoxTit
    ClearScreen
    cmbCli.SetFocus
End If

SalidaAplica:

On Error GoTo 0

Call Seguridad

Exit Sub

ErrAplica:
Dim OBJERR As ADODB.Error
For Each OBJERR In msConn.Errors
    MsgBox OBJERR.Description, vbCritical, BoxTit
Next
msConn.RollbackTrans
MsgBox "Ocurrio un ERROR en la ACTUALIZACION, Intente de nuevo", vbExclamation, BoxTit
GoTo SalidaAplica:

End Sub

Private Sub Form_Load()
Dim nLen As Integer

nCodTP = 12
txtFecIni = Format(Date, "SHORT DATE")
Set rsCli = New Recordset
rsCli.Open "SELECT NOMBRE,APELLIDO,EMPRESA,CODIGO,SALDO " & _
        " FROM CLIENTES order by 1,2,3", msConn, adOpenStatic, adLockOptimistic
Do Until rsCli.EOF
    nLen = Len(rsCli!CODIGO)
    cmbCli.AddItem rsCli!CODIGO & Space(7 - nLen) & rsCli!NOMBRE & "," & rsCli!APELLIDO & " <-> " & rsCli!EMPRESA
    rsCli.MoveNext
Loop
nCli = 0
With rsNDs
    .Fields.Append "CLIENTE", adChar, 30, adFldUpdatable
    .Fields.Append "FECHA", adChar, 10, adFldUpdatable
    .Fields.Append "NUM_DOC", adSingle, , adFldUpdatable
    .Fields.Append "MONTO", adSingle, , adFldUpdatable
    .Fields.Append "COMENTARIO", adChar, 50, adFldUpdatable
    .CursorType = adOpenKeyset
    .LockType = adLockOptimistic
    .Open
End With

Call Seguridad

End Sub
Private Function Seguridad() As String
'SETUP DE SEGURIDAD DEL SISTEMA
Dim cSeguridad As String

cSeguridad = GetSecuritySetting(npNumCaj, Me.Name)
Select Case cSeguridad
    Case "CEMV"
        'INFO: NO HAY RESTRICCIONES
    Case "CMV"
        'INFO: NO HAY RESTRICCIONES
    Case "CV"
        'INFO: NO HAY RESTRICCIONES
    Case "V"
        txtFecIni.Enabled = False
        Text1(0).Enabled = False: Text1(1).Enabled = False
        cComment.Enabled = False
        Command1.Enabled = False
        cmdBorra.Enabled = False
    Case "N"
        cmbCli.Enabled = False
        txtFecIni.Enabled = False
        Text1(0).Enabled = False: Text1(1).Enabled = False
        cComment.Enabled = False
        Command1.Enabled = False
        cmdBorra.Enabled = False
End Select
End Function

Private Sub Form_Unload(Cancel As Integer)
If Not rsNDs.EOF Then ImprimirDatos
On Error Resume Next
    If rsCli.State = adStateOpen Then rsCli.Close
    If rsNDs.State = adStateOpen Then rsNDs.Close
    Set rsCli = Nothing
    Set rsNDs = Nothing
On Error GoTo 0
End Sub

Private Sub Text1_KeyPress(Index As Integer, KeyAscii As Integer)
If KeyAscii = 13 Then
    If Index = 0 Then
        Text1(1).SetFocus
        Text1(1).SelLength = Len(Text1(1).Text)
    ElseIf Index = 1 Then
        If Not IsNumeric(Text1(1)) Then
            MsgBox "Debe ingresar un Monto en dinero", vbExclamation, BoxTit
            Text1(1) = 0#
            Text1(1).SetFocus
            Exit Sub
        Else
            If Text1(1) = 0# Then Beep: Text1(1).SetFocus: Exit Sub
        End If
        Text1(1) = Format(Text1(1), "standard")
        cComment.SetFocus
    End If
End If
End Sub

Private Sub txtFecIni_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then
    Text1(1).SetFocus
    Text1(1).SelStart = 0
    Text1(1).SelLength = Len(Text1(1).Text)
End If
End Sub

Private Sub txtFecIni_LostFocus()
Dim rsNum As Recordset
Dim TXT As String

Set rsNum = New Recordset
TXT = "SELECT CONT_ND + 1 AS ND_NUM" & _
        " FROM ORGANIZACION"
rsNum.Open TXT, msConn, adOpenStatic, adLockOptimistic
Text1(0).Enabled = True
Text1(0).SetFocus
Text1(0) = rsNum!ND_NUM
rsNum.Close
Text1(0).Enabled = False
Text1(1).SetFocus
Text1(1).SelLength = Len(Text1(1).Text)
End Sub
