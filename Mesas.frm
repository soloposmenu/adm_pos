VERSION 5.00
Begin VB.Form Mesas 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "SELECCION DE MESA"
   ClientHeight    =   8445
   ClientLeft      =   240
   ClientTop       =   135
   ClientWidth     =   11535
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   8445
   ScaleWidth      =   11535
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdCierreMesa 
      DownPicture     =   "Mesas.frx":0000
      Height          =   615
      Left            =   5160
      Picture         =   "Mesas.frx":0442
      Style           =   1  'Graphical
      TabIndex        =   9
      ToolTipText     =   "Apriete para Cambiar Mesa"
      Top             =   7750
      Width           =   735
   End
   Begin VB.CommandButton cmdOpenGaveta 
      Height          =   615
      Left            =   8640
      Picture         =   "Mesas.frx":0884
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   7750
      Width           =   855
   End
   Begin VB.Timer Timer1 
      Interval        =   5000
      Left            =   7920
      Top             =   7800
   End
   Begin VB.CommandButton Command3 
      DownPicture     =   "Mesas.frx":0CC6
      Height          =   615
      Left            =   6720
      Picture         =   "Mesas.frx":1108
      Style           =   1  'Graphical
      TabIndex        =   7
      ToolTipText     =   "Apriete para Cambiar Mesa"
      Top             =   7750
      Width           =   735
   End
   Begin VB.CommandButton Command2 
      Caption         =   "&Regresar"
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
      Left            =   9720
      TabIndex        =   6
      Top             =   7750
      Width           =   1575
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H008080FF&
      Height          =   495
      Index           =   1
      Left            =   2280
      TabIndex        =   3
      Top             =   7800
      Width           =   495
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00C0FFC0&
      Height          =   495
      Index           =   0
      Left            =   120
      TabIndex        =   2
      Top             =   7800
      Width           =   495
   End
   Begin VB.Frame Frame1 
      Caption         =   "Seleccione la Mesa que Desea"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   7455
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   11175
      Begin VB.CommandButton Command1 
         BackColor       =   &H00C0FFC0&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Index           =   0
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   240
         Width           =   1455
      End
   End
   Begin VB.Label Label1 
      Caption         =   "Ocupada"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   1
      Left            =   2880
      TabIndex        =   5
      Top             =   7920
      Width           =   1335
   End
   Begin VB.Label Label1 
      Caption         =   "Disponible"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   0
      Left            =   720
      TabIndex        =   4
      Top             =   7920
      Width           =   1455
   End
End
Attribute VB_Name = "Mesas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim nCambio As Integer
Private rsTimer As New ADODB.Recordset
Private Sub QuitaMesas()
Dim nNum As Integer, lNum As Integer

'Prepara todos los Botones Disponibles. Max. 70
nNum = rs01.RecordCount
For lNum = 1 To 70
    If Not IsObject(Command1(lNum)) Then
    Else
        Load Command1(lNum)
    End If
    Command1(1).Caption = ""
    Command1(lNum).Visible = False
Next
'Solo Deja el Primero Visible
Command1(0).Visible = True
End Sub

Private Sub cmdCierreMesa_Click()
txtInfo = "Escriba Clave CIERRE de Mesa"
AskClave.Show 1
If OkAnul = 1 Then
    txtInfo = "Escriba el número de Mesa"
    CloseMesa.Show 1
Else
    MsgBox "NO Tiene AUTORIZACION para CIERRE de Mesa", vbExclamation, BoxTit
End If
OkAnul = 0
End Sub

Private Sub cmdOpenGaveta_Click()
txtInfo = "Escriba Clave abrir GAVETA"
AskClave.Show 1
If OkAnul = 1 Then
    'abrir gaveta
    Sys_Pos.Cocash1.Claim 5000
    rc = Sys_Pos.Cocash1.OpenDrawer
Else
    MsgBox "NO Tiene AUTORIZACION para ABRIR LA GAVETA DE DINERO", vbExclamation, BoxTit
End If
End Sub

Private Sub Command1_Click(Index As Integer)
Dim RSLOCKMESA As New ADODB.Recordset

RSLOCKMESA.Open "SELECT LOCK FROM MESAS WHERE NUMERO = " & Command1(Index).Tag, msConn, adOpenStatic, adLockOptimistic
If RSLOCKMESA!LOCK = 1 Then
    MsgBox "LA MESA ESTA OCUPADA EN ESTE MOMENTO, INTENTE MAS TARDE", vbCritical, "MESA ESTA OCUPADA"
    RSLOCKMESA.Close
    Set RSLOCKMESA = Nothing
    Exit Sub
End If
RSLOCKMESA.Close
Set RSLOCKMESA = Nothing

nMesa = Command1(Index).Tag
StatMesa nMesa, 1
If nMesa = rs00!mesa_barra Then nFlag = 1 Else nFlag = 0
Unload Me
End Sub

Private Sub Command2_Click()
nMesa = 0
Unload Me
End Sub

Private Sub Command3_Click()
CambiarMesa.Show 1

Dim iTam As Integer, MiTop As Integer, MiLeft As Integer
Dim statyleft As Integer, numcajas As Integer

iTam = 0
MiTop = 240: StayLeft = 120
MiLeft = 0: numcajas = 0

rs01.Requery
Do Until rs01.EOF
    If numcajas < 1 Then
        Command1(numcajas).Caption = rs01!numero
        Command1(numcajas).Tag = rs01!numero
        'Muestra los PLUs del primer departamento
    Else
        Command1(numcajas).Visible = True
        Command1(numcajas).Caption = rs01!numero
        Command1(numcajas).Tag = rs01!numero
        Command1(numcajas).Left = MiLeft + StayLeft
        Command1(numcajas).Top = MiTop
        StayLeft = 120
    End If
    If rs01!Status = "Libre" Then
       Command1(numcajas).BackColor = &HC0FFC0
    Else
       Command1(numcajas).BackColor = &H8080FF
    End If
    numcajas = numcajas + 1
    MiLeft = MiLeft + 1560
    If numcajas = 7 Or numcajas = 14 Or numcajas = 21 Or numcajas = 28 Or numcajas = 35 Or numcajas = 42 Or numcajas = 49 Or numcajas = 56 Or numcajas = 63 Then
        MiTop = MiTop + 720
        MiLeft = 0
    End If
    If numcajas = 84 Then Exit Do
    rs01.MoveNext
Loop
End Sub

Private Sub Form_Load()
Dim iTam As Integer, MiTop As Integer, MiLeft As Integer
Dim statyleft As Integer, numcajas As Integer

'Selecciona Todas las Mesas y las Marca
Set rs01 = New Recordset

rs01.Open "SELECT numero, iif(ocupada=TRUE,'Ocupada','Libre') AS status " & _
    " FROM mesas WHERE numero > 0 " & _
    " ORDER BY NUMERO", msConn, adOpenStatic, adLockReadOnly

QuitaMesas
iTam = 0

MiTop = 240: StayLeft = 120
MiLeft = 0: numcajas = 0

'Activa los Botones segun el numero de Mesas del Sistema
'y pone el Numero de Mesa en el Valor de TAG
Do Until rs01.EOF
    If numcajas < 1 Then
        Command1(numcajas).Caption = rs01!numero
        Command1(numcajas).Tag = rs01!numero
        'Muestra los PLUs del primer departamento
    Else
        If Not IsObject(Command1(numcajas)) Then
           Load Command1(numcajas)
        End If
        Command1(numcajas).Visible = True
        Command1(numcajas).Caption = rs01!numero
        Command1(numcajas).Tag = rs01!numero
        Command1(numcajas).Left = MiLeft + StayLeft
        Command1(numcajas).Top = MiTop
        StayLeft = 120
    End If
    If rs01!Status = "Libre" Then
        If rs00!mesa_barra = Val(Command1(numcajas).Caption) Then
            Command1(numcajas).BackColor = &HFFC0FF
        Else
            Command1(numcajas).BackColor = &HC0FFC0
        End If
    Else
       Command1(numcajas).BackColor = &H8080FF
    End If
    numcajas = numcajas + 1
    MiLeft = MiLeft + 1560
    If numcajas = 7 Or numcajas = 14 Or numcajas = 21 Or numcajas = 28 Or numcajas = 35 Or numcajas = 42 Or numcajas = 49 Or numcajas = 56 Or numcajas = 63 Then
        MiTop = MiTop + 720
        MiLeft = 0
    End If
    If numcajas = 84 Then Exit Do
    rs01.MoveNext
Loop
nCambio = 0
End Sub

Private Sub Timer1_Timer()
Dim i As Integer
Dim lBusca As Boolean
lBusca = False
rsTimer.Open "SELECT NUMERO FROM MESAS WHERE LOCK = 1", msConn, adOpenDynamic, adLockReadOnly
If rsTimer.EOF Then rsTimer.Close: Exit Sub
For i = 0 To rs01.RecordCount
    If Command1(i).BackColor = &HC0FFC0 Then
        'SI ESTA LIBRE, BUSCO SI HA SIDO OCUPADA
        On Error Resume Next
            rsTimer.MoveFirst
        On Error GoTo 0
        rsTimer.Find "NUMERO = " & Val(Command1(i).Tag)
        If Not rsTimer.EOF Then
            lBusca = True
            Exit For
            'SI ENCUENTRA MESA ES TIEMPO DE HACER UNA BUSQUEDA
        End If
    End If
Next
rsTimer.Close
If lBusca = False Then Exit Sub
For i = 1 To Command1.Count - 1
    Unload Command1(i)
Next
Form_Load
End Sub
