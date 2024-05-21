VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form FacCtaPlato 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "CREACION / MODIFICACION / SELECCION DE CUENTAS"
   ClientHeight    =   7200
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11280
   ControlBox      =   0   'False
   Icon            =   "FacCtaPlato.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   7200
   ScaleWidth      =   11280
   StartUpPosition =   2  'CenterScreen
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid FlexPlatos 
      Height          =   6015
      Left            =   4200
      TabIndex        =   19
      Top             =   840
      Width           =   6855
      _ExtentX        =   12091
      _ExtentY        =   10610
      _Version        =   393216
      FixedCols       =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
   End
   Begin VB.CommandButton Command1 
      Cancel          =   -1  'True
      Caption         =   "&Incluir Cuentas a la Mesa"
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
      Left            =   240
      TabIndex        =   16
      Top             =   2040
      Width           =   1380
   End
   Begin VB.ListBox ListCtas 
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3210
      ItemData        =   "FacCtaPlato.frx":0442
      Left            =   240
      List            =   "FacCtaPlato.frx":0444
      TabIndex        =   15
      Top             =   3480
      Width           =   1815
   End
   Begin VB.CommandButton cmdCancel 
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
      Left            =   2520
      TabIndex        =   13
      Top             =   6120
      Width           =   1380
   End
   Begin VB.TextBox txtCuentas 
      Alignment       =   1  'Right Justify
      Height          =   345
      Left            =   3360
      TabIndex        =   12
      Top             =   240
      Width           =   645
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
      TabIndex        =   1
      Top             =   600
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
         Height          =   495
         Index           =   0
         Left            =   2940
         TabIndex        =   2
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
         Left            =   2220
         TabIndex        =   11
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
         Left            =   1500
         TabIndex        =   10
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
         Left            =   780
         TabIndex        =   9
         Top             =   720
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
         Left            =   60
         TabIndex        =   8
         Top             =   720
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
         Left            =   2940
         TabIndex        =   7
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
         Left            =   2220
         TabIndex        =   6
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
         Left            =   1500
         TabIndex        =   5
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
         Left            =   780
         TabIndex        =   4
         Top             =   180
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
         Left            =   60
         TabIndex        =   3
         Top             =   180
         Width           =   735
      End
   End
   Begin VB.CommandButton Borrar 
      Height          =   735
      Left            =   1800
      Picture         =   "FacCtaPlato.frx":0446
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   2040
      Width           =   615
   End
   Begin VB.Label lblLabels 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Cuentas Incluidas"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   1
      Left            =   240
      TabIndex        =   17
      Top             =   3120
      Width           =   1920
   End
   Begin VB.Label lblLabels 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Anote el Número de Cuentas"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Index           =   0
      Left            =   240
      TabIndex        =   14
      Top             =   360
      Width           =   3000
   End
   Begin VB.Label lblLabels 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Platos Asignados a Cada Cuenta"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   2
      Left            =   4200
      TabIndex        =   18
      Top             =   480
      Width           =   4200
   End
End
Attribute VB_Name = "FacCtaPlato"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim nlPase As Integer
Dim rsCuentas As Recordset
Private Sub Borrar_Click()
nlPase = 0
txtCuentas = ""
End Sub
Private Sub cmdCancel_Click()
Dim rsCuentas As New ADODB.Recordset
Dim cSQL As String

cSQL = "SELECT Cuenta FROM TMP_TRANS "
cSQL = cSQL & " WHERE MESA = " & nMesa
cSQL = cSQL & " ORDER BY CUENTA,LIN "

rsCuentas.Open cSQL, msConn, adOpenStatic, adLockOptimistic
If rsCuentas.EOF Then
    rsCuentas.Close
    Set rsCuentas = Nothing
Else
    rsCuentas.MoveFirst
    If rsCuentas!CUENTA = 0 Then
        MsgBox "DEBE ASIGNAR TODOS LOS PRODUCTOS A UNA CUENTA", vbCritical, BoxTit
        rsCuentas.Close
        Set rsCuentas = Nothing
        Exit Sub
    End If
    rsCuentas.Close
    Set rsCuentas = Nothing
End If
Unload Me
End Sub
Private Sub Command1_Click()
Dim j As Integer
On Error GoTo ErrAdm:

If txtCuentas = "" Then
    MsgBox "TIENE QUE ANOTAR EL NUMERO DE CUENTAS", vbExclamation, BoxTit
    Exit Sub
End If
If txtCuentas > 50 Then
    MsgBox "MAXIMO 50 CUENTAS POR MESA", vbExclamation, BoxTit
    Exit Sub
End If
If nMesa < 1 Then
    MsgBox "PRIMERO DEBE SELECCIONAR UNA MESA", vbInformation, BoxTit
    Unload Me
    Exit Sub
End If
ListCtas.Enabled = True
For j = 1 To txtCuentas
    msConn.Execute "INSERT INTO TMP_CUENTAS (MESA,CUENTA) " & _
            " VALUES (" & nMesa & "," & j & ")"
    ListCtas.AddItem j
Next
nCta = 1
Command1.Enabled = False
On Error GoTo 0
Exit Sub

ErrAdm:
    Resume Next
End Sub

Private Sub Command8_Click(Index As Integer)
Dim cCant As String

If nlPase = 0 Then
    txtCuentas = Command8(Index).Index
Else
    cCant = Str(txtCuentas)
    cCant = cCant & Command8(Index).Index
    txtCuentas = Val(cCant)
End If
nlPase = nlPase + 1
End Sub

Private Sub FlexPlatos_Click()

BoxPreg = "¿ ESTA CAMBIANDO LA CUENTA ASIGNADA A ESTE PRODUCTO, ESTA SEGURO ?"
BoxResp = MsgBox(BoxPreg, vbQuestion + vbYesNo, BoxTit)

If BoxResp <> vbYes Then
    Exit Sub
End If

Dim i As Integer
Dim MiArr(17) As String
Dim sqltxt As String
Dim txtTipo As String

For i = 1 To 17
    FlexPlatos.Col = i
    MiArr(i) = FlexPlatos.Text
    If i = 16 Then txtTipo = FlexPlatos.Text
Next

If Len(txtTipo) > 2 Then
    MsgBox "NO SE PUEDEN CAMBIAR DESCUENTOS, CORRECCIONES, NI ANULACIONES", vbInformation, BoxTit
    Exit Sub
End If

sqltxt = "UPDATE TMP_TRANS SET CUENTA = " & nCta & _
    " WHERE MESA = " & nMesa & " AND LIN = " & Val(MiArr(6))

msConn.BeginTrans
msConn.Execute sqltxt
msConn.CommitTrans
    
rsPlatos.Requery

Set FlexPlatos.DataSource = rsPlatos

With FlexPlatos
    .ColWidth(0) = 950: .ColWidth(7) = 3000: .ColWidth(13) = 1200:
    .ColWidth(6) = 500
    .ColAlignment(13) = flexAlignRightCenter
    .ColAlignment(0) = flexAlignCenterCenter
    For i = 1 To 5
        .ColWidth(i) = 0
    Next
    For i = 8 To 12
        .ColWidth(i) = 0
    Next
    For i = 14 To 17
        .ColWidth(i) = 0
    Next
    .ColAlignment(16) = flexAlignRightCenter
    .ColWidth(16) = 950:
End With
End Sub

Private Sub Form_Load()
Dim sqltxt As String
Dim cTabla As String

Set rsPlatos = New Recordset
Set rsCuentas = New Recordset

nlPase = 0

cTabla = "CtaMesa" & LTrim(Str(nMesa))

sqltxt = "SELECT Cuenta,CAJA,CAJERO,MESA,MESERO,VALID,LIN,DESCRIP, " & _
    " CANT,DEPTO,PLU,ENVASE,PRECIO_UNIT, " & _
    " FORMAT(precio,'##,##0.00') as Valor,FECHA,Hora,TIPO,DESCUENTO " & _
    " FROM TMP_TRANS WHERE MESA = " & nMesa & _
    " ORDER BY CUENTA,LIN "
rsPlatos.Open sqltxt, msConn, adOpenDynamic, adLockOptimistic

sqltxt = "SELECT MESA,CUENTA FROM TMP_CUENTAS " & _
    " WHERE MESA = " & nMesa & _
    " ORDER BY MESA,CUENTA"
rsCuentas.Open sqltxt, msConn, adOpenDynamic, adLockOptimistic

Set FlexPlatos.DataSource = rsPlatos

If Not rsCuentas.EOF Then
    nCta = rsCuentas!CUENTA
    Command1.Enabled = False
End If

Do Until rsCuentas.EOF
    ListCtas.AddItem rsCuentas!CUENTA
    rsCuentas.MoveNext
Loop

Dim i As Integer
With FlexPlatos
    .ColWidth(0) = 950: .ColWidth(7) = 3000: .ColWidth(13) = 1200:
    .ColWidth(6) = 500
    .ColAlignment(13) = flexAlignRightCenter
    .ColAlignment(0) = flexAlignCenterCenter
    For i = 1 To 5
        .ColWidth(i) = 0
    Next
    For i = 8 To 12
        .ColWidth(i) = 0
    Next
    For i = 14 To 17
        .ColWidth(i) = 0
    Next
    .ColAlignment(16) = flexAlignRightCenter
    .ColWidth(16) = 950:
End With

End Sub
Private Sub ListCtas_Click()
If ListCtas.Text = "" Then
    nCta = 0
Else
    nCta = Val(ListCtas.Text)
End If
End Sub
