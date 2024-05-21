VERSION 5.00
Begin VB.Form CloseMesa 
   BackColor       =   &H00B39665&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "ESCRIBA EL NUMERO DE MESA"
   ClientHeight    =   2775
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4020
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2775
   ScaleWidth      =   4020
   StartUpPosition =   1  'CenterOwner
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
      Height          =   630
      Left            =   2640
      TabIndex        =   14
      Top             =   2010
      Width           =   1260
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
      Height          =   630
      Left            =   120
      TabIndex        =   13
      Top             =   2010
      Width           =   1260
   End
   Begin VB.TextBox txtLin 
      Alignment       =   1  'Right Justify
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   465
      IMEMode         =   3  'DISABLE
      Left            =   2640
      TabIndex        =   0
      Top             =   180
      Width           =   1125
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00B39665&
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
      Left            =   120
      TabIndex        =   2
      Top             =   570
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
         TabIndex        =   12
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
      Height          =   615
      Left            =   1680
      Picture         =   "CloseMesa.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   2010
      Width           =   615
   End
   Begin VB.Label lblLabels 
      BackColor       =   &H00B39665&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   120
      TabIndex        =   15
      Top             =   120
      Width           =   2400
   End
End
Attribute VB_Name = "CloseMesa"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private nlPase As Integer

Private Sub cmdOK_Click()
Dim vResp

vResp = MsgBox("¿ DESEA CERRAR LA MESA ?", vbQuestion + vbYesNo, "CONFIRME CIERRE DE MESA")
If vResp = vbYes Then
    Dim cSQL As String
    Dim rsTempMesa As ADODB.Recordset
    Dim cTempMesero As String
    Dim cTempCajero As String
    Dim nSubTot As Single
    Dim cCadena As String
    
    cSQL = "SELECT * FROM TMP_TRANS WHERE MESA = " & Val(txtLin)
    cSQL = cSQL & " ORDER BY LIN"
    Set rsTempMesa = New ADODB.Recordset
    rsTempMesa.Open cSQL, msConn, adOpenStatic, adLockOptimistic
    
    If rsTempMesa.EOF Then
        MsgBox "LA MESA ESTA VACIA." & vbCrLf & _
            "Si desea desbloquear una mesa, vaya a Administración de SOLO ADMIN y " & vbCrLf & _
            "bajo el menú de Ventas, seleccione Desbloquear Mesas", vbOKOnly, "NO SE PUEDE CERRAR MESA"
        rsTempMesa.Close
        Set rsTempMesa = Nothing
    Else
        rsTempMesa.MoveFirst
        Sys_Pos.Coptr1.PrintTwoNormal FPTR_S_JOURNAL_RECEIPT, Space(2), Space(2)
        Sys_Pos.Coptr1.PrintTwoNormal FPTR_S_JOURNAL_RECEIPT, "CIERRE DE MESA", "CIERRE DE MESA"
        Sys_Pos.Coptr1.PrintTwoNormal FPTR_S_JOURNAL_RECEIPT, "MESA : " & txtLin, "MESA : " & txtLin
        Sys_Pos.Coptr1.PrintTwoNormal FPTR_S_JOURNAL_RECEIPT, Time & Space(2) & Date, Time & Space(2) & Date
        Sys_Pos.Coptr1.PrintTwoNormal FPTR_S_JOURNAL_RECEIPT, Space(2), Space(2)
        cTempMesero = GetStaffName("M", rsTempMesa!MESERO)
        cTempCajero = GetStaffName("C", rsTempMesa!CAJERO)
        Sys_Pos.Coptr1.PrintTwoNormal FPTR_S_JOURNAL_RECEIPT, "Mesero : " & cTempMesero, "Mesero : " & cTempMesero
        Sys_Pos.Coptr1.PrintTwoNormal FPTR_S_JOURNAL_RECEIPT, "Cajero : " & cTempCajero, "Cajero : " & cTempCajero
        Sys_Pos.Coptr1.PrintTwoNormal FPTR_S_JOURNAL_RECEIPT, "------------------------------", "------------------------------"
        Do While Not rsTempMesa.EOF
            nSubTot = nSubTot + rsTempMesa!precio
            cCadena = LlenaConSpacio(Left(rsTempMesa!DESCRIP, 15), 15) & Space(2) & LlenaConSpacio(Format(rsTempMesa!CANT, "###"), 3) & Space(2) & LlenaConSpacio(Format(rsTempMesa!precio, "#,##0.00"), 8)
            Sys_Pos.Coptr1.PrintTwoNormal FPTR_S_JOURNAL_RECEIPT, cCadena, cCadena
            rsTempMesa.MoveNext
        Loop
        Sys_Pos.Coptr1.PrintTwoNormal FPTR_S_JOURNAL_RECEIPT, "------------------------------", "------------------------------"
        Sys_Pos.Coptr1.PrintTwoNormal FPTR_S_JOURNAL_RECEIPT, "TOTAL : " & Format(nSubTot, "CURRENCY"), "TOTAL : " & Format(nSubTot, "CURRENCY")
        rsTempMesa.Close
        Set rsTempMesa = Nothing

        For i = 1 To 10
            Sys_Pos.Coptr1.PrintNormal PTR_S_RECEIPT, Space(2) & Chr(&HD) & Chr(&HA)
        Next
        Sys_Pos.Coptr1.CutPaper 100

        msConn.BeginTrans
        cSQL = "DELETE * FROM TMP_TRANS WHERE MESA = " & Val(txtLin)
        msConn.Execute cSQL

        cSQL = "DELETE * FROM TMP_CUENTAS WHERE MESA = " & Val(txtLin)
        msConn.Execute cSQL

        cSQL = "UPDATE MESAS SET OCUPADA = 0, MESERO_ACTUAL = 0, LOCK = 0 WHERE NUMERO = " & Val(txtLin)
        msConn.Execute cSQL
        msConn.CommitTrans
        Mesas.Command1(Val(txtLin) - 1).BackColor = &HC0FFC0
    End If

End If
Unload Me
End Sub

Private Sub Command8_Click(Index As Integer)
Dim cCant As String

If nlPase = 0 Then
    txtLin = Command8(Index).Index
Else
    cCant = Str(txtLin)
    cCant = cCant & Command8(Index).Index
    txtLin = cCant
End If
nlPase = nlPase + 1
End Sub

Private Sub Form_Load()
nlPase = 0
lblLabels = txtInfo
End Sub
Private Sub Borrar_Click()
nlPase = 0
txtLin = ""
End Sub
Private Function LlenaConSpacio(cTexto As String, nLargo As Integer) As String
If Len(Trim(cTexto)) = nLargo Then
    'NADA
    LlenaConSpacio = cTexto
Else
    LlenaConSpacio = cTexto & Space(nLargo - Len(Trim(cTexto)))
End If
End Function
Private Function GetStaffName(cTipo As String, nNumber As Long) As String
Dim cSQL As String
Dim rsStaff As ADODB.Recordset

Select Case cTipo
    Case "M" 'MESERO
        cSQL = "SELECT NOMBRE + space(1) + APELLIDO AS STAFF FROM MESEROS"
        cSQL = cSQL & " WHERE NUMERO = " & nNumber
    Case "C" 'CAJERO/SUPERVISOR
        cSQL = "SELECT NOMBRE + space(1) + APELLIDO AS STAFF FROM CAJEROS"
        cSQL = cSQL & " WHERE NUMERO = " & nNumber
    Case Else
        GetStaffName = "No Definido"
        Exit Function
End Select

Set rsStaff = New ADODB.Recordset
rsStaff.Open cSQL, msConn, adOpenStatic, adLockOptimistic
GetStaffName = rsStaff!STAFF
rsStaff.Close
Set rsStaff = Nothing
End Function
Private Sub cmdCancel_Click()
OkAnul = 0
Unload Me
End Sub
