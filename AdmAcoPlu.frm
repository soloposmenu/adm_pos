VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form AdmAcoPlu 
   BackColor       =   &H00B39665&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "ENLACE DE PRODUCTOS DE VENTA - ACOMPAÑANTES"
   ClientHeight    =   8580
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11865
   Icon            =   "AdmAcoPlu.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8580
   ScaleWidth      =   11865
   ShowInTaskbar   =   0   'False
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFPluAco 
      Height          =   2775
      Left            =   6600
      TabIndex        =   6
      ToolTipText     =   "Doble Click para quitar"
      Top             =   5760
      Width           =   3975
      _ExtentX        =   7011
      _ExtentY        =   4895
      _Version        =   393216
      Cols            =   7
      FixedRows       =   0
      FixedCols       =   0
      SelectionMode   =   1
      AllowUserResizing=   1
      _NumberOfBands  =   1
      _Band(0).Cols   =   7
   End
   Begin VB.CommandButton cmdSalir 
      Caption         =   "Sa&lir"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   10680
      TabIndex        =   5
      Top             =   360
      Width           =   1095
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFPlu 
      Height          =   8175
      Left            =   120
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   360
      Width           =   6375
      _ExtentX        =   11245
      _ExtentY        =   14420
      _Version        =   393216
      BackColor       =   16777215
      ForeColor       =   0
      Cols            =   7
      FixedRows       =   0
      FixedCols       =   0
      BackColorSel    =   8388608
      GridColor       =   0
      SelectionMode   =   1
      AllowUserResizing=   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _NumberOfBands  =   1
      _Band(0).Cols   =   7
      _Band(0).GridLinesBand=   1
      _Band(0).TextStyleBand=   0
      _Band(0).TextStyleHeader=   0
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFAco 
      Height          =   5055
      Left            =   6600
      TabIndex        =   1
      TabStop         =   0   'False
      ToolTipText     =   "Click para Agregar"
      Top             =   360
      Width           =   3975
      _ExtentX        =   7011
      _ExtentY        =   8916
      _Version        =   393216
      BackColor       =   16777215
      ForeColor       =   0
      Cols            =   7
      FixedRows       =   0
      FixedCols       =   0
      BackColorSel    =   8388608
      GridColor       =   0
      SelectionMode   =   1
      AllowUserResizing=   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _NumberOfBands  =   1
      _Band(0).Cols   =   7
      _Band(0).GridLinesBand=   1
      _Band(0).TextStyleBand=   0
      _Band(0).TextStyleHeader=   0
   End
   Begin VB.Label Label1 
      BackColor       =   &H00B39665&
      Caption         =   "Acompañantes del Producto de Venta"
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
      Height          =   255
      Index           =   2
      Left            =   6600
      TabIndex        =   4
      Top             =   5520
      Width           =   3375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00B39665&
      Caption         =   "Acompañantes Disponibles"
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
      Height          =   255
      Index           =   1
      Left            =   6600
      TabIndex        =   3
      Top             =   120
      Width           =   2775
   End
   Begin VB.Label Label1 
      BackColor       =   &H00B39665&
      Caption         =   "Productos de Venta"
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
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   2055
   End
End
Attribute VB_Name = "AdmAcoPlu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rsPlu As New ADODB.Recordset
Dim rsAco As New ADODB.Recordset
Dim rsPluAco As New ADODB.Recordset
Dim nPlu As Long
Private Sub SetUpPantalla()
With MSHFPlu
    .ColWidth(0) = 0: .ColWidth(1) = 0: .ColWidth(2) = 2000:: .ColWidth(3) = 4000: .ColWidth(4) = 0
End With
With MSHFAco
    .ColWidth(0) = 0: .ColWidth(1) = 3500
End With
With MSHFPluAco
    .ColWidth(0) = 0: .ColWidth(1) = 0: .ColWidth(2) = 3800
End With
End Sub

Private Sub cmdSalir_Click()
    Unload Me
End Sub

Private Sub Form_Load()
Dim cSQL As String

cSQL = "SELECT B.CODIGO,B.DEPTO,A.DESCRIP AS DEPART,B.DESCRIP, "
cSQL = cSQL & " UCASE(MID(A.DESCRIP,1,1)) AS LETRA "
cSQL = cSQL & " FROM DEPTO AS A, PLU AS B "
cSQL = cSQL & " WHERE A.CODIGO = B.DEPTO "
cSQL = cSQL & " AND B.DISPONIBLE "
cSQL = cSQL & " ORDER BY A.DESCRIP,B.DESCRIP"
    
rsPlu.Open cSQL, msConn, adOpenKeyset, adLockOptimistic

If rsPlu.EOF Then
    ShowMsg "ES NECESARIO CREAR PRODUCTOS DE VENTAS", vbYellow, vbRed
    rsPlu.Close
    Set rsPlu = Nothing
    Set rsAco = Nothing
    Exit Sub
Else
End If

SetUpPantalla

Set MSHFPlu.Recordset = rsPlu
rsAco.Open "SELECT CODIGO,DESCRIP FROM ACOMPA ORDER BY DESCRIP", msConn, adOpenStatic, adLockOptimistic
If rsAco.EOF Then
    ShowMsg "ES NECESARIO CREAR ACOMPAÑANTES ANTES DE PODER CREAR LOS ENLACES", vbYellow, vbRed
    rsPlu.Close
    rsAco.Close
    Set rsPlu = Nothing
    Set rsAco = Nothing
    'Unload Me
    'Exit Sub
Else
    Set MSHFAco.Recordset = rsAco
    MSHFPlu_EnterCell
    
    Call Seguridad
End If

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
        MSHFPluAco.Enabled = False
    Case "V"
        MSHFAco.Enabled = False
    Case "N"
        MSHFPlu.Enabled = False
        MSHFAco.Enabled = False
        MSHFPluAco.Enabled = False
End Select
End Function

Private Sub Form_Unload(Cancel As Integer)
On Error Resume Next
    If rsPlu.State = adStateOpen Then rsPlu.Close
    If rsAco.State = adStateOpen Then rsAco.Close
    If rsPluAco.State = adStateOpen Then rsPluAco.Close
    Set rsPlu = Nothing
    Set rsAco = Nothing
    Set rsPluAco = Nothing
    Set MSHFPlu.Recordset = Nothing
On Error GoTo 0
End Sub

Private Sub MSHFAco_Click()
If MSHFPluAco.Rows = MAX_ACOMP Then
    ShowMsg "TEMPORALMENTE NO ES POSIBLE TENER MAS DE " & MAX_ACOMP & " ACOMPAÑANTES POR PLATO", vbYellow, vbRed
    Exit Sub
End If

On Error Resume Next
    rsPluAco.MoveFirst
On Error GoTo 0

On Error GoTo ErrAdm:

rsPluAco.Find "PLU_ID = " & nPlu
If rsPluAco.EOF Then
    msConn.BeginTrans
    msConn.Execute "INSERT INTO PLU_ACOMP (PLU_ID,ACOMP_ID) VALUES (" & nPlu & "," & Val(MSHFAco.Text) & ")"
    msConn.CommitTrans
Else
    'ESTE NUEVO FIND ARRANCA DESDE EL ULTIMO REGISTRO
    'VALIDO ENCONTRADO
    rsPluAco.Find "ACOMP_ID = " & Val(MSHFAco.Text)
    If rsPluAco.EOF Then
        msConn.BeginTrans
        msConn.Execute "INSERT INTO PLU_ACOMP (PLU_ID,ACOMP_ID) VALUES (" & nPlu & "," & Val(MSHFAco.Text) & ")"
        msConn.CommitTrans
    Else
        MsgBox "ACOMPAÑANTE YA EXISTE PARA ESTE PRODUCTO", vbInformation, BoxTit
    End If
End If
On Error GoTo 0
MSHFPlu_EnterCell
Exit Sub

ErrAdm:
    If Err.Number = 3704 Then
        MsgBox "NO ES POSIBLE TRABAJAR CON ACOMPAÑANTES" & vbCrLf & _
            "SALGA DEL PROGRAMA y VUELVA A ENTRAR, PARA TRABAJAR NORMALMENTE", vbCritical, "Error # " & Err.Number & " - " & BoxTit
    Else
        MsgBox Err.Description, vbCritical, "Error # " & Err.Number & " - " & BoxTit
    End If
    Exit Sub
End Sub

Private Sub MSHFPlu_EnterCell()
Dim cSQL As String
nPlu = Val(MSHFPlu.Text)
If rsPluAco.State = adStateOpen Then rsPluAco.Close

cSQL = "SELECT A.*,B.DESCRIP "
cSQL = cSQL & " FROM PLU_ACOMP AS A, "
cSQL = cSQL & " ACOMPA AS B "
cSQL = cSQL & " WHERE A.PLU_ID = " & nPlu
cSQL = cSQL & " AND A.ACOMP_ID = B.CODIGO"

rsPluAco.Open cSQL, msConn, adOpenDynamic, adLockOptimistic
On Error Resume Next
    'Set MSHFPluAco.Recordset = rsPluAco
    MSHFPluAco.Clear
    Set MSHFPluAco.DataSource = Nothing
    Set MSHFPluAco.DataSource = rsPluAco
    MSHFPluAco.Refresh
On Error GoTo 0
End Sub

Private Sub MSHFPlu_KeyPress(KeyAscii As Integer)
Dim cCadena As String

On Error Resume Next
cCadena = UCase(Chr(KeyAscii))
rsPlu.MoveFirst
rsPlu.Find ("LETRA = '" & cCadena & "'")
If Not rsPlu.EOF Then
    MSHFPlu.TopRow = rsPlu.AbsolutePosition - 1
    MSHFPlu.Row = rsPlu.AbsolutePosition - 1
    MSHFPlu.Col = 0
    MSHFPlu.ColSel = MSHFPlu.Cols - 1
    MSHFPlu.RowSel = rsPlu.AbsolutePosition - 1
End If

On Error GoTo 0
End Sub

Private Sub MSHFPluAco_DblClick()
Dim nLocPLU As Long
Dim nLocACO As Long

On Error Resume Next
If MSHFPluAco.Rows = 0 Then Exit Sub

MSHFPluAco.Col = 0
nLocPLU = Val(MSHFPluAco.Text)
MSHFPluAco.Col = 1
nLocACO = Val(MSHFPluAco.Text)
MSHFPluAco.Col = 2
'BoxResp = MsgBox("¿ Desea Retirar " & UCase(MSHFPluAco.Text) & " de este Producto ?", vbQuestion + vbYesNo, BoxTit)
If ShowMsg("¿ Desea Retirar de este Producto al Acompañante ...?" & vbCrLf & vbCrLf & UCase(MSHFPluAco.Text), vbYellow, vbRed, vbYesNo) = vbYes Then BoxResp = vbYes Else BoxResp = vbNo
If BoxResp = vbYes Then
    msConn.BeginTrans
    msConn.Execute "DELETE * FROM PLU_ACOMP " & _
            " WHERE PLU_ID = " & nLocPLU & _
            " AND ACOMP_ID = " & nLocACO
    msConn.CommitTrans
End If
MSHFPluAco.Col = 0
On Error GoTo 0
MSHFPlu_EnterCell
End Sub
