VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form AdmCajero 
   BackColor       =   &H00B39665&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "CAJEROS DEL SISTEMA"
   ClientHeight    =   5820
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6765
   Icon            =   "AdmCajero.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   5820
   ScaleWidth      =   6765
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdImpresion 
      BackColor       =   &H0000FFFF&
      Height          =   495
      Left            =   5640
      Picture         =   "AdmCajero.frx":030A
      Style           =   1  'Graphical
      TabIndex        =   18
      ToolTipText     =   "Imprimir Departamentos"
      Top             =   2280
      Width           =   855
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00B39665&
      Caption         =   "Opciones Modificables"
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
      Height          =   1935
      Index           =   0
      Left            =   240
      TabIndex        =   11
      Top             =   2160
      Width           =   4695
      Begin VB.TextBox Text1 
         Enabled         =   0   'False
         Height          =   285
         IMEMode         =   3  'DISABLE
         Index           =   3
         Left            =   1440
         MaxLength       =   40
         PasswordChar    =   "*"
         TabIndex        =   4
         Top             =   1440
         Width           =   3015
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Enabled         =   0   'False
         Height          =   285
         Index           =   0
         Left            =   1440
         TabIndex        =   1
         Top             =   360
         Width           =   735
      End
      Begin VB.TextBox Text1 
         Enabled         =   0   'False
         Height          =   285
         Index           =   1
         Left            =   1440
         MaxLength       =   25
         TabIndex        =   2
         Top             =   720
         Width           =   1815
      End
      Begin VB.TextBox Text1 
         Enabled         =   0   'False
         Height          =   285
         Index           =   2
         Left            =   1440
         MaxLength       =   15
         TabIndex        =   3
         Top             =   1080
         Width           =   1815
      End
      Begin VB.Label Label1 
         BackColor       =   &H00B39665&
         Caption         =   "Número"
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
         Left            =   120
         TabIndex        =   17
         Top             =   480
         Width           =   1095
      End
      Begin VB.Label Label1 
         BackColor       =   &H00B39665&
         Caption         =   "Contraseña"
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
         TabIndex        =   15
         Top             =   1560
         Width           =   1215
      End
      Begin VB.Label Label1 
         BackColor       =   &H00B39665&
         Caption         =   "Nombre"
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
         Left            =   120
         TabIndex        =   13
         Top             =   840
         Width           =   1095
      End
      Begin VB.Label Label1 
         BackColor       =   &H00B39665&
         Caption         =   "Apellido"
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
         Index           =   3
         Left            =   120
         TabIndex        =   12
         Top             =   1200
         Width           =   1215
      End
   End
   Begin VB.CommandButton Salir 
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
      Left            =   5160
      TabIndex        =   10
      Top             =   5040
      Width           =   1335
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00B39665&
      Height          =   1695
      Left            =   240
      TabIndex        =   14
      Top             =   3960
      Width           =   4695
      Begin VB.CommandButton Command2 
         Caption         =   "&Modificar"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   0
         Left            =   120
         TabIndex        =   5
         Top             =   360
         Width           =   1300
      End
      Begin VB.CommandButton Command2 
         Caption         =   "&Nuevo"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   1
         Left            =   1680
         TabIndex        =   6
         Top             =   360
         Width           =   1300
      End
      Begin VB.CommandButton Command2 
         Caption         =   "&Eliminar"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   2
         Left            =   3240
         TabIndex        =   7
         Top             =   360
         Width           =   1300
      End
      Begin VB.CommandButton Command2 
         Caption         =   "&Salvar"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   3
         Left            =   840
         TabIndex        =   8
         Top             =   1080
         Width           =   1300
      End
      Begin VB.CommandButton Command2 
         Caption         =   "&Regresar"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   4
         Left            =   2640
         TabIndex        =   9
         Top             =   1080
         Width           =   1300
      End
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFCajeros 
      Height          =   1695
      Left            =   240
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   360
      Width           =   6255
      _ExtentX        =   11033
      _ExtentY        =   2990
      _Version        =   393216
      BackColor       =   16777215
      ForeColor       =   0
      Cols            =   7
      FixedRows       =   0
      FixedCols       =   0
      BackColorSel    =   8388608
      GridColor       =   0
      HighLight       =   2
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
      _Band(0)._NumMapCols=   7
      _Band(0)._MapCol(0)._Name=   "NUMERO"
      _Band(0)._MapCol(0)._RSIndex=   0
      _Band(0)._MapCol(0)._Alignment=   7
      _Band(0)._MapCol(1)._Name=   "CLAVE"
      _Band(0)._MapCol(1)._RSIndex=   1
      _Band(0)._MapCol(2)._Name=   "NOMBRE"
      _Band(0)._MapCol(2)._RSIndex=   2
      _Band(0)._MapCol(3)._Name=   "APELLIDO"
      _Band(0)._MapCol(3)._RSIndex=   3
      _Band(0)._MapCol(4)._Name=   "X_COUNT"
      _Band(0)._MapCol(4)._RSIndex=   4
      _Band(0)._MapCol(4)._Alignment=   7
      _Band(0)._MapCol(5)._Name=   "Z_COUNT"
      _Band(0)._MapCol(5)._RSIndex=   5
      _Band(0)._MapCol(5)._Alignment=   7
      _Band(0)._MapCol(6)._Name=   "VALOR"
      _Band(0)._MapCol(6)._RSIndex=   6
      _Band(0)._MapCol(6)._Alignment=   7
   End
   Begin VB.Label Label2 
      BackColor       =   &H00B39665&
      Caption         =   "Seleccione Cajero"
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
      Left            =   240
      TabIndex        =   16
      Top             =   120
      Width           =   1695
   End
End
Attribute VB_Name = "AdmCajero"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rsCaj As Recordset
Dim Valores(3) As String
Private Sub CambiaTablaCajeros(indice As Integer)
Dim sqltext(10) As String
Dim i As Integer, nRgistros As Integer

'0 Elimina
'1 Modifica o Agrega
Select Case indice
Case 0
    'Eliminacion
    'Solamente se borra si el cajero no tiene ventas
    sqltext(0) = "DELETE * FROM CAJEROS WHERE NUMERO = " & Text1(0) & " AND VALOR < 0.01 "

    On Error GoTo ErrorEnTrans:
        msConn.BeginTrans
        msConn.Execute sqltext(0), nRgistros
        msConn.CommitTrans
    On Error GoTo 0
    If nRgistros = 0 Then
        ShowMsg "¡ El Cajero no se puede Eliminar, ya que tiene Ventas !", vbYellow, vbRed
    Else
        rsCaj.Requery
        Set MSHFCajeros.DataSource = rsCaj
        MSHFCajeros_EnterCell
    End If
'-----------------------------------------------------
Case 1
'1 Modifica o Agrega, primero se busca en tabla cajeros

If Text1(1) = "" Or Text1(2) = "" Then
    ShowMsg "¡¡ NO HAY SUFICIENTE INFORMACION PARA GRABAR !!", vbRed, vbYellow
    MSHFCajeros_EnterCell
    MSHFCajeros.SetFocus
    Exit Sub
End If

If Text1(0) = Empty Then
    Text1(0) = 0
End If

On Error Resume Next
rsCaj.MoveFirst
On Error GoTo 0

rsCaj.Find ("numero = " & Text1(0))
If Not rsCaj.EOF Then
    'Modificacion
    sqltext(0) = "UPDATE CAJEROS SET " & _
        " NUMERO = " & Text1(0) & _
        " ,NOMBRE = '" & Text1(1) & "'" & _
        " ,APELLIDO = '" & Text1(2) & "'" & _
        " ,CLAVE = '" & IIf(Text1(3) = "", " ", Text1(3)) & "'" & _
        " WHERE NUMERO = " & Text1(0)
Else
    'Nuevo Cajero
    sqltext(0) = "INSERT INTO CAJEROS (NUMERO,NOMBRE,APELLIDO,CLAVE) " & _
        " VALUES (" & Text1(0) & ",'" & Text1(1) & "','" & _
        Text1(2) & "','" & IIf(Text1(3) = "", " ", Text1(3)) & "'" & ")"
End If
    
    On Error GoTo ErrorEnTrans:
        msConn.BeginTrans
        msConn.Execute sqltext(0), nRgistros
        msConn.CommitTrans
    On Error GoTo 0

    rsCaj.Requery
    Set MSHFCajeros.DataSource = rsCaj
    MSHFCajeros_EnterCell
End Select

MSHFCajeros.SetFocus
Exit Sub

ErrorEnTrans:
'something bad happened so rollback the transaction
  msConn.RollbackTrans
  Dim ADOError As Error
  For Each ADOError In msConn.Errors
     sError = sError & ADOError.Number & " - " & ADOError.Description _
            + vbCrLf
  Next ADOError
  EscribeLog ("Admin." & "ERROR (AdmCajero) : " & sError)
  MsgBox sError, vbCritical, BoxTit
  Exit Sub
End Sub
Private Sub SetUpPantalla()
With MSHFCajeros
    .ColWidth(0) = 800: .ColWidth(1) = 0:
    .ColWidth(2) = 1600: .ColWidth(3) = 1600:
    'INFO: ENERO2010
    'QUE NO SALGA LA INFO DEL CONTADOR NI DE LAS VENTAS
    .ColWidth(4) = 0: .ColWidth(5) = 0:
    .ColWidth(6) = 0:
End With
End Sub

Private Sub Command2_Click(Index As Integer)
Dim i As Integer

Select Case Index
'~~~~~~~~~~~~~~~~~~
Case 0 ' ~~~ MODIFICAR ~~
'~~~~~~~~~~~~~~~~~~
    Command2(0).Enabled = False
    For i = 1 To 3
        Text1(i).Enabled = True
        Command2(i).Enabled = False
        Valores(i) = Text1(i)
    Next
    For i = 3 To 4
        Command2(i).Enabled = True
    Next
    Text1(1).SetFocus
    Text1(1).SelLength = Len(Text1(1).Text)
    MSHFCajeros.BackColor = &HC0C0C0
    MSHFCajeros.Enabled = False
'~~~~~~~~~~~~~~~~~~
Case 1  ' ~~~ NUEVO ~~~
'~~~~~~~~~~~~~~~~~~
    For i = 0 To 3
        Text1(i).Enabled = True
        Valores(i) = Text1(i)
        Text1(i) = ""
        Command2(i).Enabled = False
    Next
    'Command2(i).Enabled = False
    For i = 3 To 4
        Command2(i).Enabled = True
    Next
    MSHFCajeros.BackColor = &HC0C0C0
    Text1(0).SetFocus
    Text1(0).SelLength = Len(Text1(0).Text)
'~~~~~~~~~~~~~~~~~~
Case 2 ' ~~~ ELIMINAR ~~~
'~~~~~~~~~~~~~~~~~~
    'BoxPreg = "¿ Desea Eliminar Cajero " & Text1(1) & " ?"
    'BoxResp = MsgBox(BoxPreg, vbQuestion + vbYesNo, BoxTit)
    If ShowMsg("¿ Desea Eliminar Cajero " & Text1(1) & " ?", vbYellow, vbRed, vbYesNo) = vbYes Then BoxResp = vbYes Else BoxResp = vbNo
    If BoxResp = vbYes Then
        CambiaTablaCajeros (0)
    End If
'~~~~~~~~~~~~~~~~~~
Case 3 ' ~~~ SALVAR ~~~
'~~~~~~~~~~~~~~~~~~
    'INFO: 8AGO2017. REVISANDO TRACK READER Y NUMERO DE MESERO.
    If Text1(0) = "999" Or Text1(3) = "9999" Then
        ShowMsg "NO PUEDE USAR ESE NUMERO DE CAJERO", vbYellow, vbRed
        Exit Sub
    End If
    'BoxPreg = "¿ Desea Salvar los Datos en Pantalla ?"
    'BoxResp = MsgBox(BoxPreg, vbQuestion + vbYesNo, BoxTit)
    If ShowMsg("¿ Desea Salvar los Datos en Pantalla ?", vbYellow, vbBlue, vbYesNo) = vbYes Then BoxResp = vbYes Else BoxResp = vbNo
    For i = 0 To 3
        Text1(i).Enabled = False
        Command2(i).Enabled = True
    Next
    For i = 3 To 4
        Command2(i).Enabled = False
    Next
    MSHFCajeros.Enabled = True
    If BoxResp = vbYes Then
        CambiaTablaCajeros (1)
    End If
    MSHFCajeros.BackColor = vbWhite
    Command2(2).Enabled = True
'~~~~~~~~~~~~~~~~~~~
Case 4 'REGRESAR SIN SALVAR
'~~~~~~~~~~~~~~~~~~~
    For i = 0 To 3
        Text1(i) = Valores(i)
        Text1(i).Enabled = False
        Command2(i).Enabled = True
    Next
    For i = 3 To 4
        Command2(i).Enabled = False
    Next
    MSHFCajeros.Enabled = True
    MSHFCajeros.BackColor = vbWhite
    Command2(2).Enabled = True
End Select

Call Seguridad

End Sub

Private Sub Form_Load()
Set rsCaj = New Recordset

rsCaj.Open "SELECT Numero,clave,Nombre,Apellido,x_count,z_count,Valor " & _
        " FROM CAJEROS WHERE NUMERO <> 999 ORDER BY numero", msConn, adOpenDynamic, adLockOptimistic
Set MSHFCajeros.DataSource = rsCaj
SetUpPantalla
MSHFCajeros_EnterCell

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
        Command2(2).Enabled = False
    Case "CV"
        Command2(0).Enabled = False: Command2(2).Enabled = False
    Case "V"
        Command2(0).Enabled = False: Command2(1).Enabled = False: Command2(2).Enabled = False
        Command2(3).Enabled = False: Command2(4).Enabled = False
    Case "N"
        MSHFCajeros.Enabled = False
        Command2(0).Enabled = False: Command2(1).Enabled = False: Command2(2).Enabled = False
        Command2(3).Enabled = False: Command2(4).Enabled = False
End Select
End Function

Private Sub MSHFCajeros_Click()
MSHFCajeros_EnterCell
End Sub

Private Sub MSHFCajeros_EnterCell()
Dim i As Integer
Dim nC As Integer

For i = 0 To 3
    Text1(i).Enabled = True
Next

nC = Val((MSHFCajeros.Text))
On Error Resume Next
rsCaj.MoveFirst 'ANTES DE FIND, SIEMPRE HAY QUE MANDAR
                'EL CURSOR AL INICIO DE LA TABLA
rsCaj.Find "NUMERO = " & nC
On Error GoTo 0

If Not rsCaj.EOF Then
    Text1(0) = nC
    Text1(1) = IIf(IsNull(rsCaj!NOMBRE), "", rsCaj!NOMBRE)
    Text1(2) = IIf(IsNull(rsCaj!APELLIDO), "", rsCaj!APELLIDO)
    Text1(3) = IIf(IsNull(rsCaj!CLAVE), "", rsCaj!CLAVE)
End If
    
For i = 0 To 3
    Text1(i).Enabled = False
Next
End Sub

Private Sub Salir_Click()
Unload Me
End Sub

Private Sub Text1_KeyPress(Index As Integer, KeyAscii As Integer)
Dim iCount As Integer
If KeyAscii = 13 Then
    Select Case Index
    Case 0 To 2
        If Index = 0 Then
            If Not IsNumeric(Text1(0)) Then Exit Sub
            On Error Resume Next
                rsCaj.MoveFirst
            On Error GoTo 0
            On Error GoTo ErrAdm:
            rsCaj.Find "NUMERO = " & Text1(0)
            If rsCaj.EOF Then
            Else
                MsgBox "Ya Existe Cajero con ese Número", vbExclamation, BoxTit
                Text1(Index).SetFocus
                Text1(Index).SelLength = Len(Text1(Index).Text)
                On Error GoTo 0
                Exit Sub
            End If
        End If
        Text1(Index + 1).SetFocus
        Text1(Index + 1).SelLength = Len(Text1(Index + 1).Text)
    Case 3
        If Text1(3) = "" Then
            MsgBox "¡¡ TIENE QUE ESCRIBIR UNA CONTRASEÑA PARA EL CAJERO !!", vbInformation, BoxTit
            Text1(3).SetFocus
            Text1(3).SelLength = Len(Text1(3).Text)
            On Error GoTo 0
            Exit Sub
        End If
        Command2(3).SetFocus
    End Select
End If
On Error GoTo 0
Exit Sub

ErrAdm:
iCount = iCount + 1
MsgBox Err.Description, vbCritical, "SE A DETECTADO UN DATO INVALIDO"
If iCount < 3 Then Text1(Index).SetFocus Else Exit Sub
End Sub
Private Sub cmdImpresion_Click()
Dim rsTemp As New ADODB.Recordset
Dim cSQL As String
Dim iLin  As Integer
Dim nAcum As Single

nAcum = 0
On Error Resume Next

cSQL = "SELECT A.NUMERO,A.NOMBRE,A.APELLIDO, SUM(B.PRECIO) AS TRANSACCIONES "
cSQL = cSQL & " FROM CAJEROS AS A LEFT JOIN HIST_TR AS B"
cSQL = cSQL & " ON A.NUMERO = B.CAJERO "
cSQL = cSQL & " WHERE A.NUMERO <> 999 "
cSQL = cSQL & " GROUP BY A.NUMERO,A.NOMBRE,A.APELLIDO"
cSQL = cSQL & " ORDER BY A.NOMBRE,A.APELLIDO"

Me.MousePointer = vbHourglass
rsTemp.Open cSQL, msConn, adOpenStatic, adLockOptimistic
If rsTemp.EOF Then
    'MsgBox "NO HAY CAJEROS DEFINIDOS", vbInformation
    ShowMsg "NO HAY CAJEROS DEFINIDOS", vbYellow, vbRed
    rsTemp.Close
    Set rsTemp = Nothing
    Me.MousePointer = vbDefault
    Exit Sub
End If

MainMant.spDoc.DocBegin
MainMant.spDoc.WindowTitle = "Impresión de CAJEROS"
MainMant.spDoc.FirstPage = 1
MainMant.spDoc.PageOrientation = SPOR_PORTRAIT
MainMant.spDoc.Units = SPUN_LOMETRIC

MainMant.spDoc.Page = 1

MainMant.spDoc.TextOut 300, 200, Format(Date, "long date") & "  " & Time
MainMant.spDoc.TextOut 300, 300, rs00!DESCRIP
MainMant.spDoc.TextOut 300, 400, "IMPRESION DE CAJEROS"

MainMant.spDoc.TextOut 300, 500, "NUMERO"
MainMant.spDoc.TextOut 600, 500, "NOMBRE"
MainMant.spDoc.TextOut 1250, 500, "VENTAS"
MainMant.spDoc.TextOut 300, 550, "--------------------------------------------------------------------------------------------"
iLin = 600

Do While Not rsTemp.EOF
    MainMant.spDoc.TextAlign = SPTA_LEFT
    MainMant.spDoc.TextOut 300, iLin, rsTemp!NUMERO
    MainMant.spDoc.TextOut 600, iLin, rsTemp!NOMBRE & Space(1) & rsTemp!APELLIDO
    MainMant.spDoc.TextAlign = SPTA_RIGHT
    MainMant.spDoc.TextOut 1400, iLin, Format(rsTemp!TRANSACCIONES, "###,###.00")
    iLin = iLin + 50
    rsTemp.MoveNext
Loop
MainMant.spDoc.TextAlign = SPTA_LEFT
Me.MousePointer = vbDefault
MainMant.spDoc.DoPrintPreview
rsTemp.Close
Set rsTemp = Nothing
Me.MousePointer = vbDefault
End Sub
