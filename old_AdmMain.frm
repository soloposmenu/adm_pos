VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form MainMant 
   Caption         =   "SISTEMA DE ADMINISTRACION."
   ClientHeight    =   8310
   ClientLeft      =   60
   ClientTop       =   630
   ClientWidth     =   11880
   ClipControls    =   0   'False
   Icon            =   "AdmMain.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   8310
   ScaleWidth      =   11880
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin VB.CommandButton Command5 
      Caption         =   "Command5"
      Height          =   135
      Left            =   11520
      TabIndex        =   12
      Top             =   120
      Width           =   255
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFCajas 
      Height          =   4215
      Left            =   9360
      TabIndex        =   10
      Top             =   3120
      Width           =   2415
      _ExtentX        =   4260
      _ExtentY        =   7435
      _Version        =   393216
      FixedCols       =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
      _Band(0).GridLinesBand=   1
      _Band(0).TextStyleBand=   0
      _Band(0).TextStyleHeader=   0
   End
   Begin VB.CommandButton Command4 
      Caption         =   "&Salir del Sistema"
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
      Left            =   9960
      TabIndex        =   9
      Top             =   7920
      Width           =   1815
   End
   Begin VB.PictureBox Picture3 
      BackColor       =   &H00808080&
      BorderStyle     =   0  'None
      Height          =   8655
      Left            =   6000
      ScaleHeight     =   8655
      ScaleWidth      =   3255
      TabIndex        =   6
      Top             =   0
      Width           =   3255
      Begin VB.CommandButton Command3 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   0
         Left            =   120
         TabIndex        =   8
         Top             =   240
         Width           =   1455
      End
      Begin VB.Label Label3 
         BackColor       =   &H00808080&
         Caption         =   "Informes y Listados"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   600
         TabIndex        =   7
         Top             =   0
         Width           =   2175
      End
   End
   Begin VB.PictureBox Picture2 
      BackColor       =   &H00808080&
      BorderStyle     =   0  'None
      Height          =   8655
      Left            =   3480
      ScaleHeight     =   8655
      ScaleWidth      =   2295
      TabIndex        =   3
      Top             =   0
      Width           =   2295
      Begin VB.CommandButton Command2 
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
         TabIndex        =   5
         Top             =   240
         Width           =   2055
      End
      Begin VB.Label Label2 
         BackColor       =   &H00808080&
         Caption         =   "Consultas"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   600
         TabIndex        =   4
         Top             =   0
         Width           =   1095
      End
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00808080&
      BorderStyle     =   0  'None
      Height          =   8655
      Index           =   0
      Left            =   0
      ScaleHeight     =   8655
      ScaleWidth      =   3255
      TabIndex        =   1
      Top             =   0
      Width           =   3255
      Begin VB.CommandButton Command1 
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
         TabIndex        =   0
         Top             =   240
         Width           =   1455
      End
      Begin VB.Label Label1 
         BackColor       =   &H00808080&
         Caption         =   "Mantenimiento"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   0
         Left            =   720
         TabIndex        =   2
         Top             =   0
         Width           =   1575
      End
   End
   Begin VB.Label Label4 
      Caption         =   "Ventas Actuales por Cajero"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   255
      Left            =   9360
      TabIndex        =   11
      Top             =   2760
      Width           =   2415
   End
   Begin VB.Image Image1 
      Height          =   1875
      Left            =   9600
      Picture         =   "AdmMain.frx":030A
      Stretch         =   -1  'True
      ToolTipText     =   "Logo Empresa"
      Top             =   360
      Width           =   2055
   End
   Begin VB.Menu mnuEmpresa 
      Caption         =   "Empresa"
      Begin VB.Menu mnuDeta 
         Caption         =   "Detalles de la Empresa"
      End
      Begin VB.Menu mnuRaya 
         Caption         =   "-"
      End
      Begin VB.Menu mnuMesas 
         Caption         =   "Mesas del Local"
      End
      Begin VB.Menu mnuRaya0 
         Caption         =   "-"
      End
      Begin VB.Menu mnuSalir 
         Caption         =   "Salir"
      End
   End
   Begin VB.Menu mnuVentas 
      Caption         =   "Ventas"
      Begin VB.Menu mnuEnv 
         Caption         =   "Envases para Ventas"
      End
      Begin VB.Menu mnuDepVen 
         Caption         =   "Departamentos de Ventas"
      End
      Begin VB.Menu mnuAco 
         Caption         =   "Acompañantes"
      End
      Begin VB.Menu mnuGrp 
         Caption         =   "Grupos de Ventas"
      End
      Begin VB.Menu mnuPLU 
         Caption         =   "Productos de Venta"
      End
      Begin VB.Menu mnuRaya1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuLinkAco 
         Caption         =   "Enlace de Acompañantes"
      End
      Begin VB.Menu menuRaya2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuLinkInv 
         Caption         =   "Enlace con Inventario"
      End
      Begin VB.Menu mnuRaya3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuDepOpen 
         Caption         =   "Departamentos Abiertos"
      End
   End
   Begin VB.Menu mnuInv 
      Caption         =   "Inventario"
      Begin VB.Menu mnuUMed 
         Caption         =   "Unidades de Medida"
      End
      Begin VB.Menu mnuDepInv 
         Caption         =   "Departamentos"
      End
      Begin VB.Menu mnuInvent 
         Caption         =   "Inventario Maestro"
      End
   End
   Begin VB.Menu mnuUsers 
      Caption         =   "Usuarios"
      Begin VB.Menu mnuAES 
         Caption         =   "Aministradores, Encargados y Supervisores"
      End
      Begin VB.Menu mnuCaj 
         Caption         =   "Cajeros del Sistema"
      End
      Begin VB.Menu mnuMeseros 
         Caption         =   "Meseros / Saloneros del Sistema"
      End
   End
   Begin VB.Menu mnuCli 
      Caption         =   "Clientes"
      Begin VB.Menu mnuCliOrg 
         Caption         =   "Clientes de la Empresa"
      End
      Begin VB.Menu mnuAppND 
         Caption         =   "Aplicar Notas de Debito (ND)"
      End
      Begin VB.Menu mnuRaya4 
         Caption         =   "-"
      End
      Begin VB.Menu mnuRecPag 
         Caption         =   "Recepción de Pagos"
      End
   End
   Begin VB.Menu mnuAdmProv 
      Caption         =   "Proveedores"
      Begin VB.Menu mnuProvOrg 
         Caption         =   "Proveedores de la Empresa"
      End
      Begin VB.Menu mnuRecComp 
         Caption         =   "Recepción de Compras"
      End
      Begin VB.Menu mnuRaya5 
         Caption         =   "-"
      End
      Begin VB.Menu mnuPagComp 
         Caption         =   "Pago de Compras"
      End
   End
   Begin VB.Menu mnuCons 
      Caption         =   "Consultas"
   End
   Begin VB.Menu mnuList 
      Caption         =   "Listados"
   End
   Begin VB.Menu mnuInf 
      Caption         =   "Informes"
   End
   Begin VB.Menu mnuAyuda 
      Caption         =   "Ayuda"
   End
End
Attribute VB_Name = "MainMant"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private ArregloFormas(0 To 50) As String
Private rsMCajas As Recordset
Private Sub DisplayTotales()
Dim objTable As New ADOX.Table
Dim objCat As New ADOX.Catalog

objCat.ActiveConnection = msConn

For Each objTabla In objCat.Tables
    If objTabla.Type = "TABLE" Then
        If objTabla.Name = "LOLO" Then
            msConn.BeginTrans
            msConn.Execute "DROP TABLE LOLO"
            msConn.CommitTrans
        ElseIf objTabla.Name = "LOLO1" Then
            msConn.BeginTrans
            msConn.Execute "DROP TABLE LOLO1"
            msConn.CommitTrans
        ElseIf objTabla.Name = "LOLO2" Then
            msConn.BeginTrans
            msConn.Execute "DROP TABLE LOLO2"
            msConn.CommitTrans
        ElseIf objTabla.Name = "LOLO3" Then
            msConn.BeginTrans
            msConn.Execute "DROP TABLE LOLO3"
            msConn.CommitTrans
        End If
    End If
Next

msConn.BeginTrans
msConn.Execute "SELECT cajero,sum(precio) as Valor " & _
            " INTO LOLO FROM TRANSAC " & _
            " GROUP BY CAJERO"
msConn.CommitTrans

rsMCajas.Open "SELECT b.Apellido," & _
            " format(a.valor,'standard') as Ventas " & _
            " FROM LOLO as a LEFT JOIN CAJEROS AS B " & _
            " ON A.CAJERO = B.NUMERO ", msConn, adOpenStatic, adLockOptimistic

Set MSHFCajas.DataSource = rsMCajas
SetUpPantalla
rsMCajas.Close

msConn.Execute "DROP TABLE LOLO"

End Sub
Private Sub SetUpPantalla()
With MSHFCajas
    .ColWidth(0) = 1100: .ColWidth(1) = 1000
    .ColAlignmentFixed(1) = flexAlignRightCenter
    .ColAlignment(1) = flexAlignRightCenter
End With
End Sub
Private Sub Command1_Click(Index As Integer)
'MsgBox Str(Index)
'INDEX ES UNO MENOS QUE EL VALOR EN SYS_01
If Index = 0 Then
    AdmDep.Show 1
ElseIf Index = 1 Then
    AdmEnv.Show 1
ElseIf Index = 2 Then
    AdmPlu.Show 1
ElseIf Index = 4 Then
    AdmPLU_Invent.Show 1
ElseIf Index = 5 Then
    AdmCajero.Show 1
ElseIf Index = 6 Then
    AdmMesero.Show 1
ElseIf Index = 7 Then
    AdmMesas.Show 1
ElseIf Index = 8 Then
    AdmDepInv.Show 1
ElseIf Index = 9 Then
    AdmUnid.Show 1
ElseIf Index = 10 Then
    AdmInv.Show 1
ElseIf Index = 11 Then
    AdmOrg.Show 1
ElseIf Index = 12 Then
    AdmUsers.Show 1
ElseIf Index = 13 Then
    MenuOrd.Show 1
ElseIf Index = 14 Then
    AdmGrp.Show 1
ElseIf Index = 15 Then
    AdmCli.Show 1
ElseIf Index = 16 Then
    AdmPagCli.Show 1
ElseIf Index = 17 Then
    AdmApND.Show 1
ElseIf Index = 18 Then
    AdmProv.Show 1
ElseIf Index = 19 Then
    AdmCompras.Show 1
ElseIf Index = 20 Then
    AdmDepOp.Show 1
ElseIf Index = 21 Then
    AdmAcom.Show 1
ElseIf Index = 22 Then
    AdmAcoPlu.Show 1
ElseIf Index = 23 Then
    AdmUnidCons.Show 1
ElseIf Index = 24 Then
    AdmPagProv.Show 1
ElseIf Index = 25 Then
    AjusteInv.Show 1
End If
'Dim lllevel As Long
'lllevel = msConn.BeginTrans
'msConn.CommitTrans
'MsgBox lllevel & "--->" & msConn.IsolationLevel
DisplayTotales
End Sub

Private Sub Command2_Click(Index As Integer)
If Index = 0 Then
    ConDept.Show 1
ElseIf Index = 1 Then
    conPlu.Show 1
ElseIf Index = 2 Then
    conTrans.Show 1
ElseIf Index = 3 Then
    ConCaj.Show 1
ElseIf Index = 4 Then
    ConMes.Show 1
ElseIf Index = 5 Then
    conInv.Show 1
ElseIf Index = 6 Then
    conCosto.Show 1
End If
DisplayTotales
End Sub

Private Sub Command3_Click(Index As Integer)
If Index = 0 Then
    Call PreparaImpresion(0)
ElseIf Index = 1 Then
    Call PreparaImpresion(1)
ElseIf Index = 2 Then
    Call PreparaImpresion(2)
ElseIf Index = 3 Then
    Call PreparaImpresion(3)
ElseIf Index = 5 Then
ElseIf Index = 6 Then
ElseIf Index = 7 Then
ElseIf Index = 13 Then
ElseIf Index = 14 Then
ElseIf Index = 15 Then
ElseIf Index = 16 Then
ElseIf Index = 17 Then
ElseIf Index = 18 Then
End If
DisplayTotales
End Sub

Private Sub Command4_Click()
EscribeLog ("Salida del Sistema de Administracion")
Unload Me
End
End Sub

Private Sub Command5_Click()
Dim RetVal
'Dim cVal As String
'cVal = InputBox("Enter your name")
'MsgBox "You entered: " & cVal
If npNumCaj = 1967 Then
    RetVal = Shell("NOTEPAD.EXE " & ADMIN_LOG, vbMaximizedFocus)
Else
    MsgBox "NO DISPONE DE AUTORIZACION, CONTACTE A SOLO SOFTWARE", vbCritical, BoxTit
End If
End Sub

Private Sub Form_Load()
Dim nIndice As Integer, ProxT As Integer, ProxL As Integer
Dim nIndice1 As Integer
Dim nErr As Integer

Set rs00 = New Recordset
Set rp01 = New Recordset
Set rp02 = New Recordset
Set rp03 = New Recordset
Set rsMCajas = New Recordset

nErr = 0
nIndice = 0
ProxT = 0
ProxL = 0
On Error GoTo ErrOpen:
rp01.Open "SELECT numero,descrip,vbform from Sys_01 order by numero", msConn, adOpenStatic, adLockReadOnly
rp02.Open "SELECT numero,descrip,vbform from Sys_02 order by numero", msConn, adOpenStatic, adLockReadOnly
rp03.Open "SELECT numero,descrip,vbform from Sys_03 order by numero", msConn, adOpenStatic, adLockReadOnly
rs00.Open "SELECT DESCRIP FROM ORGANIZACION ", msConn, adOpenStatic, adLockReadOnly

If Not rs00.EOF Then MainMant.Caption = MainMant.Caption & rs00!DESCRIP

Do Until rp01.EOF
    If nIndice > 0 Then
        If IsNull(rp01!DESCRIP) Or rp01!DESCRIP = "" Then
        Else
            Load Command1(nIndice)
            Command1(nIndice).Top = 240 + ProxT
            Command1(nIndice).Left = 120 + ProxL
            Command1(nIndice).Visible = True
            Command1(nIndice).Caption = IIf(IsNull(rp01!DESCRIP), "", rp01!DESCRIP)
            Command1(nIndice).Tag = IIf(IsNull(rp01!vbform), " ", rp01!vbform)
            ArregloFormas(nIndice) = Command1(nIndice).Tag
        End If
    Else
        Command1(nIndice).Caption = IIf(IsNull(rp01!DESCRIP), "", rp01!DESCRIP)
        Command1(nIndice).Tag = IIf(IsNull(rp01!vbform), "", rp01!vbform)
        ArregloFormas(nIndice) = Command1(nIndice).Tag
    End If

    nIndice = nIndice + 1
    ProxT = ProxT + 600
    If nIndice = 14 Then
        ProxT = 0
        ProxL = 1560
    End If
    rp01.MoveNext
Loop

On Error GoTo 0

ProxT = 0
nIndice1 = 0
Do Until rp02.EOF
    If nIndice1 > 0 Then
       Load Command2(nIndice1)
       Command2(nIndice1).Top = 240 + ProxT
       Command2(nIndice1).Visible = True
    End If
    Command2(nIndice1).Caption = IIf(IsNull(rp02!DESCRIP), "No hay Descripción", rp02!DESCRIP)
    Command2(nIndice1).Tag = IIf(IsNull(rp02!vbform), "", rp02!vbform)
    nIndice1 = nIndice1 + 1
    ProxT = ProxT + 600
    If nIndice1 = 14 Then
        ProxT = 0
    End If
    rp02.MoveNext
Loop

nIndice = 0
ProxT = 0
ProxL = 0

Do Until rp03.EOF
    If nIndice > 0 Then
       Load Command3(nIndice)
       Command3(nIndice).Top = 240 + ProxT
       Command3(nIndice).Left = 120 + ProxL
       Command3(nIndice).Visible = True
    End If
    Command3(nIndice).Caption = IIf(IsNull(rp03!DESCRIP), "No hay Descripcion", rp03!DESCRIP)
    Command3(nIndice).Tag = IIf(IsNull(rp03!vbform), "", rp03!vbform)
    
    nIndice = nIndice + 1
    ProxT = ProxT + 600
    If nIndice = 14 Then
        ProxT = 0
        ProxL = 1560
    End If
    rp03.MoveNext
Loop

DisplayTotales

nIndice = 1
Exit Sub

ErrOpen:
    nErr = nErr + 1
    If nErr >= 2 Then
        MsgBox "IMPOSIBLE CONTINUAR. PROGRAMA TERMINARA AHORA", vbCritical, BoxTit
        Unload Me
        End
    Else
        MsgBox "Error Abriendo Tablas de la Base de Datos, tal vez pueda continuar", vbInformation, BoxTit
        Resume Next
    End If
End Sub
