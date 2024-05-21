VERSION 5.00
Object = "{60C58221-392A-11CF-BE57-0020AF9B16AD}#1.0#0"; "COPTR.OCX"
Object = "{060154C0-3722-11CF-BE57-0020AF9B16AD}#1.0#0"; "COCASH.OCX"
Begin VB.Form Sys_Pos 
   BackColor       =   &H00B39665&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Entrada Cajeros"
   ClientHeight    =   3765
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   2925
   ClipControls    =   0   'False
   Icon            =   "Sys_pos.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3765
   ScaleWidth      =   2925
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   2520
      Top             =   0
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Salir"
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
      Left            =   1680
      TabIndex        =   6
      Top             =   2760
      Width           =   975
   End
   Begin VB.TextBox Text2 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      IMEMode         =   3  'DISABLE
      Left            =   1800
      MaxLength       =   10
      PasswordChar    =   "*"
      TabIndex        =   1
      Top             =   2160
      Width           =   855
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  'Right Justify
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "0"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   6154
         SubFormatType   =   1
      EndProperty
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   1800
      MaxLength       =   4
      TabIndex        =   0
      Top             =   1800
      Width           =   855
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Aceptar"
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
      Left            =   240
      TabIndex        =   3
      Top             =   2760
      Width           =   975
   End
   Begin COPTRLib.Coptr ImpresoraBarra 
      Left            =   2280
      Top             =   3240
      _Version        =   65536
      _ExtentX        =   1085
      _ExtentY        =   1296
      _StockProps     =   0
   End
   Begin COPTRLib.Coptr ImprCocina 
      Left            =   1680
      Top             =   3240
      _Version        =   65536
      _ExtentX        =   1085
      _ExtentY        =   1296
      _StockProps     =   0
   End
   Begin COCASHLib.Cocash Cocash1 
      Left            =   840
      Top             =   3240
      _Version        =   65536
      _ExtentX        =   1296
      _ExtentY        =   1296
      _StockProps     =   0
   End
   Begin COPTRLib.Coptr Coptr1 
      Left            =   0
      Top             =   3240
      _Version        =   65536
      _ExtentX        =   1508
      _ExtentY        =   1296
      _StockProps     =   0
   End
   Begin VB.Image Image1 
      BorderStyle     =   1  'Fixed Single
      Height          =   1185
      Left            =   120
      Picture         =   "Sys_pos.frx":0442
      Stretch         =   -1  'True
      Top             =   40
      Width           =   2655
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H00FF0000&
      Height          =   255
      Index           =   2
      Left            =   0
      TabIndex        =   7
      Top             =   3480
      Width           =   3000
   End
   Begin VB.Label Label2 
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
      Height          =   255
      Index           =   1
      Left            =   120
      TabIndex        =   5
      Top             =   2280
      Width           =   1095
   End
   Begin VB.Label Label2 
      BackColor       =   &H00B39665&
      Caption         =   "Número de Cajero"
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
      Index           =   0
      Left            =   120
      TabIndex        =   4
      Top             =   1920
      Width           =   1575
   End
   Begin VB.Label Label1 
      BackColor       =   &H00B39665&
      Caption         =   "Introduzca su número de cajero y contraseña"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   495
      Left            =   120
      TabIndex        =   2
      Top             =   1320
      Width           =   2295
   End
End
Attribute VB_Name = "Sys_Pos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private rsUsr As Recordset
Private Sub VerificaCierre()
Open App.Path & "\INOUTLOG.TXT" For Input As #1
Do Until EOF(1)
   Line Input #1, a$
Loop
Close #1
If a$ = "OK" Then
    'CERRO BIEN
Else
    Open App.Path & "\VENTALOG.TXT" For Append As #1
        Print #1, "-- SOLO POS NO CERRO BIEN --" & Date & " " & Time
    Close #1
End If
Open App.Path & "\INOUTLOG.TXT" For Output As #1
    Print #1, "NUNCA BORRAR ESTE ARCHIVO"
Close #1
End Sub

Private Sub AbrirFile()
'VERIFICA SI ES NECESARIO BORRAR TRANS LOCAL
Dim FecHost As Variant
Dim FecLoc As Variant
Dim RSLOC01 As Recordset
Dim nUpdateFlag As Integer
Dim iInt As Integer

Set rsLoc00 = New Recordset
Set RSLOC01 = New Recordset
nUpdateFlag = 0 'CERO, NO HAY QUE ACTUALIZAR
iInt = 0

On Error GoTo ErrorAdm:
   ' Open DATA_PATH + "SOLOLINE.TXT" For Input As #1
   ' Close #1
'''Open App.Path & "\" & "OrigenDB.txt" For Input As #1
'''Do Until EOF(1)
'''    Line Input #1, a$
'''    If Left(a$, 1) = "*" Then
'''        DATA_PATH = Mid(a$, 3, Len(a$) - 2)
'''    Else
'''        cDataPath = a$
'''    End If
'''Loop
'''Close #1

DATA_PATH = GetFromINI("General", "DirectorioDatos", App.Path & "\soloini.ini")
cDataPath = GetFromINI("General", "ProveedorDatos", App.Path & "\soloini.ini")

cDataPath = cDataPath & ";Jet OLEDB:Database Password=master24"
ADMIN_LOG = WindowsDirectory
ADMIN_LOG = App.Path & "\ADMLOG.SOL"

On Error GoTo 0
ON_LINE = True
If ON_LINE = True Then
    
    '\\SOLO11\ACCESS\SOLO.mdb;"
    On Error GoTo ErrDBMSOpen:
    Sys_Pos.Caption = Sys_Pos.Caption + ".ON LINE"
    msConn.Open cDataPath
    'msConn.Open "Provider=Microsoft.Jet.OLEDB.4.0;" _
            + "Data Source=\\SOLO11\ACCESS\SOLO.mdb;" _
            + "Jet OLEDB:Database Password=master24"

    On Error GoTo 0
Else
    
    Sys_Pos.Caption = Sys_Pos.Caption + ".OFF LINE"
    On Error GoTo ErrDBMSOpen:
    msConn.Open "Provider=Microsoft.Jet.OLEDB.4.0;" _
            + "Data Source=C:\SYS_POS\LOCAL\SOLO.mdb;" _
            + "Persist Security Info=False"
    On Error GoTo 0
    'msConn.Open "Provider=Microsoft.Jet.OLEDB.3.51" _
        + ";Persist Security Info=False;Data Source=" _
        + app.path & "\LOCAL\SOLO.mdb"
    MsgBox "TRABAJANDO OFF-LINE (Fuera de Linea). Puede Continuar. Presione Enter", vbInformation, BoxTit
End If

RSLOC01.Open "SELECT * FROM OPCIONES", msConn, adOpenStatic, adLockOptimistic

If RSLOC01!CHECK_UP = "Null" Then
    TipoApplicacion = " ESTE PRODUCTO ES UNA DEMOSTRACION."
Else
    TipoApplicacion = ""
End If
Me.Caption = TipoApplicacion + Me.Caption
SLIP_OK = False
If Not RSLOC01.EOF Then
    SLIP_OK = RSLOC01!SLIP_PRINTER
    REPCAJAX_OK = RSLOC01!REPORTX_OK
    MAX_DESCUENTO = RSLOC01!MAX_DESC
    OPEN_PROPINA = RSLOC01!PANTA_PROP
    If IsNull(RSLOC01!PROP_DESCR) Then
        PROPINA_DESCRIP = ""
    Else
        PROPINA_DESCRIP = Trim(RSLOC01!PROP_DESCR)
    End If
    HABITACION_OK = RSLOC01!HABITACION
    RSLOC01.Close
End If
Exit Sub

ErrorAdm:
ON_LINE = False
Resume Next

ErrorCopiaON:
    ' La BD no se pudo copiar alguien lo esta usando en la oficina
    MsgBox "ON LINE ¡ ERROR AL COPIAR BASES DE DATOS ! POSIBLEMENTE " & _
           "LA ESTEN USANDO EN LA OFICINA. EL PROGRAMA TERMINARA AHORA.", vbCritical, BoxTit
    Unload Me
    End

ErrDBMSOpen:
'Error grave NO SE ABRE DBMS
Dim OBJERR As Error
MsgBox Err.Number & " - " & Err.Description
For Each OBJERR In msConn.Errors
     MsgBox OBJERR.Number & " <-> " & OBJERR.Description, vbCritical, "Error Grave. ANOTE EL NUMERO"
Next
Unload Me
End
End Sub
Private Sub Command1_Click()

If Len(Text1) < 1 Or Len(Text2) < 1 Then Exit Sub
If Not IsNumeric(Text1) Then Exit Sub

On Error GoTo ErrorADO:
Set rs = New Recordset
Set rsUsr = New Recordset

rs.Open "SELECT numero,nombre,apellido FROM cajeros WHERE numero = " & _
    Text1 & " and clave = " & "'" & Text2 & "'", msConn, adOpenForwardOnly, adLockReadOnly

rsUsr.Open "SELECT numero,nombre FROM USUARIOS WHERE numero = " & _
    Text1 & " and clave = " & "'" & Text2 & "'", msConn, adOpenForwardOnly, adLockReadOnly

Set rs00 = New Recordset
rs00.Open "SELECT * FROM ORGANIZACION ", msConn, adOpenForwardOnly, adLockReadOnly

If rsUsr.EOF Then   'SI NO ES ADMINISTRADOR BUSCA EN CAJEROS
    If rs.EOF Then
        MsgBox "Informacion es INCORRECTA, Intente de Nuevo", vbInformation, BoxTit
        Exit Sub
    End If
Else
    RptCajas.Show 1
    Text1 = "": Text2 = ""
    Text1.SetFocus
    Exit Sub
End If

npNumCaj = rs!numero
cNomCaj = rs!nombre
nDesc01 = rs00!desc_01
nDesc02 = rs00!desc_02
nMesaBarra = rs00!mesa_barra
Text1 = "": Text2 = ""
On Error GoTo 0

'----DESPIERTA LA IMPRESORA----'
On Error Resume Next
    Coptr1.PrintNormal PTR_S_RECEIPT, "Login:" & Now & Chr(&HD) & Chr(&HA)
    For i = 1 To 10
        Coptr1.PrintNormal PTR_S_RECEIPT, " " & Chr(&HD) & Chr(&HA)
    Next
    Coptr1.CutPaper 100
    If Err.Number = 482 Then
        MsgBox "POR FAVOR ENCIENDA LA IMPRESORA", vbInformation, Err.Description
        Err.Clear
    End If
On Error GoTo 0
'------------------------------'
Text1.SetFocus
PLU.Show
Exit Sub

ErrorADO:
  Dim ADOError As Error
  For Each ADOError In msConn.Errors
     sError = sError & ADOError.Number & " - " & ADOError.Description + vbCrLf
  Next ADOError
  MsgBox sError, vbCritical, "Error Grave. ANOTE EL NUMERO"
  Resume Next
End Sub

Private Sub Command2_Click()
Dim hwnd As Integer
Dim Mifrm As Form

Label2(2).ForeColor = &HFF&
Label2(2) = "EL SISTEMA SE ESTA CERRANDO... ESPERE UNOS SEGUNDOS"

nVeriSalida = 2
Open App.Path & "\INOUTLOG.TXT" For Output As #1
    Print #1, "OK"
Close #1

'****** CIERRE DE OBJETOS OLE POS ******
'***************************************
            DoEvents
            Coptr1.Release
            Coptr1.Close
            
            DoEvents
            Cocash1.Release
            Cocash1.Close
'***************************************
'***************************************

For Each Mifrm In Forms
       Mifrm.Hide          ' hide the form
       Unload Mifrm        ' deactivate the form
       Set frm = Nothing   ' remove from memory
Next
Set rs = Nothing
msConn.Close
Unload Me
End
End Sub

Private Sub Form_Load()
Dim rs As Recordset
Dim nHandle As Integer

'Verifica si App esta abierta, para solamente cargarla una vez
If App.PrevInstance Then ActivatePrevInstance

'Call VerificaFecha

'abrir coneccion
BoxTit = "MENSAJE DEL SISTEMA DE VENTAS"
Show

DoEvents
Label2(2) = "Verificando Impresora/Gaveta ...": Label2(2).Refresh

DoEvents
Set msConn = New Connection
GENERICO_NO_JOURNAL = Chr$(27) + Chr$(99) + Chr$(48) + Chr$(0)
SOLOFAST_CUT = Chr$(27) & Chr$(105) & vbFormFeed
'==================================
'========OLE POS FIN===============
'==================================
    DoEvents
    Call OpenImpresora
    Call OpenImpresoraCocina

    Label2(2) = "Iniciando Impresora(s)": Label2(2).Refresh
    DoEvents
    Call ClaimImpresora
    Call ClaimImpresoraCocina
    Coptr1.DeviceEnabled = True
    If Coptr1.JrnEmpty = True Or Coptr1.JrnNearEnd = True Then
        MsgBox "¡¡ A T E N C I O N !!" & vbCrLf & vbCrLf & _
            "SE ESTA ACABANDO EL PAPEL DE AUDITORIA, SE LE RECOMIENDA SALIR DEL PROGRAMA PARA CAMBIAR EL PAPEL Y LUEGO REGRESAR AL PROGRAMA DE VENTAS", vbCritical, "¡¡ A T E N C I O N !!"
    End If
    
    ImprCocina.DeviceEnabled = True
    If ImprCocina.RecEmpty = True Or ImprCocina.RecNearEnd = True Then
        MsgBox "¡¡ A T E N C I O N   IMPRESORA DE COCINA  !!" & vbCrLf & vbCrLf & _
            "SE ESTA ACABANDO EL PAPEL DE LA IMPRESORA DEL COCINA, SE LE RECOMIENDA SALIR DEL PROGRAMA PARA CAMBIAR EL PAPEL Y LUEGO REGRESAR AL PROGRAMA DE VENTAS", vbCritical, "¡¡ A T E N C I O N !!"
    End If
    
    Cocash1.Close
    
    DoEvents
    Call OpenCajaRegistradora

    DoEvents
    Cocash1.DeviceEnabled = True
'==================================
'========OLE POS FIN===============
'==================================

'>>>>>>>>>> DEFAULT_PRINTER = GetCurrPrinter()
VerificaCierre
ON_LINE = True
nVeriSalida = 1

msConn.Mode = adModeShareDenyNone
Label2(2) = App.CompanyName
AbrirFile   'Verifica Conección
Call GetISC 'ITBMS

End Sub
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
If nVeriSalida = 1 Then
    MsgBox "¡ Favor utilize el boton Salir !", vbExclamation, BoxTit
    Cancel = True
End If
End Sub

Private Sub Image1_DblClick()
MsgBox "Empresa : " & App.CompanyName & Chr(13) & _
       "Derechos Reservados : " & App.LegalCopyright & Chr(13) & _
       "Nombre  : " & App.EXEName & Chr(13) & _
       "Versión : " & App.Major & "." & App.Minor & "." & App.Revision, vbInformation, "Informacion de la Aplicación"
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    If Text1 = "" Then
        Text1.SetFocus
    ElseIf Not IsNumeric(Text1) Then
        Text1.SetFocus
    Else
        Text2.SetFocus
    End If
End If
End Sub
Private Sub Text2_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    If Text2 = "" Then
        Text2.SetFocus
    ElseIf Not IsNumeric(Text2) Then
        Text2.SetFocus
    Else
        Command1.SetFocus
    End If
End If
End Sub

Private Sub Timer1_Timer()
Dim RetVal
If IsFormLoaded(PLU) = True Then
    If PLU.Hora.Caption <> CStr(Time) Then
        ' It's now a different second than the one displayed.
        PLU.Hora.Caption = Time
    End If
End If

If ON_LINE = True Then
    
    Open DATA_PATH + "SOLOLINE.TXT" For Input As #1
    Do Until EOF(1)
        Line Input #1, a$
    Loop
    Close #1
    
    If a$ = "OFF_LINE" Then
        ''''''''''''ES NECESARIO SALIR DEL PROGRAMA UN MOMENTO
        ''''''''''''YA QUE HASTA AHORA HABIAMOS TRABAJADO ON_LINE
        ''''''''''''EL PROGRAMA MSGUSER BORRA DB-LOCAL
        ''''''''''''BorraLocal
        nVeriSalida = 2
        RetVal = Shell(App.Path & "\MsgUser.exe", vbNormalFocus)
        Unload Me
        End
    End If
Else
    'NO SE PUEDE VERIFICAR DE ESTA MANERA, PONE AL SISTEMA MUY LENTO
End If

End Sub

Private Sub VerificaFecha()
Dim cMaxFecha As Date
Dim cMaxDia As String
Dim cMaxMes As String
Dim cMaxYear As String
Dim cLocalFecha As String

cMaxFecha = Date
cMaxMes = Mid(Format(cMaxFecha, "short date"), 4, 2)
cMaxDia = Mid(Format(cMaxFecha, "short date"), 1, 2)
cMaxYear = Mid(Format(cMaxFecha, "short date"), 7, 4)

cLocalFecha = cMaxYear & cMaxMes & cMaxDia
If Val(cLocalFecha) > Val("20010430") Then
    MsgBox "***** SU PERIODO DE EVALUACION A TERMINADO *****" & vbCrLf & _
            "- GRACIAS POR PROBAR PRODUCTOS DE SOLO SOFTWARE DEVELOPMENT" & vbCrLf & _
            "- CONTACTE A SU PROVEEDOR" & vbCrLf & _
            "" & vbCrLf & "El programa terminara AHORA", vbCritical, "CONTACTE A SU PROVEEDOR"
    Unload Me
    End
End If
End Sub

