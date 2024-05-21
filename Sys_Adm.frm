VERSION 5.00
Begin VB.Form Sys_Adm 
   BackColor       =   &H00B39665&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "ENTRADA AL SISTEMA DE ADMINISTRACION"
   ClientHeight    =   3030
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5730
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   FillColor       =   &H00404040&
   Icon            =   "Sys_Adm.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3030
   ScaleWidth      =   5730
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command2 
      Caption         =   "Cancelar"
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
      Left            =   2160
      TabIndex        =   6
      Top             =   2160
      Width           =   1335
   End
   Begin VB.TextBox Text2 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      IMEMode         =   3  'DISABLE
      Left            =   2520
      MaxLength       =   20
      PasswordChar    =   "*"
      TabIndex        =   1
      Top             =   1560
      Width           =   975
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  'Right Justify
      BeginProperty DataFormat 
         Type            =   0
         Format          =   "0"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   6154
         SubFormatType   =   0
      EndProperty
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Left            =   2520
      MaxLength       =   4
      TabIndex        =   0
      Top             =   1080
      Width           =   975
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
      Top             =   2160
      Width           =   1335
   End
   Begin VB.Image Image 
      Height          =   1800
      Left            =   3840
      Picture         =   "Sys_Adm.frx":0442
      Top             =   0
      Width           =   1800
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
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
      ForeColor       =   &H00C00000&
      Height          =   255
      Index           =   2
      Left            =   -360
      TabIndex        =   7
      Top             =   2760
      Width           =   6135
   End
   Begin VB.Label Label2 
      BackColor       =   &H00B39665&
      Caption         =   "Contraseña"
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
      Index           =   1
      Left            =   240
      TabIndex        =   5
      Top             =   1680
      Width           =   1335
   End
   Begin VB.Label Label2 
      BackColor       =   &H00B39665&
      Caption         =   "Número de Usuario"
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
      Left            =   240
      TabIndex        =   4
      Top             =   1200
      Width           =   2055
   End
   Begin VB.Label Label1 
      BackColor       =   &H00B39665&
      Caption         =   "Bienvenido, Introduzca su número de Usuario y Contraseña"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   615
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   3375
   End
End
Attribute VB_Name = "Sys_Adm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'=======================================
'INFO: REEMPLAZADO POR Environ("WINDIR")
'=======================================
''''Private Declare Function GetWindowsDirectory Lib "kernel32" Alias "GetWindowsDirectoryA" (ByVal lpBuffer As String, ByVal nSize As Long) As Long
Private Declare Function GetComputerName Lib "kernel32" Alias "GetComputerNameA" (ByVal lpBuffer As String, nSize As Long) As Long
Private LPARAM As Long
Dim cOldCaption As String

Private Sub Command1_Click()
Dim rsOpc As New ADODB.Recordset
Dim cSQL As String

Set rs = New Recordset

On Error GoTo ErrorADO:
cSQL = "Select numero, NOMBRE, APELLIDO, TIPO "
cSQL = cSQL & " FROM USUARIOS "
cSQL = cSQL & " WHERE numero = " & Text1
cSQL = cSQL & " AND clave = '" & Text2 & "'"
'rs.Open cSQL, msConn, adOpenForwardOnly, adLockReadOnly
rs.Open cSQL, msConn, adOpenStatic, adLockReadOnly

If rs.EOF Then
    ShowMsg "INFORMACION INCORRECTA, INTENTE DE NUEVO", vbRed, vbYellow
    If rsOpc.State = adStateOpen Then rsOpc.Close
    rs.Close
    'INFO: 13FEB2013
    Text1.SetFocus
    Text1.SelStart = 0
    Text1.SelLength = Len(Text1.Text)
    Exit Sub
End If
On Error GoTo 0

MousePointer = vbHourglass

rsOpc.Open "SELECT * FROM OPCIONES", msConn, adOpenStatic, adLockOptimistic
If Not rsOpc.EOF Then HABITACION = rsOpc!HABITACION
rsOpc.Close
Set rsOpc = Nothing

npNumCaj = rs!NUMERO
NOM_ADMINISTRADOR = rs!NOMBRE & " " & rs!APELLIDO
TIPO_ADMINISTRADOR = rs!TIPO

EscribeLog ("Admin." & "Entrada Sistema Administración (" & App.Major & "." & App.Minor & "." & App.Revision & ") - " & NOM_ADMINISTRADOR)

DoEvents
'///////////// Call Seleccion_Impresora
DoEvents
Unload Me
MainMant.Show
Exit Sub

ErrorADO:

If Err.Number = -2147217900 Then
    ShowMsg "DATOS INVALIDOS, " & vbCrLf & "FAVOR VERIFICAR", vbYellow, vbRed
Else
  Dim ADOError As Error
  For Each ADOError In msConn.Errors
     sError = sError & ADOError.Number & " - " & ADOError.Description + vbCrLf
  Next ADOError
  ShowMsg "Error Grave. ANOTE EL NUMERO" & vbCrLf & sError, vbYellow, vbRed
End If
End Sub

Private Sub Command2_Click()
Dim MiFrm As Form

On Error Resume Next
Label2(2).ForeColor = &HFF&
Label2(2) = "EL SISTEMA SE ESTA CERRANDO... ESPERE UNOS SEGUNDOS"

For Each MiFrm In Forms
    DoEvents
    MiFrm.Hide            ' hide the form
    Unload MiFrm          ' deactivate the form
    Set MiFrm = Nothing   ' remove from memory
Next

msConn.Close
Set rs = Nothing
Set msConn = Nothing
On Error GoTo 0
Unload Me
End
End Sub

Private Sub Form_Load()
Dim cTempMachineName  As String
Dim cVersionActualRegistry As String
Dim cArrayVersionActualRegistry() As String

If App.PrevInstance Then ActivatePrevInstance

'Call VerificaFecha
''Dim x As Printer
''   For Each x In Printers
''      'Debug.Print X.DeviceName
''   Next

Show
Me.MousePointer = vbHourglass
Label2(2) = "ABRIENDO BASE DE DATOS. Favor Esperar ...": Label2(2).Refresh
DoEvents
Label2(2) = "DoEvents...": Label2(2).Refresh
Set msConn = New Connection
Label2(2) = "Set msConn = New Connection": Label2(2).Refresh
'///////////DEFAULT_PRINTER = "SOLOPRN" 'GetCurrPrinter()

'''''''cTempMachineName = Space$(24)
'''''''Label2(2) = "Before GetComputerName...": Label2(2).Refresh
'''''''LPARAM = GetComputerName(cTempMachineName, Len(cTempMachineName))
'''''''Label2(2) = "After GetComputerName...": Label2(2).Refresh
'''''''cTempMachineName = UCase(Trim(cTempMachineName))
cTempMachineName = "CCAJA"

DoEvents
On Error GoTo ErrorOpen:

'''ADMIN_LOG = WindowsDirectory
Label2(2) = "Before Environ...": Label2(2).Refresh
ADMIN_LOG = Environ("WINDIR")
ADMIN_LOG = ADMIN_LOG & "\ADMLOG.SOL"
BoxTit = "Sistema de Administración"

Label2(2) = "Before INI...": Label2(2).Refresh
DATA_PATH = GetFromINI("General", "DirectorioDatos", App.Path & "\soloini.ini")
cDataPath = GetFromINI("General", "ProveedorDatos", App.Path & "\soloini.ini")

MAX_ACOMP = GetFromINI("General", "Acompañantes", App.Path & "\soloini.ini")
Label2(2) = "After INI...": Label2(2).Refresh

cDataPath = cDataPath & ";Jet OLEDB:Database Password=master24"
cShapeADOString = "PROVIDER=MSDataShape;Data " & cDataPath

Label2(2) = "Before msConn.Open...": Label2(2).Refresh
msConn.Open cDataPath
Label2(2) = "After msConn.Open...": Label2(2).Refresh

If Left(cTempMachineName, 9) = "HSEQUEIRA" Then
    'INFO: PARA MI MAQUINA SOLAMENTE.
    Text1.Text = "1967"
    Text2.Text = "6666"
ElseIf Left(cTempMachineName, 6) = "SOLOXP" Then
    'INFO: PARA MI MAQUINA SOLAMENTE. WINDOWS XP
    Text1.Text = "1967"
    Text2.Text = "6666"
End If

cFactFile = DATA_PATH + FACTURA_FILE

On Error GoTo 0
cOldCaption = App.CompanyName
Label2(2) = App.CompanyName & " está Listo"
'Abrir los recordset
'" " se usa ya que listitems tiene problemas
'para "ver" los datos"

Call GetISC 'ITBMS

'=====================================
'=====================================
'INFO: 22JUL2016
' VERIFICA SI SE ESTA ACTUALIZANDO EL SISTEMA
'=====================================
'=====================================
On Error Resume Next
If RegRead("HKLM\Software\SoloSoftware\SoloAdmin\Version") = "" Then
    cVersionActualRegistry = "0.0.0"
Else
    cVersionActualRegistry = RegRead("HKLM\Software\SoloSoftware\SoloAdmin\Version")
End If

cArrayVersionActualRegistry = Split(cVersionActualRegistry, ".")

If cArrayVersionActualRegistry(0) = "0" Then
    If ValidacionLicencia Then
        Call RegWrite("HKLM\Software\SoloSoftware\SoloAdmin\Version", App.Major & "." & App.Minor & "." & App.Revision)
    Else
        'TERMINAR PROGRAMA
        Me.Command1.Enabled = False
    End If
Else
    Select Case App.Major
        Case Is > cArrayVersionActualRegistry(0)
            If ValidacionLicencia Then
                Call RegWrite("HKLM\Software\SoloSoftware\SoloAdmin\Version", App.Major & "." & App.Minor & "." & App.Revision)
            Else
                'TERMINAR PROGRAMA
                Me.Command1.Enabled = False
            End If
        Case Else
            If cArrayVersionActualRegistry(0) & "." & cArrayVersionActualRegistry(1) > _
                App.Minor & "." & App.Revision Then
                'NO HACER NADA
            Else
                'CONTINUAR
            End If
    End Select
End If
'=====================================
'=====================================


On Error GoTo 0

'INFO: CAMBIA STRUCT DE LA TABLA DE USUARIOS
Call VerificaTabla("USUARIOS")
Call VerificaTabla2("USUARIOS")

'INFO: CAMBIA STRUCT DE LA TABLA DE CAJEROS
Call VerificaTabla("CAJEROS")

'INFO: CAMBIA STRUCT DE LA TABLA CONTEND_02 (29SEP2013)
Call VerificaTabla("CONTEND_02")

'INFO: CAMBIA STRUCT DE LA TABLA ORGANIZACION (11SEP2015)
Call VerificaTabla("ORGANIZACION")

'INFO: VERIFCA LOS TIPOS DE USUARIOS PARA VER SI ESTAN LOS 4
Call VerificaTabla("TIPO_USUARIOS")

'INFO: 30NOV2017. CAMBIA EL INDICE DE LA TABLA DE DESCUENTO A AUTONUMERICO
Call VerificaTabla("DESCUENTO")

'INFO: 23OCT2020. CREAR TABLA
Call VerificaTabla2("CLIENTES_COUNTER")

'INFO: 18ABR2022. CREAR TABLA DE AREAS
Call VerificaTabla3("AREAS")
Call VerificaTabla3("AREAS_MESAS")


'|||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||
'|||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||
Call VerificaTabla4("PAGOS")
Call VerificaTabla4("TMP_TRANS")
Call VerificaTabla4("CLIENTE_FE")

'INFO: CAMBIANDO LONGITUD DEL CAMPO DE 15 A 75 (SOLO SE NECESITA 66)
'28MAR2024
Call VerificaTabla4("TRANSAC_FISCAL")

Me.MousePointer = vbDefault
Exit Sub

ErrorOpen:
If msConn.Errors.Count > 0 Then
    Dim OBJERR As Error
    For Each OBJERR In msConn.Errors
        MsgBox OBJERR.Description, vbCritical, "ERROR DE ACCESO A DATOS. INTENTE ENTRAR MAS TARDE"
    Next
    
    Unload Me
    End
Else
    MsgBox Err.Number & " - " & Err.Description, vbCritical, "POSIBLE PROBLEMA EN SOLOINI.INI"
    EscribeLog (Err.Number & " - " & Err.Description)
    Me.MousePointer = vbDefault
End If
End Sub

'Private Sub Image1_DblClick()
''MsgBox "Empresa: " & App.CompanyName & Chr(13) & _
'       "Derechos Reservados: " & App.LegalCopyright & Chr(13) & _
'       "Nombre: " & App.Title & ".EXE" & Chr(13) & _
'       "Versión: " & App.Major & "." & App.Minor & "." & App.Revision, vbInformation, "Informacion de la Aplicación"
'ShowMsg "Empresa : " & App.LegalCopyright & Chr(13) & _
'       "Nombre  : " & App.EXEName & Chr(13) & _
'       "Versión : " & App.Major & "." & App.Minor & "." & App.Revision, vbGreen, vbBlue
'End Sub
Private Sub Image_DblClick()
'MsgBox "Empresa: " & App.CompanyName & Chr(13) & _
       "Derechos Reservados: " & App.LegalCopyright & Chr(13) & _
       "Nombre: " & App.Title & ".EXE" & Chr(13) & _
       "Versión: " & App.Major & "." & App.Minor & "." & App.Revision, vbInformation, "Informacion de la Aplicación"
ShowMsg "Empresa : " & App.LegalCopyright & Chr(13) & _
       "Nombre : " & App.EXEName & Chr(13) & _
       "Versión : " & App.Major & "." & App.Minor & "." & App.Revision, vbGreen, vbBlue
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
'If KeyAscii = 13 Then
'    If Text1 = "" Then
'        Text1.SetFocus
'    Else
'        Text2.SetFocus
'    End If
'End If
'INFO: 13FEB2013
If KeyAscii = 13 Then
    If Text1 = "" Then
        Text1.SetFocus
    'ElseIf Not IsNumeric(Text1) Then
        'Text1.SetFocus
    Else
        Text2.SetFocus
        Text2.SelStart = 0
        Text2.SelLength = Len(Text2.Text)
    End If
End If
End Sub
Private Sub Text2_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    If Text2 = "" Then
        Text2.SetFocus
    Else
        Command1.SetFocus
    End If
End If
End Sub

Private Sub VerificaFecha()
Dim cMaxFecha As Date
Dim cMaxDia As String
Dim cMaxMes As String
Dim cMaxYear As String
Dim cLocalFecha As String

cMaxFecha = Date
cMaxMes = Mid(Format(cMaxFecha, "SHORT DATE"), 4, 2)
cMaxDia = Mid(Format(cMaxFecha, "SHORT DATE"), 1, 2)
cMaxYear = Mid(Format(cMaxFecha, "SHORT DATE"), 7, 4)

cLocalFecha = cMaxYear & cMaxMes & cMaxDia
If Val(cLocalFecha) > Val("20010430") Then
    MsgBox "***** SU PERIODO DE EVALUACION A TERMINADO *****" & vbCrLf & _
            "- GRACIAS POR PROBAR PRODUCTOS DE " & App.CompanyName & vbCrLf & _
            "- CONTACTE A SU PROVEEDOR" & vbCrLf & _
            "" & vbCrLf & "El programa terminara AHORA", vbCritical, "CONTACTE A SU PROVEEDOR"
    Unload Me
    End
End If
End Sub

Public Function VerificaTabla(cTableName As String) As Boolean
Dim bDoAlter_Table As Boolean
Dim rsTempTable As ADODB.Recordset

Set rsTempTable = New ADODB.Recordset

On Error Resume Next

Select Case cTableName
    'INFO: 25ABR2016. TIPO_USUARIOS
    Case "TIPO_USUARIOS"
        rsTempTable.Open "SELECT TIPO FROM " & cTableName & " WHERE TIPO = 4", msConn, adOpenStatic, adLockOptimistic
    Case "USUARIOS"
        rsTempTable.Open "SELECT TOP 2 CLAVE FROM " & cTableName, msConn, adOpenStatic, adLockOptimistic
    Case "CAJEROS"
        rsTempTable.Open "SELECT TOP 2 CLAVE FROM " & cTableName, msConn, adOpenStatic, adLockOptimistic
    Case "CONTEND_02"
        rsTempTable.Open "SELECT TOP 2 X_COUNT FROM " & cTableName, msConn, adOpenStatic, adLockOptimistic
    Case "ORGANIZACION"
        rsTempTable.Open "SELECT TOP 2 MENSAJE FROM " & cTableName, msConn, adOpenStatic, adLockOptimistic
    Case "DESCUENTO"    '30NOV2017
        rsTempTable.Open "SELECT TOP 2 TIPO FROM " & cTableName, msConn, adOpenStatic, adLockOptimistic
End Select

'If Err.Number = -2147217904 Then
'    MsgBox Err.Description
'End If

If Err.Number = -2147217865 Then
    'TABLA NO EXISTE
    bDoAlter_Table = False
Else
    On Error GoTo ErrAdm:
    Select Case cTableName
        
        Case "TIPO_USUARIOS"
            If rsTempTable.EOF Then
                bDoAlter_Table = True
            Else
                bDoAlter_Table = False
            End If
        Case "USUARIOS"
            If rsTempTable.Fields(0).DefinedSize = 20 Then
                'NO HACER NADA TIENE LA LONGITUD CORRECTA
                bDoAlter_Table = False
            Else
                bDoAlter_Table = True
            End If
        Case "CAJEROS"
            If rsTempTable.Fields(0).DefinedSize = 20 Then
                'NO HACER NADA TIENE LA LONGITUD CORRECTA
                bDoAlter_Table = False
            Else
                bDoAlter_Table = True
            End If
        Case "CONTEND_02"
            If rsTempTable.Fields(0).DefinedSize = 4 Then
                'NO HACER NADA TIENE LA LONGITUD CORRECTA
                bDoAlter_Table = False
            Else
                bDoAlter_Table = True
            End If
        Case "ORGANIZACION"
            'CAMBIA DE 30 A 140
            If rsTempTable.Fields(0).DefinedSize = 140 Then
                'NO HACER NADA TIENE LA LONGITUD CORRECTA
                bDoAlter_Table = False
            Else
                bDoAlter_Table = True
            End If
        Case "DESCUENTO"    '30NOV2017
            'CAMBIA DE 30 A 140
            If rsTempTable.Fields(0).Type <> 2 Then
            'If rsTempTable.Fields(0).DefinedSize = 140 Then
                'NO HACER NADA TIENE LA LONGITUD CORRECTA
                bDoAlter_Table = False
            Else
                bDoAlter_Table = True
            End If
        
        Case Else
    End Select
End If
Set rsTempTable = Nothing
VerificaTabla = True

If bDoAlter_Table Then
    Select Case cTableName
        'INFO: 25ABR2016
        Case "TIPO_USUARIOS"
            msConn.Execute "INSERT INTO TIPO_USUARIOS VALUES (4,'Master Chef')"
            EscribeLog ("Admin.Se crea Tipo Usuario #4")
        Case "USUARIOS"
            msConn.Execute "ALTER TABLE USUARIOS ALTER COLUMN CLAVE TEXT(20)"
            EscribeLog ("Admin.Actualizacion Tabla USUARIOS. Campo CLAVE")
        Case "CAJEROS"
            msConn.Execute "ALTER TABLE CAJEROS ALTER COLUMN CLAVE TEXT(20)"
            EscribeLog ("Admin.Actualizacion Tabla CAJEROS. Campo CLAVE")
        Case "CONTEND_02"
            msConn.Execute "ALTER TABLE CONTEND_02 ALTER COLUMN X_COUNT LONG"
            msConn.Execute "ALTER TABLE CONTEND_02 ALTER COLUMN Z_COUNT LONG"
            EscribeLog ("Admin.Actualizacion Tabla CONTEND_02 (X_COUNT y Z_COUNT LONG CHAR)")
        Case "ORGANIZACION"
            msConn.Execute "ALTER TABLE ORGANIZACION ALTER COLUMN MENSAJE TEXT(140)"
            EscribeLog ("Admin.Actualizacion Tabla ORGANIZACION. Campo MENSAJE")
        Case "CLIENTES_COUNTER"     '23OCT2021
        
        Case "DESCUENTO"    '30NOV2017
            'ALTER TABLE TBLTEMP ADD COLUMN TEMPID AUTOINCREMENT
            
            Label2(2).ForeColor = vbRed
            Label2(2).Tag = Label2(2).Caption
            Label2(2).Caption = "PROCESO TOMA 2-10 MIN. DEBE ESPERAR QUE TERMINE"
            Label2(2).Refresh
            
            msConn.BeginTrans
            msConn.Execute "DELETE * FROM DESCUENTO"
            msConn.CommitTrans
            
            'msConn.BeginTrans
            msConn.Execute "ALTER TABLE DESCUENTO ALTER COLUMN TIPO COUNTER"
            msConn.Execute "ALTER TABLE TMP_TRANS ADD COLUMN ID_DESCUENTO LONG DEFAULT 0"
            msConn.Execute "ALTER TABLE TRANSAC ADD COLUMN ID_DESCUENTO LONG DEFAULT 0"
            msConn.Execute "ALTER TABLE HIST_TR ADD COLUMN ID_DESCUENTO LONG DEFAULT 0"
            
            msConn.BeginTrans
            msConn.Execute "INSERT INTO DESCUENTO (DESCRIP, PORCENTAJE) VALUES ('Jubilado', 25)"
            msConn.Execute "INSERT INTO DESCUENTO (DESCRIP, PORCENTAJE) VALUES ('Empleados', 50)"
            msConn.CommitTrans

            msConn.BeginTrans
            msConn.Execute "UPDATE TMP_TRANS SET ID_DESCUENTO=0"
            msConn.Execute "UPDATE TRANSAC SET ID_DESCUENTO=0"
            msConn.CommitTrans
            
            Dim rsA As New ADODB.Recordset
            
            rsA.Open "SELECT DISTINCT LEFT(FECHA,6) AS MI_FECHA FROM HIST_TR ORDER BY 1", msConn, adOpenStatic, adLockOptimistic
            'Debug.Print "Se Actualizan : " & rsA.RecordCount
            Do While Not rsA.EOF
                'Debug.Print Time & " - " & rsA!MI_FECHA
                msConn.Execute "UPDATE HIST_TR SET ID_DESCUENTO = 0  WHERE LEFT(FECHA,6) = '" & rsA!MI_FECHA & "'"
                rsA.MoveNext
            Loop
            rsA.Close
            
            EscribeLog ("Admin.Actualización DESCUENTO, ID_DESCUENTO e HISTORICOS")
            
            Label2(2).ForeColor = vbBlue
            Label2(2).Caption = Label2(2).Tag
            Label2(2).Refresh
    
    End Select
End If
On Error GoTo 0
Exit Function

ErrAdm:
ShowMsg "Error ~~~FATAL~~~ al Verificar Proceso (" & cTableName & ").  " & Err.Number & " - " & Err.Description, vbYellow, vbRed
EscribeLog ("Error ~~~FATAL~~~ al Verificar Proceso (" & cTableName & ").  " & Err.Number & " - " & Err.Description)
Resume Next
End Function


'---------------------------------------------------------------------------------------
' Procedimiento : VerificaTabla2
' Autor       : hsequeira
' Fecha       : 01/12/2014
' Proposito   : AGREGA CAMPO DE AUTORIZACION A LA TABLA USUARIOS
'---------------------------------------------------------------------------------------
'
Public Function VerificaTabla2(cTableName As String) As Boolean
Dim bDoAlter_Table As Boolean
Dim rsTempTable As ADODB.Recordset

Set rsTempTable = New ADODB.Recordset

On Error Resume Next

Select Case cTableName
    Case "USUARIOS"
        rsTempTable.Open "SELECT TOP 2 AUTORIZACION FROM " & cTableName, msConn, adOpenStatic, adLockOptimistic
    Case "CLIENTES_COUNTER"
        rsTempTable.Open "SELECT TOP 2 MESA FROM " & cTableName, msConn, adOpenStatic, adLockOptimistic
        If Err.Number = -2147217865 Or -2147217904 Then
            cSQL2 = "CREATE TABLE CLIENTES_COUNTER "
            cSQL2 = cSQL2 & "(FECHA TEXT(8), HORA TEXT(8), MESA INTEGER, CLIENTES INTEGER)  "
'            msConn.BeginTrans
'            msConn.Execute cSQL2
'            msConn.CommitTrans
            'EscribeLog ("Admin.Creacion CLIENTES_COUNTER")
        End If
End Select

If Err.Number = -2147217865 Then
    'TABLA NO EXISTE
    bDoAlter_Table = True
Else
    On Error GoTo ErrAdm:
    Select Case cTableName
        Case "USUARIOS"
            If rsTempTable.Fields(0).DefinedSize = 10 Then
                'NO HACER NADA TIENE LA LONGITUD CORRECTA
                bDoAlter_Table = False
            Else
                bDoAlter_Table = True
            End If
        Case "CLIENTES_COUNTER"
            If rsTempTable.State = 0 Then
            ''If Err.Number = -2147217904 Then
                bDoAlter_Table = True
            Else
                bDoAlter_Table = False
            End If
            
        Case Else
    End Select
End If
Set rsTempTable = Nothing
VerificaTabla2 = True

If bDoAlter_Table Then
    Select Case cTableName
        Case "USUARIOS"
            msConn.Execute "ALTER TABLE USUARIOS ADD COLUMN AUTORIZACION TEXT(10)"
            EscribeLog ("Admin.Actualizacion Tabla USUARIOS. Campo AUTORIZACION")
        Case "CLIENTES_COUNTER"
        
            msConn.BeginTrans
            msConn.Execute cSQL2
            msConn.CommitTrans
            
            msConn.BeginTrans
            msConn.Execute "INSERT INTO CLIENTES_COUNTER VALUES (0,0,0,0)"
            msConn.CommitTrans
           
            EscribeLog ("Admin.Creacion CLIENTES_COUNTER")
    End Select
End If
On Error GoTo 0
Exit Function

ErrAdm:
ShowMsg "Error En Verifica Tabla (" & cTableName & ").  " & Err.Number & " - " & Err.Description
EscribeLog ("Admin. Error al Verificar Tabla (" & cTableName & ").  " & Err.Number & " - " & Err.Description)
End Function



'---------------------------------------------------------------------------------------
' Procedimiento : VerificaTabla3
' Autor       : hsequeira
' Fecha       : 18/04/2022
' Fecha       : 1/05/2023
' Proposito   : AGREGA TABLA AREAS (ID, DESCRIPCION)
'---------------------------------------------------------------------------------------
'
Public Function VerificaTabla3(cTableName As String) As Boolean
Dim bDoAlter_Table As Boolean
Dim rsTempTable As ADODB.Recordset

Set rsTempTable = New ADODB.Recordset

On Error Resume Next

Select Case cTableName
    Case "AREAS"
        rsTempTable.Open "SELECT TOP 2 DESCRIPCION FROM " & cTableName, msConn, adOpenStatic, adLockOptimistic
        If Err.Number = -2147217865 Or Err.Number = -2147217904 Then
            cSQL2 = "CREATE TABLE AREAS "
            cSQL2 = cSQL2 & "(ID AUTOINCREMENT PRIMARY KEY, DESCRIPCION TEXT(20))  "
'            msConn.BeginTrans
'            msConn.Execute cSQL2
'            msConn.CommitTrans
            'EscribeLog ("Admin.Creacion CLIENTES_COUNTER")
        End If
    Case "AREAS_MESAS"
        rsTempTable.Open "SELECT TOP 2 MESA FROM " & cTableName, msConn, adOpenStatic, adLockOptimistic
        If Err.Number = -2147217865 Or Err.Number = -2147217904 Then
            cSQL2 = "CREATE TABLE AREAS_MESAS "
            cSQL2 = cSQL2 & "(AREA LONG, MESA LONG, FECHA TEXT(40),  CONSTRAINT AreaMesaIndex UNIQUE (AREA, MESA))  "
        End If
End Select

If Err.Number = -2147217865 Then
    'TABLA NO EXISTE
    bDoAlter_Table = True
    cOldCaption = Sys_Adm.Label2(2).Caption
    Sys_Adm.Label2(2).Caption = "ACTUALIZANDO DATOS ESPERE..."
Else
    On Error GoTo ErrAdm:
    Select Case cTableName
        Case "AREAS"
            If rsTempTable.State = 0 Then
            ''If Err.Number = -2147217904 Then
                bDoAlter_Table = True
            Else
                bDoAlter_Table = False
            End If
        Case "AREAS_MESAS"
            If rsTempTable.State = 0 Then
            ''If Err.Number = -2147217904 Then
                bDoAlter_Table = True
            Else
                bDoAlter_Table = False
            End If
        Case Else
    End Select
End If
Set rsTempTable = Nothing
VerificaTabla3 = True

If bDoAlter_Table Then
    Select Case cTableName
        Case "AREAS"
        
            msConn.BeginTrans
            msConn.Execute cSQL2
            msConn.CommitTrans
            
            ';msConn.BeginTrans
            ';msConn.Execute "INSERT INTO AREAS (DESCRIPCION) VALUES ('NO ASIGNADO')"
            ';msConn.CommitTrans
            
            
            'CREACION DEL CAMPO AREA EN (MESAS, TMP_TRANS, TRANSAC, HIST_TR)
            msConn.BeginTrans
            msConn.Execute "ALTER TABLE MESAS ADD COLUMN AREA LONG DEFAULT 0"
            msConn.Execute "ALTER TABLE TMP_TRANS ADD COLUMN AREA LONG DEFAULT 0"
            msConn.Execute "ALTER TABLE TRANSAC ADD COLUMN AREA LONG DEFAULT 0"
            msConn.Execute "ALTER TABLE HIST_TR ADD COLUMN AREA LONG DEFAULT 0"
            msConn.CommitTrans
            
            msConn.BeginTrans
            msConn.Execute "UPDATE MESAS SET AREA=0"
            msConn.Execute "UPDATE TMP_TRANS SET AREA=0"
            msConn.Execute "UPDATE TRANSAC SET AREA=0"
            msConn.Execute "UPDATE HIST_TR SET AREA=0"
            msConn.CommitTrans
                        
            msConn.BeginTrans
            msConn.Execute "INSERT INTO SYS_01 VALUES (50,'Areas del Local','Vta12')"
            msConn.CommitTrans
                        
                        
            EscribeLog ("Admin.Creacion AREAS (MESAS, TMP_TRANS, TRANSAC, HIST_TR)")
            
    Case "AREAS_MESAS"
            
            msConn.BeginTrans
            msConn.Execute cSQL2
            'msConn.Execute "CREATE INDEX AreaMesaIndex ON AREAS_MESAS (AREA, MESA)"
            msConn.CommitTrans
    
    End Select
End If
On Error GoTo 0
Sys_Adm.Label2(2).Caption = cOldCaption
Sys_Adm.Label2(2).Refresh
Exit Function

ErrAdm:
ShowMsg "Error En Verifica Tabla 3 (" & cTableName & ").  " & Err.Number & " - " & Err.Description
EscribeLog ("Admin. Error al Verificar Tabla (" & cTableName & ").  " & Err.Number & " - " & Err.Description)
Sys_Adm.Label2(2).Caption = "FALLO ACTUALIZACION. LLAMAR SUPERVISOR"
Resume
End Function



'---------------------------------------------------------------------------------------------
' Procedure : ValidacionLicencia
' Author    : hsequeira
' Date      : 24/07/2016
'---------------------------------------------------------------------------------------------
' Purpose   : PIDE LA LICENCIA SI ES UNA ACTUALIZACION DE VERSION
'---------------------------------------------------------------------------------------------
' SI ES UNA PRIMERA ACTUALIZACION DE VERSIONES ANTERIORES: ~~ Pide licencia ~~
' SI ACTUALIZA DE: 9.1.25 a 10.0.0 ~~ Pide licencia ~~
' SI ACTUALIZA DE: 9.1.25 a 9.7.25 << ~~ NO pide licencia ~~ >>
'---------------------------------------------------------------------------------------------
'
Private Function ValidacionLicencia() As Boolean
Dim nDiadel_yyyy As Integer
Dim aDiaLicencia(366, 0) As String
Dim i As Integer

aDiaLicencia(0, 0) = "+-*/-+-*/"
aDiaLicencia(1, 0) = "XLBP-5357"
aDiaLicencia(2, 0) = "HNXK-4646"
aDiaLicencia(3, 0) = "LOQB-2587"
aDiaLicencia(4, 0) = "YXUF-3933"
aDiaLicencia(5, 0) = "VVLR-6758"
aDiaLicencia(6, 0) = "SXZT-5351"
aDiaLicencia(7, 0) = "HSEN-5047"
aDiaLicencia(8, 0) = "NDUD-8575"
aDiaLicencia(9, 0) = "MAIB-8505"
aDiaLicencia(10, 0) = "TPBD-5614"
aDiaLicencia(11, 0) = "OEOX-9149"
aDiaLicencia(12, 0) = "RLNS-5433"
aDiaLicencia(13, 0) = "ELLN-8296"
aDiaLicencia(14, 0) = "KLZF-3597"
aDiaLicencia(15, 0) = "XNLO-7912"
aDiaLicencia(16, 0) = "ZCRE-2633"
aDiaLicencia(17, 0) = "FGKP-4473"
aDiaLicencia(18, 0) = "LUOE-6805"
aDiaLicencia(19, 0) = "JUDP-2523"
aDiaLicencia(20, 0) = "LSQU-5601"
aDiaLicencia(21, 0) = "QXBJ-2097"
aDiaLicencia(22, 0) = "MNHF-7175"
aDiaLicencia(23, 0) = "XZOI-4749"
aDiaLicencia(24, 0) = "EBMP-3099"
aDiaLicencia(25, 0) = "CKWQ-6675"
aDiaLicencia(26, 0) = "JQSE-6839"
aDiaLicencia(27, 0) = "YIYQ-5089"
aDiaLicencia(28, 0) = "OJHE-1107"
aDiaLicencia(29, 0) = "MSUO-4989"
aDiaLicencia(30, 0) = "WXMU-5978"
aDiaLicencia(31, 0) = "TRQQ-4242"
aDiaLicencia(32, 0) = "ZLKL-4228"
aDiaLicencia(33, 0) = "EGHU-7935"
aDiaLicencia(34, 0) = "IIMP-5262"
aDiaLicencia(35, 0) = "WNJY-7865"
aDiaLicencia(36, 0) = "OKIF-2398"
aDiaLicencia(37, 0) = "HZSD-5188"
aDiaLicencia(38, 0) = "DBSX-5131"
aDiaLicencia(39, 0) = "VOJS-9301"
aDiaLicencia(40, 0) = "LRDV-5843"
aDiaLicencia(41, 0) = "IHOI-7485"
aDiaLicencia(42, 0) = "CPLY-2149"
aDiaLicencia(43, 0) = "RBDW-1114"
aDiaLicencia(44, 0) = "KGHD-5604"
aDiaLicencia(45, 0) = "PCIQ-5394"
aDiaLicencia(46, 0) = "YZMH-2615"
aDiaLicencia(47, 0) = "QQPU-2585"
aDiaLicencia(48, 0) = "PFAN-1234"
aDiaLicencia(49, 0) = "OUMP-1465"
aDiaLicencia(50, 0) = "VDYM-4722"
aDiaLicencia(51, 0) = "WHJY-3527"
aDiaLicencia(52, 0) = "KVHS-6085"
aDiaLicencia(53, 0) = "SQNV-2009"
aDiaLicencia(54, 0) = "JMPX-3259"
aDiaLicencia(55, 0) = "NLJH-8741"
aDiaLicencia(56, 0) = "EXWX-7866"
aDiaLicencia(57, 0) = "EMBP-8435"
aDiaLicencia(58, 0) = "PPQZ-4440"
aDiaLicencia(59, 0) = "TYEO-5167"
aDiaLicencia(60, 0) = "NJTD-1523"
aDiaLicencia(61, 0) = "BEFM-1396"
aDiaLicencia(62, 0) = "PMMQ-3342"
aDiaLicencia(63, 0) = "KLBO-6624"
aDiaLicencia(64, 0) = "TYYX-5757"
aDiaLicencia(65, 0) = "EALS-5518"
aDiaLicencia(66, 0) = "ZEKQ-6552"
aDiaLicencia(67, 0) = "ZDPB-8473"
aDiaLicencia(68, 0) = "HPFE-4165"
aDiaLicencia(69, 0) = "UNPF-7081"
aDiaLicencia(70, 0) = "XTMF-8778"
aDiaLicencia(71, 0) = "KCCK-6848"
aDiaLicencia(72, 0) = "PZYV-7894"
aDiaLicencia(73, 0) = "TXQV-3108"
aDiaLicencia(74, 0) = "AZCK-5402"
aDiaLicencia(75, 0) = "OEDQ-1024"
aDiaLicencia(76, 0) = "JENJ-6087"
aDiaLicencia(77, 0) = "IIQC-9742"
aDiaLicencia(78, 0) = "SNCN-6459"
aDiaLicencia(79, 0) = "XIVR-1250"
aDiaLicencia(80, 0) = "LLHR-1939"
aDiaLicencia(81, 0) = "BSEY-2416"
aDiaLicencia(82, 0) = "MMBJ-7888"
aDiaLicencia(83, 0) = "GUWT-6700"
aDiaLicencia(84, 0) = "CSTZ-5853"
aDiaLicencia(85, 0) = "ZKKR-2465"
aDiaLicencia(86, 0) = "OSPP-5864"
aDiaLicencia(87, 0) = "WOOL-9833"
aDiaLicencia(88, 0) = "JJKA-2184"
aDiaLicencia(89, 0) = "IOEK-7429"
aDiaLicencia(90, 0) = "ADVV-5029"
aDiaLicencia(91, 0) = "YRIB-3905"
aDiaLicencia(92, 0) = "RDIE-4142"
aDiaLicencia(93, 0) = "CILM-1757"
aDiaLicencia(94, 0) = "RCSS-2332"
aDiaLicencia(95, 0) = "LFNF-3431"
aDiaLicencia(96, 0) = "OAZH-9119"
aDiaLicencia(97, 0) = "LGWE-4490"
aDiaLicencia(98, 0) = "FQMJ-2324"
aDiaLicencia(99, 0) = "MRAF-6239"
aDiaLicencia(100, 0) = "IVEU-2757"
aDiaLicencia(101, 0) = "NTMO-7590"
aDiaLicencia(102, 0) = "NDDH-1636"
aDiaLicencia(103, 0) = "IXFY-2710"
aDiaLicencia(104, 0) = "DWBU-9977"
aDiaLicencia(105, 0) = "GTZI-2473"
aDiaLicencia(106, 0) = "ERRK-8515"
aDiaLicencia(107, 0) = "HLFE-9749"
aDiaLicencia(108, 0) = "VKEI-9723"
aDiaLicencia(109, 0) = "ELUB-5655"
aDiaLicencia(110, 0) = "FSFC-7985"
aDiaLicencia(111, 0) = "TCOD-4792"
aDiaLicencia(112, 0) = "KSER-4138"
aDiaLicencia(113, 0) = "UCUV-3983"
aDiaLicencia(114, 0) = "TCAK-4912"
aDiaLicencia(115, 0) = "ATYU-4203"
aDiaLicencia(116, 0) = "DJTQ-5081"
aDiaLicencia(117, 0) = "SITQ-6892"
aDiaLicencia(118, 0) = "AEBH-2599"
aDiaLicencia(119, 0) = "XHZD-6631"
aDiaLicencia(120, 0) = "EOZZ-7327"
aDiaLicencia(121, 0) = "OAZC-5435"
aDiaLicencia(122, 0) = "TEOW-9497"
aDiaLicencia(123, 0) = "HJHN-7036"
aDiaLicencia(124, 0) = "OKMH-3472"
aDiaLicencia(125, 0) = "PGAV-4628"
aDiaLicencia(126, 0) = "QZSV-1256"
aDiaLicencia(127, 0) = "NTSN-6517"
aDiaLicencia(128, 0) = "MMIB-1347"
aDiaLicencia(129, 0) = "CEBM-8923"
aDiaLicencia(130, 0) = "BRUZ-1220"
aDiaLicencia(131, 0) = "WPJN-7031"
aDiaLicencia(132, 0) = "GUHV-2681"
aDiaLicencia(133, 0) = "VHVY-5781"
aDiaLicencia(134, 0) = "TICO-7896"
aDiaLicencia(135, 0) = "QQOE-8272"
aDiaLicencia(136, 0) = "NPSU-2785"
aDiaLicencia(137, 0) = "IFGO-5971"
aDiaLicencia(138, 0) = "FCIX-8129"
aDiaLicencia(139, 0) = "ANRM-7266"
aDiaLicencia(140, 0) = "KFFN-2002"
aDiaLicencia(141, 0) = "GTCY-2450"
aDiaLicencia(142, 0) = "ISJT-5377"
aDiaLicencia(143, 0) = "XKHI-1382"
aDiaLicencia(144, 0) = "YUFQ-6669"
aDiaLicencia(145, 0) = "CQTU-8843"
aDiaLicencia(146, 0) = "VUJV-4252"
aDiaLicencia(147, 0) = "FJRA-7157"
aDiaLicencia(148, 0) = "BQEV-2407"
aDiaLicencia(149, 0) = "SVRB-4649"
aDiaLicencia(150, 0) = "URYU-8954"
aDiaLicencia(151, 0) = "CDBY-3630"
aDiaLicencia(152, 0) = "RCVT-8610"
aDiaLicencia(153, 0) = "OTFG-5463"
aDiaLicencia(154, 0) = "FHDF-6839"
aDiaLicencia(155, 0) = "KPRD-6467"
aDiaLicencia(156, 0) = "ERDB-3726"
aDiaLicencia(157, 0) = "JCMH-1460"
aDiaLicencia(158, 0) = "GAST-4879"
aDiaLicencia(159, 0) = "SMHE-7877"
aDiaLicencia(160, 0) = "BMKG-6980"
aDiaLicencia(161, 0) = "MGWB-5143"
aDiaLicencia(162, 0) = "VSNI-4465"
aDiaLicencia(163, 0) = "YFZY-4197"
aDiaLicencia(164, 0) = "RNNC-4126"
aDiaLicencia(165, 0) = "NYBO-8918"
aDiaLicencia(166, 0) = "FXOI-1488"
aDiaLicencia(167, 0) = "NTKJ-3040"
aDiaLicencia(168, 0) = "TRDJ-8073"
aDiaLicencia(169, 0) = "XMIZ-2631"
aDiaLicencia(170, 0) = "EWPB-2851"
aDiaLicencia(171, 0) = "CKHD-8196"
aDiaLicencia(172, 0) = "NBDS-6360"
aDiaLicencia(173, 0) = "VOSK-9788"
aDiaLicencia(174, 0) = "IMVC-5517"
aDiaLicencia(175, 0) = "IYED-5939"
aDiaLicencia(176, 0) = "MEQG-8349"
aDiaLicencia(177, 0) = "GAYC-6293"
aDiaLicencia(178, 0) = "XLXG-6099"
aDiaLicencia(179, 0) = "GWOS-7291"
aDiaLicencia(180, 0) = "MZLA-7337"
aDiaLicencia(181, 0) = "MFMF-2273"
aDiaLicencia(182, 0) = "OCTM-1353"
aDiaLicencia(183, 0) = "HBXM-8781"
aDiaLicencia(184, 0) = "VDXE-7088"
aDiaLicencia(185, 0) = "CFMO-2668"
aDiaLicencia(186, 0) = "GZXS-6594"
aDiaLicencia(187, 0) = "SJDB-3253"
aDiaLicencia(188, 0) = "SOJP-5283"
aDiaLicencia(189, 0) = "TJVF-9455"
aDiaLicencia(190, 0) = "DIUC-4604"
aDiaLicencia(191, 0) = "XSMZ-7578"
aDiaLicencia(192, 0) = "SJIR-8102"
aDiaLicencia(193, 0) = "TGTN-5562"
aDiaLicencia(194, 0) = "LRMJ-3147"
aDiaLicencia(195, 0) = "HOYK-2887"
aDiaLicencia(196, 0) = "ZVWY-1323"
aDiaLicencia(197, 0) = "JFXU-3846"
aDiaLicencia(198, 0) = "KGHT-8470"
aDiaLicencia(199, 0) = "YWIX-2120"
aDiaLicencia(200, 0) = "KWYT-5135"
aDiaLicencia(201, 0) = "GJKD-5972"
aDiaLicencia(202, 0) = "UCRX-5159"
aDiaLicencia(203, 0) = "AUIL-9973"
aDiaLicencia(204, 0) = "YHVP-8026"
aDiaLicencia(205, 0) = "YHBB-1892"
aDiaLicencia(206, 0) = "KFWJ-6420"
aDiaLicencia(207, 0) = "QEQU-5720"
aDiaLicencia(208, 0) = "JVCC-6435"
aDiaLicencia(209, 0) = "LYAX-1037"
aDiaLicencia(210, 0) = "SDCA-8021"
aDiaLicencia(211, 0) = "GLCZ-6721"
aDiaLicencia(212, 0) = "PLEI-6485"
aDiaLicencia(213, 0) = "OGKK-7474"
aDiaLicencia(214, 0) = "YKXI-1371"
aDiaLicencia(215, 0) = "VJDN-8235"
aDiaLicencia(216, 0) = "OUKT-7363"
aDiaLicencia(217, 0) = "XDCW-3215"
aDiaLicencia(218, 0) = "JUPG-4281"
aDiaLicencia(219, 0) = "FDGA-6189"
aDiaLicencia(220, 0) = "WEVI-3419"
aDiaLicencia(221, 0) = "YNOD-1591"
aDiaLicencia(222, 0) = "NBIB-1364"
aDiaLicencia(223, 0) = "XUQD-9098"
aDiaLicencia(224, 0) = "RAAS-6288"
aDiaLicencia(225, 0) = "EIRL-5013"
aDiaLicencia(226, 0) = "FNVX-1111"
aDiaLicencia(227, 0) = "FMGH-7265"
aDiaLicencia(228, 0) = "ARNW-4304"
aDiaLicencia(229, 0) = "IGYD-8364"
aDiaLicencia(230, 0) = "EHRN-5571"
aDiaLicencia(231, 0) = "BFMM-4686"
aDiaLicencia(232, 0) = "HHIK-5489"
aDiaLicencia(233, 0) = "EZET-2132"
aDiaLicencia(234, 0) = "TGWP-6258"
aDiaLicencia(235, 0) = "XSLG-8037"
aDiaLicencia(236, 0) = "QARM-2953"
aDiaLicencia(237, 0) = "VQAL-4432"
aDiaLicencia(238, 0) = "RYCM-1121"
aDiaLicencia(239, 0) = "GCQH-8408"
aDiaLicencia(240, 0) = "TXTC-9938"
aDiaLicencia(241, 0) = "JVJA-4231"
aDiaLicencia(242, 0) = "ZUCK-8036"
aDiaLicencia(243, 0) = "LQPT-7113"
aDiaLicencia(244, 0) = "DKGM-2087"
aDiaLicencia(245, 0) = "KQDS-6227"
aDiaLicencia(246, 0) = "SAEH-5150"
aDiaLicencia(247, 0) = "OVKI-8525"
aDiaLicencia(248, 0) = "JLKI-5677"
aDiaLicencia(249, 0) = "XPML-8443"
aDiaLicencia(250, 0) = "BMAK-2872"
aDiaLicencia(251, 0) = "JJRP-2415"
aDiaLicencia(252, 0) = "LZGZ-7276"
aDiaLicencia(253, 0) = "UBLG-7168"
aDiaLicencia(254, 0) = "SRVI-7764"
aDiaLicencia(255, 0) = "RRBS-1348"
aDiaLicencia(256, 0) = "ATAI-9999"
aDiaLicencia(257, 0) = "FUMA-3811"
aDiaLicencia(258, 0) = "LTJI-1327"
aDiaLicencia(259, 0) = "OHYO-3578"
aDiaLicencia(260, 0) = "TIXD-8872"
aDiaLicencia(261, 0) = "WEKG-1005"
aDiaLicencia(262, 0) = "SVQL-7591"
aDiaLicencia(263, 0) = "MNVM-6879"
aDiaLicencia(264, 0) = "DKIA-5943"
aDiaLicencia(265, 0) = "TMKI-7577"
aDiaLicencia(266, 0) = "FXMA-2836"
aDiaLicencia(267, 0) = "TCSD-3200"
aDiaLicencia(268, 0) = "CPLI-9322"
aDiaLicencia(269, 0) = "PKPG-1074"
aDiaLicencia(270, 0) = "BELD-2865"
aDiaLicencia(271, 0) = "ORRT-7791"
aDiaLicencia(272, 0) = "DEPX-8266"
aDiaLicencia(273, 0) = "GJCO-4506"
aDiaLicencia(274, 0) = "RZQB-9029"
aDiaLicencia(275, 0) = "RQFR-1614"
aDiaLicencia(276, 0) = "TSZE-5747"
aDiaLicencia(277, 0) = "TJGD-2850"
aDiaLicencia(278, 0) = "NBCP-1857"
aDiaLicencia(279, 0) = "SVHL-7584"
aDiaLicencia(280, 0) = "YVHS-2172"
aDiaLicencia(281, 0) = "DXGQ-9959"
aDiaLicencia(282, 0) = "HHZZ-1796"
aDiaLicencia(283, 0) = "DTVA-6789"
aDiaLicencia(284, 0) = "GQCM-7796"
aDiaLicencia(285, 0) = "ZXZT-6646"
aDiaLicencia(286, 0) = "CVII-5088"
aDiaLicencia(287, 0) = "WFGY-9166"
aDiaLicencia(288, 0) = "NPYC-6017"
aDiaLicencia(289, 0) = "XORS-2182"
aDiaLicencia(290, 0) = "VSSM-2501"
aDiaLicencia(291, 0) = "JLTW-7152"
aDiaLicencia(292, 0) = "IIER-7476"
aDiaLicencia(293, 0) = "DJOB-9063"
aDiaLicencia(294, 0) = "NWLG-6322"
aDiaLicencia(295, 0) = "KMPD-2892"
aDiaLicencia(296, 0) = "KBCF-9602"
aDiaLicencia(297, 0) = "OSIX-4827"
aDiaLicencia(298, 0) = "XXVH-1945"
aDiaLicencia(299, 0) = "BBNC-7942"
aDiaLicencia(300, 0) = "YGCJ-1946"
aDiaLicencia(301, 0) = "HAWD-2286"
aDiaLicencia(302, 0) = "ENQS-2637"
aDiaLicencia(303, 0) = "XUZW-6354"
aDiaLicencia(304, 0) = "ZAES-1290"
aDiaLicencia(305, 0) = "IQPO-7650"
aDiaLicencia(306, 0) = "RWCH-5672"
aDiaLicencia(307, 0) = "VYZA-7711"
aDiaLicencia(308, 0) = "IVBX-7150"
aDiaLicencia(309, 0) = "TXCI-3281"
aDiaLicencia(310, 0) = "OZFV-8832"
aDiaLicencia(311, 0) = "MCWH-7860"
aDiaLicencia(312, 0) = "SYFN-9965"
aDiaLicencia(313, 0) = "RQRP-7877"
aDiaLicencia(314, 0) = "XLTZ-6496"
aDiaLicencia(315, 0) = "HBWL-3895"
aDiaLicencia(316, 0) = "JZJS-8148"
aDiaLicencia(317, 0) = "XCLF-2388"
aDiaLicencia(318, 0) = "VUXH-7609"
aDiaLicencia(319, 0) = "UWJS-3295"
aDiaLicencia(320, 0) = "WVHF-2577"
aDiaLicencia(321, 0) = "YITN-2678"
aDiaLicencia(322, 0) = "RQNF-6530"
aDiaLicencia(323, 0) = "KGML-9988"
aDiaLicencia(324, 0) = "HHLC-7553"
aDiaLicencia(325, 0) = "BHCB-3127"
aDiaLicencia(326, 0) = "MDJE-1995"
aDiaLicencia(327, 0) = "EULQ-1374"
aDiaLicencia(328, 0) = "XBRG-1072"
aDiaLicencia(329, 0) = "OZJX-2528"
aDiaLicencia(330, 0) = "JRCY-8516"
aDiaLicencia(331, 0) = "TMRT-4912"
aDiaLicencia(332, 0) = "HFYT-3803"
aDiaLicencia(333, 0) = "YAYD-3692"
aDiaLicencia(334, 0) = "JPSU-3251"
aDiaLicencia(335, 0) = "ICOM-6245"
aDiaLicencia(336, 0) = "WEPI-5991"
aDiaLicencia(337, 0) = "CYLN-4015"
aDiaLicencia(338, 0) = "XPIJ-3493"
aDiaLicencia(339, 0) = "RDPI-4341"
aDiaLicencia(340, 0) = "YIFN-9035"
aDiaLicencia(341, 0) = "BMEA-7186"
aDiaLicencia(342, 0) = "WDKP-1887"
aDiaLicencia(343, 0) = "GXAZ-3576"
aDiaLicencia(344, 0) = "YIGD-4939"
aDiaLicencia(345, 0) = "WDZH-7812"
aDiaLicencia(346, 0) = "AETJ-9616"
aDiaLicencia(347, 0) = "IFHY-7231"
aDiaLicencia(348, 0) = "NEFI-8251"
aDiaLicencia(349, 0) = "PBES-3452"
aDiaLicencia(350, 0) = "SMBW-5641"
aDiaLicencia(351, 0) = "MKXW-8456"
aDiaLicencia(352, 0) = "BFRU-8619"
aDiaLicencia(353, 0) = "COTY-2410"
aDiaLicencia(354, 0) = "KXXV-8500"
aDiaLicencia(355, 0) = "KOYC-4565"
aDiaLicencia(356, 0) = "ANOD-3347"
aDiaLicencia(357, 0) = "WLNK-2246"
aDiaLicencia(358, 0) = "VVVR-7113"
aDiaLicencia(359, 0) = "IIXL-4779"
aDiaLicencia(360, 0) = "HBWA-2073"
aDiaLicencia(361, 0) = "UKZB-8692"
aDiaLicencia(362, 0) = "ZRSJ-7649"
aDiaLicencia(363, 0) = "KFCF-3778"
aDiaLicencia(364, 0) = "GTTR-6728"
aDiaLicencia(365, 0) = "YLCP-6253"
aDiaLicencia(366, 0) = "DIAH-3412"

nDiadel_yyyy = Format(Date, "y")

Debug.Print aDiaLicencia(nDiadel_yyyy, 0)
clicencia = InputBox("INTRODUZCA SU LICENCIA - " & nDiadel_yyyy, "LICENCIA DE ACTUALIZACION")

If clicencia = aDiaLicencia(nDiadel_yyyy, 0) Then
    ValidacionLicencia = True
Else
    ValidacionLicencia = False
    ShowMsg "LICENCIA INVALIDA" & vbCrLf & vbCrLf & "Contacte a Su Proveedor", vbYellow, vbRed
End If

End Function


'---------------------------------------------------------------------------------------
' Procedure : VerificaTabla4
' Author    : hsequeira
' Date      : 29/08/2023
' Purpose   : MODIFCACION DE TABLAS PARA FACTURA ELECTRONICA
'---------------------------------------------------------------------------------------
'
Public Function VerificaTabla4(cTableName As String) As Boolean
Dim bDoAlter_Table As Boolean
Dim rsTempTable As ADODB.Recordset
Dim cCreateTable As String

Set rsTempTable = New ADODB.Recordset

On Error Resume Next

Select Case cTableName
    Case "PAGOS"
        rsTempTable.Open "SELECT TOP 2 ID_FE FROM " & cTableName, msConn, adOpenStatic, adLockOptimistic
        If Err.Number = -2147217865 Or Err.Number = -2147217904 Then
            ID = 1
        End If
    Case "TMP_TRANS"
        rsTempTable.Open "SELECT TOP 2 FE_DESCUENTO FROM " & cTableName, msConn, adOpenStatic, adLockOptimistic
        If Err.Number = -2147217865 Or Err.Number = -2147217904 Then
            ID = 1
        End If
    Case "CLIENTE_FE"
        rsTempTable.Open "SELECT TOP 2 ID FROM " & cTableName, msConn, adOpenStatic, adLockOptimistic
        If Err.Number = -2147217865 Or Err.Number = -2147217904 Then
            ID = 1
        End If
    Case "TRANSAC_FISCAL"
        rsTempTable.Open "SELECT TOP 2 ID_URL FROM " & cTableName, msConn, adOpenStatic, adLockOptimistic
        If Err.Number = -2147217865 Or Err.Number = -2147217904 Then
            ID = 1
        End If
End Select

If Err.Number = -2147217865 Then
    'TABLA NO EXISTE
    bDoAlter_Table = True
    cOldCaption = Sys_Adm.Label2(2).Caption
    Sys_Adm.Label2(2).Caption = "ACTUALIZANDO DATOS ESPERE..."
Else
    On Error GoTo ErrAdm:
    Select Case cTableName
        Case "PAGOS"
            If rsTempTable.State = 0 Then
            ''If Err.Number = -2147217904 Then
                bDoAlter_Table = True
            Else
                bDoAlter_Table = False
            End If
        Case "TMP_TRANS"
            If rsTempTable.State = 0 Then
            ''If Err.Number = -2147217904 Then
                bDoAlter_Table = True
            Else
                bDoAlter_Table = False
            End If
        Case "CLIENTE_FE"
            If rsTempTable.State = 0 Then
            ''If Err.Number = -2147217904 Then
                bDoAlter_Table = True
            Else
                bDoAlter_Table = False
            End If
        Case "TRANSAC_FISCAL"
            If rsTempTable.State = 0 Then
            ''If Err.Number = -2147217904 Then
                bDoAlter_Table = True
            Else
                bDoAlter_Table = False
            End If
        Case Else
    End Select
End If
Set rsTempTable = Nothing
VerificaTabla4 = True

If bDoAlter_Table Then
    Select Case cTableName
        Case "PAGOS"
            msConn.BeginTrans
                msConn.Execute "ALTER TABLE PAGOS ADD COLUMN ID_FE TEXT(2)  NULL"
            msConn.CommitTrans
            
            EscribeLog ("Admin.Creacion CAMPO PAGOS FACTURA ELECTRONICA")
        Case "TMP_TRANS"
            msConn.BeginTrans
                msConn.Execute "ALTER TABLE TMP_TRANS ADD COLUMN FE_DESCUENTO SINGLE DEFAULT 0"
            msConn.CommitTrans
            
            msConn.BeginTrans
                msConn.Execute "UPDATE TMP_TRANS SET FE_DESCUENTO = 0"
            msConn.CommitTrans
            
            EscribeLog ("Admin.Creacion /Update CAMPO PAGOS FACTURA ELECTRONICA")
        Case "CLIENTE_FE"
            msConn.BeginTrans
                        cCreateTable = "CREATE TABLE CLIENTE_FE "
                        cCreateTable = cCreateTable & "(ID AUTOINCREMENT PRIMARY KEY, "
                        cCreateTable = cCreateTable & "TIPO_CLIENTE TEXT(2), CONTRIBUYENTE INTEGER, "
                        cCreateTable = cCreateTable & "NOMBRE TEXT(50), CEDULA_RUC TEXT(25), "
                        cCreateTable = cCreateTable & "DV TEXT(2), RAZON_SOCIAL_NOMBRE TEXT(50), "
                        cCreateTable = cCreateTable & "DIRECCION TEXT(100), PAIS TEXT(2), "
                        cCreateTable = cCreateTable & "PROVINCIA TEXT(30), DISTRITO TEXT(40), "
                        cCreateTable = cCreateTable & "CORREGIMIENTO TEXT(40), EMAIL TEXT(50), "
                        cCreateTable = cCreateTable & "TELEFONO TEXT(15), ID_UBICACION TEXT(10) )"
            msConn.Execute cCreateTable
            msConn.CommitTrans
            
            EscribeLog ("Admin.Creacion TABLA CLIENTE_FE FACTURA ELECTRONICA")
        Case "TRANSAC_FISCAL"
            msConn.BeginTrans
            msConn.Execute "ALTER TABLE TRANSAC_FISCAL ALTER COLUMN FISCAL TEXT(75) DEFAULT 0"
            msConn.Execute "ALTER TABLE TRANSAC_FISCAL ADD COLUMN ID_URL TEXT(45) DEFAULT 0"
            msConn.CommitTrans
            EscribeLog ("Admin.Update CAMPO FISCAL GUARDAR CUFE FE y ID_URL PARA EMAIL POSTERIOR")
    End Select
End If
On Error GoTo 0
Sys_Adm.Label2(2).Caption = cOldCaption
Sys_Adm.Label2(2).Refresh
Exit Function

ErrAdm:
If cTableName = "TRANSAC_FISCAL" Then
    If Err.Number = -2147217887 Then
        'ERROR CONOCIDO IGNORAR
        Resume Next
    Else
        ShowMsg "Error En Verifica Tabla4 (" & cTableName & ").  " & Err.Number & " - " & Err.Description
        EscribeLog ("Admin. Error al Verificar (FE) Tabla (" & cTableName & ").  " & Err.Number & " - " & Err.Description)
        Sys_Adm.Label2(2).Caption = "FALLO ACTUALIZACION. LLAMAR SUPERVISOR"
    End If
Else
    ShowMsg "Error En Verifica Tabla4 (" & cTableName & ").  " & Err.Number & " - " & Err.Description, vbYellow, vbRed
    EscribeLog ("Admin. Error al Verificar (FE) Tabla (" & cTableName & ").  " & Err.Number & " - " & Err.Description)
    Sys_Adm.Label2(2).Caption = "FALLO ACTUALIZACION. LLAMAR SUPERVISOR"
End If
End Function

