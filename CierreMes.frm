VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form CierreMes 
   BackColor       =   &H00B39665&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Cierre de Mes"
   ClientHeight    =   3195
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4680
   ClipControls    =   0   'False
   Icon            =   "CierreMes.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   StartUpPosition =   2  'CenterScreen
   Begin MSComctlLib.ProgressBar ProgBar 
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   2880
      Width           =   4455
      _ExtentX        =   7858
      _ExtentY        =   450
      _Version        =   393216
      Appearance      =   1
   End
   Begin VB.CommandButton cmdAplica 
      Caption         =   "Aplicar Cierre de Mes"
      Height          =   495
      Left            =   3240
      TabIndex        =   2
      Top             =   2280
      Width           =   1215
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancelar"
      Height          =   495
      Left            =   3240
      TabIndex        =   0
      Top             =   240
      Width           =   1215
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      Left            =   120
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   1200
      Width           =   2055
   End
   Begin VB.Label Label1 
      BackColor       =   &H00B39665&
      Caption         =   "Que mes desea Cerrar ?"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   960
      Width           =   2055
   End
End
Attribute VB_Name = "CierreMes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

'msConn.Execute "UPDATE MESES SET " & cPrevMes & " = True, FECHA = '" & Format(Date, "YYYYMMDD") & "',HORA = '" & Format(Time, "HH:MM") & "', USUARIO = " & rs!NUMERO
Option Base 1
Private aMes() As String
Private nPrevMes As Integer
Private nUserNumber As Integer
Private msSolo As New ADODB.Connection
Private Declare Function GetWindowsDirectory Lib "kernel32" Alias "GetWindowsDirectoryA" (ByVal lpBuffer As String, ByVal nSize As Long) As Long
Private Declare Function SetWindowPos Lib "user32" _
         (ByVal hwnd As Long, _
         ByVal hWndInsertAfter As Long, _
         ByVal x As Long, _
         ByVal y As Long, _
         ByVal cx As Long, _
         ByVal cy As Long, _
         ByVal wFlags As Long) As Long
      Const HWND_TOPMOST = -1
      Const SWP_NOMOVE = &H2
      Const SWP_NOSIZE = &H1

Private Sub MuestraTitulo()
On Error Resume Next
    AutoRedraw = True
    ScaleMode = vbPixels

    With Font
        .Name = "Times New Roman"
        .Bold = True
        .Size = 20
    End With

    ForeColor = vbBlack
    CurrentX = 10
    CurrentY = 20
    Print "Cierre de Mes"

    ForeColor = vbWhite
    CurrentX = 10 - 3
    CurrentY = 20 - 3
    Print "Cierre de Mes"
On Error GoTo 0
End Sub

Private Sub cmdAplica_Click()
'INFO: GUARDA LOS VALORES DE EXIST1,EXIST2,COSTO,COSTO_EMPAQUE
' PARA EL MES SELECCIONADO, LA DESVENTAJA ES QUE LOS DATOS LO
' TOMA DE LA TABLA INVENT Y SON DATOS ACTUALES, LO QUE SIGNIFICA
' QUE SI CORREN LA ACTUALIZACION EL 5 DE NOVIEMBRE, SON LOS DATOS
' HASTA ESE DIA LOS QUE SE ACTUALIZAN Y SI CUADRAN LAS COMPRAS AL
' MES ANTERIOR, LOS DATOS QUE SE GUARDAN YA LLEVAN 4 DIAS DE VENTAS
' ASI QUE NO ES UN DATO REAL.

Dim rsInventario As New ADODB.Recordset
Dim rsInventHistorico As New ADODB.Recordset
Dim cSQL As String
Dim vResp
Dim i As Integer, iUpdate As Integer
Dim rsMeses As New ADODB.Recordset
Dim cLog As String

vResp = MsgBox("¿ Realmente desea aplicar el Cierre de " & Combo1.Text & " ?", vbYesNoCancel, "Cierre de Mes")
If vResp = vbYes Then
    cSQL = "SELECT ID,EXIST1,EXIST2,COSTO,COSTO_EMPAQUE "
    cSQL = cSQL & " FROM INVENT ORDER BY ID"
    
    rsInventario.Open cSQL, msSolo, adOpenStatic
    On Error Resume Next
    rsInventario.MoveFirst
    On Error GoTo 0
    
    rsInventHistorico.Open "SELECT * FROM HIST_INVENT ORDER BY ID", msSolo, adOpenKeyset, adLockOptimistic
    On Error Resume Next
    ProgBar.Max = rsInventario.RecordCount
    On Error GoTo 0
    ProgBar.Value = i + 1
    
    Do While Not rsInventario.EOF
        rsInventHistorico.Find "ID = " & rsInventario!ID
        
        If rsInventHistorico.EOF Then
            'AGREGA UN PRODUCTO NUEVO
            For i = 0 To rsInventHistorico.Fields.Count - 1
                If Right(rsInventHistorico.Fields(i).Name, 2) = Format(nPrevMes + 1, "00") Then
                    rsInventHistorico.AddNew
                    rsInventHistorico!ID = rsInventario!ID
                    rsInventHistorico.Fields(i).Value = rsInventario!EXIST1
                    rsInventHistorico.Fields(i + 1).Value = rsInventario!EXIST2
                    rsInventHistorico.Fields(i + 2).Value = rsInventario!COSTO
                    rsInventHistorico.Fields(i + 3).Value = rsInventario!COSTO_EMPAQUE
                    rsInventHistorico.Update
                    Exit For
                End If
            Next
        Else
            'ACTUALIZA LOS DATOS YA MARCADOS
            For i = 0 To rsInventHistorico.Fields.Count - 1
                If Right(rsInventHistorico.Fields(i).Name, 2) = Format(nPrevMes + 1, "00") Then
                    rsInventHistorico.Fields(i).Value = rsInventario!EXIST1
                    rsInventHistorico.Fields(i + 1).Value = rsInventario!EXIST2
                    rsInventHistorico.Fields(i + 2).Value = rsInventario!COSTO
                    rsInventHistorico.Fields(i + 3).Value = rsInventario!COSTO_EMPAQUE
                    rsInventHistorico.Update
                    Exit For
                End If
            Next
        End If
        iUpdate = iUpdate + 1
        ProgBar.Value = iUpdate
        rsInventario.MoveNext
    Loop
    rsInventario.Close
    rsInventHistorico.Close
    Set rsInventario = Nothing
    Set rsInventHistorico = Nothing
    
    rsMeses.Open "SELECT * FROM MESES", msSolo, adOpenKeyset, adLockOptimistic
''
''    For i = 0 To rsMeses.Fields.Count - 1
''        cLog = cLog & rsMeses.Fields(i).Value & "|"
''    Next
''    EscribeLog "VALOR ANTERIOR CIERRE MES : " & cLog
    
    If rsMeses.EOF Then
        EscribeLog "NO EXISTEN DATOS DE CIERRE DE MES, SE CREARA REGISTRO PARA: " & Format(Date, "YYYYMMDD")
        cSQL = "INSERT INTO MESES VALUES"
        cSQL = cSQL & "('" & Format(Date, "YYYYMMDD") & "','"
        cSQL = cSQL & Format(Time, "HH:MM") & "',1967,"
        cSQL = cSQL & "FALSE,FALSE,FALSE,FALSE,FALSE,FALSE,FALSE,FALSE,FALSE,FALSE,FALSE,FALSE)"
        rsMeses.Close
        
        msSolo.BeginTrans
        msSolo.Execute cSQL
        msSolo.CommitTrans
        
        rsMeses.Open "SELECT * FROM MESES", msSolo, adOpenKeyset, adLockOptimistic
    End If
    
    For i = 0 To rsMeses.Fields.Count - 1
        If Right(rsMeses.Fields(i).Name, 2) = Format(nPrevMes + 1, "00") Then
            rsMeses.Fields(i).Value = True
            rsMeses.Fields(0).Value = Format(Date, "YYYYMMDD")
            rsMeses.Fields(1).Value = Format(Time, "HH:MM")
            rsMeses.Fields(2).Value = nUserNumber
            rsMeses.Update
            Exit For
        End If
    Next
    
'''    For i = 0 To rsMeses.Fields.Count - 1
'''        cLog = cLog & rsMeses.Fields(i).Value & "|"
'''    Next
'''    EscribeLog "CIERRE MES : " & cLog
    
'''    If nPrevMes + 1 = "12" Then
'''        'INFO: SE ESTA CERRANDO DICIEMBRE
'''        rsMeses!FECHA = Space(8)
'''        rsMeses!HORA = Space(5)
'''        rsMeses!USUARIO = 0
'''        For i = 3 To rsMeses.Fields.Count - 1
'''            rsMeses.Fields(i).Value = False
'''        Next
'''        rsMeses.Update
'''    End If
    
    EscribeLog ("CIERRE DE MES.SE HA REALIZADO LA ACTUALIZACION HISTORICA DEL INVENTARIO")
    MsgBox "SE HA REALIZADO LA ACTUALIZACION HISTORICA DEL INVENTARIO", vbInformation, "Actualización Histórica"
    cmdCancel.Caption = "Cerrar"
    cmdAplica.Enabled = False
End If
End Sub

Private Sub cmdCancel_Click()
msSolo.Close
Set msSolo = Nothing
Unload Me
End Sub

Private Sub Combo1_Click()
nPrevMes = Combo1.ListIndex
End Sub

Private Sub Form_Load()
Dim i As Integer
Dim cOpen As String
Dim cTemp As String
Dim lngWindowPosition As Long
'Show
'******************
lngWindowPosition = SetWindowPos(CierreMes.hwnd, _
                                 HWND_TOPMOST, _
                                 0, _
                                 0, _
                                 0, _
                                 0, _
                                 SWP_NOMOVE Or SWP_NOSIZE)

On Error GoTo ErrAdm:
'INFO: GET PARAMETERS
cTemp = Command
cOpen = GetFromINI("General", "ProveedorDatos", App.Path & "\soloini.ini")
cOpen = cOpen & ";Jet OLEDB:Database Password=master24"
npos = InStr(1, cTemp, "¦")
nUserNumber = Val(Mid(cTemp, npos + 1, 4))
'cTemp = Left(cTemp, npos - 1)
nUserNumber = 1967

Call MuestraTitulo

ReDim aMes(12)
aMes(1) = "ENERO"
aMes(2) = "FEBRERO"
aMes(3) = "MARZO"
aMes(4) = "ABRIL"
aMes(5) = "MAYO"
aMes(6) = "JUNIO"
aMes(7) = "JULIO"
aMes(8) = "AGOSTO"
aMes(9) = "SEPTIEMBRE"
aMes(10) = "OCTUBRE"
aMes(11) = "NOVIEMBRE"
aMes(12) = "DICIEMBRE"

For i = 1 To 12
    Combo1.AddItem aMes(i)
Next

'muestra el mes que se desea cerrar.
'SI ES ENERO, MUESTRA DICIEMBRE
nPrevMes = Month(Date) - 2
Combo1.ListIndex = nPrevMes
If Combo1.Text = "" Then
    Combo1.ListIndex = 11
    nPrevMes = 11
End If
msSolo.Open cOpen
'msSolo.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\access\SOLO.mdb;;Jet OLEDB:Database Password=master24"
'cOpen
'Show 1
On Error GoTo 0
Exit Sub

ErrAdm:
    MsgBox Err.Number & " - " & Err.Description & vbCrLf & _
        "No es posible realizar el cierre en este momento. Contacte a Solo Software Development", vbCritical, "Imposible realizar Cierre de Mes"
    Unload Me
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
Set msSolo = Nothing
End Sub


Private Function WindowsDirectory() As String
    ' Retrieve the Windows directory.
    Dim strBuffer As String
    Dim lngLen As Long
    strBuffer = Space(dhcMaxPath)
    lngLen = dhcMaxPath
    lngLen = GetWindowsDirectory(strBuffer, lngLen)
    ' If the path is longer than dhcMaxPath, then
    ' lngLen contains the correct length. Resize the
    ' buffer and try again.
    If lngLen > dhcMaxPath Then
        strBuffer = Space(lngLen)
        lngLen = GetWindowsDirectory(strBuffer, lngLen)
    End If
    WindowsDirectory = Left$(strBuffer, lngLen)

End Function

