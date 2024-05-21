VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "Mscomctl.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form ConMes 
   BackColor       =   &H00B39665&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "CONSULTA DE MESEROS DEL SISTEMA"
   ClientHeight    =   7440
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10845
   Icon            =   "ConMes.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7440
   ScaleWidth      =   10845
   StartUpPosition =   2  'CenterScreen
   Begin VB.OptionButton opcOrden 
      BackColor       =   &H00B39665&
      Caption         =   "Ordenar por Ventas"
      Height          =   375
      Index           =   3
      Left            =   8040
      TabIndex        =   17
      Top             =   840
      Width           =   2415
   End
   Begin VB.OptionButton opcOrden 
      BackColor       =   &H00B39665&
      Caption         =   "Ordenar por Transacciones"
      Height          =   375
      Index           =   2
      Left            =   8040
      TabIndex        =   16
      Top             =   600
      Width           =   2415
   End
   Begin VB.OptionButton opcOrden 
      BackColor       =   &H00B39665&
      Caption         =   "Ordenar por Nombre"
      Height          =   375
      Index           =   1
      Left            =   8040
      TabIndex        =   15
      Top             =   360
      Width           =   2175
   End
   Begin VB.CommandButton cmdPrintTrans 
      BackColor       =   &H00FFC0FF&
      Height          =   495
      Left            =   9840
      Picture         =   "ConMes.frx":0442
      Style           =   1  'Graphical
      TabIndex        =   12
      ToolTipText     =   "Transacciones de los Meseros"
      Top             =   3360
      Width           =   855
   End
   Begin VB.CommandButton cmdPrint 
      BackColor       =   &H0000FFFF&
      Height          =   495
      Left            =   10080
      Picture         =   "ConMes.frx":0884
      Style           =   1  'Graphical
      TabIndex        =   11
      ToolTipText     =   "Envia Seleccion a la Impresora"
      Top             =   3960
      Width           =   615
   End
   Begin VB.CommandButton cmdEjec 
      Caption         =   "&Ejecutar Consulta"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   520
      Left            =   5640
      TabIndex        =   2
      Top             =   240
      Width           =   1695
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Sa&lir"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   9960
      TabIndex        =   4
      Top             =   6600
      Width           =   855
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFMes 
      Height          =   5895
      Left            =   120
      TabIndex        =   3
      Top             =   1320
      Width           =   5295
      _ExtentX        =   9340
      _ExtentY        =   10398
      _Version        =   393216
      FixedCols       =   0
      SelectionMode   =   1
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
   End
   Begin MSComCtl2.DTPicker txtFecIni 
      Height          =   345
      Left            =   1320
      TabIndex        =   0
      Top             =   360
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   609
      _Version        =   393216
      Format          =   162136065
      CurrentDate     =   36418
   End
   Begin MSComCtl2.DTPicker txtFecFin 
      Height          =   345
      Left            =   3840
      TabIndex        =   1
      Top             =   360
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   609
      _Version        =   393216
      Format          =   162136065
      CurrentDate     =   36418
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFTP 
      Height          =   1935
      Left            =   5640
      TabIndex        =   5
      Top             =   1320
      Width           =   5055
      _ExtentX        =   8916
      _ExtentY        =   3413
      _Version        =   393216
      FixedCols       =   0
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
   End
   Begin MSComctlLib.ProgressBar ProgBar 
      Height          =   175
      Left            =   120
      TabIndex        =   10
      Top             =   840
      Width           =   5295
      _ExtentX        =   9340
      _ExtentY        =   318
      _Version        =   393216
      Appearance      =   1
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid List1 
      Height          =   3855
      Left            =   5640
      TabIndex        =   13
      Top             =   3360
      Width           =   4095
      _ExtentX        =   7223
      _ExtentY        =   6800
      _Version        =   393216
      Cols            =   3
      FixedCols       =   0
      FocusRect       =   0
      HighLight       =   0
      ScrollBars      =   2
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty FontFixed {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _NumberOfBands  =   1
      _Band(0).Cols   =   3
      _Band(0).GridLinesBand=   0
      _Band(0).TextStyleBand=   0
      _Band(0).TextStyleHeader=   0
   End
   Begin VB.OptionButton opcOrden 
      BackColor       =   &H00B39665&
      Caption         =   "Ordenar por Numero"
      Height          =   375
      Index           =   0
      Left            =   8040
      TabIndex        =   14
      Top             =   120
      Width           =   2175
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H0000FFFF&
      BorderWidth     =   3
      Height          =   1180
      Left            =   7920
      Top             =   90
      Width           =   2655
   End
   Begin VB.Label Label1 
      BackColor       =   &H00B39665&
      Caption         =   "Fecha Inicial"
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
      TabIndex        =   9
      Top             =   480
      Width           =   1215
   End
   Begin VB.Label Label1 
      BackColor       =   &H00B39665&
      Caption         =   "Fecha Final"
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
      Left            =   2760
      TabIndex        =   8
      Top             =   480
      Width           =   1095
   End
   Begin VB.Label Label1 
      BackColor       =   &H00B39665&
      Caption         =   "Propinas Marcadas"
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
      Left            =   5640
      TabIndex        =   7
      Top             =   1080
      Width           =   2295
   End
   Begin VB.Label Label1 
      BackColor       =   &H00B39665&
      Caption         =   "Seleccione Mesero"
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
      TabIndex        =   6
      Top             =   1080
      Width           =   1695
   End
End
Attribute VB_Name = "ConMes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private rsConMeseros As Recordset
Private rsConTP As Recordset
Private nConMes As Integer
Private dF1 As String
Private dF2 As String
Private Type tLinea
   Descripcion As String * 15
   Cantidad As String * 5
   precio As String * 6
End Type
'INFO: ADD IMPRESION 14/9/2004
Private nPagina As Integer
Private iLin As Integer

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
MainMant.spDoc.TextOut 300, 350, rs00!DESCRIP
MainMant.spDoc.TextOut 300, 450, ConMes.Caption

''''If ConDept.opcTipo(0).Value = True Then
    MainMant.spDoc.TextOut 300, 550, "PERIODO : Desde " & txtFecIni & " Hasta " & txtFecFin
''''Else
''''    MainMant.spDoc.TextOut 300, 550, "PERIODO : REPORTE Z # " & LVZ.SelectedItem.Text
''''End If

MainMant.spDoc.TextOut 300, 650, "NUMERO"
MainMant.spDoc.TextOut 500, 650, "NOMBRE"
MainMant.spDoc.TextOut 1200, 650, "Trans"
MainMant.spDoc.TextOut 1450, 650, "Ventas"
MainMant.spDoc.TextOut 300, 700, "---------------------------------------------------------------------------------------------------------------"

iLin = 750
nPagina = nPagina + 1
End Sub

Private Sub cmdEjec_Click()
Dim nSeleccion As Integer

   On Error GoTo cmdEjec_Click_Error

dF1 = Format(txtFecIni, "YYYYMMDD")
dF2 = Format(txtFecFin, "YYYYMMDD")

Me.MousePointer = vbHourglass
ProgBar.value = 10

cSQL = "SELECT A.MESERO AS MES, B.NOMBRE, B.APELLIDO, "
cSQL = cSQL & " COUNT(A.NUM_TRANS) AS TRANS, "
''''cSQL = cSQL & " FORMAT(SUM(A.PRECIO),'STANDARD') AS VENTAS "
cSQL = cSQL & " SUM(A.PRECIO) AS VENTAS "
cSQL = cSQL & " INTO LOLO "
cSQL = cSQL & " FROM HIST_TR AS A LEFT JOIN MESEROS AS B "
cSQL = cSQL & " ON A.MESERO = B.NUMERO "
cSQL = cSQL & " WHERE A.FECHA BETWEEN '" & dF1 & "'"
cSQL = cSQL & " AND '" & dF2 & "'"
'INFO: 16FEB2013  / SE INCLUYE INFO PARA QUE CUADRE CON EL RESTO DE LOS REPORTES
'MUESTRA LAS VENTAS
cSQL = cSQL & " AND '%' NOT IN (A.DESCRIP) "
cSQL = cSQL & " AND A.DESCRIP NOT LIKE '%DESCUENTO%' "
cSQL = cSQL & " AND A.DESCRIP NOT LIKE  '%@@%' "
cSQL = cSQL & " GROUP BY A.MESERO, B.NOMBRE, B.APELLIDO "

msConn.BeginTrans
msConn.Execute cSQL
ProgBar.value = 20
msConn.CommitTrans

ProgBar.value = 30

cSQL = "SELECT A.MES,A.NOMBRE,A.APELLIDO, "
'cSQL = cSQL & " COUNT(A.TRANS) AS TRANS, "
'INFO: ANTES ESTABA COUNT, AHORA SE ESTA CORRIGIENDO POR MAX
''''cSQL = cSQL & " MAX(A.TRANS) AS TRANS, "
cSQL = cSQL & " A.TRANS, FORMAT(A.VENTAS,'STANDARD') AS MONTO, A.VENTAS "
cSQL = cSQL & " FROM LOLO AS A "
''''cSQL = cSQL & " GROUP BY A.MES,A.NOMBRE,A.APELLIDO "

nSeleccion = GetOption(ConMes.opcOrden())
Select Case nSeleccion
    Case 0
        cSQL = cSQL & " ORDER BY A.MES"
    Case 1
        cSQL = cSQL & " ORDER BY A.NOMBRE, A.APELLIDO"
    Case 2
        cSQL = cSQL & " ORDER BY 4 DESC"
    Case 3
        cSQL = cSQL & " ORDER BY 6 DESC"
    Case Else
        cSQL = cSQL & " ORDER BY A.MES"
End Select

rsConMeseros.Open cSQL, msConn, adOpenStatic, adLockOptimistic
ProgBar.value = 40
Set MSHFMes.DataSource = rsConMeseros

rsConMeseros.Close
With MSHFMes
    .ColWidth(0) = 500: .ColWidth(1) = 1300: .ColWidth(2) = 1300:
    .ColWidth(3) = 700: .ColWidth(4) = 1200: .ColWidth(5) = 0:
    .ColAlignment(4) = flexAlignRightCenter
End With

ProgBar.value = 50
msConn.BeginTrans
msConn.Execute "DROP TABLE LOLO"
msConn.CommitTrans
ProgBar.value = 60
MSHFMes_EnterCell

Call Seguridad

   On Error GoTo 0
   Exit Sub

cmdEjec_Click_Error:

    ShowMsg "Error " & Err.Number & " (" & Err.Description & ") in procedure cmdEjec_Click of Form ConMes", vbYellow, vbRed
    msConn.RollbackTrans
    Me.MousePointer = vbDefault
End Sub

Private Sub cmdPrint_Click()
''''MsgBox "NO ESTA DISPONIBLE", vbInformation, BoxTit
Dim iCtr As Integer 'Contador de Linea
Dim iCol, iFil As Integer 'Contador de Columnas
Dim cText As String
Dim ispace As Integer
Dim iLen As Integer
Dim sSubTot As Double       ''INFO: 16FEB2013

sSubTot = 0#

On Error GoTo ErrorPrn:
nPagina = 0

EscribeLog ("Admin." & "Impresión de Ventas por Mesero: " & ConMes.Caption & " PERIODO : Desde " & txtFecIni & " Hasta " & txtFecFin)
MainMant.spDoc.DocBegin
PrintTit    'Rutina de Titulos

'-ProgBar.Value = 10

For iFil = 0 To MSHFMes.Rows - 1
    If iLin > 2400 Then PrintTit
    'If ProgBar.Value < 100 Then
        '-ProgBar.Value = '-ProgBar.Value + 10
    'End If
    MSHFMes.Row = iFil
    For iCol = 0 To MSHFMes.Cols - 1
        Select Case iCol
            Case 0, 1, 2, 3, 4, 5
                MSHFMes.Col = iCol
                MainMant.spDoc.TextAlign = SPTA_LEFT
                If iFil = 0 Then
                Else
                    If IsNumeric(MSHFMes.Text) Then ispace = 10 Else ispace = 25
                    Select Case iCol
                    Case 0
                        MainMant.spDoc.TextAlign = SPTA_LEFT
                        MainMant.spDoc.TextOut 300, iLin, MSHFMes.Text
                    Case 1
                        MainMant.spDoc.TextAlign = SPTA_LEFT
                        MainMant.spDoc.TextOut 500, iLin, FormatTexto(MSHFMes.Text, ispace)
                    Case 2
                        MainMant.spDoc.TextAlign = SPTA_LEFT
                        MainMant.spDoc.TextOut 800, iLin, FormatTexto(MSHFMes.Text, ispace)
                    Case 3
                        MainMant.spDoc.TextAlign = SPTA_RIGHT
                        MainMant.spDoc.TextOut 1300, iLin, Format(MSHFMes.Text, "General Number")
                    Case 4
                        MainMant.spDoc.TextAlign = SPTA_RIGHT
                        MainMant.spDoc.TextOut 1600, iLin, Format(MSHFMes.Text, "STANDARD")
                    Case 5
                        sSubTot = sSubTot + Format(MSHFMes.Text, "standard")
                    End Select
                End If
            End Select
    Next
    iLin = iLin + 50
    '-If ProgBar.Value < 100 Then '-ProgBar.Value = '-ProgBar.Value + 5
Next

'-ProgBar.Value = 100

iLin = iLin + 100
MainMant.spDoc.TextAlign = SPTA_LEFT
MainMant.spDoc.TextOut 500, iLin, "Sub Total Ventas del Periodo : " & Format(sSubTot, "Currency")
'MainMant.spDoc.TextOut 500, iLin, "Sub Total NETO del Periodo : " & Format(sSubTot, "Currency")

MainMant.spDoc.DoPrintPreview

On Error GoTo 0

Call Seguridad

Exit Sub

100:
Exit Sub
ErrorPrn:
    'MsgBox "¡ Ocurre algún Error con la Impresora, Intente Conecterla !", vbExclamation, BoxTit
    ShowMsg "¡ Ocurre algún Error con la Impresora, Intente Conecterla !", vbRed, vbYellow
    Resume 100

End Sub

Private Sub cmdPrintTrans_Click()
'Dim aTrans() As String
Dim vFecha As String
Dim cFecha As Variant
Dim cHora As Variant
Dim cSQL As String
Dim rsT As New ADODB.Recordset
Dim rsPag As New ADODB.Recordset
Dim nTrans As Long
Dim nSubTot As Double       ''INFO: 16FEB2013
Dim tReglon As tLinea

List1.Clear
List1.Rows = 1
List1.Refresh
List1.ColWidth(0) = 2400
List1.ColWidth(1) = 400
List1.ColWidth(2) = 900
List1.ColAlignment(0) = flexAlignLeftCenter
List1.ColAlignment(1) = flexAlignRightCenter
List1.ColAlignment(2) = flexAlignRightCenter
'vFecha = Format(Date, "SHORT DATE")
vFecha = ConMes.txtFecIni
vFecha = InputBox("Escriba la fecha que desea Imprimir (dd/mm/aaaa)" & vbCrLf & _
        "(Unicamente del archivo histórico)", "Impresión de Transacciones de Meseros", vFecha)
vFecha = Right(vFecha, 4) & Mid(vFecha, 4, 2) & Left(vFecha, 2)
cSQL = "SELECT A.* "
cSQL = cSQL & " FROM HIST_TR AS A "
cSQL = cSQL & " WHERE A.MESERO = " & nConMes
cSQL = cSQL & " AND A.FECHA = '" & vFecha & "'"
cSQL = cSQL & " ORDER BY A.NUM_TRANS,A.LIN"
rsT.Open cSQL, msConn, adOpenStatic, adLockOptimistic
'Debug.Print cSQL
'Imprime las transacciones de los meseros
Do Until rsT.EOF
    nTrans = rsT!NUM_TRANS
    cFecha = Right(rsT!FECHA, 2) & "/" & Mid(rsT!FECHA, 5, 2) & "/" & Left(rsT!FECHA, 4)
    cHora = Format(rsT!HORA, "HH:MM")
    List1.AddItem "============================="
    List1.AddItem "TRANS# " & nTrans
    List1.AddItem Space(1)
    List1.AddItem "Mesa : " & rsT!MESA
    List1.AddItem "-----------------------------"
    Do While rsT!NUM_TRANS = nTrans
        tReglon.Descripcion = rsT!DESCRIP
        tReglon.Cantidad = rsT!CANT
        tReglon.precio = rsT!precio
        List1.AddItem tReglon.Descripcion & vbTab & Format(tReglon.Cantidad, "####") & vbTab & Format(tReglon.precio, "####.00")
        nSubTot = nSubTot + rsT!precio
        rsT.MoveNext
        If rsT.EOF Then Exit Do
    Loop
    List1.AddItem "-----------------------------"
    List1.AddItem "     Sub-Total :" & vbTab & vbTab & Format(nSubTot, "CURRENCY")
    List1.AddItem "     ITBMS     :" & vbTab & vbTab & Format((nSubTot * iISC), "STANDARD")
    nSubTot = 0
    
    cSQL = "SELECT B.DESCRIP, A.LIN, A.MONTO "
    cSQL = cSQL & " FROM HIST_TR_PAGO AS A, PAGOS AS B "
    cSQL = cSQL & " WHERE A.NUM_TRANS = " & nTrans
    cSQL = cSQL & " AND A.TIPO_PAGO = B.CODIGO "
    rsPag.Open cSQL, msConn, adOpenKeyset, adLockOptimistic
    
    Do While Not rsPag.EOF
        List1.AddItem rsPag!DESCRIP & vbTab & vbTab & Format(rsPag!MONTO, "####.00")
        rsPag.MoveNext
    Loop
    List1.AddItem Space(1)
    rsPag.Close
    List1.AddItem "FEC: " & cFecha & Space(2) & "HORA: " & cHora
    List1.AddItem Space(1)
Loop
rsT.Close
Set rsPag = Nothing
Set rsT = Nothing

Call Seguridad

End Sub

Private Sub Command1_Click()
Unload Me
End Sub

Private Sub Form_Load()

Set rsConMeseros = New Recordset
Set rsConTP = New Recordset

txtFecIni = Format(Date, "SHORT DATE")
txtFecFin = Format(Date, "SHORT DATE")

Call Seguridad

End Sub
Private Function Seguridad() As String
'SETUP DE SEGURIDAD DEL SISTEMA
Dim cSeguridad As String

cSeguridad = GetSecuritySetting(npNumCaj, Me.Name)
Select Case cSeguridad
    Case "CEMV"        'Crear - Eliminar - Modificar - Ver
        'INFO: NO HAY RESTRICCIONES
    Case "CMV"        'Crear - Modificar - Ver"
        'INFO: NO HAY RESTRICCIONES
    Case "CV"        'Crear - Ver
        'INFO: NO HAY RESTRICCIONES
    Case "V"        'Ver solamente
        cmdPrintTrans.Enabled = False: cmdPrint.Enabled = False
    Case "N"        'SIN DERECHOS
        txtFecIni.Enabled = False: txtFecFin.Enabled = False: cmdEjec.Enabled = False
        MSHFMes.Enabled = False: MSHFTP.Enabled = False
        cmdPrintTrans.Enabled = False: cmdPrint.Enabled = False
End Select
End Function

''''Private Sub List1_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
''''Dim vResp As Variant
''''If Button = 2 Then
''''    vResp = MsgBox("¿ Desea enviar el Reporte (Z) a la imrpresora ?", vbQuestion + vbYesNo, "Impresión de Reporte")
''''    Call Print2File
''''End If
''''End Sub
''''Private Sub Print2File()
''''Dim FACTURA_FILE As String
''''Dim DATA_PATH  As String
''''Dim cDataPath As String
''''Dim cCadena As String
''''
''''FACTURA_FILE = "SOLOFACT.TXT"
''''
''''Open App.Path & "\OrigenDB.txt" For Input As #1
''''Do Until EOF(1)
''''    Line Input #1, a$
''''    If Left(a$, 1) = "*" Then
''''        DATA_PATH = Mid(a$, 3, Len(a$) - 2)
''''    Else
''''        cDataPath = a$
''''    End If
''''Loop
''''Close #1
''''
''''cFactFile = DATA_PATH + FACTURA_FILE
''''
''''Dim i As Integer
''''Open cFactFile For Output As #2
''''For i = 0 To List1.Rows - 1
''''    List1.Row = i
''''    List1.Col = 0
''''    cCadena = List1.Text
''''    List1.Col = 1
''''    cCadena = cCadena & Space(3) & List1.Text
''''    List1.Col = 2
''''    cCadena = cCadena & Format(List1.Text, "@@@@@@@@@@")
''''    Print #2, cCadena
''''Next
''''Close #2
''''End Sub

Private Sub MSHFMes_EnterCell()
nConMes = Val(MSHFMes.Text)

MSHFTP.Clear: MSHFTP.Refresh
ProgBar.value = 50
'--------------------------------------------------

'SELECT NUM_TRANS,TIPO_PAGO,MONTO FROM HIST_TR_PROP WHERE MESERO = 804 AND NUM_tRANS IN (SELECT DISTINCT NUM_tRANS FROM HIST_tR WHERE FECHA BETWEEN '20000201' AND '20000205')
cSQL = "SELECT NUM_TRANS,TIPO_PAGO,MONTO "
cSQL = cSQL & " INTO LOLO "
cSQL = cSQL & " FROM HIST_TR_PROP "
cSQL = cSQL & " WHERE MESERO = " & nConMes
cSQL = cSQL & " AND NUM_TRANS IN "
cSQL = cSQL & "(SELECT DISTINCT NUM_TRANS FROM HIST_TR "
cSQL = cSQL & " WHERE FECHA BETWEEN '" & dF1 & "' AND '" & dF2 & "')"

msConn.BeginTrans
msConn.Execute cSQL
ProgBar.value = 80
msConn.CommitTrans

cSQL = "SELECT A.TIPO_PAGO,C.DESCRIP,COUNT(A.TIPO_PAGO) AS CANT, "
cSQL = cSQL & " format(SUM(A.MONTO),'standard') AS COBRADO "
cSQL = cSQL & " FROM LOLO AS A LEFT JOIN PAGOS AS C "
cSQL = cSQL & " ON A.TIPO_PAGO = C.CODIGO "
cSQL = cSQL & " GROUP BY A.TIPO_PAGO,C.DESCRIP "

rsConTP.Open cSQL, msConn, adOpenStatic, adLockOptimistic
ProgBar.value = 90
Set MSHFTP.DataSource = rsConTP
rsConTP.Close

With MSHFTP
    .ColWidth(0) = 0: .ColWidth(1) = 1900: .ColWidth(2) = 600:
    .ColWidth(3) = 1400:
    .ColAlignment(3) = flexAlignRightCenter
End With

msConn.BeginTrans
msConn.Execute "DROP TABLE LOLO"
msConn.CommitTrans

ProgBar.value = 0
Me.MousePointer = vbDefault
End Sub
Private Sub txtFecFin_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then cmdEjec.SetFocus
End Sub

Private Sub txtFecIni_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then txtFecFin.SetFocus
End Sub

