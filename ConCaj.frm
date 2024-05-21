VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form ConCaj 
   BackColor       =   &H00B39665&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "CONSULTA DE CAJEROS DEL SISTEMA"
   ClientHeight    =   7785
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10485
   Icon            =   "ConCaj.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7785
   ScaleWidth      =   10485
   StartUpPosition =   2  'CenterScreen
   Begin VB.OptionButton opcTipo 
      BackColor       =   &H00B39665&
      Caption         =   "Por Fecha"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   0
      Left            =   2880
      TabIndex        =   17
      ToolTipText     =   "Obtener Ventas por Fechas Seleccionadas"
      Top             =   240
      Value           =   -1  'True
      Width           =   1215
   End
   Begin VB.OptionButton opcTipo 
      BackColor       =   &H00B39665&
      Caption         =   "Por Tipo"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   1
      Left            =   4320
      TabIndex        =   14
      ToolTipText     =   "Obtener Ventas por Reporte Z"
      Top             =   240
      Width           =   1095
   End
   Begin VB.CommandButton cmdPrint 
      BackColor       =   &H0000FFFF&
      Height          =   495
      Left            =   5640
      Picture         =   "ConCaj.frx":0442
      Style           =   1  'Graphical
      TabIndex        =   13
      ToolTipText     =   "Envia Seleccion a la Impresora"
      Top             =   7080
      Width           =   735
   End
   Begin MSComctlLib.ProgressBar ProgBar 
      Height          =   180
      Left            =   120
      TabIndex        =   12
      Top             =   1320
      Width           =   5295
      _ExtentX        =   9340
      _ExtentY        =   318
      _Version        =   393216
      Appearance      =   1
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
      Left            =   8640
      TabIndex        =   3
      Top             =   480
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
      Left            =   8640
      TabIndex        =   2
      Top             =   7080
      Width           =   1695
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFCaj 
      Height          =   5895
      Left            =   120
      TabIndex        =   4
      Top             =   1800
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
      Left            =   1440
      TabIndex        =   0
      Top             =   240
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   609
      _Version        =   393216
      Format          =   134348801
      CurrentDate     =   36431
   End
   Begin MSComCtl2.DTPicker txtFecFin 
      Height          =   345
      Left            =   1440
      TabIndex        =   1
      Top             =   720
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   609
      _Version        =   393216
      Format          =   134348801
      CurrentDate     =   36418
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFTP 
      Height          =   2895
      Left            =   5640
      TabIndex        =   5
      Top             =   1800
      Width           =   4575
      _ExtentX        =   8070
      _ExtentY        =   5106
      _Version        =   393216
      FixedCols       =   0
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFExtra 
      Height          =   1695
      Left            =   5640
      TabIndex        =   6
      Top             =   5040
      Width           =   3735
      _ExtentX        =   6588
      _ExtentY        =   2990
      _Version        =   393216
      FixedCols       =   0
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
   End
   Begin MSComctlLib.ListView LVZ 
      Height          =   855
      Left            =   5400
      TabIndex        =   15
      Top             =   255
      Width           =   2415
      _ExtentX        =   4260
      _ExtentY        =   1508
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      Checkboxes      =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   0
      NumItems        =   0
   End
   Begin VB.Shape Borde1 
      BorderColor     =   &H0000FFFF&
      BorderWidth     =   2
      Height          =   1095
      Index           =   1
      Left            =   4200
      Top             =   120
      Width           =   3735
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00B39665&
      Caption         =   "Selección Reporte Z"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   5
      Left            =   4200
      TabIndex        =   16
      Top             =   645
      Width           =   1215
   End
   Begin VB.Shape Borde1 
      BorderColor     =   &H0000FFFF&
      BorderWidth     =   2
      Height          =   1095
      Index           =   0
      Left            =   120
      Top             =   120
      Width           =   4095
   End
   Begin VB.Label Label1 
      BackColor       =   &H00B39665&
      Caption         =   "Correcciones/Anulaciones/Descuentos"
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
      Index           =   4
      Left            =   5640
      TabIndex        =   11
      Top             =   4800
      Width           =   3495
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
      Left            =   240
      TabIndex        =   10
      Top             =   360
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
      Left            =   240
      TabIndex        =   9
      Top             =   840
      Width           =   1095
   End
   Begin VB.Label Label1 
      BackColor       =   &H00B39665&
      Caption         =   "Tipos de Pago por Cajero"
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
      Index           =   2
      Left            =   5640
      TabIndex        =   8
      Top             =   1560
      Width           =   2295
   End
   Begin VB.Label Label1 
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
      Height          =   255
      Index           =   3
      Left            =   120
      TabIndex        =   7
      Top             =   1560
      Width           =   1695
   End
End
Attribute VB_Name = "ConCaj"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
''''Private Declare Function GetTickCount Lib "kernel32" () As Long

Private rsConCajeros As Recordset
Private rsConTP As Recordset
Private rsConExtras As Recordset 'EC/VOID/DESC/DESC.GLOB
Private nConCaj As Integer
Private dF1 As String
Private dF2 As String
Private Sub GetReportesZ()
Dim cSQL As String
Dim rsZetas As ADODB.Recordset
Dim iLinea As Integer

On Error GoTo ErrAdm:
Set rsZetas = New ADODB.Recordset

cSQL = "SELECT TOP 100 VAL(CONTADOR) AS CONTADOR, FECHA FROM Z_COUNTER ORDER BY VAL(CONTADOR) DESC "
'cSQL = "SELECT TOP 100 CONTADOR, FECHA FROM Z_COUNTER ORDER BY CONTADOR DESC "
rsZetas.Open cSQL, msConn, adOpenStatic, adLockOptimistic

LVZ.ListItems.Clear
LVZ.ColumnHeaders.Clear

LVZ.ColumnHeaders.Add , , "Z#"
LVZ.ColumnHeaders.Add , , "Fecha"

'LV.ColumnHeaders.Item(1).Alignment = lvwColumnRight
LVZ.ColumnHeaders.Item(1).Alignment = lvwColumnLeft
LVZ.ColumnHeaders.Item(2).Alignment = lvwColumnRight
LVZ.ColumnHeaders.Item(1).Width = 700
LVZ.ColumnHeaders.Item(2).Width = 1250
iLinea = 1
Do While Not rsZetas.EOF
    LVZ.ListItems.Add , , rsZetas!CONTADOR
    
    LVZ.ListItems.Item(iLinea).ListSubItems.Add , , GetFecha(rsZetas!FECHA)
    iLinea = iLinea + 1
    rsZetas.MoveNext
Loop
On Error GoTo 0
rsZetas.Close

ErrAdm:
Set rsZetas = Nothing
End Sub
Private Sub cmdEjec_Click()
Dim i As Byte
Dim jFalseCounter As Byte
Dim cSQL As String

'INFO: 27OCT2010
Dim nMinZ As Long   'REPORTE Z INICIAL
Dim nMaxZ As Long   'REPORTE Z FINAL
Dim cArrayZ() As String

On Error GoTo ErrAdm:

For i = 0 To opcTipo.Count - 1
    Select Case opcTipo(i).value
        Case True
            Exit For
        Case False
            jFalseCounter = jFalseCounter + 1
        Case Else
    End Select
    'If opcTipo(i).Value = True Then Exit For
Next

' if 2 = jFalseCounter then NONE selected, then DEFAULT TO DATE OPTION
If jFalseCounter = 2 Then i = 0

If i = 0 Then
    'NORMAL NORMAL POR FECHA
Else
    'REPORTE POR Z#
    'INFO: ABRIL2008
    'CHECK TO SEE IF A Z HAS BEEN SELECTED
    For iZZ = 1 To LVZ.ListItems.Count
        If LVZ.ListItems(iZZ).Checked = True Then
            cZetas = cZetas & LVZ.ListItems(iZZ).Text & "','"
        End If
    Next
    On Error Resume Next
    If cZetas = "" Then
        MsgBox "Debe seleccionar al menos un Reporte Z", vbInformation, "Seleccione un reporte (Z)"
        Exit Sub
    Else
        cZetas = "'" & Mid(cZetas, 1, Len(cZetas) - 2)
    End If
    cmdEjec.Tag = cZetas
    cArrayZ = Split(Replace(cZetas, "'", ""), ",")
    On Error GoTo 0
End If

dF1 = "'" & Format(txtFecIni, "YYYYMMDD") & "'"
dF2 = "'" & Format(txtFecFin, "YYYYMMDD") & "'"

Me.MousePointer = vbHourglass

ProgBar.value = 10
sqltxt = "SELECT A.CAJERO as CAJ,B.NOMBRE,B.APELLIDO,A.NUM_TRANS, "
sqltxt = sqltxt & " COUNT(A.NUM_TRANS) AS TRANS, "
sqltxt = sqltxt & " FORMAT(SUM(A.PRECIO),'STANDARD') AS VENTAS "
sqltxt = sqltxt & " INTO LOLO FROM HIST_TR AS A LEFT JOIN CAJEROS AS B "
sqltxt = sqltxt & " ON A.CAJERO = B.NUMERO "
If i = 0 Then
    sqltxt = sqltxt & " WHERE A.FECHA >= " & dF1
    sqltxt = sqltxt & " AND A.FECHA <= " & dF2
Else
    'sqltxt = sqltxt & " WHERE A.Z_COUNTER = '" & Val(LVZ.SelectedItem.Text) & "'"
    'INFO: VALIDAR LA Z CONTRA VALORES NUMERICOS EN VEZ DE TEXTO (27OCT2010)
    sqltxt = sqltxt & " WHERE VAL(A.Z_COUNTER) BETWEEN  " & GetNumber_FromArray(cArrayZ, GetMin)
    sqltxt = sqltxt & " AND " & GetNumber_FromArray(cArrayZ, Getmax)
End If
sqltxt = sqltxt & " GROUP BY A.CAJERO,B.NOMBRE,B.APELLIDO,A.NUM_TRANS "
sqltxt = sqltxt & " ORDER BY B.NOMBRE,B.APELLIDO "

msConn.BeginTrans
msConn.Execute sqltxt
    ProgBar.value = 20
msConn.CommitTrans
    ProgBar.value = 30

sqltxt = "SELECT A.CAJ,A.NOMBRE,A.APELLIDO, "
sqltxt = sqltxt & " COUNT(TRANS) AS TRANS, "
sqltxt = sqltxt & " FORMAT(SUM(A.VENTAS),'STANDARD') AS VENTAS "
sqltxt = sqltxt & " From LOLO AS A"
sqltxt = sqltxt & " GROUP BY A.CAJ,A.NOMBRE,A.APELLIDO "
sqltxt = sqltxt & " ORDER BY A.NOMBRE,A.APELLIDO "

ProgBar.value = 40
rsConCajeros.Open sqltxt, msConn, adOpenStatic, adLockOptimistic
ProgBar.value = 50
Set MSHFCaj.DataSource = rsConCajeros

rsConCajeros.Close
With MSHFCaj
    .ColWidth(0) = 400: .ColWidth(1) = 1300: .ColWidth(2) = 1300:
    .ColWidth(3) = 700: .ColWidth(4) = 1200:
    .ColAlignment(4) = flexAlignRightCenter
End With
Me.MousePointer = vbDefault

msConn.BeginTrans
msConn.Execute "DROP TABLE LOLO"
msConn.CommitTrans

MSHFCaj_EnterCell

On Error GoTo 0

Call Seguridad

Exit Sub

ErrAdm:
Me.MousePointer = vbDefault
If Err.Number = 91 Then
    EscribeLog ("Admin." & "ConCAJ.LA OPCION DE REPORTES POR REPORTE Z, NO ESTA HABILITADA")
    MsgBox "LA OPCION DE REPORTES POR REPORTE Z, NO ESTA HABILITADA", vbCritical, "Error en Reporte"
Else
    MsgBox Err.Number & " - " & Err.Description, vbCritical, "Error en Reporte"
End If
End Sub

Private Sub cmdPrint_Click()
BoxResp = MsgBox("¿ Desea Imprimir Reporte Financiero de Caja ?", vbQuestion + vbYesNo, "Reporte Finaciero del " & txtFecIni & " al " & txtFecFin)
If BoxResp = vbYes Then
    Dim rsEmpresaLoc As New ADODB.Recordset
    Dim txtsql As String
    Dim iLen As Integer
    Dim iLen1 As Integer
    Dim nCobrado As Double
    Dim nLin As Integer
    Dim iCounter As Long
    Dim cSQL As String
   
    ''''iCounter = GetTickCount
    
    nCobrado = 0
    ProgBar.Max = 100
    ProgBar.value = 15

    dF1 = "'" & Format(txtFecIni, "YYYYMMDD") & "'"
    dF2 = "'" & Format(txtFecFin, "YYYYMMDD") & "'"
    
    On Error Resume Next
    msConn.BeginTrans
    msConn.Execute "DROP TABLE LOLO"
    msConn.Execute "DROP TABLE LOLO1"
    msConn.CommitTrans
    On Error GoTo 0
    
    Me.MousePointer = vbHourglass

    If opcTipo(0).value = True Then
        cSQL = "SELECT DISTINCT NUM_TRANS "
        cSQL = cSQL & " INTO LOLO FROM HIST_TR AS A "
        cSQL = cSQL & " WHERE A.FECHA BETWEEN  " & dF1 & " AND " & dF2
    Else
        cSQL = "SELECT DISTINCT NUM_TRANS "
        cSQL = cSQL & " INTO LOLO FROM HIST_TR AS A"
        cSQL = cSQL & " WHERE A.Z_COUNTER = '" & Val(LVZ.SelectedItem.Text) & "'"
    End If
    msConn.Execute cSQL
    
    cSQL = "SELECT A.TIPO_PAGO, SUM(A.MONTO) AS MONTO"
    cSQL = cSQL & " INTO LOLO1 "
    cSQL = cSQL & " FROM HIST_TR_PAGO AS A, LOLO AS B"
    cSQL = cSQL & " WHERE A.NUM_TRANS = B.NUM_TRANS"
    cSQL = cSQL & " GROUP BY A.TIPO_PAGO;"
    msConn.Execute cSQL

    txtsql = "SELECT A.TIPO_PAGO,C.DESCRIP,SUM(A.MONTO) AS COBRADO "
    txtsql = txtsql & " FROM LOLO1 AS A LEFT JOIN PAGOS AS C "
    txtsql = txtsql & " ON A.TIPO_PAGO = C.CODIGO "
    txtsql = txtsql & " GROUP BY A.TIPO_PAGO,C.DESCRIP"
    txtsql = txtsql & " ORDER BY 3 DESC "
    
    ''''EscribeLog "Open SQL : " & GetTickCount - iCounter
    ''''EscribeLog "SQL: " & txtsql
    
    rsEmpresaLoc.Open txtsql, msConn, adOpenDynamic, adLockReadOnly
    ProgBar.value = 30
    If rsEmpresaLoc.EOF Then
        MsgBox "No hay Información Financiera para el Periodo Seleccionado", vbInformation, "Reporte Finaciero del " & txtFecIni & " al " & txtFecFin
        rsEmpresaLoc.Close
        Set rsEmpresaLoc = Nothing
        ProgBar.value = 0
        Me.MousePointer = vbDefault
        Exit Sub
    End If
    ProgBar.value = 45
    'Open "c:\TESTFILE" For Output As #1
    
    ''''EscribeLog "Begin Print : " & GetTickCount - iCounter
    MainMant.spDoc.DocBegin
    MainMant.spDoc.TextAlign = SPTA_LEFT
    
    MainMant.spDoc.WindowTitle = "Impresión de INFORME FINANCIERO"
    MainMant.spDoc.FirstPage = 1
    MainMant.spDoc.PageOrientation = SPOR_PORTRAIT
    MainMant.spDoc.Units = SPUN_LOMETRIC
    nPage = 1
    MainMant.spDoc.Page = nPage

    MainMant.spDoc.TextOut 300, 200, Format(Date, "long date") & "  " & Time
    MainMant.spDoc.TextOut 300, 250, rs00!DESCRIP
    MainMant.spDoc.TextOut 300, 300, "INFORME FINANCIERO"
    
    If opcTipo(0).value = True Then
        MainMant.spDoc.TextOut 300, 400, "Periodo del Informe: " & txtFecIni & "-" & txtFecFin
    Else
        MainMant.spDoc.TextOut 300, 500, "PERIODO : REPORTE Z # " & LVZ.SelectedItem.Text
    End If

    MainMant.spDoc.TextOut 300, 500, Space(15) & "DESCRIPCION                              MONTO   PORCENTAJE"
    MainMant.spDoc.TextOut 300, 550, Space(15) & "---------------------------------------------------------------------------------"
    nLin = 600
    ProgBar.value = 50
    Do Until rsEmpresaLoc.EOF
        nCobrado = nCobrado + rsEmpresaLoc!cobrado
        rsEmpresaLoc.MoveNext
    Loop
    rsEmpresaLoc.MoveFirst
    ProgBar.value = 60
    
    ''''EscribeLog "Print Report Detail : " & GetTickCount - iCounter
    Do Until rsEmpresaLoc.EOF
        If Len(Format(rsEmpresaLoc!cobrado, "#,###,###.00")) > 12 Then iLen = 0 Else iLen = Len(Format(rsEmpresaLoc!cobrado, "#,###,###.00"))
        If Len(Format(rsEmpresaLoc!cobrado / nCobrado, "##.00%")) > 6 Then iLen1 = 0 Else iLen1 = Len(Format(rsEmpresaLoc!cobrado / nCobrado, "##.00%"))
        On Error Resume Next
        'MainMant.spDoc.TextAlign = SPTA_LEFT + SPTA_TOP
        MainMant.spDoc.TextAlign = SPTA_LEFT
        MainMant.spDoc.TextOut 300, nLin, Space(15) & FormatTexto(rsEmpresaLoc!DESCRIP, 20)
        MainMant.spDoc.TextAlign = SPTA_RIGHT
        MainMant.spDoc.TextOut 1100, nLin, Format(rsEmpresaLoc!cobrado, "STANDARD")
        MainMant.spDoc.TextAlign = SPTA_RIGHT
        MainMant.spDoc.TextOut 1300, nLin, Format(rsEmpresaLoc!cobrado / nCobrado, "##.00%")
        MainMant.spDoc.TextAlign = SPTA_LEFT
        nLin = nLin + 50
        rsEmpresaLoc.MoveNext
    Loop
    
    ''''EscribeLog "END Report Detail : " & GetTickCount - iCounter
    
    nLin = nLin + 100
    MainMant.spDoc.TextOut 300, nLin, "Total Cobrado en el Periodo : " & Format(nCobrado, "CURRENCY")
    ProgBar.value = 70
    rsEmpresaLoc.Close
    
    On Error Resume Next
    msConn.BeginTrans
    'msConn.Execute "DROP TABLE LOLO"
    msConn.Execute "DROP TABLE LOLO1"
    msConn.CommitTrans
    On Error GoTo 0
    
    cSQL = "SELECT A.TIPO_PAGO, SUM(A.MONTO) AS MONTO"
    cSQL = cSQL & " INTO LOLO1 "
    cSQL = cSQL & " FROM HIST_TR_PROP AS A, LOLO AS B"
    cSQL = cSQL & " WHERE A.NUM_TRANS = B.NUM_TRANS"
    cSQL = cSQL & " GROUP BY A.TIPO_PAGO;"
    msConn.Execute cSQL
    
    sqltxt = "SELECT B.TIPO_PAGO,('Propina ' + C.DESCRIP) as Descrip, "
    sqltxt = sqltxt & " COUNT (B.TIPO_PAGO) AS Cant, "
    sqltxt = sqltxt & " format(SUM(B.MONTO),'standard') as Pagos "
    sqltxt = sqltxt & " FROM LOLO1 AS B, PAGOS AS C "
    sqltxt = sqltxt & " WHERE B.TIPO_PAGO = C.CODIGO "
    sqltxt = sqltxt & " GROUP BY B.TIPO_PAGO,C.DESCRIP "
    sqltxt = sqltxt & " ORDER BY B.TIPO_PAGO "

    ''''EscribeLog "Open SQL2 : " & GetTickCount - iCounter
    ''''EscribeLog "SQL2: " & sqltxt

    rsEmpresaLoc.Open sqltxt, msConn, adOpenStatic, adLockOptimistic
    ProgBar.value = 80
    nLin = nLin + 100
    If Not rsEmpresaLoc.EOF Then
        MainMant.spDoc.TextOut 300, nLin, "Información de Propinas"
        nLin = nLin + 50
        MainMant.spDoc.TextOut 300, nLin, "--------------------------------"
    End If
    ProgBar.value = 85
    nLin = nLin + 100
    
    ''''EscribeLog "Print Propina Detail : " & GetTickCount - iCounter
    Do Until rsEmpresaLoc.EOF
        If Len(Format(rsEmpresaLoc!pagos, "standard")) > 12 Then iLen = 0 Else iLen = Len(Format(rsEmpresaLoc!pagos, "standard"))
        MainMant.spDoc.TextAlign = SPTA_LEFT
        MainMant.spDoc.TextOut 300, nLin, FormatTexto(rsEmpresaLoc!DESCRIP, 20)
        MainMant.spDoc.TextAlign = SPTA_RIGHT
        MainMant.spDoc.TextOut 1000, nLin, Format(rsEmpresaLoc!pagos, "standard")
        MainMant.spDoc.TextAlign = SPTA_LEFT
        nLin = nLin + 50
        rsEmpresaLoc.MoveNext
    Loop
    rsEmpresaLoc.Close
    
    On Error Resume Next
    msConn.BeginTrans
    msConn.Execute "DROP TABLE LOLO"
    msConn.Execute "DROP TABLE LOLO1"
    msConn.CommitTrans
    On Error GoTo 0
    
    sqltxt = "SELECT MID(TIPO,1,2) AS TIPO, COUNT(TIPO) AS CANT, "
    sqltxt = sqltxt & " format(SUM(PRECIO),'standard') AS VALOR "
    sqltxt = sqltxt & " FROM HIST_TR "
    If opcTipo(0).value = True Then
        sqltxt = sqltxt & " WHERE FECHA >= " & dF1
        sqltxt = sqltxt & " AND FECHA <= " & dF2
    Else
        sqltxt = sqltxt & " WHERE Z_COUNTER = '" & Val(LVZ.SelectedItem.Text) & "'"
    End If
    sqltxt = sqltxt & " AND (MID(TIPO,1,2) = 'VO' OR "
    sqltxt = sqltxt & " MID(TIPO,1,2) = 'EC' "
    sqltxt = sqltxt & " OR MID(TIPO,1,2) = 'DC') "
    sqltxt = sqltxt & " GROUP BY MID(TIPO,1,2) "
    
    ProgBar.value = 90
    
    ''''EscribeLog "Open SQL3 : " & GetTickCount - iCounter
    ''''EscribeLog "SQL3: " & sqltxt

    rsEmpresaLoc.Open sqltxt, msConn, adOpenStatic, adLockReadOnly
    nLin = nLin + 100
    If Not rsEmpresaLoc.EOF Then
        MainMant.spDoc.TextOut 300, nLin, "Información de Descuentos, Correcciones y Anulaciones"
        nLin = nLin + 50
        MainMant.spDoc.TextOut 300, nLin, "---------------------------------------------------------------------------"
    End If
    nLin = nLin + 100
    
    ''''EscribeLog "Print DESC/CORR/ANUL Detail : " & GetTickCount - iCounter
    Do Until rsEmpresaLoc.EOF
        If Len(Format(rsEmpresaLoc!VALOR, "standard")) > 12 Then iLen = 0 Else iLen = Len(Format(rsEmpresaLoc!VALOR, "standard"))
        MainMant.spDoc.TextAlign = SPTA_LEFT
        MainMant.spDoc.TextOut 300, nLin, FormatTexto(rsEmpresaLoc!TIPO, 20)
        MainMant.spDoc.TextAlign = SPTA_RIGHT
        MainMant.spDoc.TextOut 800, nLin, Format(rsEmpresaLoc!VALOR, "standard")
        MainMant.spDoc.TextAlign = SPTA_LEFT
        nLin = nLin + 50
        rsEmpresaLoc.MoveNext
    Loop
    rsEmpresaLoc.Close
    
    Set rsEmpresaLoc = Nothing
    
    ''''EscribeLog "Do PrintPreview: " & GetTickCount - iCounter
    
    MainMant.spDoc.DoPrintPreview
    ProgBar.value = 100
    'Close #1
    ProgBar.value = 0
End If
Me.MousePointer = vbDefault
End Sub

Private Sub Command1_Click()
Unload Me
End Sub

Private Sub Form_Load()

Set rsConCajeros = New Recordset
Set rsConTP = New Recordset
Set rsConExtras = New Recordset

txtFecIni = Format(Date, "SHORT DATE")
txtFecFin = Format(Date, "SHORT DATE")

Call GetReportesZ

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
        cmdPrint.Enabled = False
    Case "N"        'SIN DERECHOS
        txtFecIni.Enabled = False: txtFecFin.Enabled = False: opcTipo(0).Enabled = False: opcTipo(1).Enabled = False
        LVZ.Enabled = False
        cmdEjec.Enabled = False
        MSHFCaj.Enabled = False: MSHFTP.Enabled = False: MSHFExtra.Enabled = False
        cmdPrint.Enabled = False
End Select
End Function

Private Sub LVZ_ItemCheck(ByVal Item As MSComctlLib.ListItem)
opcTipo(1).value = True
End Sub

Private Sub MSHFCaj_Click()
MSHFCaj_EnterCell
End Sub

Private Sub MSHFCaj_EnterCell()
Dim iSoloErr As Integer

'INFO: 27OCT2010
Dim nMinZ As Long   'REPORTE Z INICIAL
Dim nMaxZ As Long   'REPORTE Z FINAL
Dim cArrayZ() As String

Dim i As Byte
Dim jFalseCounter As Byte

For i = 0 To opcTipo.Count - 1
    Select Case opcTipo(i).value
        Case True
            Exit For
        Case False
            jFalseCounter = jFalseCounter + 1
        Case Else
    End Select
    'If opcTipo(i).Value = True Then Exit For
Next

' if 2 = jFalseCounter then NONE selected, then DEFAULT TO DATE OPTION
If jFalseCounter = 2 Then i = 0

nConCaj = Val(MSHFCaj.Text)
iSoloErr = 0

On Error GoTo ErrAdm:
    MSHFTP.Clear
    MSHFTP.Refresh
    MSHFExtra.Clear
    MSHFExtra.Refresh
On Error GoTo 0

ProgBar.value = 60


If i = 0 Then
    'NORMAL NORMAL POR FECHA
Else
    'REPORTE POR Z#
    'INFO: ABRIL2008
    'CHECK TO SEE IF A Z HAS BEEN SELECTED
    For iZZ = 1 To LVZ.ListItems.Count
        If LVZ.ListItems(iZZ).Checked = True Then
            cZetas = cZetas & LVZ.ListItems(iZZ).Text & "','"
        End If
    Next
    On Error Resume Next
    If cZetas = "" Then
        MsgBox "Debe seleccionar al menos un Reporte Z", vbInformation, "Seleccione un reporte (Z)"
        Exit Sub
    Else
        cZetas = "'" & Mid(cZetas, 1, Len(cZetas) - 2)
    End If

    'INFO: 27OCT2010
    cmdEjec.Tag = cZetas
    cArrayZ = Split(Replace(cZetas, "'", ""), ",")
    'cmdEjec.Tag = Array(cArrayZ)
    
    On Error GoTo 0
    On Error GoTo ErrAdm:
End If

dF1 = "'" & Format(txtFecIni, "YYYYMMDD") & "'"
dF2 = "'" & Format(txtFecFin, "YYYYMMDD") & "'"
'BUSCA CORRECCIONES,ANULACIONES,ETC.

sqltxt = "SELECT MID(TIPO,1,2) AS TIPO, COUNT(TIPO) AS CANT, "
sqltxt = sqltxt & " format(SUM(PRECIO),'standard') AS VALOR "
sqltxt = sqltxt & " FROM HIST_TR "
sqltxt = sqltxt & " WHERE CAJERO = " & nConCaj
If i = 0 Then
    sqltxt = sqltxt & " AND FECHA >= " & dF1
    sqltxt = sqltxt & " AND FECHA <= " & dF2
Else
    'INFO: VALIDAR LA Z CONTRA VALORES NUMERICOS EN VEZ DE TEXTO (27OCT2010)
    sqltxt = sqltxt & " AND VAL(Z_COUNTER) BETWEEN  " & GetNumber_FromArray(cArrayZ, GetMin)
    sqltxt = sqltxt & " AND " & GetNumber_FromArray(cArrayZ, Getmax)
End If
sqltxt = sqltxt & " AND (MID(TIPO,1,2) = 'VO' OR "
sqltxt = sqltxt & " MID(TIPO,1,2) = 'EC' "
sqltxt = sqltxt & " OR MID(TIPO,1,2) = 'DC') "
sqltxt = sqltxt & " GROUP BY MID(TIPO,1,2) "

rsConExtras.Open sqltxt, msConn, adOpenStatic, adLockReadOnly
ProgBar.value = 70

Set MSHFExtra.DataSource = rsConExtras
rsConExtras.Close

With MSHFExtra
    .ColWidth(0) = 800: .ColWidth(1) = 900: .ColWidth(2) = 1400:
    .ColAlignment(2) = flexAlignRightCenter
End With
ProgBar.value = 80

msConn.BeginTrans
'EXTRAE INFO (NUMTRANS,ETC) DEL HISTORICO DE TRANSACCIONES
sqltxt = "SELECT DISTINCT "
sqltxt = sqltxt & " B.NUM_TRANS,B.TIPO_PAGO,B.CAJERO,B.MONTO,B.LIN "
sqltxt = sqltxt & " INTO LOLO "
sqltxt = sqltxt & " FROM HIST_TR AS A RIGHT JOIN HIST_TR_PAGO AS B "
sqltxt = sqltxt & " ON (A.NUM_TRANS = B.NUM_TRANS AND A.CAJERO=B.CAJERO) "
sqltxt = sqltxt & " WHERE A.CAJERO = " & nConCaj

'INFOl: VALIDA CONTADOR Z
If i = 0 Then
    sqltxt = sqltxt & " AND A.FECHA >= " & dF1
    sqltxt = sqltxt & " AND A.FECHA <= " & dF2
Else
    'INFO: VALIDAR LA Z CONTRA VALORES NUMERICOS EN VEZ DE TEXTO (27OCT2010)
    sqltxt = sqltxt & " AND VAL(A.Z_COUNTER) BETWEEN  " & GetNumber_FromArray(cArrayZ, GetMin)
    sqltxt = sqltxt & " AND " & GetNumber_FromArray(cArrayZ, Getmax)
End If

''sqltxt = sqltxt & " AND A.FECHA >= " & dF1
''sqltxt = sqltxt & " AND A.FECHA <= " & dF2

msConn.Execute sqltxt
msConn.CommitTrans

'COMBINA HIST_TR Y EXTRAE LA DESCRIPCION DEL TIPO DE PAGO
sqltxt = "SELECT A.TIPO_PAGO,C.DESCRIP,COUNT(A.TIPO_PAGO) AS CANT, "
sqltxt = sqltxt & " format(SUM(A.MONTO),'standard') AS COBRADO "
sqltxt = sqltxt & " FROM LOLO AS A LEFT JOIN PAGOS AS C "
sqltxt = sqltxt & " ON A.TIPO_PAGO = C.CODIGO "
sqltxt = sqltxt & " GROUP BY A.TIPO_PAGO,C.DESCRIP"

ProgBar.value = 90
rsConTP.Open sqltxt, msConn, adOpenStatic, adLockOptimistic
Set MSHFTP.DataSource = rsConTP
rsConTP.Close

sqltxt = "SELECT B.TIPO_PAGO,('Propina ' + C.DESCRIP) as Descrip, "
sqltxt = sqltxt & " COUNT (B.TIPO_PAGO) AS Cant, "
sqltxt = sqltxt & " format(SUM(B.MONTO),'standard') as Pagos "
sqltxt = sqltxt & " FROM HIST_TR_PROP AS B, PAGOS AS C "
sqltxt = sqltxt & " WHERE B.CAJERO = " & nConCaj
sqltxt = sqltxt & " AND B.TIPO_PAGO = C.CODIGO "
sqltxt = sqltxt & " AND B.NUM_TRANS IN (SELECT DISTINCT NUM_TRANS FROM LOLO) "
sqltxt = sqltxt & " GROUP BY B.TIPO_PAGO,C.DESCRIP "
sqltxt = sqltxt & " ORDER BY B.TIPO_PAGO "

rsConTP.Open sqltxt, msConn, adOpenStatic, adLockOptimistic

Do Until rsConTP.EOF
    MSHFTP.AddItem rsConTP!tipo_pago & Chr(9) & rsConTP!DESCRIP & Chr(9) & rsConTP!CANT & Chr(9) & Format(rsConTP!pagos * (-1#), "standard")
    rsConTP.MoveNext
Loop
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
Exit Sub

ErrAdm:
iSoloErr = iSoloErr + 1
If iSoloErr > 4 Then
    MsgBox "Favor Intente Otra vez", vbInformation, BoxTit
    Exit Sub
Else
    Resume
End If

End Sub
Private Sub txtFecFin_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then cmdEjec.SetFocus
End Sub

Private Sub txtFecIni_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then txtFecFin.SetFocus
End Sub
