VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form conCosto 
   BackColor       =   &H00B39665&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "CONSULTA DE CONSUMO POR PRODUCTO DE INVENTARIO"
   ClientHeight    =   8085
   ClientLeft      =   -1125
   ClientTop       =   4455
   ClientWidth     =   11910
   Icon            =   "conCosto.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8085
   ScaleWidth      =   11910
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.ListBox List1 
      Height          =   1815
      Left            =   240
      TabIndex        =   20
      Top             =   1800
      Width           =   3135
   End
   Begin VB.OptionButton Option1 
      BackColor       =   &H00B39665&
      Caption         =   "Por Ganancia"
      Height          =   195
      Index           =   1
      Left            =   7560
      TabIndex        =   19
      Top             =   480
      Width           =   1335
   End
   Begin VB.OptionButton Option1 
      BackColor       =   &H00B39665&
      Caption         =   "Por Nombre"
      Height          =   195
      Index           =   0
      Left            =   5880
      TabIndex        =   18
      Top             =   480
      Value           =   -1  'True
      Width           =   1215
   End
   Begin VB.CheckBox Check1 
      BackColor       =   &H00B39665&
      Caption         =   "Mostrar Detalle"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   5760
      TabIndex        =   17
      Top             =   120
      Width           =   1695
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00008000&
      Height          =   255
      Index           =   5
      Left            =   5520
      TabIndex        =   16
      Text            =   "Historial Costo Unitario"
      Top             =   2880
      Width           =   2775
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   4
      Left            =   5520
      TabIndex        =   15
      Top             =   3120
      Width           =   5175
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFFFF&
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   3
      Left            =   8760
      TabIndex        =   14
      Top             =   2400
      Width           =   1695
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFFFF&
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   2
      Left            =   9000
      TabIndex        =   13
      Top             =   2040
      Width           =   1455
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   1
      Left            =   5520
      TabIndex        =   12
      Text            =   "Porcentaje Ganacias/Perdidas"
      Top             =   2160
      Width           =   3255
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   0  'None
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
      Index           =   0
      Left            =   5520
      TabIndex        =   11
      Text            =   "Total Ganacias/Perdidas"
      Top             =   2520
      Width           =   3255
   End
   Begin VB.CommandButton cmdSalir 
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
      Height          =   555
      Left            =   10800
      TabIndex        =   5
      Top             =   7440
      Width           =   1095
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
      Left            =   9360
      TabIndex        =   2
      Top             =   240
      Width           =   1695
   End
   Begin VB.PictureBox MSCHART1 
      Height          =   255
      Left            =   2760
      ScaleHeight     =   195
      ScaleWidth      =   315
      TabIndex        =   7
      Top             =   0
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H0000FFFF&
      Height          =   495
      Left            =   120
      Picture         =   "conCosto.frx":000C
      Style           =   1  'Graphical
      TabIndex        =   4
      ToolTipText     =   "Envia Seleccion a la Impresora"
      Top             =   7440
      Width           =   735
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFAnalisis 
      Height          =   3495
      Left            =   240
      TabIndex        =   3
      Top             =   3720
      Width           =   11535
      _ExtentX        =   20346
      _ExtentY        =   6165
      _Version        =   393216
      FixedCols       =   0
      AllowBigSelection=   0   'False
      AllowUserResizing=   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
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
      Format          =   139722753
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
      Format          =   139722753
      CurrentDate     =   36418
   End
   Begin MSComctlLib.TabStrip TabStrip1 
      Height          =   6135
      Left            =   120
      TabIndex        =   8
      Top             =   1200
      Width           =   11775
      _ExtentX        =   20770
      _ExtentY        =   10821
      MultiRow        =   -1  'True
      ShowTips        =   0   'False
      Separators      =   -1  'True
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   1
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Reporte Tabular"
            ImageVarType    =   2
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSComctlLib.ProgressBar ProgBar 
      Height          =   180
      Left            =   120
      TabIndex        =   9
      Top             =   840
      Width           =   9015
      _ExtentX        =   15901
      _ExtentY        =   318
      _Version        =   393216
      Appearance      =   1
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H0000FFFF&
      BorderWidth     =   3
      Height          =   375
      Left            =   5760
      Top             =   360
      Width           =   3375
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
      TabIndex        =   6
      Top             =   480
      Width           =   1095
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
      TabIndex        =   10
      Top             =   480
      Width           =   1215
   End
End
Attribute VB_Name = "conCosto"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private rsConsulta01 As Recordset
Private msConnShape As New ADODB.Connection
Private mintCurFrame As Integer ' Current Frame visible
Private rsConDept As Recordset
Private nconDepto As Integer
Private nPluSel As Integer
Private c1PluSel As String
Private c2PluSel As String
Private cOrdSel As String
Private nPagina As Integer
Private iLin As Integer
Private nIdProducto As Long
Private cShapeOrder As String
Private Sub cmdEjec_Click()
Dim rsShape As New ADODB.Recordset
Dim sqltxt As String
Dim cTop10 As String
Dim dF1 As String
Dim dF2 As String
Dim cEldepto As String
Dim nTotal As Single
Dim i As Long
Dim j As Long
Dim nSumColVentas As Single
Dim nSumColCostos As Single
Dim TXT As String
Dim txt0 As String
Dim txt1 As String

Text1(2).Enabled = True: Text1(3).Enabled = True
Text1(2) = "": Text1(3) = ""
Text1(4) = "Información del Periodo : " & txtFecIni & " al " & txtFecFin
MSHFAnalisis.Visible = True
MSChart1.Visible = False
'-------------------------------

'If nconDepto = 0 Then cEldepto = "" Else cEldepto = " AND B.ID_DEPT_INV = " & nconDepto
dF1 = Format(txtFecIni, "YYYYMMDD")
dF2 = Format(txtFecFin, "YYYYMMDD")
If nconDepto = 0 Then cEldepto = "" Else cEldepto = " WHERE DEPTO = " & nconDepto
nTotal = 0

Me.MousePointer = vbHourglass

ProgBar.value = 10
Set MSHFAnalisis.DataSource = Nothing
MSHFAnalisis.Clear
ProgBar.value = 15
Call DropLolosTables(True)

ProgBar.value = 0
'txt = "SHAPE {SELECT MAX(DEPTO.DESCRIP) AS DESCRIP,DEPTO.CODIGO,SUM(HIST_TR.CANT) AS CANTIDAD FROM DEPTO, HIST_TR WHERE DEPTO.CODIGO=HIST_TR.DEPTO GROUP BY DEPTO.CODIGO} " & _
    " APPEND ({SELECT DEPTO,DESCRIP,FORMAT(PRECIO1*DEPTO.CANTIDAD,'####0.00') AS PRECIO*,FORMAT(VALOR,'#####0.00') AS VENTAS FROM PLU ORDER BY VALOR DESC}" & _
    " RELATE CODIGO TO DEPTO) AS Relacion"

txt0 = "SELECT "
txt0 = txt0 & "A.CODIGO,A.DEPTO,IIF(A.CODIGO=B.CODIGO,B.CONTENEDOR,0) AS ENVASE,"
txt0 = txt0 & "A.DESCRIP,IIF(A.CODIGO = B.CODIGO, B.PRECIO, A.PRECIO1) As PRECIO"
txt0 = txt0 & " Into LOLO1 "
txt0 = txt0 & " From "
txt0 = txt0 & " PLU AS A LEFT JOIN CONTEND_02 AS B ON A.CODIGO=B.CODIGO " & cEldepto


txt1 = "SELECT A.CODIGO, A.DEPTO,A.ENVASE, "
txt1 = txt1 & " IIF(A.ENVASE=B.CONTENEDOR,A.DESCRIP+ '-'+B.DESCRIP,A.DESCRIP) AS DESCRIP,"
txt1 = txt1 & " A.PRECIO INTO LOLO2 "
txt1 = txt1 & " FROM LOLO1 AS A LEFT JOIN CONTENED AS B ON A.ENVASE=B.CONTENEDOR"
    
TXT = "SELECT A.ID_PLU,C.ID AS COD_INV,MAX(B.DESCRIP) AS PLU,"
TXT = TXT & " MAX(B.PRECIO) AS PRECIO, SUM(D.CANT) AS CANT,"
TXT = TXT & " SUM(D.PRECIO) AS VENTAS,MAX(A.DESCRIP) AS DESC_INV,"
'TXT = TXT & " MAX(C.COSTO * A.CANT) AS COSTO, MAX(C.COSTO) / MAX(C.CANTIDAD2) AS COSTO_UNIT,"
'TXT = TXT & " MAX(C.COSTO * A.CANT) AS COSTO, MAX(C.COSTO * A.CANT) / MAX(C.CANTIDAD2) AS COSTO_UNIT,"
TXT = TXT & " MAX(C.COSTO * A.CANT) * SUM(D.CANT) AS COSTO, "
'TXT = TXT & " (MAX(C.COSTO * A.CANT) * SUM(D.CANT)) / MAX(C.CANTIDAD2) AS COSTO_UNIT, "
TXT = TXT & " MAX(C.COSTO) AS COSTO_UNIT, "
'TXT = TXT & " MAX (C.COSTO * A.CANT) * SUM(D.CANT) / IIF(MAX(C.CANTIDAD2)=0,1,MAX(C.CANTIDAD2)) AS CONSUMO "
'(MAX (C.COSTO * A.CANT) * SUM(D.CANT))
TXT = TXT & " MAX(C.COSTO * A.CANT) * SUM(D.CANT) AS CONSUMO "
'TXT = TXT & " (MAX (C.COSTO * A.CANT) * SUM(D.CANT)) AS CONSUMO "
TXT = TXT & " INTO LOLO3 "
TXT = TXT & " FROM "
TXT = TXT & " PLU_INVENT As A, LOLO2 As B, INVENT As C, HIST_TR As D, UNIDADES As F"
TXT = TXT & " Where "
TXT = TXT & " D.FECHA >= '" & dF1 & "'"
TXT = TXT & " And D.FECHA <= '" & dF2 & "'"
TXT = TXT & " And MID(D.DESCRIP,LEN(TRIM(D.DESCRIP)),1) <> '%' "
TXT = TXT & " And '%' NOT IN (D.DESCRIP) "
TXT = TXT & " And A.ID_PLU = B.CODIGO "
TXT = TXT & " AND A.ID_ENV = B.ENVASE "
TXT = TXT & " And A.ID_PROD_INV = C.ID "
TXT = TXT & " And A.ID_PLU = D.PLU "
TXT = TXT & " AND A.ID_ENV = D.ENVASE "
TXT = TXT & " And C.UNID_MEDIDA = F.ID "
TXT = TXT & " GROUP BY A.ID_PLU,C.ID "

ProgBar.value = 20
msConnShape.BeginTrans
msConnShape.Execute txt0
msConnShape.CommitTrans
ProgBar.value = 40
msConnShape.BeginTrans
msConnShape.Execute txt1
msConnShape.CommitTrans
ProgBar.value = 60
msConnShape.BeginTrans
msConnShape.Execute TXT
ProgBar.value = 80
msConnShape.CommitTrans
ProgBar.value = 85

TXT = "SHAPE {SELECT MAX(PLU) AS PRODUCTO,ID_PLU,"
TXT = TXT & "FORMAT(MAX(PRECIO),'###0.00') AS PRECIO,"
TXT = TXT & "MAX(CANT) AS CANT,"
TXT = TXT & "FORMAT(MAX(VENTAS),'####0.00') AS VENTAS,"
TXT = TXT & "FORMAT(SUM(CONSUMO),'####0.00') AS CONSUMO,"
TXT = TXT & "FORMAT(MAX(VENTAS) - SUM(CONSUMO),'CURRENCY') AS GANANCIA "
TXT = TXT & "FROM LOLO3 "
TXT = TXT & "GROUP BY ID_PLU ORDER BY " & cShapeOrder & " } "
TXT = TXT & " APPEND ({SELECT ID_PLU,DESC_INV,FORMAT(COSTO,'####0.00') AS COSTO, "
TXT = TXT & "FORMAT(COSTO_UNIT,'####0.00') AS COST_UNIT,"
TXT = TXT & "FORMAT(CONSUMO,'####0.00') AS CONSUMO FROM LOLO3}"
TXT = TXT & " RELATE ID_PLU TO ID_PLU) AS Relacion"

'>>>> On Error Resume Next
ProgBar.value = 90
rsShape.Open TXT, msConnShape, adOpenKeyset, adLockOptimistic
If rsShape.RecordCount = 0 Then
    rsShape.Close
    ProgBar.value = 100
    Set rsShape = Nothing
    ProgBar.value = 0
    Me.MousePointer = vbDefault
    Exit Sub
End If
Set MSHFAnalisis.DataSource = rsShape
'EL Grid MSHFAnalisis termina con 2 BANDS
    ' BANDA 1.- INFO DEL PLU
    ' BANDA 2.- INFO DE INVENTARIO
ProgBar.value = 100
'CheckBox control — 0 is Unchecked (default), 1 is Checked, and 2 is Grayed (dimmed).
If Check1.value = 0 Then
    MSHFAnalisis.CollapseAll (0)
End If
MSHFAnalisis.ColWordWrapOptionBand(0, 0) = flexWordBreak
MSHFAnalisis.ColWordWrapOptionBand(1, 1) = flexWordBreak
MSHFAnalisis.ColWidth(0, 0) = 2500
MSHFAnalisis.ColWidth(0, 1) = 0
MSHFAnalisis.ColWidth(0, 2) = 700
MSHFAnalisis.ColWidth(0, 3) = 700
MSHFAnalisis.ColWidth(0, 4) = 1100
MSHFAnalisis.ColWidth(0, 5) = 1100
MSHFAnalisis.ColWidth(0, 6) = 1500
MSHFAnalisis.ColWidth(1, 0) = 0
MSHFAnalisis.ColWidth(1, 1) = 2000
MSHFAnalisis.ColWidth(1, 2) = 600
MSHFAnalisis.ColWidth(1, 3) = 1100
MSHFAnalisis.ColWidth(1, 4) = 1000
MSHFAnalisis.ColAlignmentBand(0, 2) = flexAlignRightCenter
MSHFAnalisis.ColAlignmentBand(0, 3) = flexAlignCenterCenter
MSHFAnalisis.ColAlignmentBand(0, 4) = flexAlignRightCenter
MSHFAnalisis.ColAlignmentBand(0, 5) = flexAlignRightCenter
MSHFAnalisis.ColAlignmentBand(0, 6) = flexAlignRightCenter
MSHFAnalisis.ColAlignmentBand(1, 2) = flexAlignRightBottom
MSHFAnalisis.ColAlignmentBand(1, 3) = flexAlignRightBottom
MSHFAnalisis.ColAlignmentBand(1, 4) = flexAlignRightBottom

ProgBar.value = 20
Call DropLolosTables(False)
ProgBar.value = 90
rsShape.Close
Set rsShape = Nothing
ProgBar.value = 0
Me.MousePointer = vbDefault
On Error GoTo 0
End Sub
Private Sub cmdSalir_Click()
msConnShape.Close
Set msConnShape = Nothing
Unload Me
End Sub

Private Sub Command1_Click()
'BoxResp = MsgBox ("", vbCritical, "LISTADO DE ANALISIS DE PERDIDAS Y GANANCIAS")
'Exit Sub
Dim iCtr As Integer 'Contador de Linea
Dim iCol, iFil As Integer 'Contador de Columnas
Dim cText As String
Dim ispace As Integer
Dim iLen As Integer
Dim sSubTot As Single
Dim sqltxt As String
Dim dF1 As String
Dim dF2 As String
Dim rsMAIN As New ADODB.Recordset
Dim rsDETA As New ADODB.Recordset
Dim nlCodPlu As Long
Dim nVentas As Double
Dim nAcumPlu As Double
Dim nAcumDept As Double
Dim nGENVentas As Double
Dim nGENAcumPlu As Double
Dim nGENAcumDept As Double
Dim cSQL As String
Dim cSQL1 As String

On Error GoTo ErrAdm:

dF1 = Format(txtFecIni, "YYYYMMDD")
dF2 = Format(txtFecFin, "YYYYMMDD")
'''''Exit Sub

ProgBar.value = 10
Me.MousePointer = vbHourglass

cSQL = "SELECT MAX(PLU) AS PRODUCTO,ID_PLU, "
cSQL = cSQL & " FORMAT(MAX(PRECIO),'###0.00') AS PRECIO, "
cSQL = cSQL & " MAX(CANT) AS CANT, "
cSQL = cSQL & " FORMAT(MAX(VENTAS),'####0.00') AS VENTAS, "
cSQL = cSQL & " FORMAT(SUM(CONSUMO),'####0.00') AS CONSUMO, "
cSQL = cSQL & " FORMAT(MAX(VENTAS) - SUM(CONSUMO),'CURRENCY') AS GANANCIA "
cSQL = cSQL & " FROM LOLO3 "
cSQL = cSQL & " GROUP BY ID_PLU ORDER BY " & cShapeOrder
'>>>> On Error Resume Next
'rsMAIN.Open "SELECT * FROM LOLO3", msConn, adOpenDynamic, adLockOptimistic
rsMAIN.Open cSQL, msConn, adOpenKeyset, adLockOptimistic

ProgBar.value = 15

If rsMAIN.EOF Then
    Me.MousePointer = vbDefault
    MsgBox "No hay datos para mostrar", vbExclamation, "NO EXISTEN DATOS EN ESTE DEPARTAMENTO"
    Exit Sub
End If
sSubTot = 0#: iLin = 0: nPagina = 0

On Error GoTo 0

'---------Open "c:\mifile.txt" For Output As #3
On Error GoTo ErrorPrn:
MainMant.spDoc.DocBegin
PrintTit
EscribeLog ("Admin." & "Impresion de Listado: " & Me.Caption & " PERIODO : Desde " & txtFecIni & " Hasta " & txtFecFin)
ProgBar.value = 20

Do Until rsMAIN.EOF
    nlCodPlu = rsMAIN!ID_PLU
    ProgBar.value = 80
    'ID_PLU  COD_INV PLU PRECIO  CANT    VENTAS  DESC_INV    COSTO   COSTO_UNIT  CONSUMO
    Do While nlCodPlu = rsMAIN!ID_PLU
        ProgBar.value = 50
        MainMant.spDoc.TextOut 300, iLin, Mid(rsMAIN!PRODUCTO, 1, 25)
        MainMant.spDoc.TextOut 920, iLin, rsMAIN!precio
        MainMant.spDoc.TextAlign = SPTA_RIGHT
        MainMant.spDoc.TextOut 1120, iLin, rsMAIN!CANT
        MainMant.spDoc.TextOut 1260, iLin, rsMAIN!Ventas
        MainMant.spDoc.TextOut 1560, iLin, rsMAIN!CONSUMO
        MainMant.spDoc.TextOut 1930, iLin, rsMAIN!GANANCIA
        MainMant.spDoc.TextAlign = SPTA_LEFT
        iLin = iLin + 50
JumpNextInvent:
        cSQL1 = " SELECT ID_PLU, COD_INV, DESC_INV,FORMAT(COSTO,'####0.00') AS COSTO, "
        cSQL1 = cSQL1 & " FORMAT(COSTO_UNIT,'####0.0000') AS COST_UNIT,"
        cSQL1 = cSQL1 & " FORMAT(CONSUMO,'####0.00') AS CONSUMO "
        cSQL1 = cSQL1 & " FROM LOLO3 "
        cSQL1 = cSQL1 & " WHERE ID_PLU = " & rsMAIN!ID_PLU

        rsDETA.Open cSQL1, msConn, adOpenStatic, adLockOptimistic
        Do While Not rsDETA.EOF
            MainMant.spDoc.TextOut 600, iLin, "+ " & Mid(rsDETA!DESC_INV, 1, 20)
            MainMant.spDoc.TextAlign = SPTA_RIGHT
            MainMant.spDoc.TextOut 1260, iLin, rsDETA!COSTO
            MainMant.spDoc.TextOut 1560, iLin, rsDETA!COST_UNIT
            MainMant.spDoc.TextAlign = SPTA_LEFT
            MainMant.spDoc.TextOut 1590, iLin, GetUnidConsumo(rsDETA!ID_PLU, rsDETA!COD_INV)
            MainMant.spDoc.TextAlign = SPTA_RIGHT
            MainMant.spDoc.TextOut 1930, iLin, GetPorcentaje(rsDETA!CONSUMO, rsMAIN!Ventas)
            MainMant.spDoc.TextAlign = SPTA_LEFT
'            MainMant.spDoc.TextOut 1650, iLin, FormatPercent(rsDETA!CONSUMO / rsMAIN!Ventas)
            iLin = iLin + 50
            If iLin > 2400 Then
                MainMant.spDoc.TextAlign = SPTA_LEFT
                PrintTit
            End If
            rsDETA.MoveNext
        Loop
        rsDETA.Close
        nVentas = rsMAIN!Ventas
        nGENVentas = nGENVentas + rsMAIN!Ventas
        nAcumPlu = nAcumPlu + rsMAIN!CONSUMO
        nGENAcumPlu = nGENAcumPlu + rsMAIN!CONSUMO
        nAcumDept = nAcumDept + rsMAIN!CONSUMO
        nGENAcumDept = nGENAcumDept + rsMAIN!CONSUMO
        rsMAIN.MoveNext
        If iLin > 2400 Then
            MainMant.spDoc.TextAlign = SPTA_LEFT
            PrintTit
        End If
        If rsMAIN.EOF Then
            MainMant.spDoc.TextOut 300, iLin, "Ganancias / Perdidas   : " & Format(nVentas - nAcumPlu, "CURRENCY")
            On Error Resume Next
            MainMant.spDoc.TextOut 1100, iLin, "% Ganancias / Perdidas : " & Format(1 - (nAcumPlu / nVentas), "PERCENT")
            iLin = iLin + 100
            MainMant.spDoc.TextOut 300, iLin, "ANALISIS DEPARTAMENTAL"
            iLin = iLin + 50
            MainMant.spDoc.TextOut 300, iLin, "Ganancias / Perdidas   : " & Format(nGENVentas - nGENAcumPlu, "CURRENCY")
            MainMant.spDoc.TextOut 1100, iLin, "% Ganancias / Perdidas : " & Format(1 - (nGENAcumPlu / nGENVentas), "PERCENT")
            Exit Do
        End If
        If nlCodPlu = rsMAIN!ID_PLU Then
            GoTo JumpNextInvent:
        Else
            MainMant.spDoc.TextOut 300, iLin, "Ganancias / Perdidas   : " & Format(nVentas - nAcumPlu, "CURRENCY")
            On Error Resume Next
            MainMant.spDoc.TextOut 1100, iLin, "% Ganancias / Perdidas : " & Format(1 - (nAcumPlu / nVentas), "PERCENT")
            On Error GoTo 0
            iLin = iLin + 50
            MainMant.spDoc.TextOut 300, iLin, "---   ---   ---   ---   ---   ---   ---   ---   ---   ---   ---   ---   ---   ---   ---   ---   ---   ---   ---   ---   ---   ---   ---   ---   ---   "
            iLin = iLin + 50
            nAcumPlu = 0#
        End If
        nlCodPlu = rsMAIN!ID_PLU
    Loop
Loop
'-------Close #3
ProgBar.value = 100
Me.MousePointer = vbDefault
MainMant.spDoc.DoPrintPreview
On Error GoTo 0
ProgBar.value = 0
rsMAIN.Close

Call Seguridad

Exit Sub

ErrAdm:
Dim ADOError As Error
Dim sError As String
For Each ADOError In msConn.Errors
    If ADOError.Number = -2147217865 Then
        sError = "Favor intente imprimir nuevamente el reporte en unos segundos" + vbCrLf
    Else
        sError = sError & ADOError.Number & " - " & ADOError.Description + vbCrLf
    End If
Next ADOError
  Me.MousePointer = vbDefault
  EscribeLog ("Admin." & "ERROR (Modulo conCosto) : " & sError)
  ShowMsg "Modulo conCosto." & vbCrLf & sError, vbYellow, vbBlue
Exit Sub

ErrorPrn:
    Me.MousePointer = vbDefault
    MsgBox Err.Description, vbCritical, BoxTit
    'MsgBox "¡ Ocurre algún Error con la Impresora, Intente Conecterla !", vbExclamation, BoxTit
    ShowMsg " ¡ Ocurre algún Error con la Impresora, Intente Conecterla !", vbRed, vbYellow
    Resume
End Sub
Private Function GetPorcentaje(nConsumo As Single, nVentas As Single) As String
If nConsumo = 0 Or nVentas = 0 Then
    GetPorcentaje = FormatPercent(0 / 1)
Else
    GetPorcentaje = FormatPercent(nConsumo / nVentas)
End If
End Function
Private Function GetUnidConsumo(nIDPLU As Long, nIDInvent As Long) As String
Dim rsUnidad As ADODB.Recordset
Dim cSQL As String

cSQL = "SELECT A.CANT, C.DESCRIP, B.UNID_CONSUMO "
cSQL = cSQL & " FROM PLU_INVENT AS A, INVENT AS B, UNID_CONSUMO AS C "
cSQL = cSQL & " WHERE A.ID_PLU = " & nIDPLU
cSQL = cSQL & " AND A.ID_PROD_INV = " & nIDInvent
cSQL = cSQL & " AND A.ID_PROD_INV = B.ID "
cSQL = cSQL & " AND B.UNID_CONSUMO = C.ID "
Set rsUnidad = New ADODB.Recordset
rsUnidad.Open cSQL, msConn, adOpenStatic, adLockOptimistic

If Not rsUnidad.EOF Then
    GetUnidConsumo = "(" & rsUnidad!CANT & Space(1) & Left(rsUnidad!DESCRIP, 2) & ")"
Else
    GetUnidConsumo = " EOF"
End If
rsUnidad.Close
Set rsUnidad = Nothing
End Function
Private Sub Form_Load()
Set rsConsulta01 = New Recordset
txtFecIni = Format(Date, "SHORT DATE")
txtFecFin = Format(Date, "SHORT DATE")
nPluSel = 0
cOrdSel = " 4 ASC " 'POR PRODUCTO
nPagina = 1
cShapeOrder = "1 ASC"
If msConnShape.State = adStateOpen Then
    msConnShape.Close
End If
msConnShape.Open cShapeADOString
Call PutDepts

Call Seguridad

'msConnShape.Provider = "MSDATASHAPE"
'msConnShape.Open "Data Provider=Microsoft.Jet.OLEDB.4.0;" _
            + "Data Source=\\SOLO11\ACCESS\SOLO.mdb;" _
            + "Jet OLEDB:Database Password=master24"
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
        Command1.Enabled = False
    Case "N"        'SIN DERECHOS
        txtFecIni.Enabled = False: txtFecFin.Enabled = False: Option1(0).Enabled = False: Option1(1).Enabled = False
        cmdEjec.Enabled = False: List1.Enabled = False
        MSHFAnalisis.Enabled = False
        Command1.Enabled = False
End Select
End Function

Private Sub PutDepts()
Dim rsConDept As New ADODB.Recordset
On Error Resume Next
rsConDept.Open "SELECT CODIGO,DESCRIP FROM DEPTO " & _
            " ORDER BY DESCRIP ", msConn, adOpenStatic, adLockOptimistic
List1.AddItem "TODOS LOS DEPARTAMENTOS"
Do Until rsConDept.EOF
    List1.AddItem rsConDept!DESCRIP & Space(70) & rsConDept!CODIGO
    rsConDept.MoveNext
Loop
nconDepto = 0
rsConDept.Close
On Error GoTo 0
End Sub

Private Sub List1_Click()
Dim POSIC As Integer
'CAPTURA EL NUMERO DEL DEPTO.
POSIC = Len(List1.Text)
If IsNumeric(Val(Mid(List1.Text, POSIC - 5, 6))) Then
    nconDepto = Val(Mid(List1.Text, POSIC - 5, 6))
Else
    nconDepto = 0
End If
cmdEjec_Click
End Sub

Private Sub Option1_Click(Index As Integer)
If Index = 0 Then
    'Ordenado por Nombre
    cShapeOrder = "1 ASC"
Else
    'Ordenado por Ganancia
    cShapeOrder = "MAX(VENTAS) - SUM(CONSUMO) DESC"
End If
End Sub

Private Sub txtFecFin_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then cmdEjec.SetFocus
End Sub

Private Sub txtFecIni_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then txtFecFin.SetFocus
End Sub

Private Sub txtFecIni_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then txtFecIni.SetFocus
End Sub

Private Sub txtFecFin_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    cmdEjec.SetFocus
End If
End Sub
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
MainMant.spDoc.TextOut 300, 450, conCosto.Caption
MainMant.spDoc.TextOut 300, 500, "PERIODO : Desde " & txtFecIni & " Hasta " & txtFecFin
MainMant.spDoc.TextOut 300, 550, "Departamento Seleccionado: " & Left(List1.Text, 50)

MainMant.spDoc.TextOut 300, 600, String(136, "-")
MainMant.spDoc.TextOut 300, 650, "Producto"
MainMant.spDoc.TextOut 900, 650, "Precio"
MainMant.spDoc.TextOut 1050, 650, "Cant"
MainMant.spDoc.TextOut 1150, 650, "Ventas"
MainMant.spDoc.TextOut 1450, 650, "Consumo"
MainMant.spDoc.TextOut 1780, 650, "Ganancia"

MainMant.spDoc.TextOut 600, 700, "Inventario"
MainMant.spDoc.TextOut 1150, 700, "Costo"
MainMant.spDoc.TextOut 1450, 700, "C.Unit"
MainMant.spDoc.TextOut 1780, 700, "Cosumo"
MainMant.spDoc.TextOut 300, 750, String(136, "-")

iLin = 800
nPagina = nPagina + 1
End Sub

Private Sub DropLolosTables(nOpcion As Boolean)
'Dim objTable As New ADOX.Table
Dim objTabla As ADOX.Table
Dim objCat As New ADOX.Catalog

objCat.ActiveConnection = msConnShape

For Each objTabla In objCat.Tables
    If objTabla.Type = "TABLE" Then
        If objTabla.Name = "LOLO" Then
            ProgBar.value = 30
            msConnShape.BeginTrans
            msConnShape.Execute "DROP TABLE LOLO"
            msConnShape.CommitTrans
        ElseIf objTabla.Name = "LOLO1" Then
            ProgBar.value = 40
            msConnShape.BeginTrans
            msConnShape.Execute "DROP TABLE LOLO1"
            msConnShape.CommitTrans
        ElseIf objTabla.Name = "LOLO2" Then
            ProgBar.value = 50
            msConnShape.BeginTrans
            msConnShape.Execute "DROP TABLE LOLO2"
            msConnShape.CommitTrans
            'info: OCTUBRE 2009
        ElseIf objTabla.Name = "LOLO4" Then
            ProgBar.value = 50
            msConnShape.BeginTrans
            msConnShape.Execute "DROP TABLE LOLO4"
            msConnShape.CommitTrans
        ElseIf objTabla.Name = "LOLO3" Then
            If nOpcion = True Then
                ProgBar.value = 60
                msConnShape.BeginTrans
                msConnShape.Execute "DROP TABLE LOLO3"
                msConnShape.CommitTrans
            End If
        End If
    End If
Next
End Sub
