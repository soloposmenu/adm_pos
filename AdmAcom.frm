VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form AdmAcom 
   BackColor       =   &H00B39665&
   Caption         =   "PLATOS ACOMPAÑANTES (Sin Precio)"
   ClientHeight    =   8235
   ClientLeft      =   75
   ClientTop       =   465
   ClientWidth     =   11430
   Icon            =   "AdmAcom.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   8235
   ScaleWidth      =   11430
   Begin MSComctlLib.ProgressBar ProgBAR 
      Height          =   255
      Left            =   120
      TabIndex        =   22
      Top             =   7920
      Width           =   11175
      _ExtentX        =   19711
      _ExtentY        =   450
      _Version        =   393216
      Appearance      =   1
      Max             =   20
   End
   Begin VB.CommandButton cmdImpresion 
      BackColor       =   &H0000FFFF&
      Height          =   495
      Left            =   10560
      Picture         =   "AdmAcom.frx":030A
      Style           =   1  'Graphical
      TabIndex        =   21
      ToolTipText     =   "Imprimir Enlace de Acompañantes"
      Top             =   5400
      Width           =   735
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
      Height          =   1575
      Index           =   0
      Left            =   120
      TabIndex        =   9
      Top             =   4680
      Width           =   4695
      Begin VB.TextBox txtValConsumo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   3600
         TabIndex        =   19
         Top             =   1080
         Width           =   975
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   0
         Left            =   1440
         TabIndex        =   1
         Top             =   360
         Visible         =   0   'False
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
         Caption         =   "Valor de Consumo"
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
         Height          =   375
         Index           =   1
         Left            =   3600
         TabIndex        =   20
         Top             =   680
         Width           =   975
      End
      Begin VB.Label Label1 
         BackColor       =   &H00B39665&
         Caption         =   "Descripción"
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
         TabIndex        =   12
         Top             =   840
         Width           =   1095
      End
      Begin VB.Label Label1 
         BackColor       =   &H00B39665&
         Caption         =   "Nombre Corto"
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
         TabIndex        =   11
         Top             =   1200
         Width           =   1215
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
         Index           =   0
         Left            =   120
         TabIndex        =   10
         Top             =   480
         Visible         =   0   'False
         Width           =   735
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00B39665&
      Height          =   1695
      Left            =   120
      TabIndex        =   13
      Top             =   6120
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
         TabIndex        =   4
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
         TabIndex        =   5
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
         TabIndex        =   6
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
         TabIndex        =   7
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
         TabIndex        =   8
         Top             =   1080
         Width           =   1300
      End
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFAcompa 
      Height          =   4215
      Left            =   120
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   360
      Width           =   4695
      _ExtentX        =   8281
      _ExtentY        =   7435
      _Version        =   393216
      BackColor       =   16777215
      ForeColor       =   0
      Rows            =   0
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
   Begin MSComctlLib.ListView LV 
      Height          =   4935
      Left            =   4920
      TabIndex        =   15
      ToolTipText     =   "Haga Doble Click para Aregar Articulo de Inventario a la Formula"
      Top             =   360
      Width           =   6375
      _ExtentX        =   11245
      _ExtentY        =   8705
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   0
   End
   Begin MSComctlLib.ListView LV_ACOINV 
      Height          =   1815
      Left            =   4920
      TabIndex        =   17
      Top             =   6000
      Width           =   6375
      _ExtentX        =   11245
      _ExtentY        =   3201
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   0
   End
   Begin VB.Label Label2 
      BackColor       =   &H00B39665&
      Caption         =   "Componentes del Acompañante"
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
      Left            =   4920
      TabIndex        =   18
      Top             =   5760
      Width           =   2895
   End
   Begin VB.Label Label2 
      BackColor       =   &H00B39665&
      Caption         =   "Productos de Inventario"
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
      Left            =   4920
      TabIndex        =   16
      Top             =   120
      Width           =   2655
   End
   Begin VB.Label Label2 
      BackColor       =   &H00B39665&
      Caption         =   "Seleccione Acompañante"
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
      TabIndex        =   14
      Top             =   120
      Width           =   2655
   End
End
Attribute VB_Name = "AdmAcom"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rsMes As Recordset
Dim rsINV As New ADODB.Recordset
Dim rsAcoInvent As New ADODB.Recordset
Dim Valores(3) As String
Dim nOpc As Integer
Dim nPagina As Integer
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
MainMant.spDoc.TextOut 300, 450, "ACOMPAÑANTES / Y ENLACE CON EL INVENTARIO"

MainMant.spDoc.TextOut 300, 550, "ACOMPAÑANTE"
MainMant.spDoc.TextOut 800, 550, "INVENTARIO"
MainMant.spDoc.TextOut 1450, 550, "Se Consume"
MainMant.spDoc.TextOut 1700, 550, "VALOR.Consumo"
MainMant.spDoc.TextOut 300, 600, String(83, "=")
iLin = 650
nPagina = nPagina + 1
End Sub


Private Sub CambiaTablaAcomp(indice As Integer)
Dim sqltext(10) As String
Dim i As Integer, nRgistros As Integer

'0 Elimina
'1 Modifica o Agrega
Select Case indice
Case 0
    'Eliminacion
    'Solamente se borra si el mesero no tiene ventas
    
    sqltext(0) = "DELETE * FROM ACOMPA WHERE CODIGO =" & Val(Text1(0))

    msConn.BeginTrans
    On Error GoTo ErrorEnTrans:
        msConn.Execute sqltext(0), nRgistros
        msConn.CommitTrans
        EscribeLog ("Admin." & "ELIMINACION DE ACOMPAÑANTE: " & Text1(1).Text)
    On Error GoTo 0
    
    If nRgistros = 0 Then
        MsgBox "ACOMPAÑANTE no Puede Ser Eliminado", vbExclamation, BoxTit
    Else
        rsMes.Requery
        Set MSHFAcompa.DataSource = rsMes
        MSHFAcompa_EnterCell
    End If
'-----------------------------------------------------
Case 1  '1 Modifica o Agrega, primero se busca en tabla Meseros

    If Text1(1) = "" And Text1(2) = "" Then
        MsgBox "NO HAY SUFICIENTE INFORMACION PARA GRABAR", vbExclamation, BoxTit
        MSHFAcompa_EnterCell
        Exit Sub
    End If
        
    If Text1(0) = Empty Then
        Text1(0) = 0
    End If
    
    On Error Resume Next
    rsMes.MoveFirst
    On Error GoTo 0
    
    rsMes.Find ("CODIGO = " & Val(Text1(0)))
    If Not rsMes.EOF Then
        'Modificacion
        sqltext(0) = "UPDATE ACOMPA SET " & _
            " DESCRIP = '" & Text1(1) & "'" & _
            ",CORTO = '" & Text1(2) & "'" & _
            " WHERE CODIGO = " & Val(Text1(0))
    Else
        'Nuevo Cajero
        sqltext(0) = "INSERT INTO ACOMPA (DESCRIP,CORTO) " & _
            " VALUES ('" & Text1(1) & " ','" & Text1(2) & "')"
    End If

    msConn.BeginTrans
    On Error GoTo ErrorEnTrans:
    msConn.Execute sqltext(0), nRgistros
    msConn.CommitTrans
    On Error GoTo 0

    rsMes.Requery
    Set MSHFAcompa.DataSource = rsMes
    MSHFAcompa_EnterCell

End Select
MSHFAcompa.SetFocus
Exit Sub

ErrorEnTrans:
    'something bad happened so rollback the transaction
  Dim ADOError As Error
  'msConn.RollbackTrans
  For Each ADOError In msConn.Errors
     sError = sError & ADOError.Number & " - " & ADOError.Description _
            + vbCrLf
  Next ADOError
  EscribeLog ("Admin." & "ERROR (Modulo AdmAcom) : " & sError)
  MsgBox sError, vbCritical, BoxTit
  Resume Next
End Sub
Private Sub SetUpPantalla()
With MSHFAcompa
    .ColWidth(0) = 600: .ColWidth(1) = 2800: .ColWidth(2) = 1600:
    .ColWidth(3) = 800: .ColWidth(4) = 800:
    .ColWidth(5) = 1200:
End With
End Sub

'---------------------------------------------------------------------------------------
' Procedimiento : cmdImpresion_Click
' Autor       : hsequeira
' Fecha       : 17/10/2013
' Proposito   : Impresion de Acompañantes
'---------------------------------------------------------------------------------------
'
Private Sub cmdImpresion_Click()
Dim rsReporte As ADODB.Recordset
Dim cSQL As String
Dim nSubTotal As Single
Dim nComparacion As Long

cSQL = "SELECT A.DESCRIP AS ACOMPANANTE, C.NOMBRE AS INVENTARIO, "
cSQL = cSQL & " B.CANTIDAD_CONSUME,D.DESCRIP AS UNIDAD, B.VALOR_CONSUMO, B.ID_ACOM "
'cSQL = cSQL & " FROM ACOMPA AS A, ACOMPA_INVENT AS B, INVENT AS C, "
'cSQL = cSQL & " UNID_CONSUMO AS D WHERE B.ID_ACOM = A.CODIGO "
'cSQL = cSQL & " AND B.ID_INVENT = C.ID AND C.UNID_CONSUMO = D.ID "
cSQL = cSQL & " FROM ((((ACOMPA AS A) LEFT JOIN ACOMPA_INVENT AS B ON A.CODIGO = B.ID_ACOM)"
cSQL = cSQL & " LEFT JOIN INVENT AS C ON B.ID_INVENT = C.ID)"
cSQL = cSQL & " LEFT JOIN UNID_CONSUMO AS D ON C.UNID_CONSUMO = D.ID)"
cSQL = cSQL & " ORDER BY 1, 2 "
nPagina = 0
Set rsReporte = New ADODB.Recordset
rsReporte.Open cSQL, msConn, adOpenStatic, adLockOptimistic

If rsReporte.EOF Then
    MsgBox "NO HAY DATOS PARA GENERAR EL REPORTE", vbOKOnly, BoxTit
    rsReporte.Close
    Set rsReporte = Nothing
    Exit Sub
End If

MainMant.spDoc.DocBegin
MainMant.spDoc.TextAlign = SPTA_LEFT

Call PrintTit    'Rutina de Titulos

Do While Not rsReporte.EOF
    MainMant.spDoc.TextAlign = SPTA_LEFT
    MainMant.spDoc.TextOut 300, iLin, rsReporte!ACOMPANANTE
    nComparacion = IIf(IsNull(rsReporte!ID_ACOM), 0, rsReporte!ID_ACOM)
    MainMant.spDoc.TextOut 800, iLin, IIf(IsNull(rsReporte!INVENTARIO), "", rsReporte!INVENTARIO)
    MainMant.spDoc.TextAlign = SPTA_RIGHT
    MainMant.spDoc.TextOut 1700, iLin, "(" & IIf(IsNull(rsReporte!CANTIDAD_CONSUME), 0, rsReporte!CANTIDAD_CONSUME) & ")  " & IIf(IsNull(rsReporte!UNIDAD), "", rsReporte!UNIDAD)
    MainMant.spDoc.TextAlign = SPTA_RIGHT
    MainMant.spDoc.TextOut 1960, iLin, Format(IIf(IsNull(rsReporte!VALOR_CONSUMO), 0, rsReporte!VALOR_CONSUMO), "###0.000")
    nSubTotal = nSubTotal + IIf(IsNull(rsReporte!VALOR_CONSUMO), 0, rsReporte!VALOR_CONSUMO)
    iLin = iLin + 50
    If iLin > 2400 Then PrintTit
    rsReporte.MoveNext
    If Not rsReporte.EOF Then
        If nComparacion <> IIf(IsNull(rsReporte!ID_ACOM), 0, rsReporte!ID_ACOM) Then
            MainMant.spDoc.TextAlign = SPTA_LEFT
            MainMant.spDoc.TextOut 1050, iLin, "TOT. ACOMPAÑANTE: "
            MainMant.spDoc.TextAlign = SPTA_RIGHT
            MainMant.spDoc.TextOut 1960, iLin, Format(nSubTotal, "CURRENCY")
            iLin = iLin + 50
            nSubTotal = 0
        End If
    End If
    MainMant.spDoc.TextAlign = SPTA_LEFT
Loop
MainMant.spDoc.TextAlign = SPTA_LEFT
MainMant.spDoc.TextOut 1050, iLin, "TOT. ACOMPAÑANTE: "
MainMant.spDoc.TextAlign = SPTA_RIGHT
MainMant.spDoc.TextOut 1960, iLin, Format(nSubTotal, "CURRENCY")
rsReporte.Close
Set rsReporte = Nothing
MainMant.spDoc.DoPrintPreview
End Sub

Private Sub Command2_Click(Index As Integer)
Dim i As Integer

Select Case Index
Case 0 ' Modificar
    For i = 1 To 2
        Text1(i).Enabled = True
        Command2(i).Enabled = False
        Valores(i) = Text1(i)
    Next
    Command2(0).Enabled = False
    For i = 3 To 4
        Command2(i).Enabled = True
    Next
    'Text1(0).Enabled = False
    Text1(1).SetFocus
    MSHFAcompa.BackColor = &HC0C0C0
    MSHFAcompa.Enabled = False
    nOpc = Index
Case 1  ' Agregar
    For i = 0 To 2
        Text1(i).Enabled = True
        Valores(i) = Text1(i)
        Text1(i) = ""
        Command2(i).Enabled = False
    Next
    Command2(0).Enabled = False
    'Command2(i).Enabled = False
    For i = 3 To 4
        Command2(i).Enabled = True
    Next
    
    MSHFAcompa.BackColor = &HC0C0C0
    MSHFAcompa.Enabled = False
    
    Text1(1).SetFocus
    nOpc = Index
Case 2 'Eliminar
    'BoxPreg = "¿ Desea Eliminar Acompañante " & Text1(1) & " ?"
    'BoxResp = MsgBox(BoxPreg, vbQuestion + vbYesNo, BoxTit)
    If ShowMsg("¿ Desea Eliminar Acompañante ?" & vbCrLf & vbCrLf & Text1(1), vbYellow, vbRed, vbYesNo) = vbYes Then BoxResp = vbYes Else BoxResp = vbNo
    If BoxResp = vbYes Then
        CambiaTablaAcomp (0)
    End If
Case 3 'Salvar
    'BoxPreg = "¿ Desea Salvar los Datos en Pantalla ?"
    'BoxResp = MsgBox(BoxPreg, vbQuestion + vbYesNo, BoxTit)
    If ShowMsg("¿ Desea Salvar los Datos en Pantalla ?" & vbCrLf & vbCrLf & _
        Text1(1).Text, vbYellow, vbBlue, vbYesNo) = vbYes Then BoxResp = vbYes Else BoxResp = vbNo
    For i = 0 To 2
        Text1(i).Enabled = False
        Command2(i).Enabled = True
    Next
    Command2(0).Enabled = True
    For i = 3 To 4
        Command2(i).Enabled = False
    Next
    MSHFAcompa.Enabled = True
    MSHFAcompa.BackColor = vbWhite
    
    If BoxResp = vbYes Then
        CambiaTablaAcomp (1)
    End If
    Command2(2).Enabled = True
Case 4 'Regresar sin Salvar
    For i = 0 To 2
        Text1(i) = Valores(i)
        Text1(i).Enabled = False
        Command2(i).Enabled = True
    Next
    Command2(0).Enabled = True
    For i = 3 To 4
        Command2(i).Enabled = False
    Next
    
    MSHFAcompa.Enabled = True
    MSHFAcompa.BackColor = vbWhite
    
    Command2(2).Enabled = True
End Select

Call Seguridad

End Sub

Private Sub Form_Load()
Dim cSQL As String
Dim nFila As Integer
Set rsMes = New Recordset

Me.MousePointer = vbHourglass
rsMes.Open "SELECT CODIGO,DESCRIP,CORTO " & _
        " FROM ACOMPA ORDER BY DESCRIP", msConn, adOpenDynamic, adLockOptimistic
Set MSHFAcompa.DataSource = rsMes

'List1.ListIndex = 0
'"FORMAT((C.COSTO_EMPAQUE/A.FACT_CONSUMO),'##0.00') AS COSTO, "
cSQL = "SELECT C.COD_DEPT, A.ID, A.DESCRIP_SUB, B.DESCRIP AS CONSUMO, "
cSQL = cSQL & " C.UNID_CONSUMO, C.COSTO, "
cSQL = cSQL & " B.UNIDAD AS UNIDADES,A.FACT_CONSUMO AS CANTIDAD,"
cSQL = cSQL & " UCASE(MID(DESCRIP_SUB,1,1)) AS LETRA, A.ID_SUB "
cSQL = cSQL & " FROM INVENT_02 AS A, UNID_CONSUMO AS B, INVENT AS C "
cSQL = cSQL & " WHERE A.ID = C.ID AND C.UNID_CONSUMO = B.ID "
cSQL = cSQL & " ORDER BY A.DESCRIP_SUB"

rsINV.Open cSQL, msConn, adOpenDynamic, adLockOptimistic

nFila = 1
LV.ListItems.Clear
LV.ColumnHeaders.Clear
LV.ColumnHeaders.Add , , "Producto"
LV.ColumnHeaders.Add , , "Unidad Consumo"
LV.ColumnHeaders.Add , , "Costo"
LV.ColumnHeaders.Add , , "Cantidad"
LV.ColumnHeaders.Add , , "Id Sub"
LV.ColumnHeaders.Add , , "Id"
LV.ColumnHeaders.Item(3).Alignment = lvwColumnRight
LV.ColumnHeaders.Item(4).Alignment = lvwColumnRight

LV.ColumnHeaders.Item(1).Width = 2700
LV.ColumnHeaders.Item(2).Width = 1400
LV.ColumnHeaders.Item(3).Width = 800
LV.ColumnHeaders.Item(4).Width = 950
LV.ColumnHeaders.Item(5).Width = 0
LV.ColumnHeaders.Item(6).Width = 0

On Error Resume Next
rsINV.MoveFirst
On Error GoTo 0
Do While Not rsINV.EOF
    LV.ListItems.Add , , rsINV!DESCRIP_SUB
    LV.ListItems.Item(nFila).ListSubItems.Add , , rsINV!CONSUMO
    LV.ListItems.Item(nFila).ListSubItems.Add , , rsINV!COSTO
    LV.ListItems.Item(nFila).ListSubItems.Add , , rsINV!Cantidad
    LV.ListItems.Item(nFila).ListSubItems.Add , , rsINV!ID_SUB
    LV.ListItems.Item(nFila).ListSubItems.Add , , rsINV!ID
    nFila = nFila + 1
    rsINV.MoveNext
Loop
On Error Resume Next
rsINV.MoveFirst
On Error GoTo 0
SetUpPantalla
Me.MousePointer = vbDefault
MSHFAcompa_EnterCell
nOpc = 99
Show
Call PopulaEnlaceConInventario

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
        'Command2(3).Enabled = False: Command2(4).Enabled = False
    Case "CV"
        Command2(0).Enabled = False: Command2(2).Enabled = False
        'Command2(3).Enabled = False: Command2(4).Enabled = False
    Case "V"
        Command2(0).Enabled = False: Command2(1).Enabled = False: Command2(2).Enabled = False
        Command2(3).Enabled = False: Command2(4).Enabled = False
    Case "N"
        MSHFAcompa.Enabled = False
        Command2(0).Enabled = False: Command2(1).Enabled = False: Command2(2).Enabled = False
        Command2(3).Enabled = False: Command2(4).Enabled = False
End Select
End Function

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
rsINV.Close
Set rsINV = Nothing
Set rsAcoInvent = Nothing
End Sub
Private Sub LV_ACOINV_DblClick()
Dim vResp
Dim cSQL As String

vResp = MsgBox("¿ Desea Quitar la relacion de " & Text1(1).Text & " con " & vbCrLf & _
        LV_ACOINV.SelectedItem.Text & " ?", vbQuestion + vbYesNo, "Quitar Relación")
If vResp = vbYes Then
    cSQL = "DELETE * FROM ACOMPA_INVENT WHERE ID_INVENT = " & Val(LV_ACOINV.SelectedItem.ListSubItems.Item(2))
    msConn.BeginTrans
    msConn.Execute cSQL
    msConn.CommitTrans
    MSHFAcompa_EnterCell
End If
End Sub
Private Sub LV_DblClick()
Dim vResp
Dim nUnidEnter As Single
Dim cSQL As String
Dim cFecha As String
Dim cCadena As String

cFecha = Format(Date, "YYYYMMDD")
vResp = MsgBox("¿ Desea Enlazar " + Text1(1) + " con " & LV.SelectedItem.Text & " ?", vbQuestion + vbYesNo, "Formula de " + cNomPlu)
'For i = 1 To 5
'    cCadena = cCadena & _
'        LV.ColumnHeaders(i + 1).Text & " = " & LV.SelectedItem.ListSubItems.Item(i).Text & vbCrLf
'Next

If vResp = vbYes Then
    nUnidEnter = Val(InputBox("Escriba Cuanto se Consume en Inventario de " + vbCrLf + _
            LV.SelectedItem.Text + vbCrLf + _
            "Al Vender como Acompañante " + Text1(1) + "." + vbCrLf + vbCrLf + _
            "Consumo x " + LV.SelectedItem.ListSubItems.Item(1).Text, "ESCRIBIR UNICAMENTE DATOS NUMERICOS", "0.00"))
    If nUnidEnter <> 0 Then

        'INFO: ACTUALIZACION 8FEB2013
        cSQL = " INSERT INTO ACOMPA_INVENT "
        cSQL = cSQL & "(ID_ACOM,ID_INVENT,FECHA_IN,FECHA_MODI,"
        cSQL = cSQL & "CANTIDAD_CONSUME,USER_IN,USER_MODI, VALOR_CONSUMO) "
        cSQL = cSQL & " VALUES (" & Text1(0) & ","
        Debug.Print "LV.SelectedItem.ListSubItems.Item(5).Text: " & LV.SelectedItem.ListSubItems.Item(5).Text
        cSQL = cSQL & LV.SelectedItem.ListSubItems.Item(5).Text & ",'"
        cSQL = cSQL & cFecha & "','" & cFecha & "'," & nUnidEnter & ","
        cSQL = cSQL & npNumCaj & "," & npNumCaj & ","
        Debug.Print "Val(LV.SelectedItem.ListSubItems.Item(2).Text) * nUnidEnter : " & Val(LV.SelectedItem.ListSubItems.Item(2).Text) * nUnidEnter
        Debug.Print "Val(LV.SelectedItem.ListSubItems.Item(2).Text) " & Val(LV.SelectedItem.ListSubItems.Item(2).Text)
        Debug.Print "nUnidEnter : " & nUnidEnter
        cSQL = cSQL & Val(LV.SelectedItem.ListSubItems.Item(2).Text) * nUnidEnter & ")"
        Debug.Print cSQL
        msConn.BeginTrans
        msConn.Execute cSQL
        msConn.CommitTrans
        
        MSHFAcompa_EnterCell

    End If
End If
End Sub
Private Sub PopulaEnlaceConInventario()
'INFO: ACUALIZA LOS VALORES DE LOS ACOMPAÑANTES CON
'LOS NUEVOS COSTOS DE INVENTARIO
'19/02/2005
Dim cSQL1 As String
Dim cSQL2 As String
Dim rsInvent As ADODB.Recordset
Dim nConsumo As Single
Dim nCounter As Integer

cSQL1 = "SELECT A.ID_INVENT, A.CANTIDAD_CONSUME, B.COSTO "
cSQL1 = cSQL1 & " FROM ACOMPA_INVENT as A, INVENT AS B  "
cSQL1 = cSQL1 & " WHERE A.ID_INVENT = B.ID "
cSQL1 = cSQL1 & " ORDER BY A.ID_INVENT "

Set rsInvent = New ADODB.Recordset
rsInvent.Open cSQL1, msConn, adOpenStatic, adLockOptimistic
On Error Resume Next
ProgBar.Max = rsInvent.RecordCount
On Error GoTo 0
nCounter = 0
ProgBar.value = nCounter
Do While Not rsInvent.EOF
    nConsumo = rsInvent!CANTIDAD_CONSUME * rsInvent!COSTO
    cSQL2 = "UPDATE ACOMPA_INVENT SET VALOR_CONSUMO = " & nConsumo
    cSQL2 = cSQL2 & " WHERE ID_INVENT = " & rsInvent!COSTO
    msConn.Execute cSQL2
    nCounter = nCounter + 1
    rsInvent.MoveNext
    ProgBar.value = nCounter
Loop
rsInvent.Close
Set rsInvent = Nothing
End Sub
Private Sub MSHFAcompa_Click()
MSHFAcompa_EnterCell
End Sub
Private Sub MSHFAcompa_EnterCell()
Dim i As Integer
Dim nC As Integer
Dim cSQL As String
Dim nFila  As Integer
Dim nConsumo As Single

For i = 0 To 2
    Text1(i).Enabled = True
Next

nC = Val((MSHFAcompa.Text))

On Error Resume Next
rsMes.MoveFirst 'ANTES DE FIND, SIEMPRE HAY QUE MANDAR CURSOR AL INICIO DE LA TABLA
On Error GoTo 0

rsMes.Find "CODIGO = " & nC

If Not rsMes.EOF Then
    Text1(0) = nC
    Text1(1) = IIf(IsNull(rsMes!DESCRIP), "", rsMes!DESCRIP)
    Text1(2) = IIf(IsNull(rsMes!CORTO), "", rsMes!DESCRIP)
End If
cSQL = "SELECT '(' & B.CANTIDAD_CONSUME & ')  ' & A.DESCRIP_SUB  AS DESCRIP_SUB,"
cSQL = cSQL & " B.VALOR_CONSUMO, B.ID_INVENT "
cSQL = cSQL & " FROM INVENT_02 AS A, ACOMPA_INVENT AS B "
cSQL = cSQL & " WHERE B.ID_ACOM = " & nC
cSQL = cSQL & " AND B.ID_INVENT = A.ID "
rsAcoInvent.Open cSQL, msConn, adOpenStatic, adLockOptimistic
'rsAcoInvent.MoveFirst

LV_ACOINV.ListItems.Clear
LV_ACOINV.ColumnHeaders.Clear
LV_ACOINV.ColumnHeaders.Add , , "Producto Inventario"
LV_ACOINV.ColumnHeaders.Add , , "Costo Consumo"
LV_ACOINV.ColumnHeaders.Add , , "ID INVENT"

LV_ACOINV.ColumnHeaders.Item(2).Alignment = lvwColumnRight

LV_ACOINV.ColumnHeaders.Item(1).Width = 2700
LV_ACOINV.ColumnHeaders.Item(2).Width = 1400
LV_ACOINV.ColumnHeaders.Item(3).Width = 0
nFila = 1
Do While Not rsAcoInvent.EOF
    LV_ACOINV.ListItems.Add , , rsAcoInvent!DESCRIP_SUB
    LV_ACOINV.ListItems.Item(nFila).ListSubItems.Add , , rsAcoInvent!VALOR_CONSUMO
    LV_ACOINV.ListItems.Item(nFila).ListSubItems.Add , , rsAcoInvent!ID_INVENT
    nConsumo = nConsumo + rsAcoInvent!VALOR_CONSUMO
    nFila = nFila + 1
    rsAcoInvent.MoveNext
Loop
rsAcoInvent.Close
txtValConsumo = Format(nConsumo, "0.0000")
For i = 0 To 2
    Text1(i).Enabled = False
Next
End Sub

Private Sub Text1_KeyPress(Index As Integer, KeyAscii As Integer)
If KeyAscii = 13 Then
    Select Case Index
    Case 1
''''''''        Text1(Index + 1).SetFocus
''''''''        If nOpc = 1 Then
''''''''            If Text1(1) = "" Then
''''''''                Text1(1).SetFocus
''''''''                Exit Sub
''''''''            End If
''''''''
''''''''            On Error Resume Next
''''''''            rsMes.MoveFirst
''''''''            On Error GoTo 0
''''''''
''''''''            rsMes.Find "CODIGO = " & Val(Text1(0))
''''''''            If rsMes.EOF Then
''''''''            Else
''''''''                MsgBox "¡¡ Ya Existe Mesero con ese Número !!", vbExclamation, BoxTit
''''''''                Text1(Index).SetFocus
''''''''            End If
''''''''        End If
        If Text1(2) = "" Then Text1(2) = Text1(1)
        Text1(2).SetFocus
        Text1(2).SelLength = Len(Text1(2).Text)
    Case 2
        Command2(3).SetFocus
    End Select
End If

End Sub
