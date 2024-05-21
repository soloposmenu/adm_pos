VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form PrintCta 
   BackColor       =   &H00FF8080&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   5115
   ClientLeft      =   15
   ClientTop       =   15
   ClientWidth     =   5175
   ControlBox      =   0   'False
   Icon            =   "PrintCta.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5115
   ScaleWidth      =   5175
   StartUpPosition =   2  'CenterScreen
   Begin VB.ListBox cmbNombres 
      Height          =   3180
      Left            =   120
      TabIndex        =   6
      Top             =   1800
      Width           =   4815
   End
   Begin VB.CheckBox Check1 
      BackColor       =   &H00FF8080&
      Caption         =   "Imprimir Todos"
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
      Left            =   3360
      TabIndex        =   5
      Top             =   240
      Width           =   1695
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Regresar"
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
      Left            =   3720
      TabIndex        =   4
      Top             =   840
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Imprimir"
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
      Left            =   120
      TabIndex        =   3
      Top             =   840
      Width           =   1215
   End
   Begin MSComCtl2.DTPicker DTPicker1 
      Height          =   375
      Left            =   1440
      TabIndex        =   0
      Top             =   120
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   661
      _Version        =   393216
      Format          =   66977793
      CurrentDate     =   36616
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FF8080&
      Caption         =   "SELECCIONE"
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
      TabIndex        =   2
      Top             =   1560
      Width           =   3495
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00FF8080&
      Caption         =   "Fecha Límite"
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
      TabIndex        =   1
      Top             =   240
      Width           =   1215
   End
End
Attribute VB_Name = "PrintCta"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private rsNombre As New ADODB.Recordset
Private rsRegDebito As New ADODB.Recordset
Private rsRegCred As New ADODB.Recordset
Private rsConso As New ADODB.Recordset
Private nCodSel As Integer
Private iLin As Integer
Private nPagina As Integer
Private Sub TituloCxC()
On Error GoTo ErrAdm:
If nPagina = 0 Then
    If PRINT_OPC = True Then
        MainMant.spDoc.WindowTitle = "Impresión de ESTADO DE CUENTA de CLIENTES"
    Else
        MainMant.spDoc.WindowTitle = "Impresión de ESTADO DE CUENTA de PROVEEDORES"
    End If
    MainMant.spDoc.FirstPage = 1
    MainMant.spDoc.PageOrientation = SPOR_PORTRAIT
    MainMant.spDoc.Units = SPUN_LOMETRIC
End If
MainMant.spDoc.Page = nPagina + 1

MainMant.spDoc.TextOut 300, 200, Format(Date, "long date") & "  " & Time
MainMant.spDoc.TextOut 300, 250, "Página : " & nPagina + 1
MainMant.spDoc.TextOut 300, 350, rs00!DESCRIP
MainMant.spDoc.TextOut 300, 450, "ESTADO DE CUENTA"
   
MainMant.spDoc.TextOut 300, 500, "Nombre de la Empresa : " & rsNombre!EMPRESA
MainMant.spDoc.TextOut 300, 550, "Contacto     : " & rsNombre!NOMBRE & " " & rsNombre!APELLIDO
MainMant.spDoc.TextOut 300, 650, "Saldo Actual : " & Format(rsNombre!SALDO, "CURRENCY")
MainMant.spDoc.TextOut 300, 700, "Fecha Limite Seleccionada : " & DTPicker1
If PRINT_OPC = True Then
    MainMant.spDoc.TextOut 300, 800, "FECHA       NUM.TRANS  TP           DEBITO               CREDITO                  SALDO"
    MainMant.spDoc.TextOut 300, 850, "--------------------------------------------------------------------------------------------------------------"
Else
    MainMant.spDoc.TextOut 300, 800, "FECHA       NUM.TRANS  TP           DEBITO               CREDITO                PAGADO         PENDIENTE"
    MainMant.spDoc.TextOut 300, 850, "-----------------------------------------------------------------------------------------------------------------------------------------"
End If
iLin = 900

nPagina = nPagina + 1
On Error GoTo 0
Exit Sub

ErrAdm:
MsgBox Err.Number & " - " & Err.Description, vbCritical, BoxTit
Resume Next
End Sub

Private Sub Check1_Click()
MsgBox "AUN NO ESTA DISPONIBLE", vbInformation, BoxTit
End Sub

Private Sub cmbNombres_Click()
If cmbNombres.Text = "" Then Exit Sub
POSIC = Len(cmbNombres.Text)
nCodSel = Val(Mid(cmbNombres.Text, POSIC - 5, 6))
rsNombre.MoveFirst
rsNombre.Find "CODIGO = " & nCodSel
If rsNombre.EOF Then MsgBox "Fin de Archivo. No fue encontrado en el Archivo", vbCritical, nCodSel
End Sub
Private Sub Command1_Click()
Dim dF1 As String
Dim cString As String
Dim NSPACE As Integer
Dim cFecha As String
Dim nSumaCredito As Double
Dim nSumaDebito As Double
Dim cMesActual As String
Dim nspace1 As Integer
Dim rsSaldo As New ADODB.Recordset
Dim rsDEVOLUCIONES As New ADODB.Recordset
Dim rsNOTAS As New ADODB.Recordset
Dim XTXT As String
Dim nDBMonto As Double
Dim nDBRecibido As Double
Dim nDBSaldo As Double
Dim nCRPagado As Double

If nCodSel = 0 Then
    MsgBox "SELECCIONE UNA PERSONA o EMPRESA DE LA LISTA", vbInformation, BoxTit
    Exit Sub
End If
dF1 = Format(DTPicker1, "YYYYMMDD")
cMesActual = Mid(dF1, 1, 6)

nPag = 0: iLin = 0
rsNombre.Find "CODIGO = " & nCodSel

nPagina = 0
MainMant.spDoc.DocBegin
TituloCxC

If PRINT_OPC = True Then
    'ABRE CLIENTES. SELECCIONA TODOS LOS DB HASTA UN MES ANTES DE LA FECHA ESPECIFICADA
    'ENTONCES DA COMO RESULTADO UN ACUMULADO DE (DB)
    XTXT = "SELECT CODIGO_CLI,SUM(MONTO) AS MONTO1," & _
            " SUM(RECIBIDO) AS RECIBIDO1,SUM(SALDO) AS SALDO1 " & _
            " FROM HIST_TR_CLI " & _
            " WHERE CODIGO_CLI = " & nCodSel & _
            " AND MID(FECHA,1,6) < " & Mid(dF1, 1, 6) & _
            " GROUP BY CODIGO_CLI"
    rsSaldo.Open XTXT, msConn, adOpenStatic, adLockOptimistic
    If Not rsSaldo.EOF Then
        'MONTO1 ES EL CAMPO IMPORTANTE
        nDBMonto = rsSaldo!MONTO1
        nDBRecibido = rsSaldo!RECIBIDO1
        nDBSaldo = rsSaldo!SALDO1
        'MsgBox nDBSaldo
        'MsgBox nDBRecibido
    End If
    rsSaldo.Close
    
    XTXT = "SELECT CODIGO_CLI,SUM(VALOR_DOC) AS PAGADO " & _
        " FROM CXC_REC " & _
        " WHERE CODIGO_CLI = " & nCodSel & _
        " AND MID(FECHA_DOC,1,6) < '" & Mid(dF1, 1, 6) & "'" & _
        " GROUP BY CODIGO_CLI "
    rsSaldo.Open XTXT, msConn, adOpenStatic, adLockOptimistic
    If Not rsSaldo.EOF Then nCRPagado = rsSaldo!PAGADO
    rsSaldo.Close
    
    'DEBITO DEL MES CORRIENTE
    rsRegDebito.Open "SELECT A.CODIGO_CLI,A.NUM_TRANS,A.FECHA,A.STATUS," & _
        " A.MONTO,A.RECIBIDO,A.SALDO , A.TIPO_TRANS, A.COMMENT, A.USUARIO " & _
        " INTO LOLO From HIST_TR_CLI AS A" & _
        " WHERE A.CODIGO_CLI = " & nCodSel & _
        " AND MID(A.FECHA,1,6) = " & Mid(dF1, 1, 6) & _
        " AND MID(A.FECHA,1,8) <= " & Mid(dF1, 1, 8) & _
        " ORDER BY A.FECHA", msConn, adOpenStatic, adLockReadOnly

    'CREDITO DEL MES CORRIENTE
    rsRegCred.Open "SELECT B.CODIGO_CLI,B.NUM_DOC,B.FECHA_DOC,B.VALOR_DOC," & _
        " B.TIPO_DOC,B.USUARIO , B.COMMENT " & _
        " FROM CXC_REC AS B " & _
        " WHERE B.CODIGO_CLI = " & nCodSel & _
        " AND MID(B.FECHA_DOC,1,6) = '" & Mid(dF1, 1, 6) & "'" & _
        " AND MID(B.FECHA_DOC,1,8) <= '" & Mid(dF1, 1, 8) & "'", msConn, adOpenStatic, adLockReadOnly
    
    msConn.BeginTrans
    Do Until rsRegCred.EOF
        msConn.Execute "INSERT INTO LOLO " & _
            " (CODIGO_CLI,NUM_TRANS,FECHA," & _
            " STATUS,MONTO,RECIBIDO," & _
            " SALDO,TIPO_TRANS, COMMENT, " & _
            " USUARIO) " & _
            " VALUES (" & _
    rsRegCred!CODIGO_CLI & "," & rsRegCred!NUM_DOC & ",'" & rsRegCred!FECHA_DOC & "'," & _
    0 & "," & rsRegCred!VALOR_DOC & "," & 0# & "," & _
    0# & ",'" & rsRegCred!TIPO_DOC & "','" & rsRegCred!COMMENT & "'," & _
    rsRegCred!USUARIO & ")"
        
        rsRegCred.MoveNext
    Loop
    
    msConn.CommitTrans
    rsConso.Open "SELECT * FROM LOLO ORDER BY FECHA,TIPO_TRANS,NUM_TRANS", msConn, adOpenDynamic, adLockReadOnly
    If Not rsConso.EOF Then
        MainMant.spDoc.TextOut 300, iLin, "EL SALDO ACUMULADO ANTES DE 01/" & Mid(dF1, 5, 2) + "/" + Mid(rsConso!FECHA, 1, 4) & " ES :  " & Format(nDBMonto - nCRPagado, "STANDARD")
        iLin = iLin + 100
    End If
    'PRINTER.PRINT "123456789012345678901234567890123456789012345678901234567890123456789012345678901234567890"
    Do Until rsConso.EOF
        cFecha = (Mid(rsConso!FECHA, 7, 2) + "/" + Mid(rsConso!FECHA, 5, 2) + "/" + Mid(rsConso!FECHA, 1, 4))
        MainMant.spDoc.TextAlign = SPTA_LEFT
        MainMant.spDoc.TextOut 300, iLin, cFecha
        If rsConso!TIPO_TRANS = "FA" Or rsConso!TIPO_TRANS = "ND" Then
            nspace1 = 6 'debito
            nSumaDebito = nSumaDebito + rsConso!MONTO
        Else
             nspace1 = 19   'credito
            nSumaCredito = nSumaCredito + rsConso!MONTO
        End If

        MainMant.spDoc.TextAlign = SPTA_LEFT
        MainMant.spDoc.TextOut 300, iLin, cFecha
        MainMant.spDoc.TextOut 500, iLin, Format(Format(rsConso!NUM_TRANS, "GENERAL NUMBER"), "@@@@@@@@@")
        MainMant.spDoc.TextOut 730, iLin, rsConso!TIPO_TRANS
        MainMant.spDoc.TextAlign = SPTA_RIGHT
        MainMant.spDoc.TextOut 1000, iLin, Format(Format(rsConso!MONTO, "###,###.00"), "@@@@@@@@@@")
        If nspace1 = 6 Then
            '20
            MainMant.spDoc.TextOut 1450, iLin, Format(Format(((nDBMonto - nCRPagado) + nSumaDebito - nSumaCredito), "###,###.00"), "@@@@@@@@@@")
        Else
            '7
            MainMant.spDoc.TextOut 1200, iLin, Format(Format(((nDBMonto - nCRPagado) + nSumaDebito - nSumaCredito), "###,###.00"), "@@@@@@@@@@")
        End If
        MainMant.spDoc.TextAlign = SPTA_LEFT
        iLin = iLin + 50
AlProximo:
        rsConso.MoveNext
        If iLin > 2400 Then
'            nPagina = nPagina + 1
            MainMant.spDoc.Page = nPagina
            TituloCxC
        End If
    Loop
   
    'Printer.FontUnderline = True
    'MainMant.spDoc.SetFont
    MainMant.spDoc.TextOut 300, iLin, "Saldo al " & DTPicker1 & "  " & Format(Format((nDBMonto - nCRPagado) + nSumaDebito - nSumaCredito, "###,###.00"), "@@@@@@@@@@")
    iLin = iLin + 50
'-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+
Else
'-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+
    'ABRE PROVEEDORES
'-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+
    ' STAT = 2 SIGNIFICA QUE YA ESTAN PAGADAS
    cSQL = "SELECT A.COD_PROV,A.NUMERO,A.FECHA,A.STAT,"
    cSQL = cSQL & " A.MONTO,A.PAGADO,A.TIPO,A.USUARIO "
    cSQL = cSQL & " INTO LOLO "
    cSQL = cSQL & " From COMPRAS_HEAD AS A"
    cSQL = cSQL & " WHERE A.COD_PROV = " & nCodSel
    cSQL = cSQL & " AND A.FECHA <= '" & dF1 & "'"
    cSQL = cSQL & " AND A.TIPO='CR' "

    rsRegDebito.Open cSQL, msConn, adOpenStatic, adLockReadOnly
        
    'ABRE LAS DEVOL. QUE SE PONEN AL FINAL DEL ESTADO DE CUENTA
    rsDEVOLUCIONES.Open "SELECT PROV,NUM_DOC,FECHA_DOC,TOTAL " & _
        " FROM DEVOLUCION_HEAD " & _
        " WHERE PROV = " & nCodSel & _
        " AND FECHA_DOC <= '" & dF1 & "'" & _
        " ORDER BY FECHA_DOC ", msConn, adOpenStatic, adLockOptimistic
    
    'NOTAS DE DEBITO Y CREDITO SIMPLES
    rsNOTAS.Open "SELECT PROVEEDOR, NUMERO, FECHA, MONTO, TIPO, -333 AS NDNC_SIMPLE " & _
            " FROM MOVSPROVEEDOR WHERE PROVEEDOR = " & nCodSel & _
            " AND FECHA <= '" & dF1 & "'" & _
            " ORDER BY FECHA, NUMERO ", msConn, adOpenStatic, adLockOptimistic
    
    'Cuentas x Pagar, y se agregan a lolo
    rsRegCred.Open "SELECT B.CODIGO_PROV,B.NUM_DOC,B.FECHA_DOC,B.VALOR_DOC," & _
        " B.TIPO_DOC,B.USUARIO,B.COMMENT " & _
        " FROM CXP_REC AS B " & _
        " WHERE B.CODIGO_PROV = " & nCodSel, msConn, adOpenStatic, adLockReadOnly
    
    msConn.BeginTrans
    Do Until rsRegCred.EOF

        cSQL = "INSERT INTO LOLO "
        cSQL = cSQL & "(COD_PROV,NUMERO,FECHA,"
        cSQL = cSQL & " STAT,MONTO,PAGADO,"
        cSQL = cSQL & " TIPO,USUARIO) "
        cSQL = cSQL & " VALUES ("
        cSQL = cSQL & rsRegCred!CODIGO_PROV & "," & rsRegCred!NUM_DOC & ",'"
        cSQL = cSQL & rsRegCred!FECHA_DOC & "',0,"
        cSQL = cSQL & rsRegCred!VALOR_DOC & ",0,'"
        cSQL = cSQL & rsRegCred!TIPO_DOC & "'," & rsRegCred!USUARIO & ")"
        
        msConn.Execute cSQL
        rsRegCred.MoveNext
    
    Loop
    
    'DEVOLUCIONES
    Do Until rsDEVOLUCIONES.EOF
        
        cSQL = "INSERT INTO LOLO "
        cSQL = cSQL & " (COD_PROV,NUMERO,FECHA,"
        cSQL = cSQL & " STAT,MONTO,PAGADO,"
        cSQL = cSQL & " TIPO,USUARIO) "
        cSQL = cSQL & " VALUES ("
        cSQL = cSQL & rsDEVOLUCIONES!PROV & ",'" & rsDEVOLUCIONES!NUM_DOC & "','"
        cSQL = cSQL & rsDEVOLUCIONES!FECHA_DOC & "',"
        cSQL = cSQL & "0," & rsDEVOLUCIONES!TOTAL & "," & rsDEVOLUCIONES!TOTAL & ",'DV',0)"
        
        msConn.Execute cSQL
        rsDEVOLUCIONES.MoveNext
        
    Loop
    
    msConn.CommitTrans
    
    Dim rsNC As ADODB.Recordset
    Set rsNC = New ADODB.Recordset
    
    'NOTAS DE CREDITO
    cSQL = "SELECT PROV,ID,MAX(FECHA_DOC) AS FECHA, "
    cSQL = cSQL & "SUM(COSTO_UNIT + ITBM) AS TOTAL "
    cSQL = cSQL & " FROM NC "
    cSQL = cSQL & " WHERE PROV = " & nCodSel
    cSQL = cSQL & " AND FECHA_DOC <= '" & dF1 & "'"
    cSQL = cSQL & " GROUP BY PROV,ID "
    cSQL = cSQL & " ORDER BY 3 "
    
    rsNC.Open cSQL, msConn, adOpenStatic, adLockOptimistic
    Do While Not rsNC.EOF
        cSQL = "INSERT INTO LOLO "
        cSQL = cSQL & "(COD_PROV,NUMERO,FECHA,STAT,MONTO,PAGADO,TIPO,USUARIO)"
        cSQL = cSQL & " VALUES ("
        cSQL = cSQL & rsNC!PROV & ","
        cSQL = cSQL & rsNC!ID & ",'"
        cSQL = cSQL & rsNC!FECHA & "',"
        cSQL = cSQL & "0,"
        cSQL = cSQL & rsNC!TOTAL & ","
        cSQL = cSQL & rsNC!TOTAL & ","
        cSQL = cSQL & "'NC',0)"

        msConn.Execute cSQL
        rsNC.MoveNext
    Loop
    rsNC.Close
    Set rsNC = Nothing
    
    'NOTAS DE DEBITO y CREDITO SIMPLES
    Do While Not rsNOTAS.EOF
        cSQL = "INSERT INTO LOLO "
        cSQL = cSQL & "(COD_PROV,NUMERO,FECHA,STAT,MONTO,PAGADO,TIPO,USUARIO)"
        cSQL = cSQL & " VALUES ("
        cSQL = cSQL & rsNOTAS!PROVEEDOR & ","
        cSQL = cSQL & rsNOTAS!NUMERO & ",'"
        cSQL = cSQL & rsNOTAS!FECHA & "',"
        cSQL = cSQL & "0,"
        cSQL = cSQL & rsNOTAS!MONTO & ","
        cSQL = cSQL & rsNOTAS!MONTO & ",'"
        cSQL = cSQL & rsNOTAS!TIPO & "','" & rsNOTAS!NDNC_SIMPLE & "')"
        
        msConn.Execute cSQL
        rsNOTAS.MoveNext
    Loop
    rsNOTAS.Close
    Set rsNOTAS = Nothing

    cSQL = "SELECT * FROM LOLO ORDER BY FECHA,TIPO,NUMERO"
    rsConso.Open cSQL, msConn, adOpenDynamic, adLockReadOnly
    On Error Resume Next
    rsConso.MoveFirst
    On Error GoTo 0
    Do Until rsConso.EOF
        If rsConso!TIPO = "CR" Then
            nspace1 = 6
            If cMesActual = Mid(rsConso!FECHA, 1, 6) Then
                'SI ES MES CORRIENTE DE LA FECHA LIMITE LO SUMA y LO IMPRIME
                nSumaDebito = nSumaDebito + (rsConso!MONTO - rsConso!PAGADO)
            Else
                'NO ES MES CORRIENTE
                If rsConso!STAT = 2 Then
                    'Y YA ESTA PAGADO LO SUMA, PERO NO LO IMPRIME
                    nSumaDebito = nSumaDebito + (rsConso!MONTO - rsConso!PAGADO)
                    GoTo AlProximo2:
                Else
                    ' NO ESTA PAGADO, LO SUMA Y LO IMPRIME
                    nSumaDebito = nSumaDebito + (rsConso!MONTO - rsConso!PAGADO)
                End If
            End If
        Else
            nspace1 = 19
            If cMesActual = Mid(rsConso!FECHA, 1, 6) Then
                '+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+
                'EXCEPCION PARA NC Y ND SIMPLE, SOLAMENTE SI ES MES ACTUAL
                '+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+
                If rsConso!USUARIO = -333 Then
                    Select Case rsConso!TIPO
                        Case "ND"
                            nSumaDebito = nSumaDebito + rsConso!MONTO
                        Case "NC"
                            nSumaCredito = nSumaCredito + rsConso!MONTO
                    End Select
                Else
                    'nSumaCredito = nSumaCredito + rsConso!MONTO
                End If
            Else
                GoTo AlProximo2:
            End If
        End If
        cFecha = (Mid(rsConso!FECHA, 7, 2) + "/" + Mid(rsConso!FECHA, 5, 2) + "/" + Mid(rsConso!FECHA, 1, 4))
        NSPACE = Len(rsConso!NUMERO)
        cString = cFecha & Space(11 - NSPACE) & rsConso!NUMERO & Space(2) & rsConso!TIPO
        MainMant.spDoc.TextAlign = SPTA_LEFT
        MainMant.spDoc.TextOut 300, iLin, cFecha
        MainMant.spDoc.TextOut 500, iLin, rsConso!NUMERO
        MainMant.spDoc.TextOut 710, iLin, rsConso!TIPO

        If rsConso!TIPO = "CR" Then
            MainMant.spDoc.TextAlign = SPTA_RIGHT
            MainMant.spDoc.TextOut 1300, iLin, Format(rsConso!MONTO, "###,###.00")
            NSPACE = Len(Format(rsConso!MONTO, "###,###.00"))
            nspace1 = 6
        Else
            MainMant.spDoc.TextAlign = SPTA_RIGHT
            MainMant.spDoc.TextOut 1000, iLin, Format(rsConso!MONTO, "###,###.00")
            NSPACE = Len(Format(rsConso!MONTO, "###,###.00"))
            nspace1 = 19
'            nSumaCredito = nSumaCredito + rsConso!MONTO
        End If

        cString = cString & Space(nspace1) & Space(10 - NSPACE) & Format(rsConso!MONTO, "###,###.00")
        
        NSPACE = Len(Format((nSumaDebito - nSumaCredito), "###,###.00"))
        If rsConso!TIPO = "CR" Then
            cString = cString & Space(22 + (10 - NSPACE)) & Format((nSumaDebito - nSumaCredito), "###,###.00")
            MainMant.spDoc.TextAlign = SPTA_RIGHT
            MainMant.spDoc.TextOut 1600, iLin, Format(rsConso!PAGADO, "###,###.00")
            MainMant.spDoc.TextOut 1900, iLin, Format((nSumaDebito - nSumaCredito), "###,###.00")
        Else
            cString = cString & Space(9 + (10 - NSPACE)) & Format((nSumaDebito - nSumaCredito), "###,###.00")
            MainMant.spDoc.TextAlign = SPTA_RIGHT
            MainMant.spDoc.TextOut 1900, iLin, Format((nSumaDebito - nSumaCredito), "###,###.00")
        End If
        MainMant.spDoc.TextAlign = SPTA_LEFT
        iLin = iLin + 50
AlProximo2:
        rsConso.MoveNext
        If iLin > 2400 Then
            MainMant.spDoc.Page = nPagina
            TituloCxC
        End If
    Loop
    NSPACE = Len(Format((nSumaDebito - nSumaCredito), "###,###.00"))
    iLin = iLin + 50
    MainMant.spDoc.TextAlign = SPTA_LEFT
    MainMant.spDoc.TextOut 300, iLin, "Saldo del Periodo : " & Format(nSumaDebito - nSumaCredito, "###,###.00")
    
    Dim rsDevHistory As New ADODB.Recordset
    rsDevHistory.Open "SELECT FECHA, NUM_DOC,SUM(CANT_EXT) AS CANTIDAD,SUM (COSTO_EXTENDIDO) AS COSTO" & _
        " FROM DEV_HISTORY " & _
        " WHERE PROV = " & nCodSel & _
        " AND FECHA <= '" & dF1 & "'" & _
        " GROUP BY FECHA, NUM_DOC " & _
        " ORDER BY FECHA", msConn, adOpenStatic, adLockOptimistic
    
    If rsDevHistory.EOF Then GoTo Salida_Ajustes_a_facturas:
    iLin = iLin + 100
    MainMant.spDoc.TextOut 300, iLin, "Ajustes a Facturas"
    iLin = iLin + 50
    MainMant.spDoc.TextOut 300, iLin, "FECHA"
    MainMant.spDoc.TextOut 550, iLin, "DOCUMENTO"
    MainMant.spDoc.TextOut 900, iLin, "CANTIDAD"
    MainMant.spDoc.TextOut 1300, iLin, "MONTO"
    iLin = iLin + 50
    MainMant.spDoc.TextOut 300, iLin, "-----------------------------------------------------------------------------------------------------"
    iLin = iLin + 50
    If iLin > 2400 Then
        MainMant.spDoc.Page = nPagina
        TituloCxC
    End If
    Do While Not rsDevHistory.EOF
        MainMant.spDoc.TextAlign = SPTA_LEFT
        MainMant.spDoc.TextOut 300, iLin, (Mid(rsDevHistory!FECHA, 7, 2) + "/" + Mid(rsDevHistory!FECHA, 5, 2) + "/" + Mid(rsDevHistory!FECHA, 1, 4))
        MainMant.spDoc.TextOut 550, iLin, rsDevHistory!NUM_DOC
        MainMant.spDoc.TextAlign = SPTA_RIGHT
        MainMant.spDoc.TextOut 1070, iLin, Format(rsDevHistory!Cantidad, "#.00")
        MainMant.spDoc.TextOut 1430, iLin, Format(rsDevHistory!COSTO, "#0.00")
        iLin = iLin + 50
        If iLin > 2400 Then
            MainMant.spDoc.Page = nPagina
            TituloCxC
        End If
        rsDevHistory.MoveNext
    Loop
Salida_Ajustes_a_facturas:
    rsDevHistory.Close
    Set rsDevHistory = Nothing
End If
rsConso.Close
rsRegCred.Close
msConn.Execute "DROP TABLE LOLO"
MainMant.spDoc.TextAlign = SPTA_LEFT
MainMant.spDoc.DoPrintPreview
cmbNombres.ListIndex = -1
nCodSel = 0

Call Seguridad

End Sub
Private Sub Command2_Click()
nCodSel = 0
If rsNombre.State = adStateOpen Then rsNombre.Close
If rsRegDebito.State = adStateOpen Then rsRegDebito.Close
If rsRegCred.State = adStateOpen Then rsRegCred.Close
If rsConso.State = adStateOpen Then rsConso.Close
Unload Me
End Sub

Private Sub DTPicker1_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then cmbNombres.SetFocus
End Sub

Private Sub Form_Load()
DTPicker1 = Format(Date, "SHORT DATE")

If PRINT_OPC = True Then
    'ABRE CLIENTES
    Me.Caption = "IMPRESION  DE ESTADO DE CUENTA. CLIENTES"
    Label1(2) = Label1(2) & " CLIENTE"
    rsNombre.Open "SELECT CODIGO,NOMBRE,APELLIDO,EMPRESA,SALDO " & _
        "FROM CLIENTES ORDER BY EMPRESA,NOMBRE,APELLIDO", msConn, adOpenDynamic, adLockReadOnly
Else
    'ABRE PROVEEDORES
    Me.Caption = "IMPRESION  DE ESTADO DE CUENTA. PROVEEDORES"
    Label1(2) = Label1(2) & " PROVEEDOR"
    rsNombre.Open "SELECT CODIGO,NOMBRE,APELLIDO,EMPRESA,SALDO " & _
        "FROM PROVEEDORES ORDER BY EMPRESA,NOMBRE,APELLIDO", msConn, adOpenDynamic, adLockReadOnly
End If
Do Until rsNombre.EOF
    If PRINT_OPC = True Then
        cmbNombres.AddItem rsNombre!NOMBRE & " " & rsNombre!APELLIDO & "-" & rsNombre!EMPRESA & Space(80) & rsNombre!CODIGO
    Else
        cmbNombres.AddItem rsNombre!EMPRESA & " (" & rsNombre!NOMBRE & " " & rsNombre!APELLIDO & ") " & Space(80) & rsNombre!CODIGO
    End If
    rsNombre.MoveNext
Loop
'nPagina = 1

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
        Command1.Enabled = False
    Case "N"        'SIN DERECHOS
        DTPicker1.Enabled = False: cmbNombres.Enabled = False
        Command1.Enabled = False
End Select
End Function
