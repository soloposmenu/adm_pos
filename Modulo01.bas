Attribute VB_Name = "Modulo01"
'Constantes DRIVER ESPECIAL de impresora Epson TM-U950
Public Const EPSON_JOURNAL = 145    'Solo Journal
Public Const EPSON_RECEIPT = 146    'Solo Recibo
Public Const EPSON_AMBOS = 147      'Recibo y Journal
Public Const EPSON_REC_FULL_CUT = 157   ' Corte de Recibo Total
Public Const EPSON_REC_PART_CUT = 158   'Corte de Recibo parcial
Public Const EPSON_ABRE_CAJA = 130  'Abre la Bandeja de Dinero
'----------------------------------------------------
Public GENERICO_NO_JOURNAL As String
Public SOLOFAST_CUT As String
Public Const HOST_DB = ""
Public Const LOCAL_DB = ""
Public Const FACTURA_FILE = "SOLOFACT.TXT"
Public cFactFile As String
'-----------------------------
Public ON_LINE As Boolean       'Online/OffLine
'///////////Public DEFAULT_PRINTER As String
'Public Const ADMIN_LOG = "C:\WINDOWS\ADMLOG.SOL"
Public ADMIN_LOG As String
'''Public Const MAX_ACOMP = 8
Public MAX_ACOMP As Integer
Public DATA_PATH As String
Public cDataPath As String
Public cShapeADOString As String
''''''*******Public Const DATA_PATH = "\\SOLO11\ACCESS\"
Public SLIP_OK As Boolean
Public REPCAJAX_OK As Boolean
Public HABITACION As Boolean
Public OPC_SOLOFAST As Boolean
Public NOM_ADMINISTRADOR As String
Public TIPO_ADMINISTRADOR As Integer
Public TipoApplicacion As String
Public OPEN_PROPINA As Boolean
Public PROPINA_DESCRIP As String
Public HABITACION_OK As Boolean
Public OPOS_DevName As String

Public Function GetNEWIndice(cTablaIndice As String) As Long
Dim cSQL As String
Dim rsINDICES As ADODB.Recordset

cSQL = "SELECT INDICE FROM INDICES WHERE TABLA_CAMPO = '" & cTablaIndice & "'"
Set rsINDICES = New ADODB.Recordset

rsINDICES.Open cSQL, msConn, adOpenStatic, adLockOptimistic
If rsINDICES.EOF Then
    ShowMsg "ESTA  TABLA (" & cTablaIndice & ") NO TIENE INDICES DEFINIDOS", vbYellow, vbRed
    Exit Function
End If

GetNEWIndice = rsINDICES!indice + 1

msConn.BeginTrans
msConn.Execute "UPDATE INDICES SET INDICE = " & rsINDICES!indice + 1 & " WHERE TABLA_CAMPO = '" & cTablaIndice & "'"
msConn.CommitTrans

rsINDICES.Close
Set rsINDICES = Nothing

End Function

'---------------------------------------------------------------------------------------
' Procedure : RepiteFactura
' Author    : hsequeira
' Date      : 06/07/2012
' Purpose   : REPITE LA ULTIMA FACTURA, SI ES POSIBLE
' YA NO LO HACE EN LA IMPRESORA FISCAL, LO HACE EN
' LA IMPRESORA REGULAR
'---------------------------------------------------------------------------------------
'
Public Function RepiteFactura(ccTitulo As String, ccFile As String) As Boolean
Dim nLastTrans As Long, cLastFISCAL As String
Dim cSQL As String, cSQL2 As String, cSQL3 As String
Dim rsLAST As ADODB.Recordset
Dim rsLAST_Pago As ADODB.Recordset
Dim rsLAST_Prop As ADODB.Recordset
Dim nFreefile As Long
Dim nSubTotal As Single, nLastTax As Single, nTotPROP As Single
Dim cTexto As String
Dim i As Integer
Dim cCopiaCliente, cCopiaRUC

'INFO: CAMBIO. LA COPIA LA IMPRIME LA IMPRESORA NO-FISCAL
'05JUL2012
'DE LA TABLA ORGANIZACION. OBTENER TRANS

nFreefile = FreeFile()

On Error Resume Next
Kill ccFile
On Error GoTo 0

On Error GoTo ErrAdm:
Open ccFile For Output As #nFreefile
'21DIC2014
nLastTrans = rs00!TRANS_FAST
cLastFISCAL = GetLastFiscal(nLastTrans)

Set rsLAST = New ADODB.Recordset
cSQL = "SELECT * FROM TRANSAC WHERE NUM_TRANS = " & nLastTrans & " ORDER BY LIN"
rsLAST.Open cSQL, msConn, adOpenStatic, adLockOptimistic

If rsLAST.EOF Then
    ShowMsg "ES IMPOSIBLE IMPRIMIR COPIA DE LA ULTIMA FACTURA", vbYellow, vbRed
    EscribeLog "Ventas. ES IMPOSIBLE IMPRIMIR COPIA DE LA ULTIMA FACTURA = " & nLastTrans
    rsLAST.Close
    Set rsLAST = Nothing
    Close #nFreefile
    Exit Function
End If

Set rsLAST_Pago = New ADODB.Recordset
cSQL2 = "SELECT  A.TIPO_PAGO, B.DESCRIP, A.MONTO "
cSQL2 = cSQL2 & " FROM TRANSAC_PAGO AS A, PAGOS AS B"
cSQL2 = cSQL2 & " WHERE A.NUM_TRANS = " & nLastTrans
cSQL2 = cSQL2 & " AND A.TIPO_PAGO = B.CODIGO "
cSQL2 = cSQL2 & " ORDER BY A.LIN"
rsLAST_Pago.Open cSQL2, msConn, adOpenStatic, adLockOptimistic

If rsLAST_Pago.EOF Then
    ShowMsg "ES IMPOSIBLE IMPRIMIR COPIA DE LA ULTIMA FACTURA", vbYellow, vbRed
    EscribeLog "Ventas. ES IMPOSIBLE IMPRIMIR COPIA DE LA ULTIMA FACTURA = " & nLastTrans
    rsLAST.Close
    rsLAST_Pago.Close
    Set rsLAST = Nothing
    Set rsLAST_Pago = Nothing
    Close #nFreefile
    Exit Function
End If

Print #nFreefile, ccTitulo
'---------------------------------------------------------------------------------------
'INFO: 18SEP2012
'Imprime RUC y Nombre del Cliente
'---------------------------------------------------------------------------------------

If rsLAST!CUENTA = 0 Then
    Print #nFreefile, "Mesa: " & rsLAST!MESA
Else
    Print #nFreefile, "Mesa: " & rsLAST!MESA & " - Cuenta: " & rsLAST!CUENTA
End If
Print #nFreefile, "REFERENCIA # " & cLastFISCAL

cCopiaCliente = RegRead("HKCU\Software\SoloSoftware\SoloMix\LastCliente")
cCopiaRUC = RegRead("HKCU\Software\SoloSoftware\SoloMix\LastRUC")

If cCopiaRUC <> "" Then Print #nFreefile, "RUC/CIP: " & Left(cCopiaRUC, 20)
If cCopiaCliente <> "" Then Print #nFreefile, "Cliente: " & Left(cCopiaCliente, 20)
'---------------------------------------------------------------------------------------
'---------------------------------------------------------------------------------------

Print #nFreefile, Format(Date, "DD/MM/YYYY") & " - " & Format(Time(), "HH:MM")
Print #nFreefile, Space(2)
'Print #nFreefile, Space(2)
Print #nFreefile, ccTitulo
Print #nFreefile, String(30, "=")

Do While Not rsLAST.EOF
    cTexto = Format(Left(rsLAST!DESCRIP, 15), "!@@@@@@@@@@@@@@@") & Space(2)
    cTexto = cTexto & Format(Format(rsLAST!CANT, "###"), "@@@") & Space(2)
    cTexto = cTexto & Format(Format(rsLAST!precio, "STANDARD"), "@@@@@@@@")
    Print #nFreefile, cTexto
    nSubTotal = nSubTotal + rsLAST!precio
    nLastTax = nLastTax + rsLAST!precio * (rsLAST!CON_TAX / 100)
    rsLAST.MoveNext
Loop

Print #nFreefile, String(30, "=")
Print #nFreefile, Space(2)
Print #nFreefile, ccTitulo
Print #nFreefile, Space(2)
Print #nFreefile, Format("SUB-TOTAL: ", "!@@@@@@@@@@@@@@@") & Space(2) & Format(Format(nSubTotal, "STANDARD"), "@@@@@@@@@@")
Print #nFreefile, Format("IMPUESTO: ", "!@@@@@@@@@@@@@@@") & Space(2) & Format(Format(nLastTax, "STANDARD"), "@@@@@@@@@@")
Print #nFreefile, Format("TOTAL: ", "!@@@@@@@@@@@@@@@") & Space(2) & Format(Format(nSubTotal + nLastTax, "CURRENCY"), "@@@@@@@@@@")
Print #nFreefile, Space(2)
'Print #nFreefile, Space(2)
Print #nFreefile, ccTitulo
'Print #nFreefile, Space(2)
Print #nFreefile, "=== PAGOS ==="
'Print #nFreefile, Space(2)

Do While Not rsLAST_Pago.EOF
    cTexto = Format(Left(rsLAST_Pago!DESCRIP, 15), "!@@@@@@@@@@@@@@@") & Space(2)
    cTexto = cTexto & Format(Format(rsLAST_Pago!MONTO, "STANDARD"), "@@@@@@@@@@")
    Print #nFreefile, cTexto
    rsLAST_Pago.MoveNext
Loop

Set rsLAST_Prop = New ADODB.Recordset
cSQL3 = "SELECT  A.TIPO_PAGO, B.DESCRIP, A.MONTO "
cSQL3 = cSQL3 & " FROM TRANSAC_PROP AS A, PAGOS AS B"
cSQL3 = cSQL3 & " WHERE A.NUM_TRANS = " & nLastTrans
cSQL3 = cSQL3 & " AND A.TIPO_PAGO = B.CODIGO "
cSQL3 = cSQL3 & " ORDER BY A.LIN"

rsLAST_Prop.Open cSQL3, msConn, adOpenStatic, adLockOptimistic

If rsLAST_Prop.EOF Then
    rsLAST.Close
    rsLAST_Pago.Close
    rsLAST_Prop.Close
    Set rsLAST = Nothing
    Set rsLAST_Pago = Nothing
    Set rsLAST_Prop = Nothing
Else
    
    Print #nFreefile, Space(2)
    Print #nFreefile, ccTitulo
    Print #nFreefile, "=== " & UCase(GetFromINI("Fiscal", "TextoPropina", App.Path & "\soloini.ini")) & " ==="
    'Print #nFreefile, Space(2)
    
    Do While Not rsLAST_Prop.EOF
        cTexto = Format(Left(rsLAST_Prop!DESCRIP, 15), "!@@@@@@@@@@@@@@@") & Space(2)
        cTexto = cTexto & Format(Format(rsLAST_Prop!MONTO, "STANDARD"), "@@@@@@@@@@")
        Print #nFreefile, cTexto
        nTotPROP = nTotPROP + rsLAST_Prop!MONTO
        rsLAST_Prop.MoveNext
    Loop
    
    Print #nFreefile, Format("TOTAL: ", "!@@@@@@@@@@@@@@@") & Space(2) & _
                                  Format(Format(nSubTotal + nLastTax + nTotPROP, "CURRENCY"), "@@@@@@@@@@")
    Print #nFreefile, Space(2)
    Print #nFreefile, ccTitulo
    
    rsLAST.Close
    rsLAST_Pago.Close
    rsLAST_Prop.Close
    Set rsLAST = Nothing
    Set rsLAST_Pago = Nothing
    Set rsLAST_Prop = Nothing

End If

For i = 1 To 10
    Print #nFreefile, Space(2)
Next

Close #nFreefile
RepiteFactura = True

On Error GoTo 0
Exit Function
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
ErrAdm:
    ShowMsg "Ventas. RepiteFactura. " & Err.Number & " - " & Err.Description
    On Error Resume Next
    Close #nFreefile
    On Error GoTo 0
    RepiteFactura = False
End Function
'-------------------------------------------------------------------------------------------------------------------------
' Procedure : Print2_OPOS_Dev
' Author    : hsequeira
' Date      : 27/05/2012
' Purpose   : IMPRIME DE FORMA CORRECTA EN LOS ROLLOS DE LA IMPRESORA QUE ESTA CONECTADA.
'                    YA SEA UNA GRANDE (950) o UNA CHICA
' PARAMETROS: LOS DATOS QUE SE DESEAN IMPRIMIR
' 24OCT2016. A PARTIR DE AHORA, SI ES LA 950, UNICAMENTE VA A IMPRIMIR EN UN ROLLO
' UnRollo = True
'--------------------------------------------------------------------------------------------------------------------------
'
Public Function Print2_OPOS_Dev(cParams As String, Optional UnRollo As Boolean) As Boolean

If cParams = Space(1) Or cParams = Space(2) Then
    'VIENE UNA LINEA EN BLANCO
    cParams = Chr(&HD) & Chr(&HA)
End If

Select Case OPOS_DevName
    Case "SRP-350plus"  'INFO: 20SEP2013
        If cParams = Chr(&HD) & Chr(&HA) Then
            Sys_Pos.Coptr1.PrintNormal PtrSReceipt, Chr(&HD)
        Else
            Sys_Pos.Coptr1.PrintNormal PtrSReceipt, cParams & Chr(&HD)
        End If
    Case "LR3000", "TM-U200B", "SRP270", "MP4200TH", "TM-T20E", "TM-T20U", "TM-U220B", "TM-U200BP", "TM-U220BP"
        If cParams = Chr(&HD) & Chr(&HA) Then
            'RptCajas.Coptr1.PrintNormal FPtrSReceipt, cParams
            Sys_Pos.Coptr1.PrintNormal PtrSReceipt, cParams
        Else
            'RptCajas.Coptr1.PrintNormal FPtrSReceipt, cParams & Chr(&HD) & Chr(&HA)
            Sys_Pos.Coptr1.PrintNormal PtrSReceipt, cParams & Chr(&HD) & Chr(&HA)
        End If
    Case "TM-U950P", "TM-U950"
        UnRollo = True
        If cParams = Chr(&HD) & Chr(&HA) Then
            'RptCajas.Coptr1.PrintTwoNormal FptrSJournalReceipt , Space(1), Space(1)
            If UnRollo Then
                Sys_Pos.Coptr1.PrintNormal PtrSReceipt, cParams
            Else
                Sys_Pos.Coptr1.PrintTwoNormal PtrSJournalReceipt, Space(1), Space(1)
            End If
        Else
            'RptCajas.Coptr1.PrintTwoNormal FptrSJournalReceipt , cParams, cParams
            If UnRollo Then
                Sys_Pos.Coptr1.PrintNormal PtrSReceipt, cParams & Chr(&HD) & Chr(&HA)
            Else
                Sys_Pos.Coptr1.PrintTwoNormal PtrSJournalReceipt, cParams, cParams
            End If
        End If
    Case Else
        If cParams = Chr(&HD) & Chr(&HA) Then
            'RptCajas.Coptr1.PrintNormal FPtrSReceipt, cParams
            Sys_Pos.Coptr1.PrintNormal PtrSReceipt, cParams
        Else
            'RptCajas.Coptr1.PrintNormal FPtrSReceipt, cParams & Chr(&HD) & Chr(&HA)
            Sys_Pos.Coptr1.PrintNormal PtrSReceipt, cParams & Chr(&HD) & Chr(&HA)
        End If
End Select
Eval_OPOS_Dev (Sys_Pos.Coptr1.State)
End Function

'---------------------------------------------------------------------------------------
' Procedure : Eval_OPOS_Dev
' Author    : hsequeira
' Date      : 27/06/2012
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Function Eval_OPOS_Dev(rc As Long) As Boolean
Select Case rc
    Case OPOS_S_CLOSED
        'Debug.Print "OPOS_S_CLOSED - " & Sys_Pos.Coptr1.ResultCode
    Case OPOS_S_IDLE
        'Debug.Print "OPOS_S_IDLE - " & Sys_Pos.Coptr1.ResultCode
    Case OposSBusy
        'Debug.Print "OposSBusy - " & Sys_Pos.Coptr1.ResultCode
    Case OPOS_S_ERROR
        'Debug.Print "OPOS_S_ERROR - " & Sys_Pos.Coptr1.ResultCode
    Case Else
End Select
End Function

'---------------------------------------------------------------------------------------
' Procedure : GetLastFiscal
' Author    : hsequeira
' Date      : 05/07/2012
' Purpose   : OBTIENE EL NUMERO DE LA ULTIMA FACTURA FISCAL
'---------------------------------------------------------------------------------------
'
Private Function GetLastFiscal(nTRansaccion As Long) As String
Dim rsTransFiscal As ADODB.Recordset

Set rsTransFiscal = New ADODB.Recordset

rsTransFiscal.Open "SELECT FISCAL FROM TRANSAC_FISCAL WHERE DOC_SOLO = " & nTRansaccion, msConn, adOpenStatic, adLockOptimistic
If rsTransFiscal.EOF Then
    GetLastFiscal = "00000000"
Else
    GetLastFiscal = rsTransFiscal!FISCAL
End If
rsTransFiscal.Close
Set rsTransFiscal = Nothing
End Function
