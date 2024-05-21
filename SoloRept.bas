Attribute VB_Name = "SoloRept"
Public nLst As Integer
''''Public Sub PreparaImpresion(cSeleccion As Integer)
''''Dim MyString
''''Dim cDirecc As String
''''Dim cTexto As String
''''Dim RSREPORT As New ADODB.Recordset
''''
''''On Error Resume Next
''''cDirecc = App.Path & "\"
''''
''''Select Case cSeleccion
''''Case 0
''''nLst = 0
'''''Informacion de Inventario. Este reporte esta guardado en la Base de Datos
''''    RSREPORT.Open "SELECT C.DESCRIP AS DEPTO, a.Nombre, b.Descrip AS Envase, " & _
''''        " a.Cantidad AS Cant_Envase, a.Exist1 as Existencia_1, a.Exist2 AS Existencia_2, " & _
''''        " a.COSTO_EMPAQUE as Costo, a.ITBM" & _
''''        " FROM invent AS a, unidades AS b, DEP_INV AS C" & _
''''        " Where A.UNID_MEDIDA = b.id And A.COD_DEPT = C.CODIGO" & _
''''        " ORDER BY C.DESCRIP, b.descrip, a.nombre", msConn, adOpenStatic, adLockOptimistic
''''    Open cDirecc & "SOLOFILE.TXT" For Output As #1
''''    Print #1, "DEPARTAMENTO     NOMBRE           ENVASE  Cant_Envase  Exist 1  Exist 2     Costo   ITBM"
''''    Print #1, "----------------------------------------------------------------------------------------"
''''    Do Until RSRE?msconnPORT.EOF
''''        cTexto = Format$(Mid(StripOut(RSREPORT!depto, ","), 1, 15), "!@@@@@@@@@@@@@@@") & Space(2) & Format$(Mid(StripOut(RSREPORT!NOMBRE, ","), 1, 15), "!@@@@@@@@@@@@@@@") & Space(2) & Format$(Mid(RSREPORT!envase, 1, 10), "!@@@@@@@@@@") & Space(2) & Format$(Format(RSREPORT!cant_envase, "####0"), "@@@@@") & Space(2) & Format$(Format(RSREPORT!Existencia_1, "####.00"), "@@@@@@@@") & Space(2) & Format$(Format(RSREPORT!Existencia_2, "####.00"), "@@@@@@@@") & Space(2) & Format$(Format(RSREPORT!COSTO, "####.00"), "@@@@@@@@") & Space(2) & Format$(Format(RSREPORT!ITBM, "##.00"), "@@@@@")
''''        Print #1, cTexto
''''        RSREPORT.MoveNext
''''    Loop
''''    RSREPORT.Close
''''    Close #1
''''Case 1
''''nLst = 1
'''''Revision Inventario por Unidad
''''    RSREPORT.Open "SELECT B.DESCRIP AS DEPTO, A.NOMBRE, C.DESCRIP AS UNID_COMPRA, A.CANTIDAD AS CANT_COMPRA, D.DESCRIP AS UNID_CONSUMO, A.CANTIDAD2 AS CANT_CONSUMO " & _
''''        " FROM INVENT AS A, DEP_INV AS B, UNIDADES AS C, UNID_CONSUMO AS D " & _
''''        " Where A.COD_DEPT = B.CODIGO And A.UNID_MEDIDA = C.ID And A.UNID_CONSUMO = D.ID " & _
''''        " ORDER BY 1, 2 ", msConn, adOpenStatic, adLockOptimistic
''''    Open cDirecc & "SOLOFILE.TXT" For Output As #1
''''    Print #1, "DEPARTAMENTO     NOMBRE           UNID_COMPRA  Cant_Compra  UNID_CONSUMO  Cant_Consumo"
''''    Print #1, "--------------------------------------------------------------------------------------"
''''    Do Until RSREPORT.EOF
''''        cTexto = Format$(Mid(StripOut(RSREPORT!depto, ","), 1, 15), "!@@@@@@@@@@@@@@@") & Space(2) & Format$(Mid(StripOut(RSREPORT!NOMBRE, ","), 1, 15), "!@@@@@@@@@@@@@@@") & Space(2) & Format$(Mid(RSREPORT!UNID_COMPRA, 1, 10), "!@@@@@@@@@@") & Space(2) & Format$(Format(RSREPORT!cant_COMPRA, "####0"), "@@@@@") & Space(4) & Format$(Format(Mid(RSREPORT!UNID_CONSUMO, 1, 10)), "@@@@@@@@@@") & Space(4) & Format$(Format(RSREPORT!CANT_CONSUMO, "####.00"), "@@@@@@@@")
''''        Print #1, cTexto
''''        RSREPORT.MoveNext
''''    Loop
''''    RSREPORT.Close
''''    Close #1
''''Case 2
''''nLst = 2
'''''Revision Relacion Producto de Venta e Inventario"
''''    RSREPORT.Open "SELECT B.DESCRIP AS DEPTO, C.DESCRIP AS PRODUCTO_VENTA, D.DESCRIP AS DEPT_INV, E.NOMBRE AS ARTICULO_INVENT, CANT, F.DESCRIP AS CONSUMO" & _
''''        " FROM PLU_INVENT AS A, DEPTO AS B, PLU AS C, DEP_INV AS D, INVENT AS E, UNID_CONSUMO AS F" & _
''''        " WHERE A.ID_DEPT = B.CODIGO AND A.ID_PLU=C.CODIGO AND A.ID_DEPT_INV=D.CODIGO AND A.ID_PROD_INV=E.ID AND A.ID_UNID_CONSUMO = F.ID", msConn, adOpenStatic, adLockOptimistic
''''    Open cDirecc & "SOLOFILE.TXT" For Output As #1
''''    Print #1, "DEPARTAMENTO     PRODUCTO_VENTA    DEPT_INVENT  NOMBRE_PROD       CANT    UNID_CONSUMO"
''''    Print #1, "--------------------------------------------------------------------------------------"
''''    Do Until RSREPORT.EOF
''''        cTexto = Format$(Mid(StripOut(RSREPORT!depto, ","), 1, 15), "!@@@@@@@@@@@@@@@") & Space(2) & Format$(Mid(StripOut(RSREPORT!PRODUCTO_VENTA, ","), 1, 15), "!@@@@@@@@@@@@@@@") & Space(2) & Format$(Mid(RSREPORT!dept_inv, 1, 15), "!@@@@@@@@@@@@@@@") & Space(2) & Format$(Format(Mid(RSREPORT!ARTICULO_INVENT, 1, 15)), "!@@@@@@@@@@@@@@@") & Space(4) & Format$(Format(RSREPORT!CANT, "#####"), "@@@@@") & Space(4) & Format$(Format(Mid(RSREPORT!CONSUMO, 1, 8)), "!@@@@@@@@")
''''        Print #1, cTexto
''''        RSREPORT.MoveNext
''''    Loop
''''    RSREPORT.Close
''''    Close #1
''''Case 3
''''nLst = 3
''''    'INFO: CAMBIANDO A MOSTRAR LOS ULTIMOS 250 EVENTOS
''''    'LOG DE EVENTOS REGISTRADOS
''''    Dim nLineCounter As Long
''''    Dim nLineaReferencia As Long
''''
''''    On Error Resume Next
''''
''''    Open ADMIN_LOG For Input Shared As #66
''''    Do While Not EOF(66)
''''        Line Input #66, cTexto
''''        nLineCounter = nLineCounter + 1
''''    Loop
''''    nLineaReferencia = nLineCounter - 250
''''    If nLineaReferencia < 0 Then nLineaReferencia = 1
''''    Close #66
''''    nLineCounter = 0
''''    Open cDirecc & "SOLOFILE.TXT" For Output As #67
''''    Print #67, "˜˜˜ Se muestran los ultimos 250 eventos registrados ˜˜˜"
''''    Open ADMIN_LOG For Input Shared As #66
''''    Do While Not EOF(66)
''''        Line Input #66, cTexto
''''        nLineCounter = nLineCounter + 1
''''        If nLineCounter > nLineaReferencia Then
''''            Print #67, cTexto
''''        End If
''''    Loop
''''    Close #66
''''    Close #67
''''
''''    On Error GoTo 0
''''    ''''FileCopy ADMIN_LOG, cDirecc & "SOLOFILE.TXT"
''''End Select
''''ScreenReport.Show 1
''''End Sub

Public Function StripOut(ByVal From As String, ByVal What As String) As String
'Fuente: Tips1299.doc
    Dim i As Integer
    For i = 1 To Len(What)
        From = Replace(From, Mid$(What, i, 1), "")
    Next i
    StripOut = From
End Function

Public Function MesasPED(cOption As String) As Boolean
Dim cConnect As String
Select Case cOption
    Case "OPEN"
        Set msPED = New ADODB.Connection
        cConnect = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & DATA_PATH & "MESASPED.MDB"
        msPED.Open cConnect
    Case "CLOSE"
        msPED.Close
        Set msPED = Nothing
End Select

End Function

Public Function PrintFile(cTextFile As String, cTitulo As String) As Boolean
'INFO: LEE UN ARCHIVO DE TECTO Y LO MANDA A SWIFTPRINT
Dim nFileNumber As Integer
Dim a$
Dim nPagina As Integer
Dim iLin As Long

nFileNumber = FreeFile()
nPagina = 1
iLin = 200
MainMant.spDoc.DocBegin

MainMant.spDoc.TextAlign = SPTA_LEFT
MainMant.spDoc.WindowTitle = "Impresión de " & cTitulo
MainMant.spDoc.FirstPage = 1
MainMant.spDoc.PageOrientation = SPOR_PORTRAIT
MainMant.spDoc.Units = SPUN_LOMETRIC

Open App.Path & cTextFile For Input As #nFileNumber

Do Until EOF(nFileNumber)
    Line Input #nFileNumber, a$
    MainMant.spDoc.TextOut 300, iLin, a$
    iLin = iLin + 50
    If iLin > 2400 Then
        nPagina = nPagina + 1
        MainMant.spDoc.Page = nPagina
        MainMant.spDoc.TextAlign = SPTA_LEFT
        MainMant.spDoc.TextOut 300, 150, "Pagina # " & nPagina
        MainMant.spDoc.TextOut 300, 200, "Continua la Impresion de " & cTitulo
        MainMant.spDoc.TextOut 300, 250, String(40, "=")
        MainMant.spDoc.TextOut 300, 300, Space(1)
        iLin = 350
    End If
Loop
Close #nFileNumber
MainMant.spDoc.DoPrintPreview
End Function


'---------------------------------------------------------------------------------------
' Procedure : OpenGavetaDinero
' Author    : hsequeira
' Date      : 14/02/2017
' Date      : 4/08/2023
' Date      : 26/08/2023
' Purpose  : ABRE LA GAVETA VALIDANDO EL MODELO DE LA IMPRESORA
' UPDATE PARA LOS MODELOS 3NSTAR
'---------------------------------------------------------------------------------------
'
Public Function OpenGavetaDinero() As Boolean
Dim rc As Long

   On Error GoTo OpenGavetaDinero_Error


If Left(OPOS_DevName, 12) = "POSPrinter80" Or Left(OPOS_DevName, 10) = "SRP-E300" Or _
            Left(OPOS_DevName, 6) = "LR2000" Then
        rc = Sys_Pos.Cocash1.OpenDrawer
        Sys_Pos.Cocash1.WaitForDrawerClose 10000, 2000, 100, 1000
Else
        rc = Sys_Pos.Cocash1.DirectIO(DRW_DI_OPEN_DRAWER, 0, "")
End If

If rc <> 0 Then OpenGavetaDinero = False Else OpenGavetaDinero = True

   On Error GoTo 0
   Exit Function

OpenGavetaDinero_Error:

    EscribeLog "Error " & Err.Number & " (" & Err.Description & ") in procedure OpenGavetaDinero "

End Function


Public Function DD_SCREENREPORT(cSQL As String, cTitulo As String, Optional NumReporte As Integer) As Boolean
DD_PANTALLA.Show

DD_PANTALLA.Caption = cTitulo

Select Case NumReporte
    Case 1
        Call ReporteRecetas(cSQL)
    Case 2
    Case 3
    Case Else
End Select

End Function

Private Function ReporteRecetas(cSQL As String)
Dim rsREPORT As ADODB.Recordset
Dim nID As Integer
Dim nSumaCostoReceta As Single

Set rsREPORT = New ADODB.Recordset
rsREPORT.Open cSQL, msConn, adOpenStatic, adLockReadOnly


cData = ";;;;;" & Space(30) & "SUB TOTAL " & ";" & Format(nTotDesc, "STANDARD")
DD_PEDDETALLE.Rows.InsertAt i, sgFormatCharSeparatedValue, cData, ";"
i = i + 1
cData = ";;;;;" & Space(25) & "TOTAL GENERAL " & ";" & Format(nTotDescGEN, "STANDARD")
DD_PEDDETALLE.Rows.InsertAt i, sgFormatCharSeparatedValue, cData, ";"

nID = rsREPORT!ID_RECETA
Do While Not rsREPORT.EOF
    Do While nID = rsREPORT!ID_RECETA
        nSumaCostoReceta = nSumaCostoReceta + rsREPORT!COSTO
        rsREPORT.MoveNext
    Loop
    
    rsREPORT.MoveNext
Loop




With DD_PANTALLA.DD_PEDDETALLE
    
    .LoadArray rsREPORT.GetRows()
       ' define each column from the recordsets' fields collection
        For iLoop = 1 To rsREPORT.Fields.Count
           .Columns(iLoop).Caption = rsREPORT.Fields(iLoop - 1).Name
           .Columns(iLoop).DBField = rsREPORT.Fields(iLoop - 1).Name
           .Columns(iLoop).Key = rsREPORT.Fields(iLoop - 1).Name
        Next iLoop
    
    .ColumnClickSort = False
    .EvenOddStyle = sgEvenOddRows
    .ColorEven = vbWhite
        
    .ColorOdd = &HE0E0E0
'    .Columns(1).Width = 1000: .Columns(2).Width = 1400: .Columns(3).Width = 1000:
'    .Columns(4).Width = 0:
'    '.Columns(5).Width = 3200: .Columns(6).Width = 3200: .Columns(7).Width = 1000:
'    .Columns(1).Style.TextAlignment = sgAlignLeftCenter
'    .Columns(2).Style.TextAlignment = sgAlignLeftCenter
'    .Columns(3).Style.TextAlignment = sgAlignLeftCenter
'    '.Columns(2).Style.Format = "#,###"
'    '.Columns(3).Style.TextAlignment = sgAlignLeftCenter
'    '.Columns(3).Style.Format = "Standard"
'    .Columns(5).Style.TextAlignment = sgAlignLeftCenter
'    .Columns(6).Style.TextAlignment = sgAlignLeftCenter
'    .Columns(7).Style.TextAlignment = sgAlignRightCenter
'    .Columns(7).Style.Format = "Standard"
'    .Columns(8).Width = 0:
End With

End Function
