VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomct2.ocx"
Begin VB.Form conHistorico 
   BackColor       =   &H00B39665&
   Caption         =   "Consulta Histórica"
   ClientHeight    =   4830
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6240
   Icon            =   "conHistorico.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   4830
   ScaleWidth      =   6240
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdSeek 
      Caption         =   "Ver Historial"
      Height          =   495
      Left            =   4920
      TabIndex        =   4
      Top             =   4150
      Width           =   1215
   End
   Begin VB.ListBox List1 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2940
      Left            =   120
      TabIndex        =   3
      Top             =   1680
      Width           =   4695
   End
   Begin VB.ComboBox Combo1 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   120
      Style           =   2  'Dropdown List
      TabIndex        =   2
      Top             =   1080
      Width           =   4695
   End
   Begin MSComCtl2.DTPicker txtFecIni 
      Height          =   345
      Left            =   1320
      TabIndex        =   0
      Top             =   240
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   609
      _Version        =   393216
      Format          =   16842753
      CurrentDate     =   36418
   End
   Begin MSComCtl2.DTPicker txtFecFin 
      Height          =   345
      Left            =   3840
      TabIndex        =   1
      Top             =   240
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   609
      _Version        =   393216
      Format          =   16842753
      CurrentDate     =   36418
   End
   Begin VB.Label Label1 
      BackColor       =   &H00B39665&
      Caption         =   "Producto"
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
      TabIndex        =   8
      Top             =   1440
      Width           =   2415
   End
   Begin VB.Label Label1 
      BackColor       =   &H00B39665&
      Caption         =   "Departamento"
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
      Left            =   120
      TabIndex        =   7
      Top             =   840
      Width           =   2415
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
      Top             =   360
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
      TabIndex        =   5
      Top             =   360
      Width           =   1215
   End
End
Attribute VB_Name = "conHistorico"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private rsDepInv As New ADODB.Recordset
Private iLin As Integer
Private nPagina As Integer

Private Sub cmdSeek_Click()
Dim vResp
If List1.Text = "" Then
    MsgBox "No se ha seleccionado nada", vbInformation, Me.Caption
    Exit Sub
End If
vResp = MsgBox("¿ Desea ver historial de " & Trim(Left(List1.Text, 70)) & " ?", vbQuestion + vbYesNo, BoxTit)
If vResp = vbYes Then
    Me.MousePointer = vbHourglass
    Dim rsHistorial As New ADODB.Recordset
    Dim rsTemp As New ADODB.Recordset
    Dim rsHistInvent As New ADODB.Recordset
    Dim IDInvent As Integer
    Dim cSQL As String
    Dim dF1 As String
    Dim dF2 As String
    Dim nExist1 As Single, nExist2 As Single
    Dim nCostoUnit As Single, nCostoEmpaque As Single
    Dim nCantidadPeriodo  As Single
    Dim cPrevMes As String
    Dim nLBod1 As Single
    Dim nLBod2 As Single
    Dim nValBod1 As Single
    Dim nValBod2 As Single
    Dim nLBodegas As Single
    Dim nValBodegas As Single
    Dim nLCosto As Single
    Dim cMedida As String, cConsumo As String
    Dim nCantidadPorCompras As Single
    Dim GET_BODEGA As Integer

    dF1 = Format(txtFecIni, "YYYYMMDD")
    dF2 = Format(txtFecFin, "YYYYMMDD")

    IDInvent = Val(Right(List1.Text, 10))
    With rsHistorial
        .CursorLocation = adUseClient
        .Fields.Append "ARTICULO", adChar, 25, adFldUpdatable
        .Fields.Append "TIPO_MOV", adChar, 10, adFldUpdatable
        .Fields.Append "FECHA", adChar, 8, adFldUpdatable
        .Fields.Append "HORA", adChar, 5, adFldUpdatable
        .Fields.Append "USER", adInteger, , adFldUpdatable
        .Fields.Append "CANTIDAD", adSingle, , adFldUpdatable
        .Fields.Append "CANTIDAD_EXT", adSingle, , adFldUpdatable
        .Fields.Append "EN_BODEGA1", adSingle, , adFldUpdatable
        .Fields.Append "EN_BODEGA2", adSingle, , adFldUpdatable
        .Fields.Append "VALOR_BOD1", adSingle, , adFldUpdatable
        .Fields.Append "VALOR_BOD2", adSingle, , adFldUpdatable
        .CursorType = adOpenKeyset
        .LockType = adLockOptimistic
        .Open
        .Sort = "FECHA ASC, TIPO_MOV ASC"
    End With
    
    cPrevMes = Format(Month(conHistorico.txtFecIni) - 1, "00")
    If cPrevMes = "00" Then cPrevMes = "12"
    
    'INFO: ENE2010
    'OBTIENE LOS VALORES DEL MES ANTERIOR, SEGUN LA FECHA INCIAL SELECCIONADA
    cSQL = "SELECT ID,BOD1_" & cPrevMes & " AS BODEGA1,"
    cSQL = cSQL & "BOD2_" & cPrevMes & " AS BODEGA2,"
    cSQL = cSQL & "COSTO_" & cPrevMes & " AS COSTO,"
    cSQL = cSQL & "TOTAL_" & cPrevMes & " AS TOTAL,"
    cSQL = cSQL & "FECHA, HORA, USUARIO "
    cSQL = cSQL & " FROM HIST_INVENT, MESES"
    cSQL = cSQL & " WHERE ID = " & IDInvent

    rsHistInvent.Open cSQL, msConn, adOpenStatic, adLockOptimistic
    
    'Recontruye VENTAS (HIST_TR)
    'Y LAS GUARDA EN LA TABLA VIRTUAL (rsHistorial)
    cSQL = "SELECT HIST_TR.FECHA, SUM(HIST_TR.CANT * PLU_INVENT.CANT) AS CANTIDAD, "
    cSQL = cSQL & " SUM(HIST_TR.CANT) * MAX(INVENT.COSTO) AS VENTAS "
    cSQL = cSQL & " FROM PLU_INVENT, INVENT, HIST_TR "
    cSQL = cSQL & " WHERE PLU_INVENT.ID_PROD_INV = " & IDInvent
    cSQL = cSQL & " AND PLU_INVENT.ID_PROD_INV = INVENT.ID "
    cSQL = cSQL & " AND PLU_INVENT.ID_PLU = HIST_TR.PLU "
    cSQL = cSQL & " AND HIST_TR.FECHA >= '" & dF1 & "'"
    cSQL = cSQL & " AND HIST_TR.FECHA <= '" & dF2 & "'"
    cSQL = cSQL & " GROUP BY HIST_TR.FECHA "
    cSQL = cSQL & " ORDER BY HIST_TR.FECHA "
    
    rsTemp.Open cSQL, msConn, adOpenStatic, adLockOptimistic
    Do While Not rsTemp.EOF
        With rsHistorial
            .AddNew
            !FECHA = rsTemp!FECHA
            !TIPO_MOV = "VTA"
            !CANTIDAD_EXT = rsTemp!Cantidad
            !VALOR_BOD1 = rsTemp!Ventas
            .Update
        End With
        'Debug.Print rsTemp!FECHA, rsTemp!Cantidad, rsTemp!VENTAS
        rsTemp.MoveNext
    Loop
    rsTemp.Close
    'Debug.Print "****************************************"
    
    'Recontruye COMPRAS (COMPRAS_HEAD)
    'Y LAS GUARDA EN LA TABLA VIRTUAL (rsHistorial)
    cSQL = "SELECT FECHA, SUM(B.UNIDADES) AS CANTIDAD, SUM(B.COSTO_IN) AS COSTO "
    cSQL = cSQL & " FROM COMPRAS_HEAD AS A, COMPRAS_DETA AS B "
    cSQL = cSQL & " WHERE B.CODI_INV = " & IDInvent
    cSQL = cSQL & " AND A.FECHA >= '" & dF1 & "'"
    cSQL = cSQL & " AND A.FECHA <= '" & dF2 & "'"
    cSQL = cSQL & " AND B.NUM_COMPRA = A.INDICE "
    cSQL = cSQL & " GROUP BY A.FECHA "
    cSQL = cSQL & " ORDER BY A.FECHA "
    'Debug.Print cSQL
    
    GET_BODEGA = GetFromINI("Administracion", "ActualizaBodega", App.Path & "\soloini.ini")
    rsTemp.Open cSQL, msConn, adOpenStatic, adLockOptimistic
    Do While Not rsTemp.EOF
        With rsHistorial
            .AddNew
            !FECHA = rsTemp!FECHA
            !TIPO_MOV = "COM"
            !CANTIDAD_EXT = rsTemp!Cantidad
            !VALOR_BOD1 = rsTemp!COSTO
            .Update
        End With
        'Debug.Print rsTemp!FECHA, rsTemp!Cantidad, rsTemp!COSTO
        rsTemp.MoveNext
    Loop
    rsTemp.Close
    'Debug.Print "****************************************"
    
    'Recontruye DEVOLUCIONES (DEV_HISTORY)
    'Y LAS GUARDA EN LA TABLA VIRTUAL (rsHistorial)
    cSQL = "SELECT FECHA, SUM(CANT_EXT) AS CANTIDAD, SUM(COSTO_EXTENDIDO) AS COSTO "
    cSQL = cSQL & " FROM DEV_HISTORY "
    cSQL = cSQL & " WHERE COD_INV = " & IDInvent
    cSQL = cSQL & " AND FECHA >= '" & dF1 & "'"
    cSQL = cSQL & " AND FECHA <= '" & dF2 & "'"
    cSQL = cSQL & " GROUP BY FECHA "
    cSQL = cSQL & " ORDER BY FECHA "
    'Debug.Print cSQL
    rsTemp.Open cSQL, msConn, adOpenStatic, adLockOptimistic
    Do While Not rsTemp.EOF
        With rsHistorial
            .AddNew
            !FECHA = rsTemp!FECHA
            !TIPO_MOV = "DEV"
            !CANTIDAD_EXT = rsTemp!Cantidad
            !VALOR_BOD1 = rsTemp!COSTO
            .Update
        End With
        'Debug.Print rsTemp!FECHA, rsTemp!Cantidad, rsTemp!COSTO
        rsTemp.MoveNext
    Loop
    rsTemp.Close
    
    'Recontruye DEVOLUCIONES (DEVOLUCION_HEAD)
    'Y LAS GUARDA EN LA TABLA VIRTUAL (rsHistorial)
    cSQL = "SELECT FECHA_SISTEMA, SUM(CANT) AS CANTIDAD, SUM(B.COSTO_UNIT+B.ITBM) AS COSTO "
    cSQL = cSQL & " FROM DEVOLUCION_HEAD AS A, DEVOLUCION_DETA AS B "
    cSQL = cSQL & " WHERE B.COD_INV = " & IDInvent
    cSQL = cSQL & " AND A.FECHA_SISTEMA >= '" & dF1 & "'"
    cSQL = cSQL & " AND A.FECHA_SISTEMA <= '" & dF2 & "'"
    cSQL = cSQL & " AND A.NUM_DOC = B.NUM_DOC "
    cSQL = cSQL & " GROUP BY A.FECHA_SISTEMA "
    cSQL = cSQL & " ORDER BY A.FECHA_SISTEMA "
    'Debug.Print cSQL
    rsTemp.Open cSQL, msConn, adOpenStatic, adLockOptimistic
    Do While Not rsTemp.EOF
        With rsHistorial
            .AddNew
            !FECHA = rsTemp!FECHA_SISTEMA
            !TIPO_MOV = "DEF"
            !CANTIDAD_EXT = rsTemp!Cantidad
            !VALOR_BOD1 = rsTemp!COSTO
            .Update
        End With
        'Debug.Print rsTemp!FECHA, rsTemp!Cantidad, rsTemp!COSTO
        rsTemp.MoveNext
    Loop
    rsTemp.Close
    
    'Recontruye TRANSFERENCIAS & MERMAS (TRANSFERENCIA)
    'Y LAS GUARDA EN LA TABLA VIRTUAL (rsHistorial)
    cSQL = "SELECT FECHA, TIPO, MAX(HORA) AS LAHORA, "
    cSQL = cSQL & "SUM(CANTIDAD) AS LACANTIDAD, "
    cSQL = cSQL & "SUM(CANTIDAD_EXT) AS EXTENDIDA, "
    cSQL = cSQL & "SUM(COSTO_BODEGA1 + COSTO_BODEGA2) AS COSTO "
    cSQL = cSQL & " FROM TRANSFERENCIA "
    cSQL = cSQL & " WHERE ID = " & IDInvent
    cSQL = cSQL & " AND FECHA >= '" & dF1 & "'"
    cSQL = cSQL & " AND FECHA <= '" & dF2 & "'"
    cSQL = cSQL & " GROUP BY FECHA, TIPO "
    cSQL = cSQL & " ORDER BY 1,3 "
    'Debug.Print cSQL
    rsTemp.Open cSQL, msConn, adOpenStatic, adLockOptimistic
    Do While Not rsTemp.EOF
        With rsHistorial
            .AddNew
            !FECHA = rsTemp!FECHA
            !TIPO_MOV = rsTemp!TIPO
            !Cantidad = rsTemp!LACANTIDAD
            !CANTIDAD_EXT = rsTemp!LACANTIDAD + rsTemp!EXTENDIDA
            !VALOR_BOD1 = rsTemp!COSTO
            .Update
        End With
        'Debug.Print rsTemp!FECHA, rsTemp!TIPO, rsTemp!LACantidad, rsTemp!COSTO
        rsTemp.MoveNext
    Loop
    rsTemp.Close
    
    'Recontruye     NOTAS DE CREDITO
    'Y LAS GUARDA EN LA TABLA VIRTUAL (rsHistorial)
    cSQL = "SELECT FECHA_DOC, SUM(CANTIDAD) AS LACANTIDAD, SUM(COSTO_UNIT + ITBM) AS COSTO "
    cSQL = cSQL & " FROM NC "
    cSQL = cSQL & " WHERE INVENT = " & IDInvent
    cSQL = cSQL & " AND FECHA_DOC >= '" & dF1 & "'"
    cSQL = cSQL & " AND FECHA_DOC <= '" & dF2 & "'"
    cSQL = cSQL & " GROUP BY FECHA_DOC "
    cSQL = cSQL & " ORDER BY FECHA_DOC "
    'Debug.Print cSQL
    rsTemp.Open cSQL, msConn, adOpenStatic, adLockOptimistic
    Do While Not rsTemp.EOF
        With rsHistorial
            .AddNew
            !FECHA = rsTemp!FECHA_DOC
            !TIPO_MOV = "NC"
            !CANTIDAD_EXT = rsTemp!LACANTIDAD
            !VALOR_BOD1 = rsTemp!COSTO
            .Update
        End With
        'Debug.Print rsTemp!FECHA, rsTemp!TIPO, rsTemp!LACantidad, rsTemp!COSTO
        rsTemp.MoveNext
    Loop
    rsTemp.Close
    
    '****************

    On Error GoTo ErrAdm:
    
    'BUSCA DATOS DEL ARTICULO SELECCIONADO
    cSQL = "SELECT A.ID, B.DESCRIP AS UNID_MEDIDA, C.DESCRIP AS UNID_CONSUMO, A.CANTIDAD2 "
    cSQL = cSQL & " FROM INVENT AS A, UNIDADES AS B, UNID_CONSUMO AS C "
    cSQL = cSQL & " WHERE A.ID = " & IDInvent
    cSQL = cSQL & " AND A.UNID_MEDIDA = B.ID "
    cSQL = cSQL & " AND A.UNID_CONSUMO = C.ID "
    rsTemp.Open cSQL, msConn, adOpenStatic, adLockOptimistic
    cMedida = rsTemp!UNID_MEDIDA
    cConsumo = rsTemp!UNID_CONSUMO
    nCantidadPorCompras = rsTemp!CANTIDAD2
    rsTemp.Close
    
    nPagina = 0
    MainMant.spDoc.DocBegin
    Call PrintTit(cMedida, cConsumo)
    EscribeLog ("Admin." & "Impresión de Historial " & Trim(Left(List1.Text, 70)))
            
    MainMant.spDoc.TextAlign = SPTA_LEFT
    cFecha = Right(rsHistInvent!FECHA, 2) & "/" & Mid(rsHistInvent!FECHA, 5, 2) & "/" & Left(rsHistInvent!FECHA, 4)
    MainMant.spDoc.TextOut 300, iLin, cFecha
    MainMant.spDoc.TextOut 500, iLin, "CIERRE MES"
    MainMant.spDoc.TextAlign = SPTA_RIGHT
    
    MainMant.spDoc.TextOut 1200, iLin, Format(rsHistInvent!BODEGA1, "#.00")
    MainMant.spDoc.TextOut 1350, iLin, Format(((rsHistInvent!BODEGA1) * rsHistInvent!COSTO), "#0.00")
    
    MainMant.spDoc.TextOut 1500, iLin, Format(rsHistInvent!BODEGA2, "#.00")
    MainMant.spDoc.TextOut 1650, iLin, Format(((rsHistInvent!BODEGA2) * rsHistInvent!COSTO), "#0.00")
    
    MainMant.spDoc.TextOut 1850, iLin, Format(rsHistInvent!BODEGA1 + rsHistInvent!BODEGA2, "#.00")
    MainMant.spDoc.TextOut 2050, iLin, Format(((rsHistInvent!BODEGA1 + rsHistInvent!BODEGA2) * rsHistInvent!COSTO), "#0.00")
    MainMant.spDoc.TextAlign = SPTA_LEFT
    iLin = iLin + 50

    nLBod1 = rsHistInvent!BODEGA1
    nLBod2 = rsHistInvent!BODEGA2
    nValBod1 = rsHistInvent!BODEGA1 * rsHistInvent!COSTO
    nValBod2 = rsHistInvent!BODEGA2 * rsHistInvent!COSTO
    nLBodegas = rsHistInvent!BODEGA1 + rsHistInvent!BODEGA2
    nValBodegas = (rsHistInvent!BODEGA1 + rsHistInvent!BODEGA2) * rsHistInvent!COSTO
    nLCosto = rsHistInvent!COSTO
    
    Dim PosY1 As Integer
    Dim PosY2 As Integer
    PosY1 = 850
    PosY2 = 1050
    With rsHistorial
        On Error Resume Next
        .MoveFirst
        On Error GoTo 0
        Do While Not .EOF
            MainMant.spDoc.TextAlign = SPTA_LEFT
            cFecha = Right(!FECHA, 2) & "/" & Mid(!FECHA, 5, 2) & "/" & Left(!FECHA, 4)
            MainMant.spDoc.TextOut 300, iLin, cFecha
            MainMant.spDoc.TextOut 500, iLin, !TIPO_MOV
            MainMant.spDoc.TextAlign = SPTA_RIGHT
            MainMant.spDoc.TextOut 850, iLin, Format(!CANTIDAD_EXT, "#.00")
            MainMant.spDoc.TextOut 1010, iLin, Format(!VALOR_BOD1, "#0.00")
            Select Case Trim(!TIPO_MOV)
                Case "VTA"     'BODEGA2
                    PosY1 = 1500
                    PosY2 = 1650
                    nLBod2 = nLBod2 - !CANTIDAD_EXT
                    nLBodegas = nLBodegas - !CANTIDAD_EXT
                    MainMant.spDoc.TextOut PosY1, iLin, Format(nLBod2, "#.00")
                    MainMant.spDoc.TextOut PosY2, iLin, Format(nLBod2 * nLCosto, "#0.00")
                Case "COM"
                    'INFO: CAMBIANDO INFORMACION A DESPLEGAR YA QUE ANTES TODO APUNTABA ALA
                    'BODEGA 1, AHORA A CAMBIAR Y PONER LOS DATOS SEGUN LA BODEGA SELECCIONADA
                    If GET_BODEGA = 1 Then
                        PosY1 = 1200
                        PosY2 = 1350
                        nLCosto = Format(((!VALOR_BOD1 + (nLCosto * (nLBod1 + nLBod2)))) / (!CANTIDAD_EXT + nLBod1 + nLBod2), "#.0000")
                        nLBodegas = nLBodegas + (!CANTIDAD_EXT * nCantidadPorCompras)
                        nLBod1 = nLBod1 + (!CANTIDAD_EXT * nCantidadPorCompras)
                        MainMant.spDoc.TextOut PosY1, iLin, Format(nLBod1, "#.00")
                        MainMant.spDoc.TextOut PosY2, iLin, Format(nLBod1 * nLCosto, "#0.00")
                    Else
                        PosY1 = 1500
                        PosY2 = 1650
                        'INFO: 12OCT2011
                        'QUITANDO nCantidadPorCompras EN COMPRAS, YA QUE LAS COMPRAS YA TRAEN LA CANTIDAD EXTENDIDA
                        nLCosto = Format(((!VALOR_BOD1 + (nLCosto * (nLBod1 + nLBod2)))) / ((!CANTIDAD_EXT) + nLBod1 + nLBod2), "#.0000")
                        'nLCosto = Format(((!VALOR_BOD1 + (nLCosto * (nLBod1 + nLBod2)))) / ((!CANTIDAD_EXT * nCantidadPorCompras) + nLBod1 + nLBod2), "#.0000")
                        
                        nLBodegas = nLBodegas + (!CANTIDAD_EXT)
                        'nLBodegas = nLBodegas + (!CANTIDAD_EXT * nCantidadPorCompras)
                        
                        nLBod2 = nLBod2 + (!CANTIDAD_EXT)
                        'nLBod2 = nLBod2 + (!CANTIDAD_EXT * nCantidadPorCompras)

                        MainMant.spDoc.TextOut PosY1, iLin, Format(nLBod2, "#.00")
                        MainMant.spDoc.TextOut PosY2, iLin, Format(nLBod2 * nLCosto, "#0.00")
                    End If
                Case "TRANSFER"
                    PosY1 = 1200
                    PosY2 = 1350
                    'nLBod1 = nLBod1 - !Cantidad
                    nLBod1 = nLBod1 - !CANTIDAD_EXT
                    nLBod2 = nLBod2 + !CANTIDAD_EXT
                    'nLBodegas = nLBodegas - !Cantidad + (!CANTIDAD_EXT - !Cantidad)
                    'If !Cantidad <> 0 Then
                    MainMant.spDoc.TextOut PosY1, iLin, Format(nLBod1, "#.00")
                    MainMant.spDoc.TextOut PosY2, iLin, Format(nLBod1 * nLCosto, "#0.00")
                    'End If
                    PosY1 = 1500
                    PosY2 = 1650
                    'If (!CANTIDAD_EXT - !Cantidad) <> 0 Then
                    MainMant.spDoc.TextOut PosY1, iLin, Format(nLBod2, "#.00")
                    MainMant.spDoc.TextOut PosY2, iLin, Format(nLBod2 * nLCosto, "#0.00")
                    'End If
                Case "MERMA"
                    PosY1 = 1200
                    PosY2 = 1350
                    nLBod1 = nLBod1 - !Cantidad
                    nLBod2 = nLBod2 - (!CANTIDAD_EXT - !Cantidad)
                    nLBodegas = nLBodegas - !Cantidad - (!CANTIDAD_EXT - !Cantidad)
                    If !Cantidad <> 0 Then
                        MainMant.spDoc.TextOut PosY1, iLin, Format(nLBod1, "#.00")
                        MainMant.spDoc.TextOut PosY2, iLin, Format(nLBod1 * nLCosto, "#0.00")
                    End If
                    PosY1 = 1500
                    PosY2 = 1650
                    If (!CANTIDAD_EXT - !Cantidad) <> 0 Then
                        MainMant.spDoc.TextOut PosY1, iLin, Format(nLBod2, "#.00")
                        MainMant.spDoc.TextOut PosY2, iLin, Format(nLBod2 * nLCosto, "#0.00")
                    End If
                Case "DEV", "DEF", "NC"
                    PosY1 = 1200
                    PosY2 = 1350
                    nLBod1 = nLBod1 - !CANTIDAD_EXT
                    nLBodegas = nLBodegas - !CANTIDAD_EXT
                    MainMant.spDoc.TextOut PosY1, iLin, Format(nLBod1, "#.00")
                    MainMant.spDoc.TextOut PosY2, iLin, Format(nLBod1 * nLCosto, "#0.00")
                Case Else
            End Select
            MainMant.spDoc.TextOut 1850, iLin, Format(nLBodegas, "#.00")
            MainMant.spDoc.TextOut 2050, iLin, Format((nLBodegas * nLCosto), "#0.00")
            MainMant.spDoc.TextAlign = SPTA_LEFT
            iLin = iLin + 50
            If iLin > 2400 Then
                MainMant.spDoc.TextAlign = SPTA_LEFT
                Call PrintTit(cMedida, cConsumo)
            End If
            .MoveNext
        Loop
    End With
'''    rsTemp.Close
'''    Set rsTemp = Nothing
    MainMant.spDoc.TextAlign = SPTA_LEFT
    MainMant.spDoc.DoPrintPreview
End If
Me.MousePointer = vbDefault
On Error GoTo 0

Call Seguridad

Exit Sub

ErrAdm:
    If Err.Number = 3021 Then
        'INFO: NO HAY INFORMACION HISTORICA DE CIERRE DE MES
        MsgBox Err.Number & " - " & Err.Description & vbCrLf & _
            "NO HAY INFORMACION HISTORICA DE CIERRE DE MES" & vbCrLf & _
            "Es Posible que este articulo no tenga historial para el periodo seleccionado", vbCritical, Me.Caption
        EscribeLog ("Admin." & "conHistorico.Sin Historial (" & Left(List1.Text, 30) & ") - " & Err.Number & " - " & Err.Description)
        'Resume Next
    Else
        MsgBox Err.Number & " - " & Err.Description, vbCritical, Me.Caption
        EscribeLog ("Admin." & "conHistorico.(" & Err.Number & ") - " & Err.Description)
    End If
    Me.MousePointer = vbDefault
End Sub

Private Sub Combo1_Click()
Dim rsInvent As New ADODB.Recordset
Dim cSQL As String

rsDepInv.MoveFirst
rsDepInv.Find "DESCRIP = '" & Combo1.Text & "'"

If Combo1.Text = "" Then Exit Sub
On Error GoTo AdmErr:
rsInvent.Open "SELECT ID, NOMBRE FROM INVENT " & _
        " WHERE COD_DEPT = " & rsDepInv!CODIGO & _
        " ORDER BY NOMBRE", msConn, adOpenStatic, adLockOptimistic
List1.Clear
Do While Not rsInvent.EOF
    List1.AddItem rsInvent!NOMBRE & Space(90) & rsInvent!ID
    rsInvent.MoveNext
Loop
On Error GoTo 0

AdmErr:
rsInvent.Close
Set rsInvent = Nothing
End Sub

Private Sub Form_Load()

txtFecIni = Format(Date, "SHORT DATE")
txtFecFin = Format(Date, "SHORT DATE")
rsDepInv.Open "SELECT CODIGO,DESCRIP FROM DEP_INV ORDER BY DESCRIP", msConn, adOpenStatic, adLockOptimistic
Do While Not rsDepInv.EOF
    Combo1.AddItem rsDepInv!DESCRIP
    rsDepInv.MoveNext
Loop

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
        cmdSeek.Enabled = False
    Case "N"        'SIN DERECHOS
        txtFecIni.Enabled = False: txtFecFin.Enabled = False
        Combo1.Enabled = False: List1.Enabled = False
        cmdSeek.Enabled = False
End Select
End Function

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
On Error Resume Next
rsDepInv.Close
Set rsDepInv = Nothing
On Error GoTo 0
End Sub

Private Sub List1_DblClick()
cmdSeek_Click
End Sub
Private Sub PrintTit(cUMed As String, cUCons As String)
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
MainMant.spDoc.TextOut 300, 450, conHistorico.Caption
MainMant.spDoc.TextOut 300, 500, "Periodo desde " & txtFecIni & " hasta " & txtFecFin
MainMant.spDoc.TextOut 300, 550, "Departamento Seleccionado: " & Combo1.Text
MainMant.spDoc.TextOut 300, 600, "Producto : " & Trim(Left(List1.Text, 70))
MainMant.spDoc.TextOut 1100, 600, "Se Compra en (" & cUMed & "), se Consume por (" & cUCons & ")"

MainMant.spDoc.TextOut 300, 700, "FECHA"
MainMant.spDoc.TextOut 500, 700, "CONCEPTO"
MainMant.spDoc.TextOut 720, 700, "Cantidad"
MainMant.spDoc.TextOut 930, 700, "Monto"
MainMant.spDoc.TextOut 1100, 700, "Bodega 1"
MainMant.spDoc.TextOut 1400, 700, "Bodega 2"
MainMant.spDoc.TextOut 1750, 700, "Stock Total"
MainMant.spDoc.TextOut 1100, 750, "CANT"
MainMant.spDoc.TextOut 1250, 750, "Monto"
MainMant.spDoc.TextOut 1400, 750, "CANT"
MainMant.spDoc.TextOut 1550, 750, "Monto"
MainMant.spDoc.TextOut 1750, 750, "CANT"
MainMant.spDoc.TextOut 1950, 750, "Monto"
MainMant.spDoc.TextOut 300, 800, "----------------------------------------------------------------------------------------------------------------------------------------------------"

iLin = 850
nPagina = nPagina + 1
End Sub
