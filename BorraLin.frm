VERSION 5.00
Begin VB.Form BorraLin 
   BackColor       =   &H00C0E0FF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "ANOTE EL NUMERO DE LINEA"
   ClientHeight    =   2865
   ClientLeft      =   2835
   ClientTop       =   3480
   ClientWidth     =   3975
   ControlBox      =   0   'False
   Icon            =   "BorraLin.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   1692.736
   ScaleMode       =   0  'User
   ScaleWidth      =   3732.31
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton Borrar 
      Height          =   615
      Left            =   1560
      Picture         =   "BorraLin.frx":0442
      Style           =   1  'Graphical
      TabIndex        =   15
      Top             =   2040
      Width           =   735
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00C0E0FF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Index           =   3
      Left            =   120
      TabIndex        =   4
      Top             =   600
      Width           =   3735
      Begin VB.CommandButton Command8 
         Caption         =   "0"
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
         Index           =   0
         Left            =   2940
         TabIndex        =   14
         Top             =   720
         Width           =   735
      End
      Begin VB.CommandButton Command8 
         Caption         =   "1"
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
         Index           =   1
         Left            =   60
         TabIndex        =   13
         Top             =   180
         Width           =   735
      End
      Begin VB.CommandButton Command8 
         Caption         =   "2"
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
         Index           =   2
         Left            =   780
         TabIndex        =   12
         Top             =   180
         Width           =   735
      End
      Begin VB.CommandButton Command8 
         Caption         =   "3"
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
         Index           =   3
         Left            =   1500
         TabIndex        =   11
         Top             =   180
         Width           =   735
      End
      Begin VB.CommandButton Command8 
         Caption         =   "4"
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
         Index           =   4
         Left            =   2220
         TabIndex        =   10
         Top             =   180
         Width           =   735
      End
      Begin VB.CommandButton Command8 
         Caption         =   "5"
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
         Index           =   5
         Left            =   2940
         TabIndex        =   9
         Top             =   180
         Width           =   735
      End
      Begin VB.CommandButton Command8 
         Caption         =   "6"
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
         Index           =   6
         Left            =   60
         TabIndex        =   8
         Top             =   720
         Width           =   735
      End
      Begin VB.CommandButton Command8 
         Caption         =   "7"
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
         Index           =   7
         Left            =   780
         TabIndex        =   7
         Top             =   720
         Width           =   735
      End
      Begin VB.CommandButton Command8 
         Caption         =   "8"
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
         Index           =   8
         Left            =   1500
         TabIndex        =   6
         Top             =   720
         Width           =   735
      End
      Begin VB.CommandButton Command8 
         Caption         =   "9"
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
         Index           =   9
         Left            =   2220
         TabIndex        =   5
         Top             =   720
         Width           =   735
      End
   End
   Begin VB.TextBox txtLin 
      Alignment       =   1  'Right Justify
      Height          =   345
      Left            =   2040
      TabIndex        =   1
      Top             =   240
      Width           =   645
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "&Anular"
      Default         =   -1  'True
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   630
      Left            =   120
      TabIndex        =   2
      Top             =   2040
      Width           =   1260
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "&Cancelar"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   630
      Left            =   2520
      TabIndex        =   3
      Top             =   2040
      Width           =   1260
   End
   Begin VB.Label lblLabels 
      BackColor       =   &H00C0E0FF&
      Caption         =   "Número de Línea"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Index           =   0
      Left            =   240
      TabIndex        =   0
      Top             =   360
      Width           =   1560
   End
End
Attribute VB_Name = "BorraLin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private nlPase As Integer
Private Sub Borrar_Click()
nlPase = 0
txtLin = ""
End Sub

Private Sub cmdCancel_Click()
Unload Me
End Sub
Private Sub cmdOK_Click()
'--------------------- ANULACION DE UNA LINEA --------------------

If txtLin = "" Or Not IsNumeric(txtLin) Or Val(txtLin) > CajLin Then
    MsgBox "Numero de Linea NO es VALIDO", vbExclamation, BoxTit
Else
    Dim rsFixTmpTrans As Recordset
    Dim txto As String
    Dim rsParciales As Recordset
    Dim lParc As Integer
    Dim sqltext As String
    Dim nCtaLinAnul As Integer
    Dim SOLO_FECHA As String

    Set rsFixTmpTrans = New Recordset
    txto = "SELECT * FROM tmp_trans " & _
        " WHERE MESA = " & nMesa & " AND LIN = " & txtLin
    rsFixTmpTrans.Open txto, msConn, adOpenStatic, adLockReadOnly

    If rsFixTmpTrans.EOF = True Then
        rsFixTmpTrans.Close
        Exit Sub
    End If
    
    If rsFixTmpTrans!cant < 0 Then
        MsgBox "El Producto ya fue ANULADO", vbExclamation, BoxTit
        Exit Sub
    End If
    
    If Mid(rsFixTmpTrans!TIPO, 1, 1) = "B" Then
        MsgBox "PRODUCTO YA FUE ANULADO/CORREGIDO/SE DIO DESCUENTO EN LA LINEA " & Val(Mid(rsFixTmpTrans!TIPO, 5, 2)), vbExclamation, BoxTit
        rsFixTmpTrans.Close
        Exit Sub
    End If

    nCtaLinAnul = rsFixTmpTrans!cuenta
    CajLin = CajLin + 1

    'SOLO_FECHA = Format(Date, "MMDDYYYY")
    SOLO_FECHA = Format(Date, "YYYYMMDD")

    'CadenaSql = "INSERT INTO TMP_TRANS " & _
        "(CAJA,CAJERO,MESA,MESERO,VALID,LIN,DESCRIP,CANT," & _
          "DEPTO,PLU,ENVASE,PRECIO_UNIT,PRECIO,FECHA,HORA," & _
          "TIPO,DESCUENTO,CUENTA,IMPRESO,IMPRESORA) VALUES (" & _
        "" & cCaja & "," & npNumCaj & "," & nMesa & "," & nMesero & _
        "," & -1 & "," & CajLin & "," & "'VO-" & rsFixTmpTrans!descrip + TextEnv & "'" & "," & _
        rsFixTmpTrans!cant * (-1) & "," & rsFixTmpTrans!depto & "," & _
        rsFixTmpTrans!PLU & "," & rsFixTmpTrans!envase & "," & _
        (rsFixTmpTrans!precio_unit + (rsFixTmpTrans!precio_unit * iISC)) * (-1) & "," & _
        (rsFixTmpTrans!precio + (rsFixTmpTrans!precio * iISC)) * (-1) & _
        "," & "'" & SOLO_FECHA & "'" & "," & "'" & Time & "'" & _
        ",'VO-" & txtLin & "'," & 0# & "," & nCtaLinAnul & ",FALSE," & _
        rsFixTmpTrans!IMPRESORA & ")"
        
    CadenaSql = "INSERT INTO TMP_TRANS " & _
        "(CAJA,CAJERO,MESA,MESERO,VALID,LIN,DESCRIP,CANT," & _
          "DEPTO,PLU,ENVASE,PRECIO_UNIT,PRECIO,FECHA,HORA," & _
          "TIPO,DESCUENTO,CUENTA,IMPRESO,IMPRESORA) VALUES (" & _
        "" & cCaja & "," & npNumCaj & "," & nMesa & "," & nMesero & _
        "," & -1 & "," & CajLin & "," & "'VO-" & rsFixTmpTrans!descrip + TextEnv & "'" & "," & _
        rsFixTmpTrans!cant * (-1) & "," & rsFixTmpTrans!depto & "," & _
        rsFixTmpTrans!PLU & "," & rsFixTmpTrans!envase & "," & _
        (rsFixTmpTrans!precio_unit) * (-1) & "," & _
        (rsFixTmpTrans!precio) * (-1) & _
        "," & "'" & SOLO_FECHA & "'" & "," & "'" & Time & "'" & _
        ",'VO-" & txtLin & "'," & 0# & "," & nCtaLinAnul & ",FALSE," & _
        rsFixTmpTrans!IMPRESORA & ")"

    sqltext = "UPDATE TMP_TRANS SET VALID = 0,TIPO = 'BVO" & Str(CajLin) & _
            "' WHERE MESA = " & nMesa & _
            " AND LIN = " & txtLin
    
    msConn.BeginTrans
    msConn.Execute CadenaSql
    msConn.Execute sqltext
    msConn.CommitTrans

    rs07.Open "SELECT a.lin,a.descrip,a.cant," & _
        " format(precio_unit,'##0.00') as mPrecio_unit," & _
        " format(precio,'##0.00') as mPrecio," & _
        " a.cajero,a.mesero,a.depto,a.plu,a.mesa,a.valid, " & _
        " a.envase,a.fecha,a.hora,a.tipo,a.descuento,a.cuenta, " & _
        " a.caja " & _
        " FROM tmp_trans as a " & _
        " WHERE a.mesa = " & nMesa & _
        " AND A.CUENTA = " & nCta & _
        " ORDER BY a.lin", msConn, adOpenStatic, adLockReadOnly
        
    Set PLU.PlatosMesa.DataSource = rs07

    nLineas = PLU.PlatosMesa.Rows - 1
    
    PLU.PlatosMesa.TopRow = (PLU.PlatosMesa.Rows - 1)
    
    Set rsParciales = New Recordset
    rsParciales.Open "SELECT MESA,SUM(MONTO) AS VALOR FROM TMP_PAR_PAGO " & _
            " WHERE MESA = " & nMesa & _
            " GROUP BY MESA", msConn, adOpenDynamic, adLockOptimistic

    If rsParciales.EOF Then lParc = 0 Else lParc = 1
    
    rs07.Close
    rs07.Open "SELECT sum(a.precio) as precio FROM tmp_trans as a " & _
          " WHERE a.mesa = " & nMesa & _
          " AND A.CUENTA = " & nCta, msConn, adOpenStatic, adLockReadOnly
    PLU.SubTot = FormatCurrency(IIf(IsNull(rs07!precio), 0#, rs07!precio), 2)
    On Error Resume Next
    PLU.SubTot = FormatCurrency((PLU.SubTot + (rs07!precio * iISC)), 2)
    iISCTransaccion = rs07!precio * iISC
    SBTot = Format(PLU.SubTot, "standard")
    On Error GoTo 0
    rs07.Close
    rsFixTmpTrans.Close
    nCantidad = 1: nPase = 0
    PLU.Text1(2) = nCantidad
    
    If lParc = 1 Then
        PLU.PlatosMesa.AddItem (CajLin + 1) & Chr(9) & _
            "PAGO.PARCIAL" & Chr(9) & 1 & Chr(9) & _
            Format(rsParciales!VALOR, "STANDARD") & Chr(9) & Format(rsParciales!VALOR, "STANDARD")
        PLU.SubTot = Format(PLU.SubTot - rsParciales!VALOR, "STANDARD")
        PLU.lbMensaje.BackColor = &HFFFF00
        PLU.lbMensaje = "MESA CON PAGOS PARCIALES"
    Else
        PLU.lbMensaje.BackColor = &HFFFFFF
        PLU.lbMensaje = ""
    End If

End If
Unload Me
End Sub

Private Sub Command8_Click(Index As Integer)
Dim cCant As String

If nlPase = 0 Then
    txtLin = Command8(Index).Index
Else
    'cCant = Str(txtLin)
    cCant = CStr(txtLin)
    cCant = cCant & Command8(Index).Index
    txtLin = Val(cCant)
End If
nlPase = nlPase + 1
End Sub

Private Sub Form_Load()
nlPase = 0
End Sub

Private Sub txtLin_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then cmdOK.SetFocus
End Sub
