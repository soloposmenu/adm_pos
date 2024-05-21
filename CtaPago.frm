VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form CtaPago 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "MODULO DE FACTURACION POR CUENTAS"
   ClientHeight    =   8625
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11115
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "CtaPago.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   8625
   ScaleWidth      =   11115
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame4 
      BackColor       =   &H00008000&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   8895
      Left            =   9000
      TabIndex        =   25
      Top             =   -200
      Width           =   1935
      Begin VB.CommandButton cdmBill 
         Height          =   735
         Index           =   0
         Left            =   120
         Picture         =   "CtaPago.frx":030A
         Style           =   1  'Graphical
         TabIndex        =   36
         Tag             =   "5.00"
         Top             =   240
         Width           =   1695
      End
      Begin VB.CommandButton cdmBill 
         Height          =   735
         Index           =   1
         Left            =   120
         Picture         =   "CtaPago.frx":19E9
         Style           =   1  'Graphical
         TabIndex        =   35
         Tag             =   "10.00"
         Top             =   1020
         Width           =   1695
      End
      Begin VB.CommandButton cdmBill 
         Height          =   735
         Index           =   2
         Left            =   120
         Picture         =   "CtaPago.frx":307A
         Style           =   1  'Graphical
         TabIndex        =   34
         Tag             =   "15.00"
         Top             =   1800
         Width           =   1695
      End
      Begin VB.CommandButton cdmBill 
         Height          =   735
         Index           =   3
         Left            =   120
         Picture         =   "CtaPago.frx":4730
         Style           =   1  'Graphical
         TabIndex        =   33
         Tag             =   "20.00"
         Top             =   2580
         Width           =   1695
      End
      Begin VB.CommandButton cdmBill 
         Height          =   735
         Index           =   4
         Left            =   120
         Picture         =   "CtaPago.frx":5DF8
         Style           =   1  'Graphical
         TabIndex        =   32
         Tag             =   "25.00"
         Top             =   3380
         Width           =   1695
      End
      Begin VB.CommandButton cdmBill 
         Height          =   735
         Index           =   5
         Left            =   120
         Picture         =   "CtaPago.frx":74B7
         Style           =   1  'Graphical
         TabIndex        =   31
         Tag             =   "30.00"
         Top             =   4160
         Width           =   1695
      End
      Begin VB.CommandButton cdmBill 
         Height          =   735
         Index           =   6
         Left            =   120
         Picture         =   "CtaPago.frx":8B72
         Style           =   1  'Graphical
         TabIndex        =   30
         Tag             =   "35.00"
         Top             =   4940
         Width           =   1695
      End
      Begin VB.CommandButton cdmBill 
         Height          =   735
         Index           =   7
         Left            =   120
         Picture         =   "CtaPago.frx":A20D
         Style           =   1  'Graphical
         TabIndex        =   29
         Tag             =   "40.00"
         Top             =   5720
         Width           =   1695
      End
      Begin VB.CommandButton cdmBill 
         Height          =   735
         Index           =   8
         Left            =   120
         Picture         =   "CtaPago.frx":B8D4
         Style           =   1  'Graphical
         TabIndex        =   28
         Tag             =   "45.00"
         Top             =   6500
         Width           =   1695
      End
      Begin VB.CommandButton cdmBill 
         Height          =   735
         Index           =   9
         Left            =   120
         Picture         =   "CtaPago.frx":D00F
         Style           =   1  'Graphical
         TabIndex        =   27
         Tag             =   "50.00"
         Top             =   7280
         Width           =   1695
      End
      Begin VB.CommandButton cdmBill 
         Height          =   735
         Index           =   10
         Left            =   120
         Picture         =   "CtaPago.frx":E6CB
         Style           =   1  'Graphical
         TabIndex        =   26
         Tag             =   "100.00"
         Top             =   8080
         Width           =   1695
      End
   End
   Begin VB.CommandButton cmdDescGlob 
      BackColor       =   &H00FFC0FF&
      Caption         =   "Descuento Global"
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
      Height          =   615
      Left            =   6240
      Style           =   1  'Graphical
      TabIndex        =   23
      Top             =   5880
      Width           =   1575
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid ListaPagos 
      Height          =   1455
      Left            =   6120
      TabIndex        =   21
      Top             =   480
      Width           =   2775
      _ExtentX        =   4895
      _ExtentY        =   2566
      _Version        =   393216
      Rows            =   0
      FixedRows       =   0
      FixedCols       =   0
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
   Begin VB.CommandButton cmdSalir 
      Caption         =   "REGRESAR SIN FACTURAR"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   6720
      TabIndex        =   20
      Top             =   7800
      Width           =   1815
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00C0C0C0&
      Caption         =   "PROPINAS"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1935
      Left            =   120
      TabIndex        =   18
      Top             =   4920
      Width           =   5895
      Begin VB.CommandButton cmdPropina 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Index           =   0
         Left            =   120
         TabIndex        =   19
         Top             =   360
         Width           =   1335
      End
   End
   Begin VB.Frame Frame2 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3735
      Index           =   3
      Left            =   6120
      TabIndex        =   2
      Top             =   1920
      Width           =   2775
      Begin VB.CommandButton Clear 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   1560
         Picture         =   "CtaPago.frx":FDA5
         Style           =   1  'Graphical
         TabIndex        =   14
         Top             =   3000
         Width           =   735
      End
      Begin VB.CommandButton Command2 
         Caption         =   "00"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Index           =   10
         Left            =   840
         TabIndex        =   13
         Top             =   3000
         Width           =   735
      End
      Begin VB.CommandButton Command2 
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Index           =   0
         Left            =   120
         TabIndex        =   3
         Top             =   3000
         Width           =   735
      End
      Begin VB.CommandButton Command2 
         Caption         =   "9"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Index           =   9
         Left            =   1560
         TabIndex        =   12
         Top             =   2280
         Width           =   735
      End
      Begin VB.CommandButton Command2 
         Caption         =   "8"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Index           =   8
         Left            =   840
         TabIndex        =   11
         Top             =   2280
         Width           =   735
      End
      Begin VB.CommandButton Command2 
         Caption         =   "7"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Index           =   7
         Left            =   120
         TabIndex        =   10
         Top             =   2280
         Width           =   735
      End
      Begin VB.CommandButton Command2 
         Caption         =   "6"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Index           =   6
         Left            =   1560
         TabIndex        =   9
         Top             =   1560
         Width           =   735
      End
      Begin VB.CommandButton Command2 
         Caption         =   "5"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Index           =   5
         Left            =   840
         TabIndex        =   8
         Top             =   1560
         Width           =   735
      End
      Begin VB.CommandButton Command2 
         Caption         =   "4"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Index           =   4
         Left            =   120
         TabIndex        =   7
         Top             =   1560
         Width           =   735
      End
      Begin VB.CommandButton Command2 
         Caption         =   "3"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Index           =   3
         Left            =   1560
         TabIndex        =   6
         Top             =   840
         Width           =   735
      End
      Begin VB.CommandButton Command2 
         Caption         =   "2"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Index           =   2
         Left            =   840
         TabIndex        =   5
         Top             =   840
         Width           =   735
      End
      Begin VB.CommandButton Command2 
         Caption         =   "1"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Index           =   1
         Left            =   120
         TabIndex        =   4
         Top             =   840
         Width           =   735
      End
      Begin VB.Label lbMonto 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00404040&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0.00"
         BeginProperty DataFormat 
            Type            =   1
            Format          =   """B/."" #,##0"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   6154
            SubFormatType   =   2
         EndProperty
         BeginProperty Font 
            Name            =   "Arial Black"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0080FF80&
         Height          =   495
         Left            =   120
         TabIndex        =   15
         Top             =   240
         Width           =   2535
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0C0C0&
      Caption         =   "FORMAS DE PAGO"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2775
      Left            =   120
      TabIndex        =   0
      Top             =   1920
      Width           =   5895
      Begin VB.CommandButton cmdFPagos 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Index           =   0
         Left            =   120
         TabIndex        =   1
         Top             =   360
         Width           =   1335
      End
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFClientes 
      Height          =   1095
      Left            =   120
      TabIndex        =   39
      Top             =   7200
      Visible         =   0   'False
      Width           =   2175
      _ExtentX        =   3836
      _ExtentY        =   1931
      _Version        =   393216
      Rows            =   0
      Cols            =   3
      FixedRows       =   0
      FixedCols       =   0
      SelectionMode   =   1
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
      _Band(0).Cols   =   3
   End
   Begin VB.CheckBox chkInfo 
      Caption         =   "Información del Cliente"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   495
      Left            =   6000
      TabIndex        =   40
      Top             =   6960
      Width           =   3015
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   3600
      TabIndex        =   38
      Top             =   240
      Width           =   2415
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   3600
      TabIndex        =   37
      Top             =   960
      Width           =   2415
   End
   Begin VB.Label lbCuenta 
      BeginProperty DataFormat 
         Type            =   1
         Format          =   """B/."" #,##0"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   6154
         SubFormatType   =   2
      EndProperty
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   495
      Left            =   3120
      TabIndex        =   24
      Top             =   360
      Width           =   375
   End
   Begin VB.Label Label1 
      Caption         =   "Pagos Recibidos"
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
      Left            =   6120
      TabIndex        =   22
      Top             =   240
      Width           =   1695
   End
   Begin VB.Label lbPend 
      Caption         =   "Monto Pendiente"
      BeginProperty DataFormat 
         Type            =   1
         Format          =   """B/."" #,##0"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   6154
         SubFormatType   =   2
      EndProperty
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   120
      TabIndex        =   17
      Top             =   1080
      Width           =   3135
   End
   Begin VB.Label lbFact 
      Caption         =   "Total Cuenta #"
      BeginProperty DataFormat 
         Type            =   1
         Format          =   """B/."" #,##0"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   6154
         SubFormatType   =   2
      EndProperty
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   120
      TabIndex        =   16
      Top             =   360
      Width           =   2895
   End
End
Attribute VB_Name = "CtaPago"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim nfPase As Integer
Dim nMntOculto As String
Dim RSPAGOS As Recordset    'Pagos
Dim RSPROPINAS As Recordset   'Propinas
Dim rsCuentas As Recordset      'Cuentas
Dim OrigSB As Single
Dim nFlagParciales As Integer
Dim nCuenta As Integer
Dim nProp As Single
Dim MiPropina As Single

Private Sub Actualizador()
Dim rsAcutalizacion, rsTrans As Recordset
Dim sqltext As String, ImpText As String
Dim MiValor As Currency
Dim nValorPago As Single
Dim nTipoPago As Integer, i As Integer

Set rsAcutalizacion = New Recordset

'Actualiza los valores de la factura
'INCREMENTA EL NUMERO DE TRANSACCION EN 1

msConn.BeginTrans
msConn.Execute "UPDATE ORGANIZACION SET TRANS = TRANS + 1"
msConn.Execute "DELETE * FROM TMP_CUENTAS " & _
            " WHERE MESA = " & nMesa & _
            " AND CUENTA = " & nCuenta
'AUMENTA E INCREMENTA LOS VALORES POR DEPARTAMENTO
'AUMENTA E INCREMENTA LOS VALORES POR PLATO (PLU)
sqltext = "SELECT * FROM TMP_TRANS " & _
        " WHERE VALID AND CANT >= 0 AND " & _
        " MESA = " & nMesa & " AND CUENTA = " & nCuenta & " ORDER BY DEPTO,PLU"
rsAcutalizacion.Open sqltext, msConn, adOpenStatic, adLockReadOnly

Do Until rsAcutalizacion.EOF
    'EXISTE UN PROBLEMA QUE SI HAY DESCUENTO, ENTONCES
    'EL RECORDSET DEVUELVE 1 Y ENTRA AQUI
    
    If IsEmpty(rsAcutalizacion!precio) Then GoTo Proximo:
    
    MiValor = Format((rsAcutalizacion!CANT * rsAcutalizacion!precio_unit), "#####.00")
    'OK Para DEPTO
    sqltext = "UPDATE DEPTO SET X_COUNT = X_COUNT + " & rsAcutalizacion!CANT & _
        " , Z_COUNT = Z_COUNT + " & rsAcutalizacion!CANT & _
        " , VALOR = VALOR + " & MiValor & _
        " , X_PERIOD_CNT = X_PERIOD_CNT + " & rsAcutalizacion!CANT & _
        " , Z_PERIOD_CNT = Z_PERIOD_CNT + " & rsAcutalizacion!CANT & _
        " , PERIOD_VAL = PERIOD_VAL + " & MiValor & _
        " WHERE CODIGO = " & rsAcutalizacion!depto
    msConn.Execute sqltext

    'OK para PLU
    sqltext = "UPDATE PLU SET X_COUNT = X_COUNT + " & rsAcutalizacion!CANT & _
        " , Z_COUNT = Z_COUNT + " & rsAcutalizacion!CANT & _
        " , VALOR = VALOR + " & MiValor & _
        " , X_PERIOD_CNT = X_PERIOD_CNT + " & rsAcutalizacion!CANT & _
        " , Z_PERIOD_CNT = Z_PERIOD_CNT + " & rsAcutalizacion!CANT & _
        " , PERIOD_VAL = PERIOD_VAL + " & MiValor & _
        " WHERE CODIGO = " & rsAcutalizacion!PLU
    msConn.Execute sqltext
    
    'OK Para CONTEND_02
    sqltext = "UPDATE CONTEND_02 SET X_COUNT = X_COUNT + " & rsAcutalizacion!CANT & _
        " , Z_COUNT = Z_COUNT + " & rsAcutalizacion!CANT & _
        " , VALOR = VALOR + " & MiValor & _
        " , X_PERIOD_CNT = X_PERIOD_CNT + " & rsAcutalizacion!CANT & _
        " , Z_PERIOD_CNT = Z_PERIOD_CNT + " & rsAcutalizacion!CANT & _
        " , PERIOD_VAL = PERIOD_VAL + " & MiValor & _
        " WHERE CODIGO = " & rsAcutalizacion!PLU & " AND " & _
        " CONTENEDOR = " & rsAcutalizacion!envase
    msConn.Execute sqltext

    sqltext = "UPDATE MESAS SET VALOR = VALOR + " & MiValor & _
            " WHERE NUMERO = " & rsAcutalizacion!mesa & " OR " & _
            " NUMERO = -99 "
    msConn.Execute sqltext
    
Proximo:
    rsAcutalizacion.MoveNext
Loop

'AUMENTA E INCREMENTA LOS VALORES POR CAJERO
'OK
sqltext = "UPDATE CAJEROS SET X_COUNT = X_COUNT + 1" & _
    " , Z_COUNT = Z_COUNT + 1 " & _
    " , VALOR = VALOR + " & Format(Label3, "#0.00") & _
    " WHERE NUMERO = " & npNumCaj & " OR NUMERO = " & 999
msConn.Execute sqltext

'AUMENTA E INCREMENTA LOS VALORES POR MESEROS
'OK
sqltext = "UPDATE MESEROS SET X_COUNT = X_COUNT + 1" & _
    " , Z_COUNT = Z_COUNT + 1 " & _
    " , VALOR = VALOR + " & Format(Label3, "#0.00") & _
    " WHERE NUMERO = " & nMesero & " OR NUMERO = " & 999
msConn.Execute sqltext

For i = 0 To (ListaPagos.Rows - 1)
    On Error GoTo ErrAdm:
        ListaPagos.Row = i
        ListaPagos.Col = 0
        nTipoPago = ListaPagos.Text
        ListaPagos.Col = 2
    On Error GoTo 0
    nValorPago = Format(ListaPagos.Text, "STANDARD")
    sqltext = "UPDATE PAGOS SET X_COUNT = X_COUNT + 1" & _
        " , Z_COUNT = Z_COUNT + 1 " & _
        " , VALOR = VALOR + " & Format(nValorPago, "#0.00") & _
        " , X_PERIOD_CNT = X_PERIOD_CNT + 1" & _
        " , Z_PERIOD_CNT = Z_PERIOD_CNT + 1" & _
        " , PERIOD_VAL = PERIOD_VAL + " & Format(nValorPago, "#0.00") & _
        " WHERE CODIGO = " & nTipoPago & " OR CODIGO = " & 999
    msConn.Execute sqltext
    
    sqltext = "INSERT INTO TRANSAC_PAGO " & _
            " (NUM_TRANS,TIPO_PAGO,CAJERO,LIN,MONTO) VALUES (" & _
            rs00!TRANS & "," & nTipoPago & "," & npNumCaj & "," & _
            (i + 1) & "," & Format(nValorPago, "#0.00") & ")"
    msConn.Execute sqltext
    
    On Error GoTo ErrAdm:
        ListaPagos.Col = 1
    On Error GoTo 0
    'SI HAY PROPINAS, MARCAR PARA PAGAR A MESEROS
    If Mid(ListaPagos.Text, 1, 7) = "PROPINA" Then
        
        sqltext = "INSERT INTO TRANSAC_PROP " & _
            " (NUM_TRANS,MESERO,CAJERO,TIPO_PAGO,LIN,MONTO) VALUES (" & _
            rs00!TRANS & "," & nMesero & "," & npNumCaj & "," & nTipoPago & "," & _
            i + 1 & "," & Format(nValorPago, "#0.00") & ")"
        msConn.Execute sqltext
    
    End If
Next

'ANEXA LAS TRANSACCIONES AL MENU DE TRANSACCIONES
Set rsTrans = New Recordset

sqltext = "SELECT * FROM TMP_TRANS WHERE MESA = " & nMesa & _
        " AND CUENTA = " & nCuenta & " ORDER BY LIN "
rsTrans.Open sqltext, msConn, adOpenStatic, adLockReadOnly

Dim MiFecha As String   'PARA CLIENTES

Do Until rsTrans.EOF
    CadenaSql = "INSERT INTO TRANSAC " & _
        "(NUM_TRANS,CAJA,CAJERO,MESA,MESERO,VALID,LIN,DESCRIP,CANT,DEPTO,PLU,ENVASE,PRECIO_UNIT,PRECIO,FECHA,HORA,TIPO,DESCUENTO,CUENTA) VALUES (" & _
        "" & rs00!TRANS & "," & rsTrans!caja & "," & rsTrans!CAJERO & "," & rsTrans!mesa & "," & rsTrans!MESERO & "," & rsTrans!VALID & "," & rsTrans!LIN & ",'" & _
        rsTrans!DESCRIP & "'," & rsTrans!CANT & "," & rsTrans!depto & "," & rsTrans!PLU & "," & _
        rsTrans!envase & "," & rsTrans!precio_unit & "," & rsTrans!precio & ",'" & rsTrans!FECHA & "','" & Time & "'" & _
        ",'" & rsTrans!TIPO & "'," & rsTrans!DESCUENTO & "," & nCuenta & ")"

    msConn.Execute CadenaSql

    MiFecha = rsTrans!FECHA
    rsTrans.MoveNext
Loop

'BORRA REGISTROS DE LA TEMPORAL
sqltext = "DELETE * FROM TMP_TRANS WHERE MESA = " & nMesa & _
        "AND CUENTA = " & nCuenta
msConn.Execute sqltext

'ACTUALIZA MESAS Y BORRA CUENTAS SIN USAR
rsCuentas.Close
sqltxt = "SELECT CUENTA,SUM(PRECIO) AS VALOR FROM TMP_TRANS " & _
        " WHERE MESA = " & nMesa & _
        " GROUP BY CUENTA "
rsCuentas.Open sqltxt, msConn, adOpenStatic, adLockOptimistic

If rsCuentas.RecordCount = 0 Then
    msConn.Execute "UPDATE Mesas SET ocupada = FALSE,MESERO_ACTUAL = 0 WHERE numero = " & nMesa
    msConn.Execute "DELETE * FROM TMP_CUENTAS WHERE MESA = " & nMesa
End If

'CON PAGOS A CREDITO. INSERTA INFO. DEL GRID INVISIBLE

If MSHFClientes.Rows > 0 Then
    Dim nVal1 As Integer
    Dim nVal2 As Integer
    Dim nVal3 As Single
    Dim nVal4 As Single
    Dim rsCli As Recordset
    
    Set rsCli = New Recordset
    
    For i = 0 To (MSHFClientes.Rows - 1)
        On Error GoTo ErrAdm:
            MSHFClientes.Row = i
            MSHFClientes.Col = 0: nVal1 = MSHFClientes.Text
            MSHFClientes.Col = 1: nVal2 = MSHFClientes.Text
            MSHFClientes.Col = 2: nVal3 = MSHFClientes.Text
        On Error GoTo 0
        
        rsCli.Open "SELECT * FROM CLIENTES " & _
                " WHERE CODIGO = " & nVal2, msConn, adOpenStatic, adLockOptimistic
        nVal4 = 0#
        If Not rsCli.EOF Then
            If rsCli!saldo < 0# Then
                If Abs(rsCli!saldo) > nVal3 Then
                    nVal4 = nVal3
                ElseIf Abs(rsCli!saldo) < nVal3 Then
                    nVal4 = Abs(rsCli!saldo)
                ElseIf Abs(rsCli!saldo) = nVal3 Then
                    nVal4 = Abs(rsCli!saldo)
                End If
            End If
        End If
        
        msConn.Execute "INSERT INTO TRANSAC_CLI " & _
                " (CODIGO_TP,CODIGO_CLI,NUM_TRANS,MONTO,FECHA,RECIBIDO) " & _
                " VALUES (" & _
                nVal1 & "," & nVal2 & "," & rs00!TRANS & "," & nVal3 & ",'" & _
                MiFecha & "'," & nVal4 & ")"
        
        msConn.Execute "UPDATE CLIENTES " & _
                " SET SALDO = SALDO + " & Format(nVal3, "#0.00") & _
                " WHERE CODIGO = " & nVal2
        rsCli.Close
    Next
End If

rs07.Open "SELECT a.lin,a.descrip,a.cant," & _
        " format(precio_unit,'##0.00') as mPrecio_unit," & _
        " format(precio,'##0.00') as mPrecio," & _
        " a.cajero,a.mesero,a.depto,a.plu,a.mesa,a.valid, " & _
        " a.envase,a.fecha,a.hora,a.tipo,a.descuento,a.cuenta, " & _
        " a.caja " & _
        " FROM tmp_trans as a " & _
        " WHERE a.mesa = " & nMesa, msConn, adOpenStatic, adLockReadOnly
CajLin = rs07.RecordCount

Set PLU.PlatosMesa.DataSource = rs07

rs07.Close
rs07.Open "SELECT sum(a.precio) as precio FROM tmp_trans as a WHERE a.mesa = " & nMesa, msConn, adOpenStatic, adLockReadOnly
PLU.SubTot = FormatCurrency(IIf(IsNull(rs07!precio), 0#, rs07!precio), 2)
On Error Resume Next
PLU.SubTot = FormatCurrency((PLU.SubTot + (rs07!precio * iISC)), 2)
iISCTransaccion = rs07!precio * iISC
SBTot = Format(PLU.SubTot, "standard")
On Error GoTo 0
rs07.Close

'LA INSTRUCCION CommitTrans EJECUTA TODAS ESTAS ACTUALIZACIONES
msConn.Execute "DELETE * FROM TMP_PAR_PAGO WHERE MESA = " & nMesa
msConn.Execute "DELETE * FROM TMP_PAR_PROP WHERE MESA = " & nMesa
msConn.CommitTrans

Exit Sub

ErrAdm:

If iError < 7 Then
    iError = iError + 1
    Resume
Else
    MsgBox Err.Description, vbCritical, "OCURRIO UN ERROR, ANOTE LOS DATOS EN PANTALLA"
    Exit Sub
End If


End Sub
Private Sub ImprFactura(iSoloJournal As Byte)
Dim i As Integer
Dim nMiSub As Single
Dim nCodigoPago As Integer
Dim sqltext As String
Dim LinTx As String
Dim rsCuenta As Recordset
Dim MiMatriz(0, 3) As String
Dim MiLen1, Milen2 As Integer
Dim n1 As Single
Dim n2 As Single
Dim nImp As Integer
Dim iSlip As Integer
Dim nEspacio As Integer
Dim nLinDetalle As Integer
Dim STATION_2PRINT As Integer
Dim vResp As Variant
Dim LOCAL_ISC As Single

LOCAL_ISC = iISCTransaccion

'DEJA DE OCUPAR LA MESA
StatMesa nMesa, 0
On Error GoTo ErrAdm:
nImp = 0: nEspacio = 0

If SLIP_OK = True Then
    'EN DBMS ESTA MARCADO QUE ESTE CLIENTE IMPRIME CON SLIP
    If iSoloJournal = 0 Then
        STATION_2PRINT = FPTR_S_SLIP
    Else
        STATION_2PRINT = FPTR_S_RECEIPT
    End If
Else
    STATION_2PRINT = FPTR_S_RECEIPT
End If

'SON 2 CICLOS UNO PARA RECEIPT/SLIP EL OTRO PARA JOURNAL
For nImp = 0 To 1
    Set rsCuenta = New Recordset
    nMiSub = 0#

    sqltext = "SELECT * FROM TMP_TRANS WHERE MESA = " & nMesa & "AND CUENTA = " & nCuenta & " ORDER BY lin "
    rsCuenta.Open sqltext, msConn, adOpenStatic, adLockReadOnly
    If nImp = 0 Then
        If SLIP_OK = True Then
            If iSoloJournal = 0 Then
                vResp = MsgBox("COLOQUE EL PAPEL EN LA RANURA DE LA IMPRESORA y PRESIONE ENTER", vbInformation + vbYesNoCancel, "PREPARANDOSE PARA IMPRIMIR EN EL SLIP PRINTER")
                If vbresp = vbNo Or vbresp = vbCancel Then
                    MsgBox "SE CANCELO LA IMPRESION EN EL SLIP PRINTER", vbExclamation, "IMPRESION CANCELADA"
                    Exit Sub
                End If
                Sys_Pos.Coptr1.BeginInsertion (5000)
                nEspacio = 8    '16
                For iSlip = 0 To 11     '16
                    Sys_Pos.Coptr1.PrintNormal STATION_2PRINT, Space(2) & Chr(&HD) & Chr(&HA)
                Next
            End If
        End If
        Sys_Pos.Cocash1.OpenDrawer
        Sys_Pos.Coptr1.PrintNormal STATION_2PRINT, Space(2) & Chr(&HD) & Chr(&HA)
        If iSoloJournal <> 0 Then
            'isolojpurnal es 1, asi que voy a imprimir en la factura normal
            Sys_Pos.Coptr1.PrintNormal STATION_2PRINT, rs00!DESCRIP & Chr(&HD) & Chr(&HA)
            Sys_Pos.Coptr1.PrintNormal STATION_2PRINT, rs00!RAZ_SOC & Chr(&HD) & Chr(&HA)
            Sys_Pos.Coptr1.PrintNormal STATION_2PRINT, "RUC:" & rs00!RUC & Chr(&HD) & Chr(&HA)
            Sys_Pos.Coptr1.PrintNormal STATION_2PRINT, Mid$(rs00!Direccion, 1, 25) & Chr(&HD) & Chr(&HA)
        End If
    Else
        If iSoloJournal <> 0 Then
            For i = 1 To 10
                Sys_Pos.Coptr1.PrintNormal STATION_2PRINT, Space(2) & Chr(&HD) & Chr(&HA)
            Next
            Sys_Pos.Coptr1.CutPaper 100
        End If
        nEspacio = 0
        STATION_2PRINT = FPTR_S_JOURNAL
    End If

IntentarOtraVez:
    rc = Sys_Pos.Coptr1.PrintNormal(STATION_2PRINT, Space(2) & Chr(&HD) & Chr(&HA))
    If rc <> OPOS_SUCCESS Then
        nIntentos = nIntentos + 1
        If nIntentos > 3 Then
            If SLIP_OK = True Then
                MsgBox "FAVOR RETIRE EL PAPEL DEL SLIP DE LA IMPRESORA, PARA PODER CONTINUAR", vbInformation, "RETIRE EL PAPEL!!!"
                iSoloJournal = 1
            Else
                MsgBox "POSIBLE PROBLEMA DE IMPRESION. VERIFIQUE", vbInformation, BoxTit
            End If
        Else
            GoTo IntentarOtraVez:
        End If
    End If

    Sys_Pos.Coptr1.PrintNormal STATION_2PRINT, Space(nEspacio) & "SERIAL:" & rs00!SERIAL & Chr(&HD) & Chr(&HA)
    Sys_Pos.Coptr1.PrintNormal STATION_2PRINT, Space(nEspacio) & "TRANS# " & rs00!TRANS + 1 & Chr(&HD) & Chr(&HA)
    Sys_Pos.Coptr1.PrintNormal STATION_2PRINT, Space(nEspacio) & "Mesero : " & cNomMesero & Chr(&HD) & Chr(&HA)
    Sys_Pos.Coptr1.PrintNormal STATION_2PRINT, Space(nEspacio) & "Cajero : " & cNomCaj & Chr(&HD) & Chr(&HA)

    Sys_Pos.Coptr1.PrintNormal STATION_2PRINT, Space(nEspacio) & "Cuenta # " & nCuenta & Chr(&HD) & Chr(&HA)
    
    If cNombreCliente <> "" Then Sys_Pos.Coptr1.PrintNormal STATION_2PRINT, Space(nEspacio) & "Cliente : " & cNombreCliente & Chr(&HD) & Chr(&HA)
    Sys_Pos.Coptr1.PrintNormal STATION_2PRINT, Space(2) & Chr(&HD) & Chr(&HA)
    Sys_Pos.Coptr1.PrintNormal STATION_2PRINT, Space(nEspacio) & "Mesa : " & nMesa & Chr(&HD) & Chr(&HA)
    Sys_Pos.Coptr1.PrintNormal STATION_2PRINT, Space(nEspacio) & "------------------------------" & Chr(&HD) & Chr(&HA)
    'IMPRESION DETALLE DE LOS PLATOS
    Do Until rsCuenta.EOF
        If SLIP_OK = True And nEspacio = 8 Then     '16
            MiMatriz(0, 0) = FormatTexto(rsCuenta!DESCRIP, 30)  '35
        Else
            MiMatriz(0, 0) = FormatTexto(rsCuenta!DESCRIP, 15)
        End If
        If iSoloJournal = 0 Then
            MiMatriz(0, 1) = Format(rsCuenta!CANT, "general number")
            MiMatriz(0, 2) = Format(rsCuenta!precio_unit, "#,###.00")
            MiMatriz(0, 3) = Format(rsCuenta!precio, "#,###.00")
        Else
            MiMatriz(0, 1) = Format(rsCuenta!CANT, "general number")
            MiMatriz(0, 2) = Format(rsCuenta!precio, "#,###.00")
        End If

        nMiSub = nMiSub + rsCuenta!precio
        MiLen1 = Len(MiMatriz(0, 1))
        Milen2 = Len(MiMatriz(0, 2))
        Milen3 = Len(MiMatriz(0, 3))
        If iSoloJournal = 0 Then
            LinTx = MiMatriz(0, 0) & Space(5 - MiLen1) & _
                MiMatriz(0, 1) & Space(12 - Milen2) & _
                MiMatriz(0, 2) & Space(11 - Milen3) & _
                MiMatriz(0, 3)
        Else
            LinTx = MiMatriz(0, 0) & Space(5 - MiLen1) & _
                MiMatriz(0, 1) & Space(10 - Milen2) & _
                MiMatriz(0, 2)
        End If
        Sys_Pos.Coptr1.PrintNormal STATION_2PRINT, Space(nEspacio) & LinTx & Chr(&HD) & Chr(&HA)
        nLinDetalle = nLinDetalle + 1
        'IMPRESION DE LA PROXIMA PAGINA
        If SLIP_OK And nImp = 0 Then
            If iSoloJournal = 0 Then
                If nLinDetalle = 20 Or nLinDetalle = 40 Or nLinDetalle = 60 Then
                    Sys_Pos.Coptr1.PrintNormal STATION_2PRINT, Space(nEspacio) & Chr(&HD) & Chr(&HA)
                    Sys_Pos.Coptr1.PrintNormal STATION_2PRINT, Space(nEspacio) & "Impresión continua en la siguiente página..." & Chr(&HD) & Chr(&HA)
                    'agregar el manejo de error
                    Sys_Pos.Coptr1.BeginRemoval (OPOS_FOREVER)

                    MsgBox "INSERTE EL PROXIMO FORMULARIO PARA PODER CONTINUAR" & vbCrLf & _
                           "PRESIONE ENTER PARA CONTINUAR", vbInformation, BoxTit

                    Sys_Pos.Coptr1.BeginInsertion (OPOS_FOREVER)
                    nEspacio = 8
                    For iSlip = 0 To 11
                        Sys_Pos.Coptr1.PrintNormal STATION_2PRINT, Space(2) & Chr(&HD) & Chr(&HA)
                    Next
                    Sys_Pos.Coptr1.PrintNormal STATION_2PRINT, Space(nEspacio) & "Viene de la página anterior" & Chr(&HD) & Chr(&HA)
                End If
            End If
        End If
        rsCuenta.MoveNext
    Loop

    'FORMAS DE PAGOS. FALTA MANEJO DE ERROR PARA CUANDO SE ACABA EL PAPEL.
    'Y AUN ESTA IMPRIMIENDO LAS FORMAS DE PAGO
    On Error Resume Next

    Sys_Pos.Coptr1.PrintNormal STATION_2PRINT, Space(nEspacio) & "------------------------------" & Chr(&HD) & Chr(&HA)
    Milen2 = Len(Format(nMiSub, "CURRENCY"))
    Sys_Pos.Coptr1.PrintNormal STATION_2PRINT, Space(nEspacio) & "     Sub-Total :" & Space(14 - Milen2) & Format(nMiSub, "CURRENCY") & Chr(&HD) & Chr(&HA)

    Milen2 = Len(Format(LOCAL_ISC, "STANDARD"))
    'txtString = "ITBMS (5%): " & Space(18 - Milen2) & Format(iISCTransaccion, "STANDARD")
    txtString = Space(nEspacio) & "ITBMS (5%): " & Space(18 - Milen2) & Format(LOCAL_ISC, "STANDARD")
    Sys_Pos.Coptr1.PrintNormal STATION_2PRINT, txtString & Chr(&HD) & Chr(&HA)
    
    Call PutISC(Format(iISCTransaccion, "STANDARD"))

    Milen2 = Len(Format(nMiSub + FormatCurrency(LOCAL_ISC, 2), "STANDARD"))
    'txtString = "TOTAL     : " & Space(18 - Milen2) & Format(nMiSub + FormatCurrency(iISCTransaccion, 2), "STANDARD")
    txtString = Space(nEspacio) & "TOTAL     : " & Space(18 - Milen2) & Format(nMiSub + FormatCurrency(LOCAL_ISC, 2), "STANDARD")
    Sys_Pos.Coptr1.PrintNormal STATION_2PRINT, txtString & Chr(&HD) & Chr(&HA)
    'Sys_Pos.Coptr1.PrintNormal STATION_2PRINT, Space(2) & Chr(&HD) & Chr(&HA)
    iISCTransaccion = 0
    On Error GoTo 0
    '*******************************************
    '*******************************************
    
    If OPEN_PROPINA = False Then
        NLEN = Len(PROPINA_DESCRIP)
        Sys_Pos.Coptr1.PrintNormal STATION_2PRINT, Space(nEspacio) & PROPINA_DESCRIP & " " & Space(25 - NLEN) & Format(nProp, "##0.00") & Chr(&HD) & Chr(&HA)
    End If
    If SLIP_OK = False Then Sys_Pos.Coptr1.PrintNormal STATION_2PRINT, Space(3) & Chr(&HD) & Chr(&HA)
    If nFlagParciales = 1 Then Sys_Pos.Coptr1.PrintNormal STATION_2PRINT, Space(nEspacio) & "** Mesa tiene Pago Parcial" & Chr(&HD) & Chr(&HA)
    For i = 0 To (ListaPagos.Rows - 1)
        ListaPagos.Row = i
        ListaPagos.Col = 0
        nCodigoPago = ListaPagos.Text
        ListaPagos.Col = 1
        MiMatriz(0, 0) = FormatTexto(ListaPagos.Text, 15)
        ListaPagos.Col = 2
        If nCodigoPago = 99 Then
            MiMatriz(0, 1) = Format(ListaPagos.Text * (-1), "##,##0.00")
        Else
            ListaPagos.Col = 2
            n1 = Format(ListaPagos.Text, "##,##0.00")
            ListaPagos.Col = 3
            n2 = Format(ListaPagos.Text, "##,##0.00")
            If n1 <> n2 Then
                MiMatriz(0, 1) = Format(n2, "##,##0.00")
            Else
                MiMatriz(0, 1) = Format(n1, "##,##0.00")
            End If
        End If
        Milen2 = Len(MiMatriz(0, 1))
        Sys_Pos.Coptr1.PrintNormal STATION_2PRINT, Space(nEspacio) & MiMatriz(0, 0) & Space(15 - Milen2) & MiMatriz(0, 1) & Chr(&HD) & Chr(&HA)
    Next
    Sys_Pos.Coptr1.PrintNormal STATION_2PRINT, Space(2) & Chr(&HD) & Chr(&HA)
    Milen2 = Len(Format(nCambio, "currency"))
    Sys_Pos.Coptr1.PrintNormal STATION_2PRINT, Space(nEspacio) & "Su CAMBIO : " & Space(18 - Milen2) & Format(nCambio, "currency") & Chr(&HD) & Chr(&HA)
    Sys_Pos.Coptr1.PrintNormal STATION_2PRINT, Space(3) & Chr(&HD) & Chr(&HA)
    Sys_Pos.Coptr1.PrintNormal STATION_2PRINT, Space(nEspacio) & "FEC: " & Format(Date, "short date") & " HORA: " & Mid(Time, 1, 5) & Mid(Time, 10, 4) & Chr(&HD) & Chr(&HA)
    
    If nImp = 0 Then
        If SLIP_OK = False Then Sys_Pos.Coptr1.PrintNormal STATION_2PRINT, Space(2) & Chr(&HD) & Chr(&HA)
        If SLIP_OK = False Then Sys_Pos.Coptr1.PrintNormal STATION_2PRINT, rs00!MENSAJE & Chr(&HD) & Chr(&HA)
        If SLIP_OK = False Then Sys_Pos.Coptr1.PrintNormal STATION_2PRINT, Space(2) & Chr(&HD) & Chr(&HA)

        If CtaPago.chkInfo.Value = 1 Then
            Sys_Pos.Coptr1.PrintNormal STATION_2PRINT, "NOMBRE : __________________" & Chr(&HD) & Chr(&HA)
            Sys_Pos.Coptr1.PrintNormal STATION_2PRINT, "CEDULA : ____________" & Chr(&HD) & Chr(&HA)
            Sys_Pos.Coptr1.PrintNormal STATION_2PRINT, Space(2) & Chr(&HD) & Chr(&HA)
        End If
    
    End If
    If nImp = iSoloJournal Then
        If SLIP_OK = True Then Sys_Pos.Coptr1.BeginRemoval (5000)
    Else
        Sys_Pos.Coptr1.PrintNormal STATION_2PRINT, "==============================" & Chr(&HD) & Chr(&HA)
    End If
    rsCuenta.Close
Next
cNombreCliente = ""
On Error GoTo 0
nFlagParciales = 0
Exit Sub

ErrAdm:
Resume Next
End Sub
Private Sub SetupPantalla()
    With ListaPagos
        .ColWidth(0) = 0: .ColWidth(1) = 2000: .ColWidth(2) = 800:
        .ColWidth(3) = 0
    End With
End Sub
'Private Sub CargaFormasPago()
'Dim MiTop As Integer, MiLeft As Integer, StayLeft As Integer
'Dim numplu As Integer
'Dim sqltext As String
'
'Set RSPAGOS = New Recordset
'Set RSPROPINAS = New Recordset
'
'sqltext = "SELECT * FROM pagos WHERE CODIGO <> 999 AND CODIGO <> 99 ORDER BY CODIGO"
'RSPAGOS.Open sqltext, msConn, adOpenStatic, adLockOptimistic
'
'sqltext = "SELECT * FROM pagos WHERE TIPO = 'TJ' ORDER BY CODIGO"
'RSPROPINAS.Open sqltext, msConn, adOpenStatic, adLockOptimistic
'
'For i = 1 To 12
'    Load cmdFPagos(i)
'Next
'
'For i = 1 To 8
'    Load cmdPropina(i)
'Next
'
'MiTop = 360: StayLeft = 120
'MiLeft = 0: numplu = 0
'
''codigo,tipo,descrip
'Do Until RSPAGOS.EOF
'    If numplu < 1 Then
'        cmdFPagos(numplu).Caption = RSPAGOS!descrip
'        cmdFPagos(numplu).Tag = RSPAGOS!codigo
'    Else
'        If Not IsObject(cmdFPagos(numplu)) Then
'           Load cmdFPagos(numplu)
'        End If
'        cmdFPagos(numplu).Visible = True
'        cmdFPagos(numplu).Caption = RSPAGOS!descrip
'        cmdFPagos(numplu).Tag = RSPAGOS!codigo
'        cmdFPagos(numplu).Left = MiLeft + StayLeft
'        cmdFPagos(numplu).Top = MiTop
'        StayLeft = 120
'    End If
'    numplu = numplu + 1
'    MiLeft = MiLeft + 1440
'    If numplu = 4 Or numplu = 8 Or numplu = 12 Then
'        MiTop = MiTop + 800
'        MiLeft = 0
'    End If
'    If numplu = 12 Then Exit Do
'    RSPAGOS.MoveNext
'Loop
'
'MiTop = 360: StayLeft = 120
'MiLeft = 0: numplu = 0
'
'Do Until RSPROPINAS.EOF
'    If numplu < 1 Then
'        cmdPropina(numplu).Caption = RSPROPINAS!descrip
'        cmdPropina(numplu).Tag = RSPROPINAS!codigo
'    Else
'        If Not IsObject(cmdPropina(numplu)) Then
'           Load cmdPropina(numplu)
'        End If
'        cmdPropina(numplu).Visible = True
'        cmdPropina(numplu).Caption = RSPROPINAS!descrip
'        cmdPropina(numplu).Tag = RSPROPINAS!codigo
'        cmdPropina(numplu).Left = MiLeft + StayLeft
'        cmdPropina(numplu).Top = MiTop
'        StayLeft = 120
'    End If
'    numplu = numplu + 1
'    MiLeft = MiLeft + 1440
'    If numplu = 4 Or numplu = 8 Or numplu = 12 Then
'        MiTop = MiTop + 800
'        MiLeft = 0
'    End If
'    If numplu = 8 Then Exit Do
'    RSPROPINAS.MoveNext
'Loop
'End Sub

Private Sub cdmBill_Click(Index As Integer)
nMntOculto = cdmBill(Index).Tag
lbMonto = Format(Val(nMntOculto), "standard")
cmdFPagos_Click (0)
End Sub

Private Sub Clear_Click()
nfPase = 0
lbMonto = Format(0#, "standard")
nMntOculto = ""
End Sub

Private Sub cmdDescGlob_Click()
Dim nMiDesc As Integer
Dim nDescAplicado As Single

nDescAplicado = Format(lbMonto, "STANDARD")

If nDescAplicado < 0.01 Then
    MsgBox "¡¡ NO PUEDE APLICAR ESE DESCUENTO !!", vbInformation, BoxTit
    OKGlobal = 0
    Exit Sub
ElseIf OrigSB <> rsCuentas!VALOR Then
    MsgBox "¡¡ ES IMPOSIBLE APLICAR DESCUENTO GLOBAL !!", vbInformation, BoxTit
    OKGlobal = 0
    Exit Sub
ElseIf nDescAplicado > rsCuentas!VALOR Then
    MsgBox "¡¡ ES IMPOSIBLE APLICAR ESTE DESCUENTO GLOBAL !!", vbInformation, BoxTit
    OKGlobal = 0
    Exit Sub
End If

txtInfo = "Escriba Clave para Descuento Global"
AskClave.Show 1

If OKGlobal = 1 Then
    OKGlobal = 0
    BoxPreg = "¿ DESEA APLICAR DESCUENTO DE " & Format(nDescAplicado, "currency") & "  ?"
    BoxResp = MsgBox(BoxPreg, vbQuestion + vbYesNo, BoxTit)
    If BoxResp = vbYes Then
        ListaPagos.AddItem 99 & Chr(9) & "DESC.GLOBAL" & Chr(9) & Format(nDescAplicado, "STANDARD") & Chr(9) & Format(nDescAplicado, "STANDARD")
        SBTot = rsCuentas!VALOR - nDescAplicado
        Label3.BackColor = &HFFC0FF
        Label2.BackColor = &HFFC0FF
        Label3 = Format(SBTot, "currency")
        Label2 = Format(SBTot, "currency")
        'OrigSB = 1
    End If
    nfPase = 0
    lbMonto = Format(0#, "standard")
    nMntOculto = ""
Else
    MsgBox "¡¡ USTED NO ESTA AUTORIZADO PARA HACER DESCUENTOS !!", vbInformation, BoxTit
End If
End Sub

Private Sub cmdFPagos_Click(Index As Integer)

If lbMonto < 0.01 Then Exit Sub

If PLU.PlatosMesa.Rows < 1 Then
    MsgBox "NO HAY NADA MARCADO, FAVOR MARQUE PRODUCTOS", vbInformation, BoxTit
    cmdSalir_Click
    Exit Sub
End If

RSPAGOS.MoveFirst
RSPAGOS.Find "CODIGO = " & cmdFPagos(Index).Tag
If Not RSPAGOS.EOF Then
    If RSPAGOS!CLIENTES = True Then
        MisClientes.Show 1
        If nCliNum = 0 Then
            'NO HACE NADA YA QUE NO SE MARCO UN CLIENTE
            Exit Sub
        Else
            MSHFClientes.AddItem cmdFPagos(Index).Tag & Chr(9) & nCliNum & Chr(9) & lbMonto
        End If
    End If
End If

nfPase = 0

SBTot = SBTot - lbMonto
If SBTot < 0# Then
    RSPROPINAS.MoveFirst
    RSPROPINAS.Find "CODIGO = " & cmdFPagos(Index).Tag
    If Not RSPROPINAS.EOF Then
        MsgBox "No puede cargar mas del SALDO A ESTA TARJETA", vbInformation, BoxTit
        SBTot = SBTot + lbMonto
        '''ListaPagos.RemoveItem (ListaPagos.Rows)
        Exit Sub
    End If
    Label2 = Format(0#, "currency")
    nCambio = SBTot * (-1)
    SBTot = 0#
Else
    Label2 = Format(SBTot, "currency")
End If
ListaPagos.AddItem cmdFPagos(Index).Tag & Chr(9) & cmdFPagos(Index).Caption & Chr(9) & Format(lbMonto - nCambio, "STANDARD") & Chr(9) & Format(lbMonto, "STANDARD")
nMntOculto = ""
lbMonto = Format(0#, "standard")
If SBTot = 0# Then
    'ImpresionFactura y Propinas
    If SLIP_OK = True Then
        vResp = MsgBox("¿ Desea Imprimir la FACTURA con el SLIP ?", vbQuestion + vbYesNo, "¿ Imprimir Factura Pre-Impresa ?")
    End If
    If vResp = vbYes Then
        Call ImprFactura(0)
    Else
        Call ImprFactura(1)
    End If
    
    'ImprFactura
    Actualizador
    If nCambio <> 0# Then
        Vuelto.Show 1
    End If
    nCambio = 0#
    Unload Me
End If
End Sub

Private Sub cmdPropina_Click(Index As Integer)
Dim nMasCxc As Single

nMasCxc = Format(lbMonto, "#####.00")

RSPROPINAS.MoveFirst
RSPROPINAS.Find "CODIGO = " & cmdPropina(Index).Tag
If Not RSPAGOS.EOF Then
    If RSPROPINAS!CLIENTES = True Then
        SBTot = SBTot + nMasCxc
        Label3.BackColor = &HFFC0FF
        Label2.BackColor = &HFFC0FF
        Label3 = Format(SBTot, "currency")
        Label2 = Format(SBTot, "currency")
        'OrigSB = 1
    End If
End If
ListaPagos.AddItem cmdPropina(Index).Tag & Chr(9) & "PROPINA " & cmdPropina(Index).Caption & Chr(9) & Format(lbMonto, "standard") & Chr(9) & Format(lbMonto, "standard")
nfPase = 0
nMntOculto = ""
lbMonto = Format(0#, "standard")

'ListaPagos.AddItem cmdPropina(Index).Tag & Chr(9) & "PROPINA " & cmdPropina(Index).Caption & Chr(9) & Format(lbMonto, "standard") & Chr(9) & Format(lbMonto, "standard")
'nfPase = 0
'nMntOculto = ""
'lbMonto = Format(0#, "standard")
End Sub

Private Sub cmdSalir_Click()
Unload Me
End Sub
Private Sub Command2_Click(Index As Integer)
Dim cChar As String
If nfPase = 0 Then
    nMntOculto = Command2(Index).Caption
Else
    nMntOculto = nMntOculto & Command2(Index).Caption
End If
lbMonto = Format(Val(nMntOculto) / 100, "standard")
nfPase = nfPase + 1
End Sub

Private Sub Form_Load()
Dim nErrCnt As Integer

On Error GoTo ErrorAdm:
Set rsCuentas = New Recordset

sqltxt = "SELECT CUENTA,SUM(PRECIO) AS VALOR " & _
        " FROM TMP_TRANS " & _
        " WHERE MESA = " & nMesa & _
        " AND CUENTA = " & nCta & _
        " GROUP BY CUENTA "
rsCuentas.Open sqltxt, msConn, adOpenStatic, adLockOptimistic

nfPase = 0
OKGlobal = 0
lbCuenta = nCta
' SE ESTA SELECCIONANDO LA CUENTA ACTUAL
    'lbCuenta = Str(rsCuentas!cuenta)
    'nCuenta = rsCuentas!cuenta
nCuenta = nCta

If OPEN_PROPINA = True Then
    OrigSB = Format(rsCuentas!VALOR + iISCTransaccion, "currency")
    SBTot = Format(rsCuentas!VALOR + iISCTransaccion, "STANDARD")
    Label3 = Format(rsCuentas!VALOR + iISCTransaccion, "currency")
    Label2 = Format(rsCuentas!VALOR + iISCTransaccion, "currency")
Else
    'DU LIBAN, SI ES MESA_BARRA, NO SE CALCULA EL % DE SERVICIO
    If nMesa <> rs00!mesa_barra Then
        MiPropina = RoundToNearest(SBTot * 0.1, 0.05, 1)
        nProp = MiPropina * 100
        nProp = nProp / 100
        OrigSB = Format(rsCuentas!VALOR + nProp + iISCTransaccion, "currency")
        SBTot = Format(rsCuentas!VALOR + nProp + iISCTransaccion, "STANDARD")
        Label3 = Format(rsCuentas!VALOR + nProp + iISCTransaccion, "currency")
        Label2 = Format(rsCuentas!VALOR + nProp + iISCTransaccion, "currency")
    Else
        OrigSB = Format(rsCuentas!VALOR + iISCTransaccion, "currency")
        SBTot = Format(rsCuentas!VALOR + iISCTransaccion, "STANDARD")
        Label3 = Format(rsCuentas!VALOR + iISCTransaccion, "currency")
        Label2 = Format(rsCuentas!VALOR + iISCTransaccion, "currency")
    End If
End If

nMntOculto = ""
Call CargaFormasPago(RSPAGOS, RSPROPINAS, CtaPago)
Call SetupPantalla
Exit Sub

ErrorAdm:
MsgBox Err.Description, vbCritical, "OCURRIO UN ERROR"
End Sub

Private Sub Label2_Click()
    nMntOculto = Label2.Caption
    lbMonto = Format(nMntOculto, "standard")
End Sub
