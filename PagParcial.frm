VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form PagParcial 
   BackColor       =   &H80000018&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "MODULO DE FACTURACION POR PAGOS PARCIALES"
   ClientHeight    =   8625
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11310
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "PagParcial.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   8625
   ScaleWidth      =   11310
   StartUpPosition =   2  'CenterScreen
   Begin VB.CheckBox chkInfo 
      BackColor       =   &H80000018&
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
      TabIndex        =   43
      Top             =   7080
      Width           =   3015
   End
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
      Left            =   9120
      TabIndex        =   29
      Top             =   -200
      Width           =   1935
      Begin VB.CommandButton cdmBill 
         Height          =   735
         Index           =   0
         Left            =   120
         Picture         =   "PagParcial.frx":030A
         Style           =   1  'Graphical
         TabIndex        =   40
         Tag             =   "5.00"
         Top             =   240
         Width           =   1695
      End
      Begin VB.CommandButton cdmBill 
         Height          =   735
         Index           =   1
         Left            =   120
         Picture         =   "PagParcial.frx":19E9
         Style           =   1  'Graphical
         TabIndex        =   39
         Tag             =   "10.00"
         Top             =   1020
         Width           =   1695
      End
      Begin VB.CommandButton cdmBill 
         Height          =   735
         Index           =   2
         Left            =   120
         Picture         =   "PagParcial.frx":307A
         Style           =   1  'Graphical
         TabIndex        =   38
         Tag             =   "15.00"
         Top             =   1800
         Width           =   1695
      End
      Begin VB.CommandButton cdmBill 
         Height          =   735
         Index           =   3
         Left            =   120
         Picture         =   "PagParcial.frx":4730
         Style           =   1  'Graphical
         TabIndex        =   37
         Tag             =   "20.00"
         Top             =   2580
         Width           =   1695
      End
      Begin VB.CommandButton cdmBill 
         Height          =   735
         Index           =   4
         Left            =   120
         Picture         =   "PagParcial.frx":5DF8
         Style           =   1  'Graphical
         TabIndex        =   36
         Tag             =   "25.00"
         Top             =   3380
         Width           =   1695
      End
      Begin VB.CommandButton cdmBill 
         Height          =   735
         Index           =   5
         Left            =   120
         Picture         =   "PagParcial.frx":74B7
         Style           =   1  'Graphical
         TabIndex        =   35
         Tag             =   "30.00"
         Top             =   4160
         Width           =   1695
      End
      Begin VB.CommandButton cdmBill 
         Height          =   735
         Index           =   6
         Left            =   120
         Picture         =   "PagParcial.frx":8B72
         Style           =   1  'Graphical
         TabIndex        =   34
         Tag             =   "35.00"
         Top             =   4940
         Width           =   1695
      End
      Begin VB.CommandButton cdmBill 
         Height          =   735
         Index           =   7
         Left            =   120
         Picture         =   "PagParcial.frx":A20D
         Style           =   1  'Graphical
         TabIndex        =   33
         Tag             =   "40.00"
         Top             =   5720
         Width           =   1695
      End
      Begin VB.CommandButton cdmBill 
         Height          =   735
         Index           =   8
         Left            =   120
         Picture         =   "PagParcial.frx":B8D4
         Style           =   1  'Graphical
         TabIndex        =   32
         Tag             =   "45.00"
         Top             =   6500
         Width           =   1695
      End
      Begin VB.CommandButton cdmBill 
         Height          =   735
         Index           =   9
         Left            =   120
         Picture         =   "PagParcial.frx":D00F
         Style           =   1  'Graphical
         TabIndex        =   31
         Tag             =   "50.00"
         Top             =   7280
         Width           =   1695
      End
      Begin VB.CommandButton cdmBill 
         Height          =   735
         Index           =   10
         Left            =   120
         Picture         =   "PagParcial.frx":E6CB
         Style           =   1  'Graphical
         TabIndex        =   30
         Tag             =   "100.00"
         Top             =   8080
         Width           =   1695
      End
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  'Right Justify
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
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   570
      Index           =   2
      Left            =   3720
      TabIndex        =   27
      Top             =   240
      Width           =   2415
   End
   Begin VB.CommandButton cmdAplic 
      BackColor       =   &H0000C000&
      Caption         =   "Aplicar Pagos"
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
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   26
      Top             =   7560
      Width           =   1815
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
      TabIndex        =   25
      Top             =   5880
      Width           =   1575
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  'Right Justify
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
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   570
      Index           =   0
      Left            =   3720
      TabIndex        =   18
      Top             =   840
      Width           =   2415
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid ListaPagos 
      Height          =   1575
      Left            =   6240
      TabIndex        =   23
      Top             =   480
      Width           =   2775
      _ExtentX        =   4895
      _ExtentY        =   2778
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
      Left            =   6600
      TabIndex        =   22
      Top             =   7800
      Width           =   1935
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H0000C0C0&
      Caption         =   "PROPINAS PARCIALES"
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
      TabIndex        =   20
      Top             =   5280
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
         TabIndex        =   21
         Top             =   360
         Width           =   1335
      End
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  'Right Justify
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
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   570
      Index           =   1
      Left            =   3720
      TabIndex        =   19
      Top             =   1440
      Width           =   2415
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H80000018&
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
      Left            =   6240
      TabIndex        =   2
      Top             =   2040
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
         Picture         =   "PagParcial.frx":FDA5
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
      BackColor       =   &H0000C0C0&
      Caption         =   "FORMAS DE PAGO PARCIALES"
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
      Top             =   2280
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
      Left            =   2400
      TabIndex        =   41
      Top             =   7440
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
   Begin VB.Label LbMesa 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   26.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   735
      Left            =   2040
      TabIndex        =   42
      Top             =   7515
      Width           =   3975
   End
   Begin VB.Label lbFact 
      BackColor       =   &H80000018&
      Caption         =   "Pagos Recibidos"
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
      Index           =   1
      Left            =   120
      TabIndex        =   28
      Top             =   360
      Width           =   3135
   End
   Begin VB.Label Label1 
      BackColor       =   &H80000018&
      Caption         =   "Pagos Parciales Recibidos"
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
      Left            =   6240
      TabIndex        =   24
      Top             =   240
      Width           =   2415
   End
   Begin VB.Label lbPend 
      BackColor       =   &H80000018&
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
      Top             =   1560
      Width           =   3135
   End
   Begin VB.Label lbFact 
      BackColor       =   &H80000018&
      Caption         =   "Total Factura"
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
      Index           =   0
      Left            =   120
      TabIndex        =   16
      Top             =   960
      Width           =   3135
   End
End
Attribute VB_Name = "PagParcial"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim nfPase As Integer
Dim nMntOculto As String
Dim RSPAGOS As Recordset    'Pagos
Dim RSPROPINAS As Recordset   'Propinas
Dim OrigSB As Single
'''''''''''Dim nProp As Single
'''''''''''Dim MiPropina As Single
Private Sub Actualizador()
Dim sqltext As String, ImpText As String
Dim MiValor As Currency
Dim nValorPago As Single
Dim nTipoPago As Integer, i As Integer
Dim iError As Integer

Set rsAcutalizacion = New Recordset
iError = 0
'Actualiza los valores de la factura
'INCREMENTA EL NUMERO DE TRANSACCION EN 1
msConn.BeginTrans
msConn.Execute "UPDATE ORGANIZACION SET TRANS = TRANS + 1"
        
''''''''''''''------------- msConnLoc.BeginTrans
        
For i = 0 To (ListaPagos.Rows - 1)
    On Error GoTo ErrAdm:
        ListaPagos.Row = i
        ListaPagos.Col = 0
        nTipoPago = ListaPagos.Text
        ListaPagos.Col = 2
        nValorPago = Format(ListaPagos.Text, "STANDARD")
        ListaPagos.Col = 1
    On Error GoTo 0

    'SI HAY PROPINAS, MARCAR PARA PAGAR A MESEROS
    If Mid(ListaPagos.Text, 1, 7) = "PROPINA" Then
        
        sqltext = "INSERT INTO TMP_PAR_PROP " & _
            " (CAJERO,MESA,MESERO,TIPO_PAGO,LIN,MONTO) VALUES (" & _
            npNumCaj & "," & nMesa & "," & nMesero & "," & _
            nTipoPago & "," & (i + 1) & "," & Format(nValorPago, "#0.00") & ")"
        
        msConn.Execute sqltext
        
        ''''''''''''''------------- msConnLoc.Execute sqltext
        
        GoTo Proximo:
    End If
    
    On Error GoTo ErrAdm:
        ListaPagos.Row = i
        ListaPagos.Col = 0
        nTipoPago = ListaPagos.Text
        ListaPagos.Col = 2
        nValorPago = Format(ListaPagos.Text, "STANDARD")
    On Error GoTo 0
    
    sqltext = "INSERT INTO TMP_PAR_PAGO " & _
            " (CAJERO,MESA,MESERO,TIPO_PAGO,LIN,MONTO) VALUES (" & _
            npNumCaj & "," & nMesa & "," & nMesero & "," & _
            nTipoPago & "," & (i + 1) & "," & Format(nValorPago, "#0.00") & ")"
    msConn.Execute sqltext
    
    ''''''''''''''------------- msConnLoc.Execute sqltext
Proximo:
Next

'CON PAGOS A CREDITO. INSERTA INFO. DEL GRID INVISIBLE

If MSHFClientes.Rows > 0 Then
    Dim nVal1 As Integer
    Dim nVal2 As Integer
    Dim nVal3 As Single

    For i = 0 To (MSHFClientes.Rows - 1)
        On Error GoTo ErrAdm:
            MSHFClientes.Row = i
            MSHFClientes.Col = 0: nVal1 = MSHFClientes.Text
            MSHFClientes.Col = 1: nVal2 = MSHFClientes.Text
            MSHFClientes.Col = 2: nVal3 = MSHFClientes.Text
        On Error GoTo 0
        
        msConn.Execute "INSERT INTO TMP_CLI " & _
            " (CODIGO_TP,CODIGO_CLI,MESA,MONTO) " & _
            " VALUES (" & _
            nVal1 & "," & nVal2 & "," & nMesa & "," & nVal3 & ")"
    Next
End If

msConn.CommitTrans
''''''''''''''------------- msConnLoc.CommitTrans

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

Exit Sub

ErrAdm:

If iError < 6 Then
    iError = iError + 1
    Resume
Else
    MsgBox Err.Description, vbCritical, "OCURRIO UN ERROR, ANOTE LOS DATOS EN PANTALLA"
    Exit Sub
End If

End Sub
Private Sub ImprFactura()
Dim i As Integer
Dim nMiSub As Single
Dim nCodigoPago As Integer
Dim sqltext As String
Dim LinTx As String
Dim MiMatriz(0, 3) As String
Dim MiLen1, Milen2 As Integer
Dim nImp As Integer
Dim iSlip As Integer
Dim nEspacio As Integer
Dim STATION_2PRINT As Integer
Dim vResp As Variant

'DEJA DE OCUPAR LA MESA
StatMesa nMesa, 0
nImp = 0: nEspacio = 0

STATION_2PRINT = FPTR_S_RECEIPT

For nImp = 0 To 1
    nMiSub = 0#
    
    If nImp = 0 Then
        'nEspacio = 16
        Sys_Pos.Cocash1.OpenDrawer
        Sys_Pos.Coptr1.PrintNormal STATION_2PRINT, Space(2) & Chr(&HD) & Chr(&HA)
        'Printer.Print Chr$(EPSON_RECEIPT)
        'If SLIP_OK = False Then Printer.Print rs00!descrip & Space(8) & Date
        Sys_Pos.Coptr1.PrintNormal STATION_2PRINT, rs00!DESCRIP & Chr(&HD) & Chr(&HA)
        Sys_Pos.Coptr1.PrintNormal STATION_2PRINT, rs00!RAZ_SOC & Chr(&HD) & Chr(&HA)
        Sys_Pos.Coptr1.PrintNormal STATION_2PRINT, "RUC:" & rs00!RUC & Chr(&HD) & Chr(&HA)
        Sys_Pos.Coptr1.PrintNormal STATION_2PRINT, Mid$(rs00!Direccion, 1, 25) & Chr(&HD) & Chr(&HA)
    Else
        For i = 1 To 10
            Sys_Pos.Coptr1.PrintNormal STATION_2PRINT, Space(2) & Chr(&HD) & Chr(&HA)
        Next
        Sys_Pos.Coptr1.CutPaper 100
        nEspacio = 0
        STATION_2PRINT = FPTR_S_JOURNAL
    End If
       
    rc = Sys_Pos.Coptr1.PrintNormal(STATION_2PRINT, Space(2) & Chr(&HD) & Chr(&HA))

    Sys_Pos.Coptr1.PrintNormal STATION_2PRINT, Space(nEspacio) & "SERIAL:" & rs00!SERIAL & " TRANS# " & rs00!TRANS + 1 & Chr(&HD) & Chr(&HA)
    Sys_Pos.Coptr1.PrintNormal STATION_2PRINT, Space(nEspacio) & "Mesero : " & cNomMesero & Chr(&HD) & Chr(&HA)
    Sys_Pos.Coptr1.PrintNormal STATION_2PRINT, Space(nEspacio) & "Cajero : " & cNomCaj & Chr(&HD) & Chr(&HA)

    If cNombreCliente <> "" Then Sys_Pos.Coptr1.PrintNormal STATION_2PRINT, Space(nEspacio) & "Cliente : " & cNombreCliente & Chr(&HD) & Chr(&HA)
    Sys_Pos.Coptr1.PrintNormal STATION_2PRINT, Space(nEspacio) & "PAGO PARCIAL" & Chr(&HD) & Chr(&HA)
    Sys_Pos.Coptr1.PrintNormal STATION_2PRINT, Space(2) & Chr(&HD) & Chr(&HA)
    Sys_Pos.Coptr1.PrintNormal STATION_2PRINT, Space(nEspacio) & "Mesa : " & nMesa & Chr(&HD) & Chr(&HA)
    Sys_Pos.Coptr1.PrintNormal STATION_2PRINT, Space(nEspacio) & "------------------------------" & Chr(&HD) & Chr(&HA)

    Sys_Pos.Coptr1.PrintNormal STATION_2PRINT, Space(2) & Chr(&HD) & Chr(&HA)
    Milen2 = Len(Format(OrigSB, "CURRENCY"))
    Sys_Pos.Coptr1.PrintNormal STATION_2PRINT, Space(nEspacio) & "Total Cuenta   :" & Space(14 - Milen2) & Format(OrigSB, "CURRENCY") & Chr(&HD) & Chr(&HA)
    Sys_Pos.Coptr1.PrintNormal STATION_2PRINT, Space(2) & Chr(&HD) & Chr(&HA)
    
    'IMPRESION DE PAGOS
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
            MiMatriz(0, 1) = Format(ListaPagos.Text, "##,##0.00")
        End If
        Milen2 = Len(MiMatriz(0, 1))
        Sys_Pos.Coptr1.PrintNormal STATION_2PRINT, Space(nEspacio) & MiMatriz(0, 0) & Space(15 - Milen2) & MiMatriz(0, 1) & Chr(&HD) & Chr(&HA)
    Next
    Sys_Pos.Coptr1.PrintNormal STATION_2PRINT, Space(2) & Chr(&HD) & Chr(&HA)
    Milen2 = Len(Format(nCambio, "currency"))
    Sys_Pos.Coptr1.PrintNormal STATION_2PRINT, Space(nEspacio) & "Su CAMBIO : " & Space(18 - Milen2) & Format(nCambio, "currency") & Chr(&HD) & Chr(&HA)
    Milen2 = Len(Format(SBTot, "currency"))
    Sys_Pos.Coptr1.PrintNormal STATION_2PRINT, Space(3) & Chr(&HD) & Chr(&HA)
    Sys_Pos.Coptr1.PrintNormal STATION_2PRINT, Space(nEspacio) & "SALDO PENDIENTE : " & Space(12 - Milen2) & Format(SBTot, "currency") & Chr(&HD) & Chr(&HA)
    Sys_Pos.Coptr1.PrintNormal STATION_2PRINT, Space(3) & Chr(&HD) & Chr(&HA)
    Sys_Pos.Coptr1.PrintNormal STATION_2PRINT, Space(nEspacio) & "FEC: " & Format(Date, "short date") & " HORA: " & Mid(Time, 1, 5) & Mid(Time, 10, 4) & Chr(&HD) & Chr(&HA)
    If nImp = 0 Then
        Sys_Pos.Coptr1.PrintNormal STATION_2PRINT, Space(2) & Chr(&HD) & Chr(&HA)
        Sys_Pos.Coptr1.PrintNormal STATION_2PRINT, rs00!MENSAJE & Chr(&HD) & Chr(&HA)
        Sys_Pos.Coptr1.PrintNormal STATION_2PRINT, Space(2) & Chr(&HD) & Chr(&HA)

        If PagParcial.chkInfo.Value = 1 Then
            Sys_Pos.Coptr1.PrintNormal STATION_2PRINT, "NOMBRE : __________________" & Chr(&HD) & Chr(&HA)
            Sys_Pos.Coptr1.PrintNormal STATION_2PRINT, "CEDULA : ____________" & Chr(&HD) & Chr(&HA)
            Sys_Pos.Coptr1.PrintNormal STATION_2PRINT, Space(2) & Chr(&HD) & Chr(&HA)
        End If
    Else
        Sys_Pos.Coptr1.PrintNormal STATION_2PRINT, "==============================" & Chr(&HD) & Chr(&HA)
    End If
Next
cNombreCliente = ""
End Sub
Private Sub SetupPantalla()
    With ListaPagos
        .ColWidth(0) = 0: .ColWidth(1) = 2000: .ColWidth(2) = 800:
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

Private Sub cmdAplic_Click()
'ImpresionFactura y Propinas
If ListaPagos.Rows < 1 Then
    MsgBox "¡¡ NO HAY NINGUN PAGO PARA APLICAR !!", vbExclamation, BoxTit
Else
    ImprFactura
    Actualizador
    Vuelto.Show 1
    nCambio = 0#
End If
Unload Me
End Sub

Private Sub cmdDescGlob_Click()
Dim nMiDesc As Integer
Dim nDescAplicado As Single

nDescAplicado = Format(lbMonto, "STANDARD")

If nDescAplicado < 0.01 Then
    MsgBox "¡¡ NO PUEDE APLICAR ESE DESCUENTO !!", vbExclamation, BoxTit
    OKGlobal = 0
    Exit Sub
ElseIf OrigSB <> SBTot Then
    MsgBox "¡¡ ES IMPOSIBLE APLICAR DESCUENTO GLOBAL !!", vbExclamation, BoxTit
    OKGlobal = 0
    Exit Sub
ElseIf nDescAplicado > SBTot Then
    MsgBox "¡¡ ES IMPOSIBLE APLICAR ESTE DESCUENTO GLOBAL !!", vbExclamation, BoxTit
    OKGlobal = 0
    Exit Sub
End If

txtInfo = "Escriba Clave para Descuento Global"
AskClave.Show 1

If OKGlobal = 1 Then
    OKGlobal = 0
    BoxPreg = "¿ DESEA APLICAR DESCUENTO DE " & Format(nDescAplicado, "standard") & "  ?"
    BoxResp = MsgBox(BoxPreg, vbQuestion + vbYesNo, BoxTit)
    If BoxResp = vbYes Then
        ListaPagos.AddItem 99 & Chr(9) & "DESC.GLOBAL" & Chr(9) & Format(nDescAplicado, "STANDARD")
        SBTot = SBTot - nDescAplicado
        Text1(0).BackColor = &HFFC0FF
        Text1(1).BackColor = &HFFC0FF
        Text1(0) = Format(SBTot, "currency")
        Text1(1) = Format(SBTot, "currency")
        'OrigSB = 1
    End If
    nfPase = 0
    lbMonto = Format(0#, "standard")
    nMntOculto = ""
Else
    MsgBox "USTED NO ESTA AUTORIZADO PARA HACER DESCUENTOS", vbExclamation, BoxTit
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
If lbMonto > SBTot Then
    MsgBox "NO SE PUEDE RECIBIR MAS DEL MONTO DE LA FACTURA", vbExclamation, BoxTit
    cmdSalir_Click
    Exit Sub
End If
SBTot = SBTot - lbMonto
If SBTot < 0# Then
    RSPROPINAS.MoveFirst
    RSPROPINAS.Find "CODIGO = " & cmdFPagos(Index).Tag
    If Not RSPROPINAS.EOF Then
        MsgBox "NO puede cargar mas del SALDO A ESTA TARJETA", vbInformation, BoxTit
        SBTot = SBTot + lbMonto
        '''ListaPagos.RemoveItem (ListaPagos.Rows)
        Exit Sub
    End If
    Text1(1) = Format(0#, "currency")
    nCambio = SBTot * (-1)
    SBTot = 0#
Else
    Text1(1) = Format(SBTot, "currency")
End If
ListaPagos.AddItem cmdFPagos(Index).Tag & Chr(9) & cmdFPagos(Index).Caption & Chr(9) & Format(lbMonto - nCambio, "STANDARD")
nMntOculto = ""
lbMonto = Format(0#, "standard")
End Sub

Private Sub cmdPropina_Click(Index As Integer)
Dim nMasCxc As Single

nMasCxc = Format(lbMonto, "#####.00")
On Error Resume Next
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
On Error GoTo 0
'ListaPagos.AddItem cmdPropina(Index).Tag & Chr(9) & "PROPINA " & cmdPropina(Index).Caption & Chr(9) & lbMonto
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
nfPase = 0
OKGlobal = 0

'If OPEN_PROPINA = True Then
'Else
'    MiPropina = Format(Round(SBTot * 0.1, 1) * 1#, "STANDARD")
'    nProp = MiPropina * 100
'    nProp = nProp / 100
'    SBTot = SBTot + nProp
'End If

Text1(0) = Format(SBTot, "currency")
Text1(1) = Format(SBTot, "currency")
OrigSB = Format(SBTot, "currency")

nMntOculto = ""
CargaFormasPago RSPAGOS, RSPROPINAS, PagParcial
SetupPantalla

Dim rsParciales As Recordset
Dim rsParcPro As Recordset
Dim lParc As Integer
Dim nPagosRec As Single

LbMesa = "Mesa # " & nMesa
nPagosRec = Format(0#, "standard")

Set rsParciales = New Recordset
rsParciales.Open "SELECT CAJERO,MESA,MESERO,TIPO_PAGO,LIN,MONTO " & _
            " FROM TMP_PAR_PAGO " & _
            " WHERE MESA = " & nMesa, msConn, adOpenDynamic, adLockOptimistic

Set rsParcPro = New Recordset
rsParcPro.Open "SELECT CAJERO,MESA,MESERO,TIPO_PAGO,LIN,MONTO FROM TMP_PAR_PROP " & _
            " WHERE MESA = " & nMesa, msConn, adOpenDynamic, adLockOptimistic

nFlagParciales = 0

Do Until rsParciales.EOF
    RSPAGOS.MoveFirst
    RSPAGOS.Find "CODIGO = " & rsParciales!TIPO_PAGO
    'ListaPagos.AddItem rsParciales!TIPO_PAGO & Chr(9) & rsPagos!descrip & Chr(9) & Format(rsParciales!MONTO, "STANDARD") & Chr(9) & Format(rsParciales!MONTO, "STANDARD")
    SBTot = SBTot - rsParciales!MONTO
    nPagosRec = nPagosRec + rsParciales!MONTO
    rsParciales.MoveNext
    nFlagParciales = 1
Loop

Text1(1) = Format(SBTot, "currency")
Text1(2) = Format(nPagosRec, "currency")

'Do Until rsParcPro.EOF
'    rsPagos.MoveFirst
'    rsPagos.Find "CODIGO = " & rsParcPro!TIPO_PAGO
'    'ListaPagos.AddItem rsParcPro!TIPO_PAGO & Chr(9) & "PROPINA " & rsPagos!descrip & Chr(9) & Format(rsParcPro!MONTO, "STANDARD") & Chr(9) & Format(rsParcPro!MONTO, "STANDARD")
'    rsParcPro.MoveNext
'    nFlagParciales = 1
'Loop

End Sub

