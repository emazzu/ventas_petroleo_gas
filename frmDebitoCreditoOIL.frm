VERSION 5.00
Object = "{F856EC8B-F03C-4515-BDC6-64CBD617566A}#6.0#0"; "fpSpr60.ocx"
Begin VB.Form frmDebitoCreditoOIL 
   BackColor       =   &H80000018&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Débito - Crédito (Oil)"
   ClientHeight    =   10155
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10665
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   10155
   ScaleWidth      =   10665
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txt_Libre 
      BackColor       =   &H80000018&
      Height          =   1545
      Left            =   135
      MultiLine       =   -1  'True
      TabIndex        =   26
      Top             =   4965
      Width           =   10410
   End
   Begin VB.TextBox txtTitulo 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000018&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      ForeColor       =   &H00808080&
      Height          =   285
      Index           =   22
      Left            =   180
      TabIndex        =   68
      Text            =   "Texto Libre"
      Top             =   4650
      Width           =   10350
   End
   Begin FPSpreadADO.fpSpread spdDet 
      Height          =   1425
      Left            =   120
      TabIndex        =   67
      Top             =   8640
      Width           =   10425
      _Version        =   393216
      _ExtentX        =   18389
      _ExtentY        =   2514
      _StockProps     =   64
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      SpreadDesigner  =   "frmDebitoCreditoOIL.frx":0000
   End
   Begin VB.TextBox txtFactura 
      BackColor       =   &H80000018&
      ForeColor       =   &H00808080&
      Height          =   315
      Left            =   1710
      Locked          =   -1  'True
      TabIndex        =   1
      Top             =   690
      Width           =   3675
   End
   Begin VB.TextBox txtTitulo 
      Appearance      =   0  'Flat
      BackColor       =   &H80000018&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      ForeColor       =   &H00808080&
      Height          =   240
      Index           =   4
      Left            =   270
      TabIndex        =   43
      Text            =   "Comprobante Nro."
      Top             =   720
      Width           =   1350
   End
   Begin VB.ComboBox cboCliente 
      Appearance      =   0  'Flat
      BackColor       =   &H80000018&
      ForeColor       =   &H00808080&
      Height          =   315
      ItemData        =   "frmDebitoCreditoOIL.frx":01FE
      Left            =   1710
      List            =   "frmDebitoCreditoOIL.frx":0208
      Sorted          =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   5
      Top             =   1770
      Width           =   3675
   End
   Begin VB.ComboBox cboMoneda 
      Appearance      =   0  'Flat
      BackColor       =   &H80000018&
      ForeColor       =   &H00808080&
      Height          =   315
      ItemData        =   "frmDebitoCreditoOIL.frx":021D
      Left            =   1680
      List            =   "frmDebitoCreditoOIL.frx":0227
      Style           =   2  'Dropdown List
      TabIndex        =   8
      Top             =   2850
      Width           =   3675
   End
   Begin VB.ComboBox cboEmpresa 
      Appearance      =   0  'Flat
      BackColor       =   &H80000018&
      ForeColor       =   &H00808080&
      Height          =   315
      ItemData        =   "frmDebitoCreditoOIL.frx":023C
      Left            =   1710
      List            =   "frmDebitoCreditoOIL.frx":0246
      Style           =   2  'Dropdown List
      TabIndex        =   4
      Top             =   1410
      Width           =   3675
   End
   Begin VB.TextBox txtTitulo 
      Appearance      =   0  'Flat
      BackColor       =   &H80000018&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      ForeColor       =   &H00808080&
      Height          =   240
      Index           =   14
      Left            =   270
      TabIndex        =   40
      Text            =   "Empresa"
      Top             =   1440
      Width           =   950
   End
   Begin VB.TextBox txtConceptoVenta 
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H00808080&
      Height          =   870
      Left            =   135
      MultiLine       =   -1  'True
      TabIndex        =   27
      Top             =   7485
      Width           =   10410
   End
   Begin VB.TextBox txtIvaImp 
      Alignment       =   1  'Right Justify
      BackColor       =   &H80000018&
      ForeColor       =   &H00808080&
      Height          =   330
      Left            =   6720
      Locked          =   -1  'True
      TabIndex        =   22
      Top             =   3435
      Width           =   1275
   End
   Begin VB.TextBox txtIvaPje 
      Alignment       =   1  'Right Justify
      BackColor       =   &H80000018&
      ForeColor       =   &H00808080&
      Height          =   330
      Left            =   3840
      TabIndex        =   19
      Top             =   3435
      Width           =   780
   End
   Begin VB.TextBox txtSubtotal 
      Alignment       =   1  'Right Justify
      BackColor       =   &H80000018&
      ForeColor       =   &H00808080&
      Height          =   330
      Left            =   390
      Locked          =   -1  'True
      TabIndex        =   18
      Top             =   4155
      Width           =   1590
   End
   Begin VB.TextBox txtTotal 
      Alignment       =   1  'Right Justify
      BackColor       =   &H80000018&
      ForeColor       =   &H00808080&
      Height          =   330
      Left            =   8730
      Locked          =   -1  'True
      TabIndex        =   25
      Top             =   4155
      Width           =   1590
   End
   Begin VB.TextBox txtTitulo 
      Appearance      =   0  'Flat
      BackColor       =   &H80000018&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      ForeColor       =   &H00808080&
      Height          =   240
      Index           =   6
      Left            =   270
      TabIndex        =   30
      Text            =   "Cliente"
      Top             =   1800
      Width           =   950
   End
   Begin VB.TextBox txtTitulo 
      Appearance      =   0  'Flat
      BackColor       =   &H80000018&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      ForeColor       =   &H00808080&
      Height          =   240
      Index           =   3
      Left            =   270
      TabIndex        =   29
      Text            =   "Moneda Tipo"
      Top             =   2880
      Width           =   950
   End
   Begin VB.TextBox txtTitulo 
      Appearance      =   0  'Flat
      BackColor       =   &H80000018&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      ForeColor       =   &H00808080&
      Height          =   240
      Index           =   10
      Left            =   270
      TabIndex        =   28
      Text            =   "Operación Tipo"
      Top             =   2520
      Width           =   1170
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H80000018&
      Height          =   3210
      Left            =   135
      TabIndex        =   31
      Top             =   90
      Width           =   5370
      Begin VB.TextBox txtTitulo 
         Appearance      =   0  'Flat
         BackColor       =   &H80000018&
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         ForeColor       =   &H00808080&
         Height          =   240
         Index           =   25
         Left            =   150
         TabIndex        =   73
         Text            =   "Provincia"
         Top             =   2070
         Width           =   950
      End
      Begin VB.ComboBox cboProvincia 
         Appearance      =   0  'Flat
         BackColor       =   &H80000018&
         ForeColor       =   &H00808080&
         Height          =   315
         ItemData        =   "frmDebitoCreditoOIL.frx":025B
         Left            =   1560
         List            =   "frmDebitoCreditoOIL.frx":0262
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   6
         Top             =   2010
         Width           =   3675
      End
      Begin VB.TextBox txtTitulo 
         Appearance      =   0  'Flat
         BackColor       =   &H80000018&
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         ForeColor       =   &H00808080&
         Height          =   285
         Index           =   21
         Left            =   120
         TabIndex        =   61
         Text            =   "Fecha Factura"
         Top             =   975
         Width           =   1365
      End
      Begin VB.TextBox txtFecha 
         BackColor       =   &H80000018&
         ForeColor       =   &H00808080&
         Height          =   300
         Left            =   1560
         Locked          =   -1  'True
         TabIndex        =   2
         Top             =   960
         Width           =   1140
      End
      Begin VB.TextBox txtContable 
         BackColor       =   &H80000018&
         ForeColor       =   &H00808080&
         Height          =   300
         Left            =   4005
         TabIndex        =   3
         Top             =   960
         Width           =   1230
      End
      Begin VB.TextBox txtTitulo 
         Appearance      =   0  'Flat
         BackColor       =   &H80000018&
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         ForeColor       =   &H00808080&
         Height          =   285
         Index           =   20
         Left            =   2775
         TabIndex        =   60
         Text            =   "Fecha Contable"
         Top             =   1005
         Width           =   1170
      End
      Begin VB.ComboBox cboOperacion 
         Appearance      =   0  'Flat
         BackColor       =   &H80000018&
         ForeColor       =   &H00808080&
         Height          =   315
         ItemData        =   "frmDebitoCreditoOIL.frx":026B
         Left            =   1545
         List            =   "frmDebitoCreditoOIL.frx":0275
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   7
         Top             =   2385
         Width           =   3690
      End
      Begin VB.TextBox txtTitulo 
         Appearance      =   0  'Flat
         BackColor       =   &H80000018&
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         ForeColor       =   &H00808080&
         Height          =   240
         Index           =   2
         Left            =   135
         TabIndex        =   42
         Text            =   "Comprobante Tipo"
         Top             =   270
         Width           =   1395
      End
      Begin VB.ComboBox cboComprobante 
         Appearance      =   0  'Flat
         BackColor       =   &H80000018&
         ForeColor       =   &H00808080&
         Height          =   315
         ItemData        =   "frmDebitoCreditoOIL.frx":028B
         Left            =   1575
         List            =   "frmDebitoCreditoOIL.frx":0298
         Style           =   2  'Dropdown List
         TabIndex        =   0
         Top             =   225
         Width           =   3660
      End
   End
   Begin VB.CommandButton cmdGuardar 
      Caption         =   "&Guardar"
      Enabled         =   0   'False
      Height          =   330
      Left            =   8865
      TabIndex        =   32
      Top             =   6720
      Width           =   1500
   End
   Begin VB.CommandButton cmdSalir 
      Caption         =   "&Salir"
      Height          =   330
      Left            =   5625
      TabIndex        =   33
      Top             =   6720
      Width           =   1500
   End
   Begin VB.CommandButton cmdCalcular 
      Caption         =   "&Calcular"
      Height          =   330
      Left            =   7245
      TabIndex        =   34
      Top             =   6720
      Width           =   1500
   End
   Begin VB.TextBox txtTitulo 
      Appearance      =   0  'Flat
      BackColor       =   &H80000018&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      ForeColor       =   &H00808080&
      Height          =   240
      Index           =   12
      Left            =   8805
      TabIndex        =   35
      Text            =   "Total"
      Top             =   3885
      Width           =   405
   End
   Begin VB.TextBox txtTitulo 
      Appearance      =   0  'Flat
      BackColor       =   &H80000018&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      ForeColor       =   &H00808080&
      Height          =   240
      Index           =   11
      Left            =   5100
      TabIndex        =   36
      Text            =   "Iva Importe"
      Top             =   3480
      Width           =   1470
   End
   Begin VB.TextBox txtTitulo 
      Appearance      =   0  'Flat
      BackColor       =   &H80000018&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      ForeColor       =   &H00808080&
      Height          =   240
      Index           =   9
      Left            =   1290
      TabIndex        =   37
      Text            =   "Subtotal"
      Top             =   3885
      Width           =   750
   End
   Begin VB.TextBox txtTitulo 
      Appearance      =   0  'Flat
      BackColor       =   &H80000018&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      ForeColor       =   &H00808080&
      Height          =   240
      Index           =   0
      Left            =   2535
      TabIndex        =   38
      Text            =   "Iva %"
      Top             =   3480
      Width           =   435
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H80000018&
      Height          =   3210
      Left            =   5580
      TabIndex        =   39
      Top             =   90
      Width           =   4965
      Begin VB.TextBox txtOperacionOrigen 
         BackColor       =   &H80000018&
         ForeColor       =   &H00808080&
         Height          =   300
         Left            =   4725
         TabIndex        =   66
         Top             =   1305
         Visible         =   0   'False
         Width           =   105
      End
      Begin VB.TextBox txtCotizaOrigen 
         BackColor       =   &H80000018&
         ForeColor       =   &H00808080&
         Height          =   300
         Left            =   4590
         TabIndex        =   65
         Top             =   1305
         Visible         =   0   'False
         Width           =   105
      End
      Begin VB.TextBox txtSubtotalOrigen 
         BackColor       =   &H80000018&
         ForeColor       =   &H00808080&
         Height          =   300
         Left            =   4455
         TabIndex        =   64
         Top             =   1305
         Visible         =   0   'False
         Width           =   105
      End
      Begin VB.TextBox txtTitulo 
         Appearance      =   0  'Flat
         BackColor       =   &H80000018&
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         ForeColor       =   &H00808080&
         Height          =   285
         Index           =   15
         Left            =   3015
         TabIndex        =   63
         Text            =   "Ajuste"
         Top             =   1350
         Width           =   480
      End
      Begin VB.TextBox txtAjuste 
         BackColor       =   &H80000018&
         ForeColor       =   &H00808080&
         Height          =   300
         Left            =   3600
         TabIndex        =   14
         Top             =   1305
         Width           =   645
      End
      Begin VB.TextBox txtIDcotizacionDolar 
         BackColor       =   &H80000018&
         ForeColor       =   &H00808080&
         Height          =   300
         Left            =   4320
         TabIndex        =   13
         Top             =   1305
         Visible         =   0   'False
         Width           =   105
      End
      Begin VB.CommandButton cmdCotiza 
         Caption         =   "Cotización"
         Height          =   300
         Left            =   1890
         TabIndex        =   62
         Top             =   1305
         Width           =   945
      End
      Begin VB.CommandButton cmdCondicion 
         Caption         =   "New"
         Height          =   280
         Left            =   4275
         TabIndex        =   59
         Top             =   225
         Width           =   500
      End
      Begin VB.CommandButton cmdVer 
         Caption         =   "Comprobante Origen"
         Height          =   330
         Left            =   3150
         TabIndex        =   58
         Top             =   2385
         Width           =   1635
      End
      Begin VB.TextBox txtTitulo 
         Appearance      =   0  'Flat
         BackColor       =   &H80000018&
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         ForeColor       =   &H00808080&
         Height          =   240
         Index           =   19
         Left            =   135
         TabIndex        =   57
         Text            =   "Comprobante"
         Top             =   2430
         Width           =   1035
      End
      Begin VB.TextBox txtOrigen 
         BackColor       =   &H80000018&
         ForeColor       =   &H00808080&
         Height          =   330
         Left            =   1170
         Locked          =   -1  'True
         TabIndex        =   17
         Top             =   2385
         Width           =   1950
      End
      Begin VB.ComboBox cboCotizacion 
         Appearance      =   0  'Flat
         BackColor       =   &H80000018&
         ForeColor       =   &H00808080&
         Height          =   315
         ItemData        =   "frmDebitoCreditoOIL.frx":02B6
         Left            =   1170
         List            =   "frmDebitoCreditoOIL.frx":02C0
         Style           =   2  'Dropdown List
         TabIndex        =   15
         Top             =   1665
         Width           =   3075
      End
      Begin VB.TextBox txtTitulo 
         Appearance      =   0  'Flat
         BackColor       =   &H80000018&
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         ForeColor       =   &H00808080&
         Height          =   285
         Index           =   5
         Left            =   135
         TabIndex        =   54
         Text            =   "Cotizacion"
         Top             =   1710
         Width           =   1260
      End
      Begin VB.CommandButton cmdNewCotizacionTexto 
         Caption         =   "New"
         Height          =   285
         Left            =   4275
         TabIndex        =   53
         Top             =   1665
         Width           =   500
      End
      Begin VB.TextBox txtTitulo 
         Appearance      =   0  'Flat
         BackColor       =   &H80000018&
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         ForeColor       =   &H00808080&
         Height          =   240
         Index           =   1
         Left            =   135
         TabIndex        =   52
         Text            =   "Concepto"
         Top             =   2070
         Width           =   1035
      End
      Begin VB.ComboBox cboConcepto 
         Appearance      =   0  'Flat
         BackColor       =   &H80000018&
         ForeColor       =   &H00808080&
         Height          =   315
         ItemData        =   "frmDebitoCreditoOIL.frx":02D1
         Left            =   1170
         List            =   "frmDebitoCreditoOIL.frx":02DB
         Style           =   2  'Dropdown List
         TabIndex        =   16
         Top             =   2025
         Width           =   3615
      End
      Begin VB.CommandButton cmdFormaPago 
         Caption         =   "New"
         Height          =   280
         Left            =   4275
         TabIndex        =   51
         Top             =   945
         Width           =   500
      End
      Begin VB.TextBox txtTipoCambio 
         BackColor       =   &H80000018&
         ForeColor       =   &H00808080&
         Height          =   300
         Left            =   1170
         TabIndex        =   12
         Top             =   1305
         Width           =   645
      End
      Begin VB.TextBox txtTitulo 
         Appearance      =   0  'Flat
         BackColor       =   &H80000018&
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         ForeColor       =   &H00808080&
         Height          =   240
         Index           =   13
         Left            =   135
         TabIndex        =   50
         Text            =   "Tipo Cambio"
         Top             =   1350
         Width           =   990
      End
      Begin VB.TextBox txtTitulo 
         Appearance      =   0  'Flat
         BackColor       =   &H80000018&
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         ForeColor       =   &H00808080&
         Height          =   240
         Index           =   17
         Left            =   135
         TabIndex        =   49
         Text            =   "Forma Pago"
         Top             =   990
         Width           =   950
      End
      Begin VB.ComboBox cboCuentaBancaria 
         Appearance      =   0  'Flat
         BackColor       =   &H80000018&
         ForeColor       =   &H00808080&
         Height          =   315
         ItemData        =   "frmDebitoCreditoOIL.frx":02F0
         Left            =   1170
         List            =   "frmDebitoCreditoOIL.frx":02FA
         Style           =   2  'Dropdown List
         TabIndex        =   11
         Top             =   945
         Width           =   3075
      End
      Begin VB.TextBox txtTitulo 
         Appearance      =   0  'Flat
         BackColor       =   &H80000018&
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         ForeColor       =   &H00808080&
         Height          =   240
         Index           =   8
         Left            =   135
         TabIndex        =   48
         Text            =   "Vencimiento"
         Top             =   630
         Width           =   990
      End
      Begin VB.TextBox txtVencimiento 
         BackColor       =   &H80000018&
         ForeColor       =   &H00808080&
         Height          =   300
         Left            =   1170
         TabIndex        =   10
         Top             =   585
         Width           =   3615
      End
      Begin VB.TextBox txtTitulo 
         Appearance      =   0  'Flat
         BackColor       =   &H80000018&
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         ForeColor       =   &H00808080&
         Height          =   240
         Index           =   16
         Left            =   135
         TabIndex        =   47
         Text            =   "Condicion"
         Top             =   270
         Width           =   950
      End
      Begin VB.ComboBox cboCondicion 
         Appearance      =   0  'Flat
         BackColor       =   &H80000018&
         ForeColor       =   &H00808080&
         Height          =   315
         ItemData        =   "frmDebitoCreditoOIL.frx":030F
         Left            =   1170
         List            =   "frmDebitoCreditoOIL.frx":0319
         Style           =   2  'Dropdown List
         TabIndex        =   9
         Top             =   225
         Width           =   3030
      End
   End
   Begin VB.Frame Frame4 
      BackColor       =   &H80000018&
      Height          =   645
      Left            =   135
      TabIndex        =   41
      Top             =   6540
      Width           =   10410
      Begin VB.CheckBox chk_Resguardo 
         BackColor       =   &H80000018&
         Caption         =   "Emisiòn de Resguardo"
         Height          =   285
         Left            =   1665
         TabIndex        =   69
         Top             =   225
         Width           =   2085
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H80000018&
      Height          =   1320
      Left            =   135
      TabIndex        =   44
      Top             =   3255
      Width           =   10410
      Begin VB.TextBox txt_IDprovincia_Ent 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000018&
         ForeColor       =   &H00808080&
         Height          =   285
         Left            =   270
         Locked          =   -1  'True
         TabIndex        =   72
         Top             =   405
         Visible         =   0   'False
         Width           =   315
      End
      Begin VB.TextBox txtIIBBPje 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000018&
         ForeColor       =   &H00808080&
         Height          =   330
         Left            =   3690
         TabIndex        =   21
         Top             =   900
         Width           =   780
      End
      Begin VB.TextBox txtTitulo 
         Appearance      =   0  'Flat
         BackColor       =   &H80000018&
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         ForeColor       =   &H00808080&
         Height          =   240
         Index           =   24
         Left            =   2385
         TabIndex        =   71
         Text            =   "IIBB %"
         Top             =   945
         Width           =   1080
      End
      Begin VB.TextBox txtIIBBImp 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000018&
         ForeColor       =   &H00808080&
         Height          =   330
         Left            =   6570
         Locked          =   -1  'True
         TabIndex        =   24
         Top             =   900
         Width           =   1275
      End
      Begin VB.TextBox txtTitulo 
         Appearance      =   0  'Flat
         BackColor       =   &H80000018&
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         ForeColor       =   &H00808080&
         Height          =   240
         Index           =   23
         Left            =   4950
         TabIndex        =   70
         Text            =   "IIBB Importe"
         Top             =   945
         Width           =   1470
      End
      Begin VB.TextBox txtTitulo 
         Appearance      =   0  'Flat
         BackColor       =   &H80000018&
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         ForeColor       =   &H00808080&
         Height          =   240
         Index           =   18
         Left            =   4965
         TabIndex        =   56
         Text            =   "Iva Rg 3337 Importe"
         Top             =   585
         Width           =   1470
      End
      Begin VB.TextBox txtRg3337Imp 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000018&
         ForeColor       =   &H00808080&
         Height          =   330
         Left            =   6585
         Locked          =   -1  'True
         TabIndex        =   23
         Top             =   540
         Width           =   1275
      End
      Begin VB.TextBox txtTitulo 
         Appearance      =   0  'Flat
         BackColor       =   &H80000018&
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         ForeColor       =   &H00808080&
         Height          =   240
         Index           =   7
         Left            =   2400
         TabIndex        =   55
         Text            =   "Iva Rg 3337 %"
         Top             =   585
         Width           =   1080
      End
      Begin VB.TextBox txtRg3337Pje 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000018&
         ForeColor       =   &H00808080&
         Height          =   330
         Left            =   3705
         TabIndex        =   20
         Top             =   540
         Width           =   780
      End
      Begin VB.Line Line1 
         BorderColor     =   &H80000000&
         BorderWidth     =   2
         Index           =   3
         X1              =   4710
         X2              =   4725
         Y1              =   180
         Y2              =   1305
      End
      Begin VB.Line Line1 
         BorderColor     =   &H80000000&
         BorderWidth     =   2
         Index           =   2
         X1              =   8235
         X2              =   8235
         Y1              =   90
         Y2              =   1305
      End
      Begin VB.Line Line1 
         BorderColor     =   &H80000000&
         BorderWidth     =   2
         Index           =   1
         X1              =   4725
         X2              =   4725
         Y1              =   -225
         Y2              =   1260
      End
      Begin VB.Line Line1 
         BorderColor     =   &H80000000&
         BorderWidth     =   2
         Index           =   0
         X1              =   2115
         X2              =   2130
         Y1              =   90
         Y2              =   1275
      End
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Detalle"
      ForeColor       =   &H00808080&
      Height          =   285
      Left            =   135
      TabIndex        =   46
      Top             =   8340
      Width           =   10410
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Concepto"
      ForeColor       =   &H00808080&
      Height          =   285
      Left            =   135
      TabIndex        =   45
      Top             =   7215
      Width           =   10410
   End
End
Attribute VB_Name = "frmDebitoCreditoOIL"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim strNuevoNumero As String

Private Function Muestro_Origen()
  
  Dim rs As ADODB.Recordset
  
  'lleno con detalle de comprobante
  strSQL = "select * from oil_nd_nc_detalle_vw where factura = '" & txtOrigen & "'"
  Set rs = adoGetRS(strSQL)
    
  'chequeo errores
  If Not lngAdoErrNum = -1 Then
    adoError
    Exit Function
  End If
    
  'asocio vista a grilla cond etalle
  Set Me.spdDet.DataSource = rs
    
  'set filas maximas en grilla
  Me.spdDet.MaxRows = rs.RecordCount
    
  'aculto columnas
  spdDet.Col = 2
  spdDet.ColHidden = True
  
  spdDet.Col = 9
  spdDet.ColHidden = True
  
  spdDet.Col = 10
  spdDet.ColHidden = True
  
  For intI = 12 To 22
    spdDet.Col = intI
    spdDet.ColHidden = True
  Next
  
  'set todas las filas
  Me.spdDet.Row = -1
  
  'set decimales volumen 15
  Me.spdDet.Col = 3
  Me.spdDet.TypeCurrencyDecPlaces = 3
  
  'set decimales volumen 1556
  Me.spdDet.Col = 4
  Me.spdDet.TypeCurrencyDecPlaces = 3
  
  'set decimales barriles
  Me.spdDet.Col = 5
  Me.spdDet.TypeCurrencyDecPlaces = 3
  
  'set decimales precio
  Me.spdDet.Col = 6
  Me.spdDet.TypeNumberDecPlaces = 3
    
  'bloqueo item
  Me.spdDet.Col = 1
  Me.spdDet.Lock = True
  
  'bloqueo volumen 1556
  Me.spdDet.Col = 4
  Me.spdDet.Lock = True
    
  'bloqueo operacion en
  Me.spdDet.Col = 7
  Me.spdDet.Lock = True
    
  'bloqueo importe
  Me.spdDet.Col = 8
  Me.spdDet.Lock = True
    
  'bloqueo tipo de dato
  Me.spdDet.Col = 11
  Me.spdDet.Lock = True
    
  'pongo limites a volumenes
  For intI = 1 To Me.spdDet.MaxRows
        
    'set fila
    Me.spdDet.Row = intI
    
    'set col y maximo
    Me.spdDet.Col = 3
    Me.spdDet.TypeCurrencyMax = Me.spdDet.Value
    
    'set col y maximo
    Me.spdDet.Col = 5
    Me.spdDet.TypeCurrencyMax = Me.spdDet.Value
    
  Next
    
End Function

Private Function tomoNumeracionNueva()
  
    Dim rs As ADODB.Recordset

    'valido que se haya seleccionado tipo Comprobante y Empresa
    If Me.cboEmpresa.ListIndex = -1 Or Me.cboComprobante.ListIndex = -1 Or Me.cboOperacion.ListIndex = -1 Then
        Exit Function
    End If
    
    ' tomo valores de empresas para numero de factura automatico
    strSQL = "SELECT * FROM empresas_Puntos_Venta " & _
             "WHERE IDempresa = " & Me.cboEmpresa.ItemData(Me.cboEmpresa.ListIndex) & _
             " AND " & _
             "Operacion = '" & UCase(Left(Me.cboOperacion, 3)) & "'" & _
             " AND " & _
             "Comprobante = '" & UCase(Left(Me.cboComprobante, 3)) & "'"
             
  
    Set rs = adoGetRS(strSQL)
  
    'chequeo errores
    If Not lngAdoErrNum = -1 Then
      adoError
      Exit Function
    End If
    
    
    'CHECK si encontro algo
    If Not rs.EOF Then
    
        '   20/07/2015  -   Emazzu
        '                   Agregue Tipo Comprobante y Operaciòn al nùmero de factura
        '
        Me.txtFactura = rs!Letra & Format(rs!Punto, "0000") & "-" & _
                        Format(rs!Numero + 1, "00000000") & "-" & _
                        UCase(Left(Me.cboOperacion, 3)) & "-" & UCase(Left(Me.cboComprobante, 3))
     

      ' 20/07/2015 - Emazzu - Esto no se usa mas, pero lo dejamos
      
      ' guardo nuevo numero automatico, esto lo hago por que si despues
      ' me lo modifican a mano, e ingresan un numero valido, pero que
      ' no es el ultimo disponible, no lo tengo que actualizar
      strNuevoNumero = Me.txtFactura
  
    End If
    
    rs.Close

End Function
  
'   11/06/2015
'   Edu Mazzu   -   Llena lista desplegable con provincias de entrega
'
Private Function getProvinciasEntrega()

  strSQL = "SELECT * FROM Clientes_Impuesto_Provincia_vw WHERE IDcliente = " & _
           Me.cboCliente.ItemData(Me.cboCliente.ListIndex)
    
    Set rs = adoGetRS(strSQL)
    
    'chequeo errores
    If Not lngAdoErrNum = -1 Then
      adoError
      Exit Function
    End If
        
      'CLEAR lista desplegable
      Me.cboProvincia.Clear
    
      ' recorro recordset y lleno Combo Box
      
      While Not rs.EOF
    
          Me.cboProvincia.AddItem rs!Provincia
          Me.cboProvincia.ItemData(Me.cboProvincia.NewIndex) = rs!IDprovincia
        
        rs.MoveNext
        
      Wend
    
      rs.Close

End Function
  
  

Private Function tomoPorcentajeIVA()
  
  'valido que se haya seleccionado cliente
  If Me.cboCliente.ListIndex >= 0 Then
      
    Dim rs As ADODB.Recordset
    
 'GET informacion cliente seleccionado, con provincia de entrega
  strSQL = "SELECT * FROM ViewVentasClientes WHERE Cliente = '" & _
           Me.cboCliente & "'"
    
    Set rs = adoGetRS(strSQL)
    
    'chequeo errores
    If Not lngAdoErrNum = -1 Then
      adoError
      Exit Function
    End If
    
    'CHECK si encontro algo
    If Not rs.EOF Then
     
      'Determina IVA segun Cliente
      Select Case rs!condIVA
      Case "Responsable Inscripto"
        Me.txtIvaPje = Format(getParam("ivaResponsable"), "#0.00")
      
      Case "Responsable no Inscripto"
        Me.txtIvaPje = Format(getParam("ivaNoInscripto"), "#0.00")
      
      Case "Consumidor Final"
        Me.txtIvaPje = Format(getParam("ivaFinal"), "#0.00")
      
      Case "Exento"
        Me.txtIvaPje = Format(getParam("ivaExento"), "#0.00")
      
      Case "Exportación"
        Me.txtIvaPje = Format(getParam("ivaExportacion"), "#0.00")
      
      End Select
    
      'determina IVA RG3337
      If rs!ivaRG3337 = "Si" Then
        Me.txtRg3337Pje = Format(getParam("ivaRG3337"), "#0.00")
      Else
        Me.txtRg3337Pje = Format(0, "#0.00")
      End If
    
    End If
    
    'cierro rs
    rs.Close
      
  End If
  
End Function



Private Sub cboCliente_Click()

    intRes = tomoPorcentajeIVA()
    
    '   11/06/2015
    '   Edu Mazzu   -   Llena lista desplegable con provincias de entrega
    '
    intRes = getProvinciasEntrega()

End Sub

Private Sub cboComprobante_Click()

  intRes = tomoNumeracionNueva()
  
End Sub



Private Sub cboConcepto_Click()
  Dim rs As ADODB.Recordset
  
  'traigo concepto segun ID seleccionado
  strSQL = "select concepto from ventasConceptos where IDIdentifica='" & cboConcepto & "'"
  Set rs = adoGetRS(strSQL)

  If Not rs.EOF Then
    Me.txtConceptoVenta = rs!concepto
  End If

End Sub


Private Sub cboCondicion_Click()
  
  If Me.cboCondicion <> "Toma fecha vencimiento" Then
    Me.txtVencimiento = "01/01/1900"
  Else
    Me.txtVencimiento = ""
  End If
  
End Sub

Private Sub cboEmpresa_Click()

  intRes = tomoNumeracionNueva()
  
End Sub




Private Sub cboOperacion_Click()
    
    '   21/07/2015
    '   Emazzu - Deshabilitado
    '
    '
    '  If cboOperacion = "Dcg" Or cboOperacion = "Dco" Or cboOperacion = "Dcv" Then
    '
    '    'fuerzo moneda a $
    '    Me.cboMoneda = "$"
    '    Me.cboMoneda.Enabled = False
    '
    '    'fuerzo condicion a contado
    '    Me.cboCondicion = "Contado"
    '
    '    'fuerzo cotizacion a Ninguno
    '    Me.cboCotizacion = "Ninguno"
    '
    '    'fuerzo fecha de vencimiento a fecha de factura
    '    Me.txtVencimiento = Me.txtFecha
    '
    '  Else
    '    cboMoneda.Enabled = True
    '  End If

 intRes = tomoNumeracionNueva()

  
End Sub


Private Sub cboProvincia_Click()

'   11/06/2015
'   Edu Mazzu   -   GET porcentaje de IIBB y ID provincia
'
  strSQL = "SELECT * FROM Clientes_Impuesto_Provincia_vw WHERE IDprovincia = " & _
           Me.cboProvincia.ItemData(Me.cboProvincia.ListIndex)
    
    Set rs = adoGetRS(strSQL)
    
    'chequeo errores
    If Not lngAdoErrNum = -1 Then
      adoError
      Exit Sub
    End If
        
    '   CHECK si trajo algo
    If Not rs.EOF Then
    
        'GET IIBB Porcentaje
        Me.txtIIBBPje = Format(rs!IIBB, "#0.00")
    
        'SAVE IDprovincia en un textBox oculto y bloqueado, para luego guardarlo
        Me.txt_IDprovincia_Ent = rs!IDprovincia
        
    End If

End Sub

Private Sub cmdCalcular_Click()
  Dim rs As ADODB.Recordset
  
  ' validaciones
  If Not DataValidate(cboEmpresa, , True) Then Exit Sub
  If Not DataValidate(cboOperacion, , True) Then Exit Sub
  If Not DataValidate(cboComprobante, , True) Then Exit Sub
  If Not DataValidate(cboMoneda, , True) Then Exit Sub
  If Not DataValidate(txtFecha, "dd/mm/yyyy", True) Then Exit Sub
  If Not DataValidate(txtContable, "dd/mm/yyyy", True) Then Exit Sub
  If Not DataValidate(cboCondicion, , True) Then Exit Sub
  If Not DataValidate(txtVencimiento, "dd/mm/yyyy", True) Then Exit Sub
  If Not DataValidate(cboCliente, , True) Then Exit Sub
  If Not DataValidate(cboProvincia, , True) Then Exit Sub
  If Not DataValidate(cboCuentaBancaria, , True) Then Exit Sub
  If Not DataValidate(cboCotizacion, , True) Then Exit Sub
  If Not DataValidate(txtTipoCambio, "##.###", False) Then Exit Sub
  
  'declare
  Dim intSucu, intNumero As Integer
  Dim strTipo As String
  Dim rsEmpComp As ADODB.Recordset
    
  'get numero de comprobante por separado
  strTipo = Left(txtFactura, 1)
  intSucu = Mid(txtFactura, 2, 4)
  intNumero = Mid(txtFactura, 7, 8)
  
    'CHECK si facturaciòn de resguardo = true, valida que este el preimpreso autorizado
    If chk_Resguardo Then
  
            'build Query para buscar fecha de vencimiento de comprobante
            strSQL = "SELECT vencimiento FROM empresasComprobantes " & _
                     "WHERE IDempresa = " & Me.cboEmpresa.ItemData(Me.cboEmpresa.ListIndex) & " and " & _
                        "tipo = '" & strTipo & "' and " & _
                        "sucu = " & intSucu & " and " & _
                        intNumero & " between desde and hasta"
            
            'get vencimiento comprobante
            Set rsEmpComp = adoGetRS(strSQL)
            
            'check si rs vacio
            If rsEmpComp.EOF Then
              
              intRes = MsgBox("El comprobante no se encuentra registrado.", vbCritical + vbApplicationModal, "Atención...")
              rsEmpComp.Close
              Exit Sub
              
            End If
            
            If CDate(txtFecha) > CDate(rsEmpComp!vencimiento) Then
              
              intRes = MsgBox("El comprobante venció el: " & rsEmpComp!vencimiento, vbCritical + vbApplicationModal, "Atención...")
              rsEmpComp.Close
              Exit Sub
              
            End If
  
    End If
  
  'vacio totales de facturacion
  Me.txtSubtotal = Format(0, "########0.00")
  Me.txtIvaImp = Format(0, "########0.00")
  Me.txtRg3337Imp = Format(0, "########0.00")
  Me.txtIIBBImp = Format(0, "########0.00")
  Me.txtTotal = Format(0, "########0.00")
  
  'calculo subtotal auxiliar para calcularle el iva, lo hago
  'en forma separada porque algun item puede no llevar iva
  Dim curSubtotalParaIva As Currency
  Dim intFila As Integer
  Dim varConIva, varSubtotal As Variant
    
  curSubtotalParaIva = 0
  
  'recorro detalle, sumo los items que llevan iva
  For intFila = 1 To Me.spdDet.MaxRows
        
    'get conIva
    Me.spdDet.GetText 10, intFila, varConIva
    
    'get subtotal
    Me.spdDet.GetText 8, intFila, varSubtotal
    
    'sumo subtotal
    Me.txtSubtotal = Format(CSng(Me.txtSubtotal) + varSubtotal, "##########0.00")
    
    If varConIva = 1 Then
      curSubtotalParaIva = curSubtotalParaIva + varSubtotal
    End If
  
  Next
  
  If Val(Me.txtSubtotal) <> 0 Then
    Me.txtIvaImp = Format(curSubtotalParaIva * Val(Me.txtIvaPje) / 100, "##########0.00")
    Me.txtRg3337Imp = Format(curSubtotalParaIva * Val(Me.txtRg3337Pje) / 100, "##########0.00")
    Me.txtIIBBImp = Format(curSubtotalParaIva * Val(Me.txtIIBBPje) / 100, "##########0.00")
  End If
  
  'calulo subtotal
  
  'calculo total
  Me.txtTotal = Format(CCur(Me.txtSubtotal) + CCur(Me.txtIvaImp) + CCur(Me.txtRg3337Imp) + CCur(Me.txtIIBBImp), "##########0.00")
  
  ' abilito Guardar
  Me.cmdGuardar.Enabled = True

End Sub

Private Sub cmdCondicion_Click()
  Dim strAux As String
        
  ' carga formulario
  Load frmAddCondiciones
    
  ' muestra formulario
  frmAddCondiciones.Show vbModal
    
  ' si hace clicn en aceptar
  If blnAceptar Then
    
    With frmAddCondiciones
    
    strSQL = "EXEC spCondicionesInsert '" & .txtIdentificacion & "','" & .txtDetalle & "'"
    intRes = adoExecSQL(strSQL)
    
    End With
    
    ' refresh ComboBox
    strSQL = "SELECT * FROM ViewCondiciones"
    intRes = ComboBoxRefresh(cboCondicion, strSQL)
    
    ' hubico listindex en elemento agregado
    cboCondicion.ListIndex = ComboBoxFindItem(cboCondicion, frmAddCondiciones.txtIdentificacion)
    
    ' descarga formulario
    Unload frmAddCondiciones
      
  End If

End Sub

Private Sub cmdCotiza_Click()

  'muestro frm
  frmCotizacion.Show vbModal

End Sub

Private Sub cmdFormaPago_Click()
  Dim strAux As String
        
  ' carga formulario
  Load frmAddCuentasBancarias
    
  ' muestra formulario
  frmAddCuentasBancarias.Show vbModal
    
  ' si hace clicn en aceptar
    
  If blnAceptar Then
    
    With frmAddCuentasBancarias
    
    strSQL = "EXEC spCuentasBancariasInsert '" & .txtIdentificacion & "','" & .txtDetalle & "'"
    intRes = adoExecSQL(strSQL)
    
    End With
    
    ' refresh ComboBox
  
    strSQL = "SELECT * FROM ViewCuentasBancarias"
    intRes = ComboBoxRefresh(cboCuentaBancaria, strSQL)
    
    ' hubico listindex en elemento agregado
    
    cboCuentaBancaria.ListIndex = ComboBoxFindItem(cboCuentaBancaria, frmAddCuentasBancarias.txtIdentificacion)
    
    ' descarga formulario
      
    Unload frmAddCuentasBancarias
      
  End If

End Sub

Private Sub cmdGuardar_Click()
  Dim intUltimaVenta, intItem  As Integer
  Dim rs As ADODB.Recordset
  
  'chequeo fecha factura que no sea menor al ultimo periodo cerrado
  Set rs = adoGetRS("select max(fecha) as fechaCierre From stockCierre where proceso = 'cerr' and status = 2")
  If Not rs.EOF And Not IsNull(rs!fechaCierre) Then
    If Format(Me.txtFecha, "dd/mm/yyyy") <= rs!fechaCierre Then
      intRes = MsgBox("La fecha del comprobante corresponde a un período cerrado.", vbApplicationModal + vbInformation + vbOKOnly, "informacion...")
      Exit Sub
    End If
  End If
  rs.Close
  
  
'   08/05/2015
'   Edu Mazzu - Deshabilitado
'   Como a partir de ahora, no se puede cambiar el nùmero de factura a mano, esta validaciòn, ya no se aplica.
'
  ' chequeo que el numero de factura ingresado este libre
'  Set rs = adoGetRS("SELECT factura FROM ViewVentas where Factura = '" & Me.txtFactura & "'")
'  If Not rs.EOF And Not IsNull(rs!factura) Then
'    intRes = MsgBox("El numero de comprobante ingresado ya fue utilizado, debe cambiarlo para poder guardar el comprobante.", vbApplicationModal + vbInformation + vbOKOnly, "informacion...")
'    rs.Close
'    Exit Sub
'  End If
'  rs.Close
'  End If
  
  
    
  'get IDentregaCli, antes de grabar encabezado
  Dim varIDentregaCli, varIDterminal As Variant
  Me.spdDet.GetText 14, 1, varIDentregaCli
  Me.spdDet.GetText 13, 1, varIDterminal
  
  ' guardo ventas
  strSQL = "EXEC spVentasInsert " & _
           "'" & Me.txtFactura & "'," & _
           "'" & dateToIso(Me.txtFecha) & "','" & dateToIso(Me.txtContable) & "'," & _
           Val(Me.cboEmpresa.ItemData(Me.cboEmpresa.ListIndex)) & "," & _
           "'" & Me.cboOperacion.List(Me.cboOperacion.ListIndex) & "'," & _
           "'" & Left(Me.cboComprobante.List(Me.cboComprobante.ListIndex), 3) & "'," & _
           Val(Me.cboMoneda.ItemData(Me.cboMoneda.ListIndex)) & "," & _
           "'" & dateToIso(Me.txtVencimiento) & "'," & _
           Val(Me.cboCliente.ItemData(Me.cboCliente.ListIndex)) & "," & Val(Me.txt_IDprovincia_Ent) & "," & _
           varIDentregaCli & "," & _
           Val(Me.cboCuentaBancaria.ItemData(Me.cboCuentaBancaria.ListIndex)) & "," & _
           Val(Me.cboCondicion.ItemData(Me.cboCondicion.ListIndex)) & ","
  strSQL = strSQL & _
           varIDterminal & "," & _
           "''" & ",'" & _
           txtOrigen & "'," & _
           Val(Me.txtSubtotal) & "," & _
           Val(Me.txtIvaPje) & "," & _
           Val(Me.txtIvaImp) & "," & _
           Val(Me.txtRg3337Pje) & "," & _
           Val(Me.txtRg3337Imp) & "," & _
           Val(Me.txtIIBBPje) & "," & _
           Val(Me.txtIIBBImp) & "," & _
           Val(Me.txtTotal) & "," & _
           Val(Me.txtTipoCambio) & "," & _
           Val(Me.cboCotizacion.ItemData(Me.cboCotizacion.ListIndex)) & "," & _
           "'" & Me.txtConceptoVenta & "'," & _
           "'','" & Replace(Me.txt_Libre, "'", Chr(34)) & "'"

  
  intResul = adoExecSQL(strSQL)
  
  'chequeo errores
  If Not lngAdoErrNum = -1 Then
    adoError
    Exit Sub
  End If
    
  ' guardo VentasDetalle
  For intCuenta = 1 To Me.spdDet.MaxRows
                 
    strSQL = "EXEC spVentasDetalleInsert "
    strSQL = strSQL & "'" & Me.txtFactura & "',"

    'set fila actual
    Me.spdDet.Row = intCuenta
    
    Me.spdDet.Col = 1   'item
    strSQL = strSQL & Me.spdDet.Text & ","
                 
    Me.spdDet.Col = 15  'IDcontrato
    strSQL = strSQL & Me.spdDet.Text & ","
                     
    Me.spdDet.Col = 22  'FechaEntregaCli
    strSQL = strSQL & "'" & dateToIso(Me.spdDet.Text) & "',"
                 
    Me.spdDet.Col = 3  'cantidad
    strSQL = strSQL & Me.spdDet.Value & ","
    
    Me.spdDet.Col = 4  'cantidadInfo
    strSQL = strSQL & Me.spdDet.Value & ","
    
    Me.spdDet.Col = 5  'cantidadInfo1
    strSQL = strSQL & Me.spdDet.Value & ","
                 
    Me.spdDet.Col = 6  'precio
    strSQL = strSQL & Me.spdDet.Value & ","
                 
    Me.spdDet.Col = 8  'importe
    strSQL = strSQL & Me.spdDet.Value & ","
                 
    Me.spdDet.Col = 2  'concepto
    strSQL = strSQL & "'" & Me.spdDet.Text & "',"
                 
    Me.spdDet.Col = 10  'conIva
    strSQL = strSQL & Me.spdDet.Text & ","
                 
    Me.spdDet.Col = 11  'ventasTipo
    strSQL = strSQL & "'" & Me.spdDet.Text & "',"
                 
    Me.spdDet.Col = 9  'IDunidad
    strSQL = strSQL & Me.spdDet.Text
                 
    intResul = adoExecSQL(strSQL)
    
    'chequeo errores
    If Not lngAdoErrNum = -1 Then
      adoError
      Exit Sub
    End If
    
  Next
  
  
    'UPDATE numero de comprobante
    strSQL = "EXEC Empresas_Puntos_Venta_sp " & Me.cboEmpresa.ItemData(cboEmpresa.ListIndex) & "," & _
                                                "'" & UCase(Left(Me.cboOperacion, 3)) & "'," & _
                                                "'" & UCase(Left(Me.cboComprobante, 3)) & "'," & _
                                                Val(Mid(Me.txtFactura, 7, 8))
    
    intResul = adoExecSQL(strSQL)
    
    'chequeo errores
    If Not lngAdoErrNum = -1 Then
        adoError
        Exit Sub
    End If
    
    
    
  
  'oculto frm
  blnAceptar = True
  blnCancelar = False
  Me.Hide
  
End Sub

Private Sub cmdNewEmpresa_Click()

  Dim strStore, strView, strDato As String

  strStore = "spEmpresasInsert"
  strView = "SELECT * FROM ViewEmpresas"
  strDato = ComboBoxAddItem(Me, cboEmpresa, "@50", strStore, strView)

End Sub

Private Sub cmdNewCotizacionTexto_Click()
  Dim strAux As String
        
  ' carga formulario
  Load frmAddCotizacionesTexto
    
  ' muestra formulario
  frmAddCotizacionesTexto.Show vbModal
    
  ' si hace clicn en aceptar
  If blnAceptar Then
    
    With frmAddCotizacionesTexto
    
    strSQL = "EXEC spCotizacionesTextoInsert '" & .txtIdentificacion & "','" & .txtDetalle & "'," & .chkCotizacion
    intRes = adoExecSQL(strSQL)
    
    End With
    
    ' refresh ComboBox
    strSQL = "SELECT * FROM cotizacionesTexto_vw"
    intRes = ComboBoxRefresh(cboCotizacion, strSQL)
    
    ' hubico listindex en elemento agregado
    cboCotizacion.ListIndex = ComboBoxFindItem(cboCotizacion, frmAddCotizacionesTexto.txtIdentificacion)
    
    ' descarga formulario
    Unload frmAddCotizacionesTexto
      
  End If


End Sub

Private Sub cmdSalir_Click()
  
  blnAceptar = False
  blnCancelar = True
  Unload Me

End Sub

Private Sub cmdVer_Click()
      
  'chequeo que se haya seleccionado alguna empresa o cliente
  If Me.cboEmpresa.ListIndex = -1 Or Me.cboCliente.ListIndex = -1 Then
    intRes = MsgBox("Debe seleccionar Empresa y Cliente.", vbCritical + vbOKOnly, "atención")
    Exit Sub
  End If
      
  'muestro frm
  oil_NC_ND_origenFRM.Show vbModal
    
  'muestro detalle de comprobante origen
  If blnAceptar Then
    blnB = Muestro_Origen()
  End If
  
End Sub

    

Private Sub Form_Load()
  
  Dim strAux As String
  Dim intI As Integer
  Dim rs As ADODB.Recordset
  
  ' lleno combos
  strSQL = "SELECT * FROM ViewEmpresas"
  intRes = ComboBoxRefresh(cboEmpresa, strSQL)
  
  'default vintage
'  Me.cboEmpresa = "Vintage Oil Argentina Inc."
  
  strSQL = "SELECT * FROM ViewVentasClientes"
  intRes = ComboBoxRefresh(cboCliente, strSQL)

  strSQL = "SELECT * FROM ViewMonedas"
  intRes = ComboBoxRefresh(cboMoneda, strSQL)

  strSQL = "SELECT * FROM ViewCuentasBancarias"
  intRes = ComboBoxRefresh(cboCuentaBancaria, strSQL)

  strSQL = "SELECT * FROM ViewCondiciones"
  intRes = ComboBoxRefresh(cboCondicion, strSQL)

  strSQL = "SELECT * FROM ventasConceptos_vw"
  intRes = ComboBoxRefresh(cboConcepto, strSQL)

  strSQL = "SELECT * FROM cotizacionesTexto_vw"
  intRes = ComboBoxRefresh(cboCotizacion, strSQL)

  'toma parametro de cotizacion
  txtCotizaActual = Format(getParam("cotizU$S"), "#0.000")
  
  'guardo nombre tabla actual para ini
  strAux = strTableNameActual
  strTableNameActual = "viewVentasDetalle"
      
  'cambio apariencia
  Me.spdDet.EditModeReplace = True
        
  Me.spdDet.GrayAreaBackColor = Me.spdDet.BackColor
        
  'muestro origen
  blnB = Muestro_Origen()
  
  ' recupero nombre tabla actual para ini
  strTableNameActual = strAux
    
    'DEFECTO - OIL - Unico valor
    Me.cboOperacion.ListIndex = 0
    
    
    '   POR DEFECTO fecha del dìa para facturar
    Me.txtFecha = DateValue(Now)


End Sub

Private Sub spdDet_Change(ByVal Col As Long, ByVal Row As Long)
      
  Dim varM3, varM31556, varBarriles, varPrecio, varUnidad As Variant
      
  'si modifica m3
  If Col = 3 Then
    
    'get m3 modificado
    Me.spdDet.GetText Col, Row, varM3
    
    'calculo m3 1556
    Me.spdDet.SetText 4, Row, Round(varM3 * CSng(getParam("m315TOm31556")), 3)
      
    'tomo m3 1556
    Me.spdDet.GetText 4, Row, varM31556
       
    'calculo barriles
    Me.spdDet.SetText 5, Row, Round(varM31556 * CSng(getParam("m31556TObarr1556")), 3)
        
  End If
  
  'si modifica barriles
  If Col = 5 Then
    
    'get barriles modificado
    Me.spdDet.GetText Col, Row, varBarriles
    
    'calculo m3 1556
    Me.spdDet.SetText 4, Row, Round(varBarriles / CSng(getParam("m31556TObarr1556")), 3)
      
    'tomo m3 1556
    Me.spdDet.GetText 4, Row, varM31556
       
    'calculo m3
    Me.spdDet.SetText 3, Row, Round(varM31556 / CSng(getParam("m315TOm31556")), 3)
    
  End If
  
  'get barriles modificado
  Me.spdDet.GetText 5, Row, varBarriles
  
  'get precio
  Me.spdDet.GetText 6, Row, varPrecio
  
  'get unidad
  Me.spdDet.GetText 7, Row, varUnidad
  
  'calculo subtotal
  Me.spdDet.SetText 8, Row, Round(varBarriles * varPrecio, 2)
    
End Sub


Private Sub spdDet_ComboCloseUp(ByVal Col As Long, ByVal Row As Long, ByVal SelChange As Integer)
  
  'check si DBclick en columna tipo de venta
  If Col = 11 Then
        
    'set puntero grilla
    Me.spdDet.Col = Col
    Me.spdDet.Row = Row
    
    'change tipo de celda
    Me.spdDet.CellType = CellTypeEdit
    
    'set lock true
    Me.spdDet.Lock = True
    
  End If
  
End Sub

Private Sub spdDet_DblClick(ByVal Col As Long, ByVal Row As Long)
  
  Dim rs, rsT As ADODB.Recordset
  Dim strT, strTant As String
  
  'check si DBclick en columna tipo de venta
  If Col = 11 Then
      
    'set puntero grilla
    Me.spdDet.Col = Col
    Me.spdDet.Row = Row
        
    'save tipo de venta anterior
    strTant = Me.spdDet.Text
      
    'get tipos de venta
    strSQL = "select * from VentasTipo"
    Set rsT = adoGetRS(strSQL)
    
    'check errores
    If Not lngAdoErrNum = -1 Then
      adoError
      Exit Sub
    End If
    
    'inicializo
    strT = ""
    
    'while rs
    While Not rsT.EOF
          
      'add id y nombre
      strT = strT & rsT!ventasTipoCorto & Chr(9)
      
      'next puntero
      rsT.MoveNext
      
    Wend
    
    'change tipo de celda
    Me.spdDet.CellType = CellTypeComboBox
    
    'add elementos al combo
    Me.spdDet.TypeComboBoxList = strT
    
    'set tipo de venta anterior
    Me.spdDet.Text = strTant
    
    'set lock = false
    Me.spdDet.Lock = False
    
  End If
  
End Sub

