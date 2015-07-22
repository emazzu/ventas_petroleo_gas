VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmDebitoCredito 
   BackColor       =   &H80000018&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Facturas - Débito - Crédito (Varios)"
   ClientHeight    =   10185
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11940
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   10185
   ScaleWidth      =   11940
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtTitulo 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000018&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      ForeColor       =   &H00808080&
      Height          =   285
      Index           =   24
      Left            =   135
      TabIndex        =   72
      Text            =   "Texto Libre"
      Top             =   4605
      Width           =   11640
   End
   Begin VB.TextBox txt_Libre 
      BackColor       =   &H80000018&
      Height          =   1635
      Left            =   135
      MultiLine       =   -1  'True
      TabIndex        =   24
      Top             =   4920
      Width           =   11670
   End
   Begin VB.Frame Frame5 
      BackColor       =   &H80000018&
      Height          =   3180
      Left            =   5535
      TabIndex        =   71
      Top             =   90
      Width           =   1320
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
      TabIndex        =   42
      Text            =   "Comprobante Nro."
      Top             =   720
      Width           =   1350
   End
   Begin VB.ComboBox cboCliente 
      Appearance      =   0  'Flat
      BackColor       =   &H80000018&
      ForeColor       =   &H00808080&
      Height          =   315
      ItemData        =   "frmDebitoCredito.frx":0000
      Left            =   1710
      List            =   "frmDebitoCredito.frx":000A
      Sorted          =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   5
      Top             =   1770
      Width           =   3135
   End
   Begin VB.ComboBox cboMoneda 
      Appearance      =   0  'Flat
      BackColor       =   &H80000018&
      ForeColor       =   &H00808080&
      Height          =   315
      ItemData        =   "frmDebitoCredito.frx":001F
      Left            =   1680
      List            =   "frmDebitoCredito.frx":0029
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
      ItemData        =   "frmDebitoCredito.frx":003E
      Left            =   1710
      List            =   "frmDebitoCredito.frx":0048
      Style           =   2  'Dropdown List
      TabIndex        =   4
      Top             =   1410
      Width           =   3675
   End
   Begin MSComctlLib.ListView lvwDetalle 
      Height          =   1455
      Left            =   135
      TabIndex        =   39
      Top             =   8655
      Width           =   11700
      _ExtentX        =   20638
      _ExtentY        =   2566
      View            =   3
      LabelWrap       =   0   'False
      HideSelection   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   0
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
      TabIndex        =   38
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
      TabIndex        =   25
      Top             =   7530
      Width           =   11670
   End
   Begin VB.TextBox txtIvaImp 
      Alignment       =   1  'Right Justify
      BackColor       =   &H80000018&
      ForeColor       =   &H00808080&
      Height          =   330
      Left            =   7710
      Locked          =   -1  'True
      TabIndex        =   20
      Top             =   3435
      Width           =   1140
   End
   Begin VB.TextBox txtIvaPje 
      Alignment       =   1  'Right Justify
      BackColor       =   &H80000018&
      ForeColor       =   &H00808080&
      Height          =   330
      Left            =   3930
      TabIndex        =   19
      Top             =   3435
      Width           =   1095
   End
   Begin VB.TextBox txtSubtotal 
      Alignment       =   1  'Right Justify
      BackColor       =   &H80000018&
      ForeColor       =   &H00808080&
      Height          =   330
      Left            =   435
      Locked          =   -1  'True
      TabIndex        =   18
      Top             =   4155
      Width           =   1635
   End
   Begin VB.TextBox txtTotal 
      Alignment       =   1  'Right Justify
      BackColor       =   &H80000018&
      ForeColor       =   &H00808080&
      Height          =   330
      Left            =   9855
      Locked          =   -1  'True
      TabIndex        =   23
      Top             =   4155
      Width           =   1815
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
      TabIndex        =   28
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
      TabIndex        =   27
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
      TabIndex        =   26
      Text            =   "Operación Tipo"
      Top             =   2550
      Width           =   1170
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H80000018&
      Height          =   3180
      Left            =   135
      TabIndex        =   29
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
         TabIndex        =   75
         Text            =   "Provincia"
         Top             =   2100
         Width           =   950
      End
      Begin VB.ComboBox cboProvincia 
         Appearance      =   0  'Flat
         BackColor       =   &H80000018&
         ForeColor       =   &H00808080&
         Height          =   315
         ItemData        =   "frmDebitoCredito.frx":005D
         Left            =   1560
         List            =   "frmDebitoCredito.frx":0073
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   6
         Top             =   2040
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
         Height          =   315
         Left            =   1560
         Locked          =   -1  'True
         TabIndex        =   2
         Top             =   960
         Width           =   1155
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
         ItemData        =   "frmDebitoCredito.frx":0098
         Left            =   1560
         List            =   "frmDebitoCredito.frx":00A8
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   7
         Top             =   2400
         Width           =   3675
      End
      Begin VB.CommandButton cmdNewCliente 
         Caption         =   "New"
         Height          =   280
         Left            =   4725
         TabIndex        =   44
         Top             =   1695
         Width           =   500
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
         TabIndex        =   41
         Text            =   "Comprobante Tipo"
         Top             =   270
         Width           =   1395
      End
      Begin VB.ComboBox cboComprobante 
         Appearance      =   0  'Flat
         BackColor       =   &H80000018&
         ForeColor       =   &H00808080&
         Height          =   315
         ItemData        =   "frmDebitoCredito.frx":00CE
         Left            =   1575
         List            =   "frmDebitoCredito.frx":00DB
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
      Left            =   8550
      TabIndex        =   30
      Top             =   6765
      Width           =   1500
   End
   Begin VB.CommandButton cmdSalir 
      Caption         =   "&Salir"
      Height          =   330
      Left            =   10170
      TabIndex        =   31
      Top             =   6765
      Width           =   1500
   End
   Begin VB.CommandButton cmdCalcular 
      Caption         =   "&Calcular"
      Height          =   330
      Left            =   6930
      TabIndex        =   32
      Top             =   6765
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
      Left            =   11250
      TabIndex        =   33
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
      Left            =   6090
      TabIndex        =   34
      Text            =   "Iva Importe"
      Top             =   3480
      Width           =   975
   End
   Begin VB.TextBox txtTitulo 
      Appearance      =   0  'Flat
      BackColor       =   &H80000018&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      ForeColor       =   &H00808080&
      Height          =   240
      Index           =   9
      Left            =   1440
      TabIndex        =   35
      Text            =   "Subtotal"
      Top             =   3900
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
      Left            =   2700
      TabIndex        =   36
      Text            =   "Iva %"
      Top             =   3480
      Width           =   480
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H80000018&
      Height          =   3180
      Left            =   6870
      TabIndex        =   37
      Top             =   90
      Width           =   4980
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
      Begin VB.TextBox txtComprobante 
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
         ItemData        =   "frmDebitoCredito.frx":00F9
         Left            =   1170
         List            =   "frmDebitoCredito.frx":0103
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
         ItemData        =   "frmDebitoCredito.frx":0114
         Left            =   1170
         List            =   "frmDebitoCredito.frx":011E
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
         ItemData        =   "frmDebitoCredito.frx":0133
         Left            =   1170
         List            =   "frmDebitoCredito.frx":013D
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
         ItemData        =   "frmDebitoCredito.frx":0152
         Left            =   1170
         List            =   "frmDebitoCredito.frx":015C
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
      TabIndex        =   40
      Top             =   6585
      Width           =   11670
      Begin VB.CheckBox chk_Resguardo 
         BackColor       =   &H80000018&
         Caption         =   "Emisiòn de Resguardo"
         Height          =   285
         Left            =   2295
         TabIndex        =   73
         Top             =   225
         Width           =   2085
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H80000018&
      Height          =   1320
      Left            =   135
      TabIndex        =   43
      Top             =   3255
      Width           =   11670
      Begin VB.TextBox txt_IDprovincia_Ent 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000018&
         ForeColor       =   &H00808080&
         Height          =   285
         Left            =   270
         Locked          =   -1  'True
         TabIndex        =   74
         Top             =   270
         Visible         =   0   'False
         Width           =   315
      End
      Begin VB.TextBox txtIIBBimp 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000018&
         ForeColor       =   &H00808080&
         Height          =   330
         Left            =   7560
         Locked          =   -1  'True
         TabIndex        =   70
         Top             =   900
         Width           =   1140
      End
      Begin VB.TextBox txtIIBBPje 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000018&
         ForeColor       =   &H00808080&
         Height          =   330
         Left            =   3780
         TabIndex        =   69
         Top             =   900
         Width           =   1140
      End
      Begin VB.TextBox txtTitulo 
         Appearance      =   0  'Flat
         BackColor       =   &H80000018&
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         ForeColor       =   &H00808080&
         Height          =   240
         Index           =   23
         Left            =   5985
         TabIndex        =   68
         Text            =   "IIBB Importe"
         Top             =   945
         Width           =   1080
      End
      Begin VB.TextBox txtTitulo 
         Appearance      =   0  'Flat
         BackColor       =   &H80000018&
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         ForeColor       =   &H00808080&
         Height          =   240
         Index           =   22
         Left            =   2565
         TabIndex        =   67
         Text            =   "IIBB (%)"
         Top             =   945
         Width           =   975
      End
      Begin VB.TextBox txtTitulo 
         Appearance      =   0  'Flat
         BackColor       =   &H80000018&
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         ForeColor       =   &H00808080&
         Height          =   240
         Index           =   18
         Left            =   5955
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
         Left            =   7575
         Locked          =   -1  'True
         TabIndex        =   22
         Top             =   540
         Width           =   1140
      End
      Begin VB.TextBox txtTitulo 
         Appearance      =   0  'Flat
         BackColor       =   &H80000018&
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         ForeColor       =   &H00808080&
         Height          =   240
         Index           =   7
         Left            =   2580
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
         Left            =   3795
         TabIndex        =   21
         Top             =   540
         Width           =   1110
      End
      Begin VB.Line Line4 
         BorderColor     =   &H80000000&
         BorderWidth     =   2
         X1              =   9360
         X2              =   9360
         Y1              =   90
         Y2              =   1260
      End
      Begin VB.Line Line2 
         BorderColor     =   &H80000000&
         BorderWidth     =   2
         X1              =   5490
         X2              =   5490
         Y1              =   90
         Y2              =   1305
      End
      Begin VB.Line Line1 
         BorderColor     =   &H80000000&
         BorderWidth     =   2
         X1              =   2130
         X2              =   2115
         Y1              =   105
         Y2              =   1305
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
      Top             =   8385
      Width           =   11670
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
      Top             =   7260
      Width           =   11670
   End
   Begin VB.Menu mnuDebitoCredito 
      Caption         =   "DebitoCredito"
      Visible         =   0   'False
      Begin VB.Menu mnuAgregar 
         Caption         =   "Agregar"
      End
      Begin VB.Menu mnuModificar 
         Caption         =   "Modificar"
      End
   End
End
Attribute VB_Name = "frmDebitoCredito"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim strNuevoNumero As String

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
             "Operacion = '" & UCase(Right(Me.cboOperacion, 3)) & "'" & _
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
                        UCase(Right(Me.cboOperacion, 3)) & "-" & UCase(Left(Me.cboComprobante, 3))
     

      ' 20/07/2015 - Emazzu - Esto no se usa mas, pero lo dejamos
      
      ' guardo nuevo numero automatico, esto lo hago por que si despues
      ' me lo modifican a mano, e ingresan un numero valido, pero que
      ' no es el ultimo disponible, no lo tengo que actualizar
      strNuevoNumero = Me.txtFactura
  
    End If
    
    rs.Close
    

End Function
  
Private Function tomoPorcentajeIVA()
  
    Dim rs As ADODB.Recordset
    
  
  'valido que se haya seleccionado cliente
  If Me.cboCliente.ListIndex < 0 Then
    
    'CLOSE rs
    rs.Close
    
    'EXIT function
    Exit Function
  
  End If
      
      
  'GET informacion cliente seleccionado, con provincia de entrega
  strSQL = "SELECT * FROM ViewVentasClientes WHERE Cliente = '" & Me.cboCliente & "'"
  
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
    
      'GET IVA RG3337
      If rs!ivaRG3337 = "Si" Then
        Me.txtRg3337Pje = Format(getParam("ivaRG3337"), "#0.00")
      Else
        Me.txtRg3337Pje = Format(0, "#0.00")
      End If
    
    End If
  
  
End Function



Private Sub cboCliente_Click()


    intRes = tomoPorcentajeIVA()
    
    '   11/06/2015
    '   Edu Mazzu   -   Llena lista desplegable con provincias de entrega
    '
    intRes = getProvinciasEntrega()


End Sub


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
'   Emazzu  -   Deshabilitado
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
  intNumero = Val(Mid(txtFactura, 7, 8))
  
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
  
  
  
  ' CLEAR totales de facturaciòn
  Me.txtSubtotal = Format(0, "########0.00")
  Me.txtIvaImp = Format(0, "########0.00")
  Me.txtRg3337Imp = Format(0, "########0.00")
  Me.txtIIBBImp = Format(0, "########0.00")
  Me.txtTotal = Format(0, "########0.00")
  
 
  
  'calculo subtotal
  Me.txtSubtotal = Format(objSumColumn(lvwDetalle, "importe"), "##########0.00")
  
  
  'calculo subtotal auxiliar para calcularle el iva, lo hago
  'en forma separada porque algun item puede no llevar iva
  Dim curSubtotalParaIva As Currency
  Dim intCuenta As Integer
  
  
  curSubtotalParaIva = 0
  
  'recorro detalle, so selecciono, si lleva iva sumo
  For intCuenta = 1 To Me.lvwDetalle.ListItems.Count
    Me.lvwDetalle.ListItems(intCuenta).Selected = True
    If lvwGetValue(lvwDetalle, "iva") = "Si" Then
      curSubtotalParaIva = curSubtotalParaIva + Val(lvwGetValue(lvwDetalle, "importe"))
    End If
    Me.lvwDetalle.ListItems(intCuenta).Selected = False
  Next
  
  If Val(Me.txtSubtotal) <> 0 Then
    Me.txtIvaImp = Format(curSubtotalParaIva * Val(Me.txtIvaPje) / 100, "##########0.00")
    Me.txtRg3337Imp = Format(curSubtotalParaIva * Val(Me.txtRg3337Pje) / 100, "##########0.00")
    Me.txtIIBBImp = Format(Val(curSubtotalParaIva) * Val(Me.txtIIBBPje) / 100, "##########0.00")
  End If
  
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
'    Set rs = adoGetRS("SELECT factura FROM ViewVentas where Factura = '" & Me.txtFactura & "'")
'    If Not rs.EOF And Not IsNull(rs!factura) Then
'      intRes = MsgBox("El numero de comprobante ingresado ya fue utilizado, debe cambiarlo para poder guardar el comprobante.", vbApplicationModal + vbInformation + vbOKOnly, "informacion...")
'      rs.Close
'      Exit Sub
'    End If
'    rs.Close
'   End If
  
  
  'guardo ventas
  strSQL = "EXEC spVentasInsert " & _
           "'" & Me.txtFactura & "'," & _
           "'" & dateToIso(Me.txtFecha) & "','" & dateToIso(Me.txtContable) & "'," & _
           Val(Me.cboEmpresa.ItemData(Me.cboEmpresa.ListIndex)) & "," & _
           "'" & Me.cboOperacion.List(Me.cboOperacion.ListIndex) & "'," & _
           "'" & Left(Me.cboComprobante.List(Me.cboComprobante.ListIndex), 3) & "'," & _
           Val(Me.cboMoneda.ItemData(Me.cboMoneda.ListIndex)) & "," & _
           "'" & dateToIso(Me.txtVencimiento) & "'," & _
           Val(Me.cboCliente.ItemData(Me.cboCliente.ListIndex)) & "," & Val(Me.txt_IDprovincia_Ent) & "," & _
           0 & "," & _
           Val(Me.cboCuentaBancaria.ItemData(Me.cboCuentaBancaria.ListIndex)) & "," & _
           Val(Me.cboCondicion.ItemData(Me.cboCondicion.ListIndex)) & "," & _
           0 & "," & _
           "''" & ",'" & _
           Me.txtComprobante & "',"
  strSQL = strSQL & _
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
    
  'guardo VentasDetalle
  intItem = 0
  For intCuenta = 1 To lvwDetalle.ListItems.Count
    
    ' ubico puntero en cada item
    lvwDetalle.ListItems(intCuenta).Selected = True
         
      ' acumulador de numero de item
      intItem = intItem + 1
            
            
      'armo string
      strSQL = "EXEC spVentasDetalleInsert " & _
        "'" & Me.txtFactura & "'," & _
        intItem & "," & _
        0 & "," & _
        "''," & _
        Val(lvwGetValue(lvwDetalle, "cantidad")) & "," & _
        0 & "," & _
        0 & "," & _
        Val(lvwGetValue(lvwDetalle, "precio")) & "," & _
        Val(lvwGetValue(lvwDetalle, "importe")) & "," & _
        "'" & lvwGetValue(lvwDetalle, "concepto") & "'," & _
        IIf(lvwGetValue(lvwDetalle, "iva") = "Si", 1, 0) & "," & _
        "'" & lvwGetValue(lvwDetalle, "idVentasTipo") & "'," & _
        Val(lvwGetValue(lvwDetalle, "IDunidad"))
                
      'exec
      intResul = adoExecSQL(strSQL)
      
      'chequeo errores
      If Not lngAdoErrNum = -1 Then
        adoError
        Exit Sub
      End If
    
  Next
  
  
    'UPDATE numero de comprobante
    strSQL = "EXEC Empresas_Puntos_Venta_sp " & Me.cboEmpresa.ItemData(cboEmpresa.ListIndex) & "," & _
                                                "'" & UCase(Right(Me.cboOperacion, 3)) & "'," & _
                                                "'" & UCase(Left(Me.cboComprobante, 3)) & "'," & _
                                                Val(Mid(Me.txtFactura, 7, 8))
    
    intResul = adoExecSQL(strSQL)
    
    'chequeo errores
    If Not lngAdoErrNum = -1 Then
        adoError
        Exit Sub
    End If
 
  
 
 
  ' oculto frm
  blnAceptar = True
  blnCancelar = False
  Me.Hide

End Sub

Private Sub cmdNewCliente_Click()
  Dim strAux As String
        
  ' cargo formulario
  ' lo muestro
  ' si hago click en aceptar grabo
  '   ejecuto un store procedure
  '   vuelvo a cargar combo
  '   busco elemento agregado
  ' descargo formulario
        
  Load frmAddCliente
  frmAddCliente.Show vbModal
  
  If blnAceptar Then
    
    With frmAddCliente
    strSQL = "EXEC spClientesInsert '" & .txtCliente & "','" & _
            .txtDomicilio & "','" & .txtCodigoPostal & "','" & .txtLocalidad & "','" & _
            .txtPais & "','" & .txtCuit & "','" & .cboCondicionIva.List(.cboCondicionIva.ListIndex) & "','" & _
            .cboRg3337.List(.cboRg3337.ListIndex) & "'," & _
            .txtDiasVentas & ",'" & .cboExportacion.List(.cboExportacion.ListIndex) & "'"
    intRes = adoExecSQL(strSQL)
    End With
    
    strSQL = "SELECT * FROM ViewClientes"
    intRes = ComboBoxRefresh(cboCliente, strSQL)
    
    cboCliente.ListIndex = ComboBoxFindItem(cboCliente, frmAddCliente.txtCliente)
      
    Unload frmAddCliente
      
  End If

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
  If frmDebitoCredito.cboEmpresa.ListIndex = -1 Or frmDebitoCredito.cboCliente.ListIndex = -1 Then
    intRes = MsgBox("Debe seleccionar Empresa y Cliente.", vbCritical + vbOKOnly, "atención")
    Exit Sub
  End If
      
  'muestro frm
  frmComprobantes.Show vbModal
  
End Sub

Private Sub Form_Load()
  Dim strAux As String

  ' lleno combos
  strSQL = "SELECT * FROM ViewEmpresas"
  intRes = ComboBoxRefresh(cboEmpresa, strSQL)
  
  'default vintage
 'Me.cboEmpresa = "Vintage Oil Argentina Inc."
  
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

  ' toma parametro de cotizacion
  txtCotizaActual = Format(getParam("cotizU$S"), "#0.000")

  ' guardo nombre tabla actual para ini
  strAux = strTableNameActual
  strTableNameActual = "viewVentasDetalle"
  
  ' apariencia lvw
  intRes = ListViewAppearanceChange(lvwDetalle)
  
  ' lleno con detalle de comprobante
  strSQL = "select * from ViewVentasDetalle where factura = '" & txtOrigen & "'"
  intRes = ListViewRefresh(lvwDetalle, strSQL)
  intRes = lvwHideColumn(lvwDetalle, "iditem")
  intRes = lvwHideColumn(lvwDetalle, "ContratoID")
  intRes = lvwHideColumn(lvwDetalle, "factura")
  intRes = lvwHideColumn(lvwDetalle, "cantidadinfo")
  intRes = lvwHideColumn(lvwDetalle, "cantidadinfo1")
  intRes = lvwHideColumn(lvwDetalle, "idVentasTipo")

  ' recupero nombre tabla actual para ini
  strTableNameActual = strAux

    '   POR DEFECTO fecha del dìa para facturar
    Me.txtFecha = DateValue(Now)


End Sub

Private Sub lvwDetalle_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

  If Button = 2 Then
    PopupMenu mnuDebitoCredito
  End If

End Sub

Private Sub mnuAgregar_Click()
  Dim item As ListItem

  'validar que se haya ingresado lo minimo para poder calcular automaticamente
  If frmDebitoCredito.cboOperacion = "Dcg" Or frmDebitoCredito.cboOperacion = "Dco" Or frmDebitoCredito.cboOperacion = "Dcv" Then
    
    If frmDebitoCredito.txtComprobante = "" Then
      intRes = MsgBox("Debe seleccionar comprobante origen", vbCritical + vbOKOnly, "atención...")
      Exit Sub
    End If
    
    If frmDebitoCredito.txtIDcotizacionDolar = "" Then
      intRes = MsgBox("Debe seleccionar alguna cotización", vbCritical + vbOKOnly, "atención...")
      Exit Sub
    End If
      
  End If

  ' cargo form
  Load frmDebitoCreditoUpdate
  
  ' le paso los valores a modificar
  frmDebitoCreditoUpdate.txtConcepto = ""
  frmDebitoCreditoUpdate.txtCantidad = "0.000"
  frmDebitoCreditoUpdate.txtPrecio = "0.000000000000"
  frmDebitoCreditoUpdate.txtImporte = "0.00"
  frmDebitoCreditoUpdate.txtCantidad.SelLength = Len(frmDebitoCreditoUpdate.txtCantidad)
  frmDebitoCreditoUpdate.txtPrecio.SelLength = Len(frmDebitoCreditoUpdate.txtPrecio)
  frmDebitoCreditoUpdate.txtImporte.SelLength = Len(frmDebitoCreditoUpdate.txtImporte)
  
  'si cliente con iva marco el check con true
  If Val(Me.txtIvaPje) <> 0 Then
    frmDebitoCreditoUpdate.chkIva = 1
  Else
    frmDebitoCreditoUpdate.chkIva = 0
  End If
  
  ' lo muestro modal
  frmDebitoCreditoUpdate.Show vbModal
  
  ' si acepto update
  If blnAceptar Then
    
    Set item = lvwDetalle.ListItems.Add
    item.Selected = True
    intRes = lvwSetValue(lvwDetalle, "Concepto", frmDebitoCreditoUpdate.txtConcepto)
    intRes = lvwSetValue(lvwDetalle, "Cantidad", Format(Val(frmDebitoCreditoUpdate.txtCantidad), "########0.000"))
    intRes = lvwSetValue(lvwDetalle, "Precio", Format(Val(frmDebitoCreditoUpdate.txtPrecio), "########0.000000000000"))
    intRes = lvwSetValue(lvwDetalle, "Importe", Format(Val(frmDebitoCreditoUpdate.txtImporte), "########0.00"))
    intRes = lvwSetValue(lvwDetalle, "CantidadInfo", "0.000")
    intRes = lvwSetValue(lvwDetalle, "CantidadInfo1", "0.000")
    intRes = lvwSetValue(lvwDetalle, "Iva", IIf(Val(frmDebitoCreditoUpdate.chkIva) = 1, "Si", "No"))
    intRes = lvwSetValue(lvwDetalle, "Tipo Item", frmDebitoCreditoUpdate.cboTipoItem)
    intRes = lvwSetValue(lvwDetalle, "idVentasTipo", frmDebitoCreditoUpdate.cboTipoItem.List(frmDebitoCreditoUpdate.cboTipoItem.ListIndex))
    intRes = lvwSetValue(lvwDetalle, "Unidad", frmDebitoCreditoUpdate.cboUnidad)
    intRes = lvwSetValue(lvwDetalle, "IDunidad", frmDebitoCreditoUpdate.cboUnidad.ItemData(frmDebitoCreditoUpdate.cboUnidad.ListIndex))
    item.Selected = False

  End If
  
  ' descargo form
  Unload frmDebitoCreditoUpdate

End Sub

Private Sub mnuModificar_Click()
  
  If lvwDetalle.ListItems.Count = 0 Then
    intRes = MsgBox("No se selecciono ningún item.", vbInformation + vbOKOnly, "modificando detalle...")
    Exit Sub
  End If
  
  'chequeo que se haya seleccionado algo
  If Not lvwDetalle.SelectedItem.Checked Then
    intRes = MsgBox("No se selecciono ningún item.", vbInformation + vbOKOnly, "modificando detalle...")
    Exit Sub
  End If
  
  'cargo form
  Load frmDebitoCreditoUpdate
  
  'le paso los valores a modificar
  frmDebitoCreditoUpdate.txtConcepto = lvwGetValue(lvwDetalle, "concepto")
  frmDebitoCreditoUpdate.txtCantidad = lvwGetValue(lvwDetalle, "cantidad")
  frmDebitoCreditoUpdate.txtPrecio = lvwGetValue(lvwDetalle, "precio")
  frmDebitoCreditoUpdate.txtImporte = lvwGetValue(lvwDetalle, "importe")
  frmDebitoCreditoUpdate.chkIva = IIf(lvwGetValue(lvwDetalle, "iva") = "Si", 1, 0)
  
  frmDebitoCreditoUpdate.txtCantidad.SelLength = Len(frmDebitoCreditoUpdate.txtCantidad)
  frmDebitoCreditoUpdate.txtPrecio.SelLength = Len(frmDebitoCreditoUpdate.txtPrecio)
  frmDebitoCreditoUpdate.txtImporte.SelLength = Len(frmDebitoCreditoUpdate.txtImporte)
  frmDebitoCreditoUpdate.cboTipoItem.ListIndex = ComboBoxFindItem(frmDebitoCreditoUpdate.cboTipoItem, lvwGetValue(lvwDetalle, "Tipo Item"))
  frmDebitoCreditoUpdate.cboUnidad.ListIndex = ComboBoxFindItem(frmDebitoCreditoUpdate.cboUnidad, lvwGetValue(lvwDetalle, "unidad"))
  
  
  ' lo muestro modal
  frmDebitoCreditoUpdate.Show vbModal
  
  ' si acepto update
  If blnAceptar Then
    intRes = lvwSetValue(lvwDetalle, "Concepto", frmDebitoCreditoUpdate.txtConcepto)
    intRes = lvwSetValue(lvwDetalle, "Cantidad", frmDebitoCreditoUpdate.txtCantidad)
    intRes = lvwSetValue(lvwDetalle, "Precio", frmDebitoCreditoUpdate.txtPrecio)
    intRes = lvwSetValue(lvwDetalle, "Importe", frmDebitoCreditoUpdate.txtImporte)
    intRes = lvwSetValue(lvwDetalle, "Iva", IIf(Val(frmDebitoCreditoUpdate.chkIva) = 1, "Si", "No"))
    intRes = lvwSetValue(lvwDetalle, "Tipo Item", frmDebitoCreditoUpdate.cboTipoItem)
    intRes = lvwSetValue(lvwDetalle, "idVentasTipo", frmDebitoCreditoUpdate.cboTipoItem.List(frmDebitoCreditoUpdate.cboTipoItem.ListIndex))
    intRes = lvwSetValue(lvwDetalle, "Unidad", frmDebitoCreditoUpdate.cboUnidad)
    intRes = lvwSetValue(lvwDetalle, "IDunidad", frmDebitoCreditoUpdate.cboUnidad.ItemData(frmDebitoCreditoUpdate.cboUnidad.ListIndex))
  End If
  
  ' descargo form
  Unload frmDebitoCreditoUpdate

End Sub

