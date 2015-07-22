VERSION 5.00
Object = "{0E59F1D2-1FBE-11D0-8FF2-00A0D10038BC}#1.0#0"; "msscript.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmVentasInfo 
   BackColor       =   &H80000018&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Facturas Oil"
   ClientHeight    =   10110
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11010
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   10110
   ScaleWidth      =   11010
   StartUpPosition =   2  'CenterScreen
   Begin VB.ComboBox cboProvincia 
      Appearance      =   0  'Flat
      BackColor       =   &H80000018&
      ForeColor       =   &H00808080&
      Height          =   315
      Left            =   1620
      Sorted          =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   6
      Top             =   2010
      Width           =   3705
   End
   Begin VB.TextBox txtTitulo 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000018&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      ForeColor       =   &H00808080&
      Height          =   285
      Index           =   20
      Left            =   135
      TabIndex        =   58
      Text            =   "Texto Libre"
      Top             =   4380
      Width           =   10755
   End
   Begin VB.TextBox txt_Libre 
      BackColor       =   &H80000018&
      Height          =   1680
      Left            =   90
      MultiLine       =   -1  'True
      TabIndex        =   46
      Top             =   4650
      Width           =   10820
   End
   Begin VB.TextBox txtTitulo 
      Appearance      =   0  'Flat
      BackColor       =   &H80000018&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      ForeColor       =   &H00808080&
      Height          =   285
      Index           =   19
      Left            =   2835
      TabIndex        =   57
      Text            =   "Fecha Contable"
      Top             =   1020
      Width           =   1170
   End
   Begin VB.TextBox txtContable 
      BackColor       =   &H80000018&
      ForeColor       =   &H00808080&
      Height          =   300
      Left            =   4095
      TabIndex        =   3
      Top             =   975
      Width           =   1230
   End
   Begin VB.Frame Frame4 
      BackColor       =   &H80000018&
      Height          =   645
      Left            =   90
      TabIndex        =   48
      Top             =   6405
      Width           =   10815
      Begin VB.TextBox txt_IDprovincia_Ent 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000018&
         ForeColor       =   &H00808080&
         Height          =   285
         Left            =   135
         Locked          =   -1  'True
         TabIndex        =   62
         Top             =   225
         Visible         =   0   'False
         Width           =   315
      End
      Begin VB.CheckBox chk_Resguardo 
         BackColor       =   &H80000018&
         Caption         =   "Emisiòn de Resguardo"
         Height          =   285
         Left            =   900
         TabIndex        =   59
         Top             =   225
         Width           =   2085
      End
      Begin VB.CommandButton cmdGuardar 
         Caption         =   "&Guardar"
         Enabled         =   0   'False
         Height          =   330
         Left            =   9030
         TabIndex        =   52
         Top             =   180
         Width           =   1635
      End
      Begin VB.CommandButton cmdSoporte 
         Caption         =   "&Soporte"
         Enabled         =   0   'False
         Height          =   330
         Left            =   7290
         TabIndex        =   51
         Top             =   180
         Width           =   1635
      End
      Begin VB.CommandButton cmdFacturar 
         Caption         =   "&Facturar"
         Height          =   330
         Left            =   5580
         TabIndex        =   50
         Top             =   180
         Width           =   1635
      End
      Begin VB.CommandButton cmdSalir 
         Caption         =   "&Salir"
         Height          =   330
         Left            =   3870
         TabIndex        =   49
         Top             =   180
         Width           =   1635
      End
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H80000018&
      Height          =   1185
      Left            =   90
      TabIndex        =   34
      Top             =   3165
      Width           =   10815
      Begin VB.TextBox txtTitulo 
         Appearance      =   0  'Flat
         BackColor       =   &H80000018&
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         ForeColor       =   &H00808080&
         Height          =   285
         Index           =   22
         Left            =   5355
         TabIndex        =   61
         Text            =   "IIBB ($)"
         Top             =   855
         Width           =   1170
      End
      Begin VB.TextBox txtIIBBImp 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000018&
         ForeColor       =   &H00808080&
         Height          =   285
         Left            =   6570
         Locked          =   -1  'True
         TabIndex        =   41
         Top             =   810
         Width           =   1290
      End
      Begin VB.TextBox txtTitulo 
         Appearance      =   0  'Flat
         BackColor       =   &H80000018&
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         ForeColor       =   &H00808080&
         Height          =   285
         Index           =   21
         Left            =   2475
         TabIndex        =   60
         Text            =   "IIBB (%)"
         Top             =   855
         Width           =   1125
      End
      Begin VB.TextBox txtIIBBPje 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000018&
         ForeColor       =   &H00808080&
         Height          =   285
         Left            =   3735
         TabIndex        =   38
         Top             =   810
         Width           =   855
      End
      Begin VB.TextBox txtRg3337Imp 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000018&
         ForeColor       =   &H00808080&
         Height          =   285
         Left            =   6570
         Locked          =   -1  'True
         TabIndex        =   40
         Top             =   495
         Width           =   1290
      End
      Begin VB.TextBox txtTitulo 
         Appearance      =   0  'Flat
         BackColor       =   &H80000018&
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         ForeColor       =   &H00808080&
         Height          =   285
         Index           =   18
         Left            =   5355
         TabIndex        =   56
         Text            =   "Iva Rg 3337 ($)"
         Top             =   540
         Width           =   1170
      End
      Begin VB.TextBox txtRg3337Pje 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000018&
         ForeColor       =   &H00808080&
         Height          =   285
         Left            =   3735
         TabIndex        =   37
         Top             =   495
         Width           =   855
      End
      Begin VB.TextBox txtTitulo 
         Appearance      =   0  'Flat
         BackColor       =   &H80000018&
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         ForeColor       =   &H00808080&
         Height          =   285
         Index           =   17
         Left            =   2475
         TabIndex        =   55
         Text            =   "Iva Rg 3337 (%)"
         Top             =   540
         Width           =   1215
      End
      Begin VB.TextBox txtTitulo 
         Appearance      =   0  'Flat
         BackColor       =   &H80000018&
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         ForeColor       =   &H00808080&
         Height          =   240
         Index           =   12
         Left            =   8865
         TabIndex        =   47
         Text            =   "Total"
         Top             =   540
         Width           =   405
      End
      Begin VB.TextBox txtTitulo 
         Appearance      =   0  'Flat
         BackColor       =   &H80000018&
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         ForeColor       =   &H00808080&
         Height          =   285
         Index           =   11
         Left            =   5355
         TabIndex        =   44
         Text            =   "Iva ($)"
         Top             =   225
         Width           =   990
      End
      Begin VB.TextBox txtTitulo 
         Appearance      =   0  'Flat
         BackColor       =   &H80000018&
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         ForeColor       =   &H00808080&
         Height          =   285
         Index           =   1
         Left            =   2475
         TabIndex        =   42
         Text            =   "Iva (%)"
         Top             =   225
         Width           =   900
      End
      Begin VB.TextBox txtIvaImp 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000018&
         ForeColor       =   &H00808080&
         Height          =   285
         Left            =   6570
         Locked          =   -1  'True
         TabIndex        =   39
         Top             =   180
         Width           =   1305
      End
      Begin VB.TextBox txtIvaPje 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000018&
         ForeColor       =   &H00808080&
         Height          =   285
         Left            =   3735
         TabIndex        =   36
         Top             =   180
         Width           =   855
      End
      Begin VB.TextBox txtTotal 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000018&
         ForeColor       =   &H00808080&
         Height          =   285
         Left            =   8820
         Locked          =   -1  'True
         TabIndex        =   45
         Top             =   810
         Width           =   1845
      End
      Begin VB.TextBox txtSubtotal 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000018&
         ForeColor       =   &H00808080&
         Height          =   285
         Left            =   135
         Locked          =   -1  'True
         TabIndex        =   43
         Top             =   810
         Width           =   1485
      End
      Begin VB.TextBox txtTitulo 
         Appearance      =   0  'Flat
         BackColor       =   &H80000018&
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         ForeColor       =   &H00808080&
         Height          =   285
         Index           =   9
         Left            =   180
         TabIndex        =   35
         Text            =   "Subtotal"
         Top             =   540
         Width           =   630
      End
      Begin VB.Line Line3 
         BorderColor     =   &H80000000&
         BorderWidth     =   2
         X1              =   8235
         X2              =   8235
         Y1              =   90
         Y2              =   1170
      End
      Begin VB.Line Line2 
         BorderColor     =   &H80000000&
         BorderWidth     =   2
         X1              =   4995
         X2              =   4995
         Y1              =   90
         Y2              =   1125
      End
      Begin VB.Line Line1 
         BorderColor     =   &H80000000&
         BorderWidth     =   2
         X1              =   2070
         X2              =   2070
         Y1              =   90
         Y2              =   1170
      End
   End
   Begin VB.TextBox txtTipoCambio 
      BackColor       =   &H80000018&
      ForeColor       =   &H00808080&
      Height          =   300
      Left            =   7065
      TabIndex        =   13
      Top             =   1755
      Width           =   3705
   End
   Begin VB.TextBox txtTitulo 
      Appearance      =   0  'Flat
      BackColor       =   &H80000018&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      ForeColor       =   &H00808080&
      Height          =   285
      Index           =   15
      Left            =   5625
      TabIndex        =   33
      Text            =   "Tipo de Cambio"
      Top             =   1755
      Width           =   1260
   End
   Begin VB.TextBox txtTitulo 
      Appearance      =   0  'Flat
      BackColor       =   &H80000018&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      ForeColor       =   &H00808080&
      Height          =   285
      Index           =   14
      Left            =   5625
      TabIndex        =   32
      Text            =   "Facturacion en"
      Top             =   1395
      Width           =   1350
   End
   Begin VB.ComboBox cboBase 
      Appearance      =   0  'Flat
      BackColor       =   &H80000018&
      ForeColor       =   &H00808080&
      Height          =   315
      ItemData        =   "frmVentasInfo.frx":0000
      Left            =   7065
      List            =   "frmVentasInfo.frx":000A
      Style           =   2  'Dropdown List
      TabIndex        =   12
      Top             =   1395
      Width           =   3705
   End
   Begin MSScriptControlCtl.ScriptControl ScriptControl1 
      Left            =   120
      Top             =   9420
      _ExtentX        =   1005
      _ExtentY        =   1005
   End
   Begin VB.ComboBox cboCondicion 
      Appearance      =   0  'Flat
      BackColor       =   &H80000018&
      ForeColor       =   &H00808080&
      Height          =   315
      ItemData        =   "frmVentasInfo.frx":001B
      Left            =   7065
      List            =   "frmVentasInfo.frx":001D
      Style           =   2  'Dropdown List
      TabIndex        =   9
      Top             =   270
      Width           =   3120
   End
   Begin VB.CommandButton cmdNewCondicion 
      Caption         =   "New"
      Height          =   285
      Left            =   10215
      TabIndex        =   31
      Top             =   270
      Width           =   525
   End
   Begin VB.TextBox txtTitulo 
      Appearance      =   0  'Flat
      BackColor       =   &H80000018&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      ForeColor       =   &H00808080&
      Height          =   285
      Index           =   13
      Left            =   5625
      TabIndex        =   30
      Text            =   "Condición"
      Top             =   315
      Width           =   1350
   End
   Begin VB.TextBox txtTitulo 
      Appearance      =   0  'Flat
      BackColor       =   &H80000018&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      ForeColor       =   &H00808080&
      Height          =   285
      Index           =   10
      Left            =   180
      TabIndex        =   27
      Text            =   "TipoOperación"
      Top             =   2415
      Width           =   1350
   End
   Begin MSComctlLib.ListView lvwEntregasCli 
      Height          =   1410
      Left            =   90
      TabIndex        =   25
      Top             =   7125
      Width           =   10815
      _ExtentX        =   19076
      _ExtentY        =   2487
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      Checkboxes      =   -1  'True
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
      Height          =   285
      Index           =   8
      Left            =   5625
      TabIndex        =   24
      Text            =   "FechaVencimiento"
      Top             =   675
      Width           =   1350
   End
   Begin VB.TextBox txtFechaVencimiento 
      BackColor       =   &H80000018&
      ForeColor       =   &H00808080&
      Height          =   300
      Left            =   7065
      TabIndex        =   10
      Top             =   675
      Width           =   3705
   End
   Begin VB.CommandButton cmdNewFormaPago 
      Caption         =   "New"
      Height          =   285
      Left            =   10215
      TabIndex        =   23
      Top             =   1035
      Width           =   525
   End
   Begin VB.ComboBox cboCuentaBancaria 
      Appearance      =   0  'Flat
      BackColor       =   &H80000018&
      ForeColor       =   &H00808080&
      Height          =   315
      ItemData        =   "frmVentasInfo.frx":001F
      Left            =   7065
      List            =   "frmVentasInfo.frx":0029
      Style           =   2  'Dropdown List
      TabIndex        =   11
      Top             =   1035
      Width           =   3120
   End
   Begin VB.TextBox txtTitulo 
      Appearance      =   0  'Flat
      BackColor       =   &H80000018&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      ForeColor       =   &H00808080&
      Height          =   285
      Index           =   7
      Left            =   5625
      TabIndex        =   22
      Text            =   "Forma Pago"
      Top             =   1035
      Width           =   1350
   End
   Begin VB.CommandButton cmdNewTipoMoneda 
      Caption         =   "New"
      Height          =   285
      Left            =   4770
      TabIndex        =   21
      Top             =   2715
      Width           =   525
   End
   Begin VB.TextBox txtFecha 
      BackColor       =   &H80000018&
      ForeColor       =   &H00808080&
      Height          =   300
      Left            =   1620
      Locked          =   -1  'True
      TabIndex        =   2
      Top             =   975
      Width           =   1140
   End
   Begin VB.TextBox txtFactura 
      BackColor       =   &H80000018&
      ForeColor       =   &H00808080&
      Height          =   300
      Left            =   1620
      Locked          =   -1  'True
      TabIndex        =   1
      Top             =   630
      Width           =   3705
   End
   Begin VB.ComboBox cboTipoMoneda 
      Appearance      =   0  'Flat
      BackColor       =   &H80000018&
      ForeColor       =   &H00808080&
      Height          =   315
      ItemData        =   "frmVentasInfo.frx":0037
      Left            =   1620
      List            =   "frmVentasInfo.frx":0041
      Style           =   2  'Dropdown List
      TabIndex        =   8
      Top             =   2715
      Width           =   3120
   End
   Begin VB.ComboBox cboTipoComprobante 
      Appearance      =   0  'Flat
      BackColor       =   &H80000018&
      ForeColor       =   &H00808080&
      Height          =   315
      ItemData        =   "frmVentasInfo.frx":004F
      Left            =   1620
      List            =   "frmVentasInfo.frx":0056
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   270
      Width           =   3705
   End
   Begin VB.ComboBox cboOperacion 
      Appearance      =   0  'Flat
      BackColor       =   &H80000018&
      ForeColor       =   &H00808080&
      Height          =   315
      ItemData        =   "frmVentasInfo.frx":0063
      Left            =   1620
      List            =   "frmVentasInfo.frx":0070
      Style           =   2  'Dropdown List
      TabIndex        =   7
      Top             =   2355
      Width           =   3705
   End
   Begin VB.ComboBox cboCliente 
      Appearance      =   0  'Flat
      BackColor       =   &H80000018&
      ForeColor       =   &H00808080&
      Height          =   315
      Left            =   1620
      Sorted          =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   5
      Top             =   1665
      Width           =   3705
   End
   Begin VB.ComboBox cboEmpresa 
      Appearance      =   0  'Flat
      BackColor       =   &H80000018&
      ForeColor       =   &H00808080&
      Height          =   315
      Left            =   1620
      Sorted          =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   4
      Top             =   1305
      Width           =   3705
   End
   Begin VB.TextBox txtTitulo 
      Appearance      =   0  'Flat
      BackColor       =   &H80000018&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      ForeColor       =   &H00808080&
      Height          =   285
      Index           =   0
      Left            =   180
      TabIndex        =   20
      Text            =   "Empresa"
      Top             =   1305
      Width           =   1710
   End
   Begin VB.TextBox txtTitulo 
      Appearance      =   0  'Flat
      BackColor       =   &H80000018&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      ForeColor       =   &H00808080&
      Height          =   285
      Index           =   2
      Left            =   180
      TabIndex        =   19
      Text            =   "Comprobante Tipo"
      Top             =   270
      Width           =   1305
   End
   Begin VB.TextBox txtTitulo 
      Appearance      =   0  'Flat
      BackColor       =   &H80000018&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      ForeColor       =   &H00808080&
      Height          =   285
      Index           =   3
      Left            =   180
      TabIndex        =   18
      Text            =   "TipoMoneda"
      Top             =   2775
      Width           =   1380
   End
   Begin VB.TextBox txtTitulo 
      Appearance      =   0  'Flat
      BackColor       =   &H80000018&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      ForeColor       =   &H00808080&
      Height          =   285
      Index           =   4
      Left            =   180
      TabIndex        =   17
      Text            =   "Comprobante Nro."
      Top             =   630
      Width           =   1755
   End
   Begin VB.TextBox txtTitulo 
      Appearance      =   0  'Flat
      BackColor       =   &H80000018&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      ForeColor       =   &H00808080&
      Height          =   285
      Index           =   5
      Left            =   180
      TabIndex        =   16
      Text            =   "Fecha Factura"
      Top             =   990
      Width           =   1395
   End
   Begin VB.TextBox txtTitulo 
      Appearance      =   0  'Flat
      BackColor       =   &H80000018&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      ForeColor       =   &H00808080&
      Height          =   285
      Index           =   6
      Left            =   180
      TabIndex        =   15
      Text            =   "Cliente"
      Top             =   1695
      Width           =   1710
   End
   Begin MSComctlLib.ListView lvwContratos 
      Height          =   1500
      Left            =   90
      TabIndex        =   26
      Top             =   8520
      Width           =   10815
      _ExtentX        =   19076
      _ExtentY        =   2646
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      Checkboxes      =   -1  'True
      FullRowSelect   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   0
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H80000018&
      Height          =   3105
      Left            =   90
      TabIndex        =   28
      Top             =   45
      Width           =   5370
      Begin VB.TextBox txtTitulo 
         Appearance      =   0  'Flat
         BackColor       =   &H80000018&
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         ForeColor       =   &H00808080&
         Height          =   285
         Index           =   23
         Left            =   90
         TabIndex        =   63
         Text            =   "Provincia"
         Top             =   2010
         Width           =   1710
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H80000018&
      Height          =   3090
      Left            =   5535
      TabIndex        =   29
      Top             =   45
      Width           =   5370
      Begin VB.CommandButton cmdNewCotizacionTexto 
         Caption         =   "New"
         Height          =   285
         Left            =   4680
         TabIndex        =   54
         Top             =   2070
         Width           =   525
      End
      Begin VB.TextBox txtTitulo 
         Appearance      =   0  'Flat
         BackColor       =   &H80000018&
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         ForeColor       =   &H00808080&
         Height          =   285
         Index           =   16
         Left            =   90
         TabIndex        =   53
         Text            =   "Cotizacion Texto"
         Top             =   2115
         Width           =   1260
      End
      Begin VB.ComboBox cboCotizacion 
         Appearance      =   0  'Flat
         BackColor       =   &H80000018&
         ForeColor       =   &H00808080&
         Height          =   315
         ItemData        =   "frmVentasInfo.frx":0083
         Left            =   1530
         List            =   "frmVentasInfo.frx":008D
         Style           =   2  'Dropdown List
         TabIndex        =   14
         Top             =   2070
         Width           =   3120
      End
   End
   Begin VB.Menu mnuSeleccionar 
      Caption         =   "Seleccionar"
      Visible         =   0   'False
      Begin VB.Menu mnuBarrels 
         Caption         =   "Agregar volumen Barrels"
      End
      Begin VB.Menu mnuEmbarque 
         Caption         =   "Seleccionar Fecha Embarque"
      End
      Begin VB.Menu mnuPrecios 
         Caption         =   "Seleccionar rango de Precios"
      End
   End
End
Attribute VB_Name = "frmVentasInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim intCuenta, intCantidad, intPromedio, intCual, intCantPrecios As Integer
Dim intPrevios, intPosteriores, intTerminal As Integer
Dim strFecha, strFechaDesde, strFechaHasta, strIVACliente, strNuevoNumero As String
Dim blnVolumen, blnFechaEntregaCli As Boolean
Dim strID(), strInfo() As String
Dim curMin(), curMAx(), curAvg(), curSuma, curAPIGravity, curAPIMin, curAPIMax As Currency
Dim curAZUMin, curAZUMax, curDTO, curAvgPrecioDesc As Currency
Dim sngBarrelsTo1556, sngTotalBarrels, sngTotal1556 As Double
Dim strConceptoVenta, strContratosRango As String
Dim strConceptoSoporte As String

'
' segun empresa y tipo de comprobante tomo ultimo numero disponible
'
Private Function tomoNumeracionNueva()
    
    Dim rs As ADODB.Recordset

    'valido que se haya seleccionado tipo Comprobante y Empresa
    If Me.cboEmpresa.ListIndex = -1 Or Me.cboTipoComprobante.ListIndex = -1 Or Me.cboOperacion.ListIndex = -1 Then
        Exit Function
    End If
    
    ' tomo valores de empresas para numero de factura automatico
    strSQL = "SELECT * FROM empresas_Puntos_Venta " & _
             "WHERE IDempresa = " & Me.cboEmpresa.ItemData(Me.cboEmpresa.ListIndex) & _
             " AND " & _
             "Operacion = '" & UCase(Left(Me.cboOperacion, 3)) & "'" & _
             " AND " & _
             "Comprobante = '" & UCase(Left(Me.cboTipoComprobante, 3)) & "'"
             
  
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
                        UCase(Left(Me.cboOperacion, 3)) & "-" & UCase(Left(Me.cboTipoComprobante, 3))
     

      ' 20/07/2015 - Emazzu - Esto no se usa mas, pero lo dejamos
      
      ' guardo nuevo numero automatico, esto lo hago por que si despues
      ' me lo modifican a mano, e ingresan un numero valido, pero que
      ' no es el ultimo disponible, no lo tengo que actualizar
      strNuevoNumero = Me.txtFactura
  
    End If
    
    rs.Close
    

End Function

Private Sub cboCliente_Click()

    a = lvwBuscaEntregasCliContratos()

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



Private Sub cboEmpresa_Click()
  
  intRes = tomoNumeracionNueva()
  intRes = lvwBuscaEntregasCliContratos()

End Sub




Private Sub cboOperacion_Click()

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


Private Sub cboTipoComprobante_Click()

  intRes = tomoNumeracionNueva()

End Sub


Private Sub cmdFacturar_Click()
    
  'chequeo error
  On Error GoTo controlError
  
  Dim rsPrecios As ADODB.Recordset
  Dim rsFormula As ADODB.Recordset
  Dim strT As String
  
  ' inicializo variables
  intCuenta = 0
  intCantidad = 0
  intPromedio = 0
  intCual = 0
  intCantPrecios = 0
  intPrevios = 0
  intPosteriores = 0
  intRes = 0
  
  strFecha = ""
  strIVACliente = ""
  strSQL = ""
  strConceptoSoporte = ""
  
  blnVolumen = False
  blnFechaEmbarque = False
  
  ' calcular promedio de precios
  ReDim strID(0)
  ReDim strInfo(0)
  ReDim curMin(0)
  ReDim curMAx(0)
  ReDim curAvg(0)
  
  curSuma = 0
  curAPIGravity = 0
  curAPIMin = 0
  curAPIMax = 0
  curAZUMin = 0
  curAZUMax = 0
  curDTO = 0
  curAvgPrecioDesc = 0
  curIVARes = 0
  curIVAFin = 0
  curIVAExe = 0
  curIVAExp = 0
  curCOEAju = 0
  
  sngBarrelsTo1556 = 0
  sngTotalBarrels = 0
  sngTotal1556 = 0
  sngTotal15 = 0
  
  ' validaciones
  If Not DataValidate(cboEmpresa, , True) Then Exit Sub
  If Not DataValidate(cboOperacion, , True) Then Exit Sub
  If Not DataValidate(cboTipoComprobante, , True) Then Exit Sub
  If Not DataValidate(cboTipoMoneda, , True) Then Exit Sub
  If Not DataValidate(txtFecha, "dd/mm/yyyy", True) Then Exit Sub
  If Not DataValidate(txtContable, "dd/mm/yyyy", True) Then Exit Sub
  If Not DataValidate(txtFechaVencimiento, "dd/mm/yyyy", True) Then Exit Sub
  If Not DataValidate(cboCondicion, , True) Then Exit Sub
  If Not DataValidate(cboCliente, , True) Then Exit Sub
  If Not DataValidate(cboProvincia, , True) Then Exit Sub
  If Not DataValidate(cboCuentaBancaria, , True) Then Exit Sub
  If Not DataValidate(cboBase, , True) Then Exit Sub
  If Not DataValidate(cboCotizacion, , True) Then Exit Sub
  If Not DataValidate(txtTipoCambio, "###.###", True) Then Exit Sub
    
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
  
  
  'solo se puede seleccionar un solo item de entrega clientes
  intCantidad = 0
  For intCuenta = 1 To lvwEntregasCli.ListItems.Count
    If lvwEntregasCli.ListItems(intCuenta).Checked Then
      intCual = intCuenta
      intCantidad = intCantidad + 1
    End If
  Next
  
  If intCantidad <> 1 Then
    intRes = MsgBox("Para poder  facturar debe seleccionar un Item Entrega Clientes.", vbInformation + vbOKOnly, "Información")
    Exit Sub
  End If
  
  ' selecciono el item que se le hizo el tilde de entregasCli
  lvwEntregasCli.ListItems(intCual).Selected = True
  
  ' valido seleccion de 1 o mas contratos y asignacion de volumen
  ' si hay algun contrato seleccionado, le asigno la fecha de Entrega Clientes
  ' seleccionado.
  ' Tambien chequeo que la fecha de entregaClientes asignada a cada contrato
  ' sea valida y tenga cotizacion en precios
  intPromedio = 1
  blnVolumen = True
  blnFechaEntregaCli = True
  
  ' recorro lvw de contratos
  For intCuenta = 1 To lvwContratos.ListItems.Count
    
    ' chequeo si el item tiene un tilde
    If lvwContratos.ListItems(intCuenta).Checked Then
      intCantidad = intCantidad + 1
      
      ' me paro sobre el item para poder tomar los valores
      lvwContratos.ListItems(intCuenta).Selected = True
      
      ' chequeo si columna Barrels esta en 0, o sea si se ingreso volumen
      If Val(lvwGetValue(lvwContratos, "barrelss")) = 0 Then ' Barrels
        blnVolumen = False
      End If
            
      ' chequeo si columna entregaCli esta en blanco, o sea no se agrego la fecha de embarque al contrato
      If lvwGetValue(lvwContratos, "entregaCli") = "" Then
      
        ' agrego fecha de entregasCli a cada contrato que se le hizo un tilde
        intRes = lvwSetValue(lvwContratos, "entregacli", lvwGetValue(lvwEntregasCli, "fecha"))
      End If
    
      'chequeo que exista una fecha valida para entregasClientes
      strSQL = "SELECT * FROM ViewVentasPrecios " & _
               "WHERE Fecha = '" & dateToIso(lvwGetValue(lvwContratos, "entregaCli")) & "' AND " & _
               "Operacion = '" & Me.cboOperacion.List(Me.cboOperacion.ListIndex) & "' AND " & _
               "PrecioTipoID = " & lvwGetValue(lvwContratos, "precioTipo") & " AND " & _
               "ValorMin > 0"
      
      Set rsPrecios = adoGetRS(strSQL)
      If rsPrecios.EOF Then
        blnFechaEntregaCli = False
      End If
      rsPrecios.Close
    
    End If
  
  Next
  
  If intCantidad = 0 Then
    intRes = MsgBox("Para facturar debe seleccionar por lo menos un contrato.", vbInformation + vbOKOnly, "Información")
    Exit Sub
  End If
  
  If Not blnVolumen Then
    intRes = MsgBox("Para facturar debe asignar a cada contrato seleccionado el volumen barrels.", vbInformation + vbOKOnly, "Información")
    Exit Sub
  End If
  
  If Not blnFechaEntregaCli Then
    intRes = MsgBox("Para la fecha de embarque seleccionada no existe cotizacion.", vbInformation + vbOKOnly, "Información")
    Exit Sub
  End If
  
  ' comienzo de facturacion
  ' tomo valores de parametros para calculo
  ' tomo valores de clientes para calculo
  ' tomo valores de embarque para calculo APIGravity
  ' recorro contratos
  ' tomo solo los contratos seleccionados
  ' tomo valores para calculo APIMin, APIMax
  ' armo SQL con rango de fechas segun definicion para calcular promedios
  ' recorro recordset y comienzo a calcular el WTI
  ' guardo todos los registros es 5 arrays identificandolos por el array
  ' strID que corrresponde al numero de contrato, cuando termino
  ' agrego al array 4 lineas mas que corresponden a
  ' Wti, Discount, API Adjust, Wti Final y el ultimo elemento cuando finaliza
  ' el recordset corresponde al Wti final para ese contrato, por lo tanto
  ' se lo agrego al lvwContratos.
  
  
  ' CLEAR totales de facturaciòn
  Me.txtSubtotal = Format(0, "########0.00")
  Me.txtIvaPje = Format(0, "########0.00")
  Me.txtRg3337Pje = Format(0, "########0.00")
  
  ' En este unico caso, no lo vacio, porque viene del txt, cuando se selecciona una provincia
  Me.txtIIBBPje = Format(Me.txtIIBBPje, "########0.00")
  Me.txtIvaImp = Format(0, "########0.00")
  Me.txtRg3337Imp = Format(0, "########0.00")
  Me.txtIIBBImp = Format(0, "########0.00")
  Me.txtTotal = Format(0, "########0.00")
  
  
  'GET informacion cliente seleccionado, con provincia de entrega
  strSQL = "SELECT * FROM ViewVentasClientes WHERE Cliente = '" & _
           Me.cboCliente & "'"
    
    
  'DECLARE
  Dim strRg3337 As String
  Dim sngIngBrutos As Single
  
  'ASSIGN valores
  strIVACliente = ""
  strRg3337 = ""
  sngIngBrutos = 0
  
  
  'GET impuestos
  Set rs = adoGetRS(strSQL)
  
    'chequeo errores
    If Not lngAdoErrNum = -1 Then
      adoError
      Exit Sub
    End If
  
  'CHECK si encontro algo
  If Not rs.EOF Then
  
    strIVACliente = rs!condIVA
    strRg3337 = rs!ivaRG3337
    
  End If
  
    'CLOSE
  rs.Close
 
 
  'ASSIGN IIBB segun provincia, esto viene directo ahora del txt
  sngIngBrutos = Val(Me.txtIIBBPje)
  
  
  ' tomo valores de EntregasCli para calculo
  
  ' tomo apigravity
  curAPIGravity = lvwGetValue(lvwEntregasCli, "Napigravity")
  
  ' tomo Terminal
  intTerminal = lvwGetValue(lvwEntregasCli, "terminalID")
  
  intPromedio = 0
  For intCuenta = 1 To lvwContratos.ListItems.Count
    
    'si se tildo el contrato
    If lvwContratos.ListItems(intCuenta).Checked Then
      
      'hubico puntero en contrato para poder tomar los datos
      lvwContratos.ListItems(intCuenta).Selected = True
      
      'tomo valores de contrato para calculo
      curAPIMin = Val(lvwGetValue(lvwContratos, "apimin"))
      curAPIMax = Val(lvwGetValue(lvwContratos, "apimax"))
      curAZUMin = Val(lvwGetValue(lvwContratos, "azumin"))
      curAZUMax = Val(lvwGetValue(lvwContratos, "azumax"))
      curDTO = Val(lvwGetValue(lvwContratos, "dto"))
      
      'curAvgPrecioDesc = Val(lvwGetValue(lvwContratos, "avgPrecioDesc"))
         
      '28/09/2004: tomo la tabla con los rangos para calcular el curAvgPrecioDesc
      strContratosRango = lvwGetValue(lvwContratos, "tablaRangos")
      
      'determino el sistema utilizado para el calculo
      Select Case Format(lvwGetValue(lvwContratos, "tipoCalculo"), "<")
      
      'TIPO DIAS tomo el rango de diasPrevios y diasPosteriores ------------------------------------------------------
      Case Is = "dias"
        
        'tomo rango de dias previos y posteriores
        intPrevios = Val(lvwGetValue(lvwContratos, "diasprevios")) + 1
        intPosteriores = Val(lvwGetValue(lvwContratos, "diasposteriores")) + 1
          
        ' Armo SQL para determinar la fecha desde
        ' Tomo desde la fecha de embarque para atras o tomo la cantidad
        ' de registros como dias previos esten definidos, las cotizaciones
        ' sin valores no las tomo
        strFecha = lvwGetValue(lvwContratos, "entregaCli")
        strSQL = "SELECT TOP " & str(intPrevios) & " * FROM ViewVentasPrecios " & _
                 "WHERE Fecha <= '" & dateToIso(strFecha) & "' AND " & _
                 "Operacion = '" & frmVentasInfo.cboOperacion.List(frmVentasInfo.cboOperacion.ListIndex) & "' AND " & _
                 "PrecioTipoID = " & lvwGetValue(lvwContratos, "precioTipo") & " AND " & _
                 "ValorMin > 0 " & _
                 "ORDER BY Fecha DESC"
                   
        ' si encontro algun rango de fechas tomo fecha desde
        Set rsPrecios = adoGetRS(strSQL)  ' abro recordset
        If Not rsPrecios.EOF Then         ' Chequeo que exista registro
          rsPrecios.MoveLast              ' puntero en fecha desde
          strFechaDesde = rsPrecios!fecha
        End If
        rsPrecios.Close
        
        ' Armo SQL para determinar la fecha hasta
        ' Tomo desde la fecha de embarque para adelante o tomo la cantidad
        ' de registros como dias posteriores esten definidos, las cotizaciones
        ' sin valores no las tomo
        strFecha = lvwGetValue(lvwContratos, "entregaCli")
        strSQL = "SELECT TOP " & str(intPosteriores) & " * FROM ViewVentasPrecios " & _
                 "WHERE Fecha >='" & dateToIso(strFecha) & "' AND " & _
                 "Operacion = '" & frmVentasInfo.cboOperacion.List(frmVentasInfo.cboOperacion.ListIndex) & "' AND " & _
                 "PrecioTipoID = " & lvwGetValue(lvwContratos, "precioTipo") & " AND " & _
                 "ValorMin > 0 " & _
                 "ORDER BY Fecha"
                        
        Set rsPrecios = adoGetRS(strSQL)
        If Not rsPrecios.EOF Then         ' chequeo que existan registros
          rsPrecios.MoveLast              ' puntero en fecha hasta
          strFechaHasta = rsPrecios!fecha
        End If
        rsPrecios.Close
          
        strSQL = "SELECT * FROM ViewVentasPrecios " & _
                 "WHERE Fecha BETWEEN '" & dateToIso(strFechaDesde) & "' AND '" & dateToIso(strFechaHasta) & "' AND " & _
                 "Operacion = '" & frmVentasInfo.cboOperacion.List(frmVentasInfo.cboOperacion.ListIndex) & "' AND " & _
                 "PrecioTipoID = " & lvwGetValue(lvwContratos, "precioTipo") & " AND " & _
                 "ValorMin > 0" & _
                 "ORDER BY Fecha"
                
      ' TIPO RANGO se toma el rango de fecha precioDesde y precioHasta ------------------------------------------------
      Case Is = "rango"
        
        'valido que se haya ingresado rango de fechas
        If strFechaDesde = "" And strFechaHasta = "" Then
          intRes = MsgBox("Debe seleccionar un rango de fechas.", vbInformation + vbOKOnly, "Información")
          Exit Sub
        End If
        
        strSQL = "SELECT * FROM ViewVentasPrecios " & _
                 "WHERE Fecha BETWEEN '" & dateToIso(strFechaDesde) & "' AND '" & dateToIso(strFechaHasta) & "' AND " & _
                 "Operacion = '" & Me.cboOperacion.List(Me.cboOperacion.ListIndex) & "' AND " & _
                 "PrecioTipoID = " & lvwGetValue(lvwContratos, "preciotipo") & " AND " & _
                 "ValorMin > 0 " & _
                 "ORDER BY Fecha"
      
      ' TIPO MENSUAL puede ser el promedio del mes anterior o actual --------------------------------------------------
      Case Is = "mensual"
      
        ' promedio del mes ANTERIOR ---------------------------------
        If Format(lvwGetValue(lvwContratos, "mespromedio"), "<") = "anterior" Then
        End If
      
        ' promedio del mes ACTUAL ---------------------------------
        If Format(lvwGetValue(lvwContratos, "mespromedio"), "<") = "actual" Then
          
          ' tomo fecha entregaCli actual, y tomo el primer dia y
          '  el ultimo del mes para poder calcular el promedio
          strFecha = lvwGetValue(lvwContratos, "entregaCli")
          strFechaDesde = dateToIso(dateToFirstDay(strFecha))
          strFechaHasta = dateToIso(dateToLastDay(strFecha))
                    
          strSQL = "SELECT * FROM ViewVentasPrecios " & _
                   "WHERE Fecha BETWEEN '" & strFechaDesde & "' AND '" & strFechaHasta & "' AND " & _
                   "Operacion = '" & Me.cboOperacion.List(Me.cboOperacion.ListIndex) & "' AND " & _
                   "PrecioTipoID = " & lvwGetValue(lvwContratos, "preciotipo") & " AND " & _
                   "ValorMin > 0 " & _
                   "ORDER BY Fecha"
        
        End If
      
      End Select
     
      ' levanta el rango de precios armardo en los pasos
      ' anteriores en strSQL segun opciones de calculo
      Set rsPrecios = adoGetRS(strSQL)
      If Not rsPrecios.EOF Then
        
        ' vacio acumuladores de promedio
        curSuma = 0
        intCantPrecios = 0
        
        rsPrecios.MoveFirst
        While Not rsPrecios.EOF
          
          ' redimenciono segun cantidad de items
          ReDim Preserve strID(intPromedio)
          ReDim Preserve strInfo(intPromedio)
          ReDim Preserve curMin(intPromedio)
          ReDim Preserve curMAx(intPromedio)
          ReDim Preserve curAvg(intPromedio)
          
          ' calculo promedios de precios Min y Max
          strID(intPromedio) = lvwGetValue(lvwContratos, "contrato")
          strInfo(intPromedio) = rsPrecios!fecha
          curMin(intPromedio) = rsPrecios!ValorMin
          curMAx(intPromedio) = rsPrecios!valorMAx
          curAvg(intPromedio) = Round((rsPrecios!ValorMin + rsPrecios!valorMAx) / 2, 3)
            
          ' suma parcial para promedio
          curSuma = curSuma + curAvg(intPromedio)
          
          ' acumulador de cantidad de precios
          intCantPrecios = intCantPrecios + 1
                      
          ' valor utilizado para redimencionar array
          intPromedio = intPromedio + 1
          rsPrecios.MoveNext
            
        Wend
        rsPrecios.Close
      
        ' levantamos la formula para cada contrato
        Dim strFormulaPrecio, strFormulaAPI, strFormulaVarios As String
        
        ' tomo formula precio
        strFormulaPrecio = ""
        strSQL = "select parametro from viewParametros where referencia = '" & lvwGetValue(lvwContratos, "formulaajuprecio") & "'"
        Set rsFormula = adoGetRS(strSQL)
        If Not rsFormula.EOF Then
          strFormulaPrecio = rsFormula!parametro
        End If
     
        ' tomo formula api
        strFormulaAPI = ""
        strSQL = "select parametro from viewParametros where referencia = '" & lvwGetValue(lvwContratos, "formulaajuapi") & "'"
        Set rsFormula = adoGetRS(strSQL)
        If Not rsFormula.EOF Then
          strFormulaAPI = rsFormula!parametro
        End If
        
        ' tomo formula varios
        strFormulaVarios = ""
        strSQL = "select parametro from viewParametros where referencia = '" & lvwGetValue(lvwContratos, "formulaajuvarios") & "'"
        Set rsFormula = adoGetRS(strSQL)
        If Not rsFormula.EOF Then
          strFormulaVarios = rsFormula!parametro
        End If
     
        ' cuando termino de recorrer precios disponibles
        ' redimenciono por 6 para guardar average, discount, adjust, net price
        intPromedio = intPromedio + 6
        ReDim Preserve strID(intPromedio)
        ReDim Preserve strInfo(intPromedio)
        ReDim Preserve curMin(intPromedio)
        ReDim Preserve curMAx(intPromedio)
        ReDim Preserve curAvg(intPromedio)
        
        Dim sngPrecio As Single
        
        'determinando precio segun si se factura por m3 o barrels
        Dim intRedondeo As Integer
        If cboBase = "M3" Then
          intRedondeo = 2
          sngPrecio = Format(Round((curSuma / intCantPrecios), 3) * getParam("m31556TObarr1556"), "." & String(Val(lvwGetValue(lvwContratos, "redondeo4")), "#"))
        Else
          intRedondeo = 5
          sngPrecio = Format(curSuma / intCantPrecios, "." & String(Val(lvwGetValue(lvwContratos, "redondeo4")), "#"))
        End If
        
        'guardo average
        strID(intPromedio - 6) = lvwGetValue(lvwContratos, "contrato")
        strInfo(intPromedio - 6) = "Average"
        curMin(intPromedio - 6) = 0
        curMAx(intPromedio - 6) = 0
        curAvg(intPromedio - 6) = sngPrecio
                         
        'calculo descuento1 -------------------------------------------------------------------------
        
        'get descuento
        strT = lvwGetValue(lvwContratos, "dto")
        
        'reemplazo variables por valores
        strT = Replace(strT, "AVG", curAvg(intPromedio - 6))
        
        'si se utilizo variable COEF reemplazo por 0, para que el DESC tome valor
        'para poder calcular tabla de rangos y luego vuelvo a calcular el descuento
        If InStr(strT, "COEF") <> 0 Then
          
          strT = Replace(strT, "COEF", 1)
          
        End If
          
        'guardo discount
        strID(intPromedio - 5) = lvwGetValue(lvwContratos, "contrato")
        strInfo(intPromedio - 5) = "Discount"
        curMin(intPromedio - 5) = 0
        curMAx(intPromedio - 5) = 0
        curAvg(intPromedio - 5) = -Val(Format(Me.ScriptControl1.Eval(strT), "#0." & String(intRedondeo, "#")))
        
        '----------------------------------------------------------------------------------
        'antes el descuento era un valor, ahora puede ser valor o formula
        'por eso utilizo el control script
        'curAvg(intPromedio - 5) = -Val(Format(Val(lvwGetValue(lvwContratos, "dto")), "#0." & String(intredondeo, "#")))
        '----------------------------------------------------------------------------------
               
        '----------------------------------------------------------------------------------
        '07/10/2004: antes de resolvia despues de la tabla de rangos, ahora antes
        '----------------------------------------------------------------------------------
        
        'reemplazo variables por valores para formula AjusteAPIXX
        strFormulaAPI = Replace(strFormulaAPI, "AVG", curAvg(intPromedio - 6))
        strFormulaAPI = Replace(strFormulaAPI, "DISC", curAvg(intPromedio - 5) * -1)
        strFormulaAPI = Replace(strFormulaAPI, "APICLI", lvwGetValue(lvwEntregasCli, "napigravity"))
        strFormulaAPI = Replace(strFormulaAPI, "APIMAXCON", lvwGetValue(lvwContratos, "apimax"))
        strFormulaAPI = Replace(strFormulaAPI, "AZUMAXCON", lvwGetValue(lvwContratos, "azumax"))
        strFormulaAPI = Replace(strFormulaAPI, "AZUCLI", lvwGetValue(lvwEntregasCli, "azufre"))
        
        'guardo AjusteAPIXX
        strID(intPromedio - 3) = lvwGetValue(lvwContratos, "contrato")
        strInfo(intPromedio - 3) = "Api Adjust"
        curMin(intPromedio - 3) = 0
        curMAx(intPromedio - 3) = 0
        curAvg(intPromedio - 3) = Val(Format(ScriptControl1.Eval(strFormulaAPI), "#0." & String(Val(lvwGetValue(lvwContratos, "redondeo2")), "#")))
                       
        'reemplazo variables para formula AjusteVariosXX
        strFormulaVarios = Replace(strFormulaVarios, "AVG", curAvg(intPromedio - 6))
        strFormulaVarios = Replace(strFormulaVarios, "DISC", curAvg(intPromedio - 5) * -1)
        strFormulaVarios = Replace(strFormulaVarios, "APICLI", lvwGetValue(lvwEntregasCli, "napigravity"))
        strFormulaVarios = Replace(strFormulaVarios, "APIMAXCON", lvwGetValue(lvwContratos, "apimax"))
        strFormulaVarios = Replace(strFormulaVarios, "AZUMAXCON", lvwGetValue(lvwContratos, "azumax"))
        strFormulaVarios = Replace(strFormulaVarios, "AZUCLI", lvwGetValue(lvwEntregasCli, "azufre"))
        
        'guardo AjusteVariosXX
        strID(intPromedio - 2) = lvwGetValue(lvwContratos, "contrato")
        strInfo(intPromedio - 2) = "Other Adjust"
        curMin(intPromedio - 2) = 0
        curMAx(intPromedio - 2) = 0
        'curAvg(intPromedio - 2) = Round(ScriptControl1.Eval(strFormulaVarios), Val(lvwGetValue(lvwContratos, "redondeo3")))
        curAvg(intPromedio - 2) = Val(Format(ScriptControl1.Eval(strFormulaVarios), "#0." & String(Val(lvwGetValue(lvwContratos, "redondeo3")), "#")))

        '----------------------------------------------------------------------------------
        '28/09/2004: el PNETPRICECON se calcula segun tabla contratosRangos
        '----------------------------------------------------------------------------------
        
        Dim rsRan As ADODB.Recordset
        Dim strFormula As String
        Dim sngValor, sngCOEF  As Single
        
        curAvgPrecioDesc = 0
        
        strSQL = "select * from contratosRangos where nombre = '" & strContratosRango & "'"
        Set rsRan = adoGetRS(strSQL)
                                
        'chequeo errores
        If Not lngAdoErrNum = -1 Then
          adoError
          Exit Sub
        End If
                                
        'recorro
        Do While Not rsRan.EOF
                    
          'le doy valor a formula, luego chequeo en que rango encaja
          strFormula = rsRan!Formula
          strFormula = Replace(strFormula, "AVG", curAvg(intPromedio - 6))
          strFormula = Replace(strFormula, "DISC", curAvg(intPromedio - 5) * -1)
          sngValor = CSng(ScriptControl1.Eval(strFormula))
          
          'si valor a testear esta dentro del rango
          If sngValor >= rsRan!valor1 And sngValor <= rsRan!valor2 Then
          
            'le doy valor al COEFICIENTE
            strFormula = rsRan!coeficiente
            strFormula = Replace(strFormula, "AJUAPI", curAvg(intPromedio - 3))
            strFormula = Replace(strFormula, "AJUVAR", curAvg(intPromedio - 2))
            strFormula = Replace(strFormula, "AVG", curAvg(intPromedio - 6))
            strFormula = Replace(strFormula, "DISC", curAvg(intPromedio - 5) * -1)
            sngCOEF = CSng(ScriptControl1.Eval(strFormula))
                    
            'le doy valor al IIBB
            strFormula = rsRan!coef_IIBB
            strFormula = Replace(strFormula, "AJUAPI", curAvg(intPromedio - 3))
            strFormula = Replace(strFormula, "AJUVAR", curAvg(intPromedio - 2))
            strFormula = Replace(strFormula, "AVG", curAvg(intPromedio - 6))
            strFormula = Replace(strFormula, "DISC", curAvg(intPromedio - 5) * -1)
            sngIIBB = CSng(ScriptControl1.Eval(strFormula))
                        
            Exit Do
                        
          End If
          
          
          rsRan.MoveNext
          
        Loop
                                
        'cierro rs
        rsRan.Close
       
        'calculo descuento2 -------------------------------------------------------------------------
        
        'get descuento
        strT = lvwGetValue(lvwContratos, "dto")
        
        'reemplazo variables por valores
        strT = Replace(strT, "AVG", curAvg(intPromedio - 6))
        strT = Replace(strT, "COEF", sngCOEF)
          
        'guardo discount
        strID(intPromedio - 5) = lvwGetValue(lvwContratos, "contrato")
        strInfo(intPromedio - 5) = "Discount"
        curMin(intPromedio - 5) = 0
        curMAx(intPromedio - 5) = 0
        curAvg(intPromedio - 5) = -Val(Format(Me.ScriptControl1.Eval(strT), "#0." & String(intRedondeo, "#")))
        
        'reemplazo variables por valores para formula price adjust
        strFormulaPrecio = Replace(strFormulaPrecio, "COEF", sngCOEF)
        strFormulaPrecio = Replace(strFormulaPrecio, "IIBB", sngIIBB)
        strFormulaPrecio = Replace(strFormulaPrecio, "AJUAPI", curAvg(intPromedio - 3))
        strFormulaPrecio = Replace(strFormulaPrecio, "AJUVAR", curAvg(intPromedio - 2))
        strFormulaPrecio = Replace(strFormulaPrecio, "AVG", curAvg(intPromedio - 6))
        strFormulaPrecio = Replace(strFormulaPrecio, "DISC", curAvg(intPromedio - 5) * -1)
        strFormulaPrecio = Replace(strFormulaPrecio, "APICLI", lvwGetValue(lvwEntregasCli, "napigravity"))
        strFormulaPrecio = Replace(strFormulaPrecio, "APIMAXCON", lvwGetValue(lvwContratos, "apimax"))
        strFormulaPrecio = Replace(strFormulaPrecio, "AZUMAXCON", lvwGetValue(lvwContratos, "azumax"))
        strFormulaPrecio = Replace(strFormulaPrecio, "AZUCLI", lvwGetValue(lvwEntregasCli, "azufre"))
        strFormulaPrecio = Replace(strFormulaPrecio, "REDONDEO", str(intRedondeo))
                
        'chequeo si en la formula vienen un IF
        If InStr(1, strFormulaPrecio, "IF") <> 0 Then
            
          'saco el string IF
          strFormulaPrecio = Replace(strFormulaPrecio, "IF", "")
          Dim strF As Variant
          strF = separateText(strFormulaPrecio)
          If ScriptControl1.Eval(strF(1)) Then
            strFormulaPrecio = strF(2)
          Else
            strFormulaPrecio = strF(3)
          End If
                      
        End If
        
        'guardo price adjust --------------------------------------------------------------------------
        strID(intPromedio - 4) = lvwGetValue(lvwContratos, "contrato")
        strInfo(intPromedio - 4) = "Price Adjust"
        curMin(intPromedio - 4) = 0
        curMAx(intPromedio - 4) = 0
        'curAvg(intPromedio - 4) = Format(ScriptControl1.Eval(strFormulaPrecio),, "." & String(Val(lvwGetValue(lvwContratos, "redondeo1")), "#"))
        curAvg(intPromedio - 4) = Val(Format(ScriptControl1.Eval(strFormulaPrecio), "#0." & String(Val(lvwGetValue(lvwContratos, "redondeo1")), "#")))
        
        ' guardo y calculo Net Price
        strID(intPromedio - 1) = lvwGetValue(lvwContratos, "contrato")
        strInfo(intPromedio - 1) = "Total Price"
        curMin(intPromedio - 1) = 0
        curMAx(intPromedio - 1) = 0
        'curAvg(intPromedio - 1) = Round(curAvg(intPromedio - 6) + curAvg(intPromedio - 5) + curAvg(intPromedio - 4) + curAvg(intPromedio - 3) + curAvg(intPromedio - 2), intRedondeo)
        curAvg(intPromedio - 1) = Val(Format(curAvg(intPromedio - 6) + curAvg(intPromedio - 5) + curAvg(intPromedio - 4) + curAvg(intPromedio - 3) + curAvg(intPromedio - 2), "#0." & String(intRedondeo, "#")))
      
      End If
     
      ' Determina IVA segun Cliente
      Select Case strIVACliente
      
      Case "Responsable Inscripto"
        intRes = lvwSetValue(lvwContratos, "ivaPje", Format(getParam("ivaResponsable"), "#0.00"))
      
      Case "Responsable No Inscripto"
        intRes = lvwSetValue(lvwContratos, "ivaPje", Format(getParam("ivaNoInscripto"), "#0.00"))
      
      Case "Consumidor Final"
        intRes = lvwSetValue(lvwContratos, "ivaPje", Format(getParam("ivaFinal"), "#0.00"))
      
      Case "Exento"
        intRes = lvwSetValue(lvwContratos, "ivaPje", Format(getParam("ivaExento"), "#0.00"))
      
      Case "Exportación"
        intRes = lvwSetValue(lvwContratos, "ivaPje", Format(getParam("ivaExportacion"), "#0.00"))
      
      End Select
      
      
      'CHECK si tiene percepcion Rg3337
      If strRg3337 = "Si" Then
        Me.txtRg3337Pje = strRg3337
        intRes = lvwSetValue(lvwContratos, "Rg3337Pje", Format(getParam("ivaRg3337"), "#0.00"))
      Else
        intRes = lvwSetValue(lvwContratos, "Rg3337Pje", Format(0, "#0.00"))
      End If
      
      
      'SAVE IIBB Porcentaje en listView
      intRes = lvwSetValue(lvwContratos, "IIBB_Pje", Format(sngIngBrutos, "#0.00"))
      
      
      
      ' calculo totales para cada contrato
      intRes = lvwSetValue(lvwContratos, "wti", Round(curAvg(intPromedio - 1), 3))
      If Me.cboBase = "Barrels" Then
        intRes = lvwSetValue(lvwContratos, "subtotal", Round(curAvg(intPromedio - 1) * Val(lvwGetValue(lvwContratos, "barrelss")), 2))
      Else
        intRes = lvwSetValue(lvwContratos, "subtotal", Round(curAvg(intPromedio - 1) * Val(lvwGetValue(lvwContratos, "volss15")), 2))
      End If
      
      'CALC y SAVE impuestos en List View
      intRes = lvwSetValue(lvwContratos, "ivaValor", Round(Val(lvwGetValue(lvwContratos, "subtotal")) * Val(lvwGetValue(lvwContratos, "ivaPje")) / 100, 2))
      intRes = lvwSetValue(lvwContratos, "rg3337Valor", Round(Val(lvwGetValue(lvwContratos, "subtotal")) * Val(lvwGetValue(lvwContratos, "rg3337Pje")) / 100, 2))
      intRes = lvwSetValue(lvwContratos, "IIBB_Valor", Round(Val(lvwGetValue(lvwContratos, "subtotal")) * Val(lvwGetValue(lvwContratos, "IIBB_Pje")) / 100, 2))
      
      'CALC y SAVE Total en ListView
      intRes = lvwSetValue(lvwContratos, "total", Round(Val(lvwGetValue(lvwContratos, "subtotal")) + Val(lvwGetValue(lvwContratos, "ivaValor")) + Val(lvwGetValue(lvwContratos, "rg3337Valor")) + Val(lvwGetValue(lvwContratos, "IIBB_Valor")), 2))
   
   
      'CALC subtotal de factura
      Me.txtSubtotal = Format(Val(Me.txtSubtotal) + Val(lvwGetValue(lvwContratos, "subtotal")), "########0.00")
      
      'GET impuestos Porcentaje y ASSIGN en textBox
      Me.txtIvaPje = Format(Val(lvwGetValue(lvwContratos, "ivaPje")), "########0.00")
      Me.txtRg3337Pje = Format(Val(lvwGetValue(lvwContratos, "Rg3337Pje")), "########0.00")
      Me.txtIIBBPje = Format(Val(lvwGetValue(lvwContratos, "IIBB_Pje")), "########0.00")
      
      'GET impuestos
      Me.txtIvaImp = Format(Val(Me.txtIvaImp) + Val(lvwGetValue(lvwContratos, "ivaValor")), "########0.00")
      Me.txtRg3337Imp = Format(Val(Me.txtRg3337Imp) + Val(lvwGetValue(lvwContratos, "Rg3337Valor")), "########0.00")
      Me.txtIIBBImp = Format(Val(Me.txtIIBBImp) + Val(lvwGetValue(lvwContratos, "IIBB_Valor")), "########0.00")
      
      'CALC total de factura
      Me.txtTotal = Format(Val(Me.txtTotal) + Val(lvwGetValue(lvwContratos, "total")), "########0.00")
   
   End If
    
  Next

  'Armo Concepto Venta para Oil
  'If lvwGetValue(lvwEntregasCli, "entregacliTipo") = "Barco" Then
  '  strConceptoVenta = ""
  '  strConceptoVenta = strConceptoVenta & "Por la venta de petroleo crudo según el siguiente detalle:" & vbCrLf & vbCrLf
  '  strConceptoVenta = strConceptoVenta & "Petroleo Crudo Tipo:    " & lvwGetValue(lvwEntregasCli, "tipoOil") & vbCrLf
  '  strConceptoVenta = strConceptoVenta & "B/T:                    " & lvwGetValue(lvwEntregasCli, "barco") & vbCrLf
  '  strConceptoVenta = strConceptoVenta & "Fecha B/L:              " & lvwGetValue(lvwEntregasCli, "fecha") & vbCrLf
  '  strConceptoVenta = strConceptoVenta & "Inspector:              " & lvwGetValue(lvwEntregasCli, "inspeccion") & vbCrLf
  '  strConceptoVenta = strConceptoVenta & "Api:                    " & Format(Val(lvwGetValue(lvwEntregasCli, "napigravity")), "##.000") & vbCrLf
  '  strConceptoVenta = strConceptoVenta & "m3 a 15°C Seco-Seco:    " & Format(Val(lvwGetValue(lvwEntregasCli, "ngsv")), "###,###.000")
  'Else
  '  strConceptoVenta = ""
  '  strConceptoVenta = strConceptoVenta & "Por la venta de petroleo crudo según siguiente detalle:" & vbCrLf & vbCrLf
  '  strConceptoVenta = strConceptoVenta & "Petroleo Crudo Tipo:    " & lvwGetValue(lvwEntregasCli, "tipoOil") & vbCrLf
  '  strConceptoVenta = strConceptoVenta & "Certificado Nro:        " & lvwGetValue(lvwEntregasCli, "entreganro") & vbCrLf
  '  strConceptoVenta = strConceptoVenta & "Fecha de Entrega:       " & "Desde el " & lvwGetValue(lvwEntregasCli, "certifDesde") & " hasta el " & lvwGetValue(lvwEntregasCli, "certifHasta") & vbCrLf
  '  strConceptoVenta = strConceptoVenta & "m3 a 15°C Seco-Seco:    " & Format(Val(lvwGetValue(lvwEntregasCli, "ngsv")), "###,###.000")
  'End If
  'MsgBox (strConceptoVenta)
  
  ' Armo Concepto Soporte para Oil
  ' empresa , cliente, factura
  strConceptoSoporte = strConceptoSoporte & "Compañía   : " & Me.cboEmpresa.List(Me.cboEmpresa.ListIndex) & vbCrLf
  strConceptoSoporte = strConceptoSoporte & "Cliente    : " & Me.cboCliente.List(Me.cboCliente.ListIndex) & vbCrLf
  strConceptoSoporte = strConceptoSoporte & "Comprobante: " & Me.txtFactura & vbCrLf & vbCrLf
  strConceptoSoporte = strConceptoSoporte & "Tipo Cambio: " & Me.txtTipoCambio & vbCrLf & vbCrLf

  ' embarque
  strConceptoSoporte = strConceptoSoporte & "Barco         : " & lvwGetValue(lvwEntregasCli, "barco") & vbCrLf
  strConceptoSoporte = strConceptoSoporte & "Fecha Entrega : " & lvwGetValue(lvwEntregasCli, "fecha") & vbCrLf
  strConceptoSoporte = strConceptoSoporte & "°API          : " & lvwGetValue(lvwEntregasCli, "Napigravity") & vbCrLf & vbCrLf
  
  ' encabezado de contrato
  Dim strMoneda, strTipoVol, strPrecio As String
  
  strMoneda = Me.cboTipoMoneda
  strTipoVol = IIf(cboBase.List(cboBase.ListIndex) = "M3", "m3 ", "Bbl")
  strPrecio = strMoneda & "/" & strTipoVol
  
  strConceptoSoporte = strConceptoSoporte & " Cnto" & Space(2) & Format(strPrecio, "@@@@@@") & Space(2) & "      Mts 15°" & Space(2) & "   Mts 15.56°" & Space(2) & "     Barriles" & Space(2) & "   Subtotal" & Space(2) & "         Iva" & Space(2) & "       Total" & vbCrLf
  
  ' contratos
  intCuenta = 0
  sngTotal15 = 0
  sngTotal1556 = 0
  sngTotalBarrels = 0
  
  ' recorro lvw contratos
  For intCuenta = 1 To lvwContratos.ListItems.Count
    
    ' trabajo con los que tienen tilde
    If lvwContratos.ListItems(intCuenta).Checked Then
      
      ' puntero arriba del contrato
      lvwContratos.ListItems(intCuenta).Selected = True
      
      strConceptoSoporte = strConceptoSoporte & Format(lvwGetValue(lvwContratos, "contrato"), "@@@@@") & Space(2) & Format(Format(Val(lvwGetValue(lvwContratos, "wti")), "##0.000"), "@@@@@@@") & Space(2) & Format(Format(Val(lvwGetValue(lvwContratos, "volss15")), "###,###,##0.000"), "@@@@@@@@@@@@@") & Space(2) & Format(Format(Val(lvwGetValue(lvwContratos, "vol1556")), "###,###,##0.000"), "@@@@@@@@@@@@@") & Space(2) & Format(Format(Val(lvwGetValue(lvwContratos, "barrelss")), "###,###,##0.000"), "@@@@@@@@@@@@@") & " " & Format(Format(Val(lvwGetValue(lvwContratos, "subtotal")), "###,###,##0.00"), "@@@@@@@@@@@@") & " " & Format(Format(Val(lvwGetValue(lvwContratos, "ivaValor")), "###,###,##0.00"), "@@@@@@@@@@@@") & " " & Format(Format(Val(lvwGetValue(lvwContratos, "total")), "###,###,##0.00"), "@@@@@@@@@@@@") & vbCrLf
      sngTotal15 = sngTotal15 + Val(lvwGetValue(lvwContratos, "volss15"))
      sngTotal1556 = sngTotal1556 + CDbl(lvwGetValue(lvwContratos, "vol1556"))
      sngTotalBarrels = sngTotalBarrels + CDbl(lvwGetValue(lvwContratos, "barrelss"))
    
    End If
  Next
  
  ' totales
  strConceptoSoporte = strConceptoSoporte & Space(16) & "--------------------------------------------------------------------------------" & vbCrLf
  strConceptoSoporte = strConceptoSoporte & Space(16) & Format(Format(sngTotal15, "###,###,##0.000"), "@@@@@@@@@@@@@") & Space(2) & Format(Format(sngTotal1556, "###,###,###.000"), "@@@@@@@@@@@@@") & Space(2) & Format(Format(sngTotalBarrels, "###,###,###.000"), "@@@@@@@@@@@@@") & Space(2) & Format(Format(Me.txtSubtotal, "###,###,###.00"), "@@@@@@@@@@@@") & " " & Format(Format(Me.txtIvaImp, "###,###,###.00"), "@@@@@@@@@@@@@") & Space(2) & Format(Format(Me.txtTotal, "###,###,###.00"), "@@@@@@@@@@@@") & vbCrLf & vbCrLf

  ' calculos
  For intRes = 0 To intPromedio
    strConceptoSoporte = strConceptoSoporte & Format(strID(intRes), "@@@@@") & vbTab & Format(strInfo(intRes), "!@@@@@@@@@@@@@") & vbTab & Format(Format(curMin(intRes), "##.###"), "@@@@@@") & vbTab & Format(Format(curMAx(intRes), "##.###"), "@@@@@@") & vbTab & Format(Format(curAvg(intRes), "##.000000"), "@@@@@@@@@") & vbCrLf
  Next

  'Armo Concepto Venta para Oil
  If lvwGetValue(lvwEntregasCli, "entregacliTipo") = "Barco" Then
    
    strConceptoVenta = ""
        
    'si exportaciones
    If cboOperacion = "Exp" Then
      strConceptoVenta = strConceptoVenta & "Crude oil sale under the following conditions:" & vbCrLf & vbCrLf
      strConceptoVenta = strConceptoVenta & "Crude Oil Type:         " & lvwGetValue(lvwEntregasCli, "tipoOil") & vbCrLf
      strConceptoVenta = strConceptoVenta & "M/T:                    " & lvwGetValue(lvwEntregasCli, "barco") & vbCrLf
      strConceptoVenta = strConceptoVenta & "B/L Date:               " & lvwGetValue(lvwEntregasCli, "fecha") & vbCrLf
      strConceptoVenta = strConceptoVenta & "Inspector:              " & lvwGetValue(lvwEntregasCli, "inspeccion") & vbCrLf
      strConceptoVenta = strConceptoVenta & "Api:                    " & Format(Val(lvwGetValue(lvwEntregasCli, "napigravity")), "##.000") & vbCrLf
      strConceptoVenta = strConceptoVenta & "m3 15°C:                " & Format(sngTotal15, "###,###.000")
    Else
      strConceptoVenta = strConceptoVenta & "Por la venta de petroleo crudo según el siguiente detalle:" & vbCrLf & vbCrLf
      strConceptoVenta = strConceptoVenta & "Petroleo Crudo Tipo:    " & lvwGetValue(lvwEntregasCli, "tipoOil") & vbCrLf
      strConceptoVenta = strConceptoVenta & "B/T:                    " & lvwGetValue(lvwEntregasCli, "barco") & vbCrLf
      strConceptoVenta = strConceptoVenta & "Fecha B/L:              " & lvwGetValue(lvwEntregasCli, "fecha") & vbCrLf
      strConceptoVenta = strConceptoVenta & "Inspector:              " & lvwGetValue(lvwEntregasCli, "inspeccion") & vbCrLf
      strConceptoVenta = strConceptoVenta & "Api:                    " & Format(Val(lvwGetValue(lvwEntregasCli, "napigravity")), "##.000") & vbCrLf
      strConceptoVenta = strConceptoVenta & "m3 15°C Seco-Seco:      " & Format(sngTotal15, "###,###.000")
    End If
      
  Else
    strConceptoVenta = ""
    
    'si exportaciones
    If cboOperacion = "Exp" Then
      strConceptoVenta = strConceptoVenta & "Crude oil sale under the following conditions:" & vbCrLf & vbCrLf
      strConceptoVenta = strConceptoVenta & "Crude Oil Type:        " & lvwGetValue(lvwEntregasCli, "tipoOil") & vbCrLf
      strConceptoVenta = strConceptoVenta & "Certificate Nbr:       " & lvwGetValue(lvwEntregasCli, "entreganro") & vbCrLf
      strConceptoVenta = strConceptoVenta & "B/L Date:              " & "Desde el " & lvwGetValue(lvwEntregasCli, "certifDesde") & " hasta el " & lvwGetValue(lvwEntregasCli, "certifHasta") & vbCrLf
      strConceptoVenta = strConceptoVenta & "m3 15°C:               " & Format(sngTotal15, "###,###.000")
    Else
      strConceptoVenta = strConceptoVenta & "Por la venta de petroleo crudo según el siguiente detale:" & vbCrLf & vbCrLf
      strConceptoVenta = strConceptoVenta & "Tipo de Crudo:         " & lvwGetValue(lvwEntregasCli, "tipoOil") & vbCrLf
      strConceptoVenta = strConceptoVenta & "Certificado Nro:       " & lvwGetValue(lvwEntregasCli, "entreganro") & vbCrLf
      strConceptoVenta = strConceptoVenta & "Fecha de Entrega:      " & "Desde el " & lvwGetValue(lvwEntregasCli, "certifDesde") & " hasta el " & lvwGetValue(lvwEntregasCli, "certifHasta") & vbCrLf
      strConceptoVenta = strConceptoVenta & "m3 15°C Seco-Seco:     " & Format(sngTotal15, "###,###.000")
    End If
    
  End If
  
  MsgBox (strConceptoVenta)

  ' habilito Guardar y Soporte
  
  Me.cmdGuardar.Enabled = True
  Me.cmdSoporte.Enabled = True
  
  Exit Sub
  
controlError:
  
 intRes = MsgBox(Err.Number & " - " & Err.Description, vbCritical + vbOKOnly, "atención...")
 Exit Sub
  
End Sub

Private Sub cmdGuardar_Click()
  Dim intUltimaVenta, intItem  As Integer
  Dim strConcepto As String
  Dim curTotalBarrelsVendido As Currency
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
  'chequeo que el numero de factura ingresado este libre
'  Set rs = adoGetRS("SELECT factura FROM ViewVentas where Factura = '" & Me.txtFactura & "'")
'  If Not rs.EOF And Not IsNull(rs!factura) Then
'    intRes = MsgBox("El numero de comprobante ingresado ya fue utilizado, debe cambiarlo para poder guardar el comprobante.", vbApplicationModal + vbInformation + vbOKOnly, "informacion...")
'    rs.Close
'    Exit Sub
'  End If
'  rs.Close
  
  
  
  ' guardo ventas
  strSQL = "EXEC spVentasInsert " & _
           "'" & Me.txtFactura & "'," & _
           "'" & dateToIso(Me.txtFecha) & "'," & _
           "'" & dateToIso(Me.txtContable) & "'," & _
           Val(Me.cboEmpresa.ItemData(Me.cboEmpresa.ListIndex)) & "," & _
           "'" & Me.cboOperacion.List(Me.cboOperacion.ListIndex) & "'," & _
           "'" & Left(Me.cboTipoComprobante.List(Me.cboTipoComprobante.ListIndex), 3) & "'," & _
           Val(Me.cboTipoMoneda.ItemData(Me.cboTipoMoneda.ListIndex)) & "," & _
           "'" & dateToIso(Me.txtFechaVencimiento) & "'," & _
           Val(Me.cboCliente.ItemData(Me.cboCliente.ListIndex)) & "," & Val(Me.txt_IDprovincia_Ent) & "," & _
           Val(lvwGetValue(lvwEntregasCli, "entregaCli")) & "," & _
           Val(Me.cboCuentaBancaria.ItemData(Me.cboCuentaBancaria.ListIndex)) & "," & _
           Val(Me.cboCondicion.ItemData(Me.cboCondicion.ListIndex)) & "," & _
           Val(intTerminal) & "," & _
           "'" & Me.cboBase & "'," & _
           "''" & "," & _
           Val(Me.txtSubtotal) & "," & _
           Val(Me.txtIvaPje) & "," & Val(Me.txtIvaImp) & "," & _
           Val(Me.txtRg3337Pje) & "," & Val(Me.txtRg3337Imp) & "," & _
           Val(Me.txtIIBBPje) & "," & Val(Me.txtIIBBImp) & "," & _
           Val(Me.txtTotal) & "," & _
           Val(Me.txtTipoCambio) & "," & _
           Val(Me.cboCotizacion.ItemData(Me.cboCotizacion.ListIndex)) & "," & _
           "'" & strConceptoVenta & "'," & _
           "'" & strConceptoSoporte & "','" & Replace(Me.txt_Libre, "'", Chr(34)) & "'"
           
           'En la linea de arriba, en el texto libre, reemplazo comilla simple, por doble, para que no de error.
            
  intResul = adoExecSQL(strSQL)
  
  'chequeo errores
  If Not lngAdoErrNum = -1 Then
    adoError
    Exit Sub
  End If

  ' guardo VentasDetalle
  intItem = 0
  curTotalBarrelsVendido = 0
  For intCuenta = 1 To lvwContratos.ListItems.Count
    
    ' recorro contratos
    If lvwContratos.ListItems(intCuenta).Checked Then
    
      ' ubico puntero en cada contrato
      lvwContratos.ListItems(intCuenta).Selected = True
    
      ' acumulador de numero de item
      intItem = intItem + 1
      
      strSQL = "EXEC spVentasDetalleInsert " & _
               "'" & Me.txtFactura & "'," & _
                intItem & "," & _
                Val(lvwGetValue(lvwContratos, "contrato")) & "," & _
                "'" & dateToIso(lvwGetValue(lvwContratos, "entregacli")) & "'," & _
                Val(lvwGetValue(lvwContratos, "volss15")) & "," & _
                Val(lvwGetValue(lvwContratos, "vol1556")) & "," & _
                Val(lvwGetValue(lvwContratos, "barrelss")) & "," & _
                Val(lvwGetValue(lvwContratos, "wti")) & "," & _
                Val(lvwGetValue(lvwContratos, "subtotal")) & "," & _
                "'" & strConcepto & "'," & _
                IIf(Val(lvwGetValue(lvwContratos, "ivaPje")) <> 0, 1, 0) & "," & _
                "'OIL'," & _
                IIf(Me.cboBase = "M3", 1, 2)
      
      intResul = adoExecSQL(strSQL)
      
      'chequeo errores
      If Not lngAdoErrNum = -1 Then
        adoError
        Exit Sub
      End If
    
      ' Actualizo volumen vendido en cada contrato
      strSQL = "EXEC spVentasContratoUpdate " & Val(lvwGetValue(lvwContratos, "contrato")) & "," & _
               Val(lvwGetValue(lvwContratos, "barrelss"))
      
      intResul = adoExecSQL(strSQL)
      
      'chequeo errores
      If Not lngAdoErrNum = -1 Then
        adoError
        Exit Sub
      End If
      
      
      ' Sumo volumen vendido en cada contrato
      curTotalBarrelsVendido = curTotalBarrelsVendido + Val(lvwGetValue(lvwContratos, "barrelss"))
      
    End If
    
  Next
  
  ' Actualizo volumen vendido en Ent regasClientes
  strSQL = "EXEC spVentasEntregaCliUpdate " & Val(lvwGetValue(lvwEntregasCli, "entregaCli")) & "," & _
            curTotalBarrelsVendido
  
    intRes = adoExecSQL(strSQL)
    
    'chequeo errores
    If Not lngAdoErrNum = -1 Then
        adoError
        Exit Sub
    End If
  
  
    'UPDATE numero de comprobante
    strSQL = "EXEC Empresas_Puntos_Venta_sp " & Me.cboEmpresa.ItemData(cboEmpresa.ListIndex) & "," & _
                                                "'" & UCase(Left(Me.cboOperacion, 3)) & "'," & _
                                                "'" & UCase(Left(Me.cboTipoComprobante, 3)) & "'," & _
                                                Val(Mid(Me.txtFactura, 7, 8))
    
    intResul = adoExecSQL(strSQL)
    
    'chequeo errores
    If Not lngAdoErrNum = -1 Then
        adoError
        Exit Sub
    End If
  
  
  
  ' Oculto formulario
  blnAceptar = True
  blnCancelar = False
  Me.Hide

End Sub

Private Sub cmdNewCondicion_Click()
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

Private Sub cmdNewFormaPago_Click()
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

Private Sub cmdNewTipoMoneda_Click()
  Dim strStore, strView, strDato As String

  strStore = "spMonedasInsert"
  strView = "SELECT * FROM ViewMonedas"
  strDato = ComboBoxAddItem(Me, cboTipoMoneda, "@50", strStore, strView)

End Sub

Private Sub cmdSalir_Click()
  
  blnAceptar = False
  blnCancelar = True
  Unload Me

End Sub

Private Sub cmdSoporte_Click()

  Load frmVentasSoporte
  frmVentasSoporte.txtSoporte = strConceptoSoporte
  frmVentasSoporte.Show 1

End Sub

Private Sub Form_Load()
  
  strSQL = "SELECT * FROM ViewEmpresas"
  intRes = ComboBoxRefresh(cboEmpresa, strSQL)
  
  strSQL = "SELECT * FROM ViewVentasClientes"
  intRes = ComboBoxRefresh(cboCliente, strSQL)

  strSQL = "SELECT * FROM ViewMonedas"
  intRes = ComboBoxRefresh(cboTipoMoneda, strSQL)

  strSQL = "SELECT * FROM ViewCuentasBancarias"
  intRes = ComboBoxRefresh(cboCuentaBancaria, strSQL)

  strSQL = "SELECT * FROM ViewCondiciones"
  intRes = ComboBoxRefresh(cboCondicion, strSQL)

  strSQL = "SELECT * FROM cotizacionesTexto_vw"
  intRes = ComboBoxRefresh(cboCotizacion, strSQL)

  ' toma parametro de cotizacion
  txtTipoCambio = Format(getParam("cotizU$S"), "#0.000")

    '   POR DEFECTO fecha del dìa para facturar
    Me.txtFecha = DateValue(Now)

End Sub

Public Function lvwBuscaEntregasCliContratos()
  Dim strAux As String
  
  If cboEmpresa.ListIndex <> -1 And cboCliente.ListIndex <> -1 Then
  
    ' armo SQL EntregasCli
    ' cambio apariencia y llena lista con datos
    ' oculto columnas
    ' armo SQL Contratos
    ' cambio apariencia y llena lista con datos
    ' oculto columnas
  
    strSQL = "SELECT * FROM ViewVentasEntregasCli " & _
             "WHERE Empresa = " & cboEmpresa.ItemData(cboEmpresa.ListIndex) & " AND " & _
             "Cliente = " & cboCliente.ItemData(cboCliente.ListIndex) & " AND " & _
             "Vendido < NBarrels60" & " AND " & _
             "fecha >= '20030101'"
           
    ' guardo nombre tabla actual
    strAux = strTableNameActual
    
    ' paso nombre de vista entregasClientes para buscar en ini ancho de columnas
    strTableNameActual = "ViewVentasEntregasCli"
           
    intRes = ListViewAppearanceChange(lvwEntregasCli)
    intRes = ListViewRefresh(lvwEntregasCli, strSQL)
  
    ' ocultar columnas
    intRes = lvwHideColumn(lvwEntregasCli, "legajo")
    intRes = lvwHideColumn(lvwEntregasCli, "density")
    intRes = lvwHideColumn(lvwEntregasCli, "density15")
    intRes = lvwHideColumn(lvwEntregasCli, "empresa")
    intRes = lvwHideColumn(lvwEntregasCli, "cliente")
    intRes = lvwHideColumn(lvwEntregasCli, "certificado")
    intRes = lvwHideColumn(lvwEntregasCli, "inspeccion")
    intRes = lvwHideColumn(lvwEntregasCli, "terminalID")
    intRes = lvwHideColumn(lvwEntregasCli, "tipoOil")
    intRes = lvwHideColumn(lvwEntregasCli, "entregacliTipo")
    intRes = lvwHideColumn(lvwEntregasCli, "certifDesde")
    intRes = lvwHideColumn(lvwEntregasCli, "certifHasta")
    intRes = lvwHideColumn(lvwEntregasCli, "entregaNro")
  
    strSQL = "SELECT * FROM ViewVentasContratos " & _
             "WHERE Empresa = " & cboEmpresa.ItemData(cboEmpresa.ListIndex) & " AND " & _
             "Cliente = " & cboCliente.ItemData(cboCliente.ListIndex) & " AND " & _
             "Vendido < barrels60"

'hasta el 02/06/2004, tomaba esto: "Vendido < ((datediff(day,Desde,Hasta)+1)*barrels60)"
           
    ' paso nombre de vista Contratos para buscar en ini ancho de columnas
    strTableNameActual = "ViewVentasContratos"
    
    intRes = ListViewAppearanceChange(lvwContratos)
    intRes = ListViewRefresh(lvwContratos, strSQL)
    
    ' ocultar columnas
    intRes = lvwHideColumn(lvwContratos, "m315")
    intRes = lvwHideColumn(lvwContratos, "m31556")
    intRes = lvwHideColumn(lvwContratos, "barrels60")
    intRes = lvwHideColumn(lvwContratos, "apimin")
    intRes = lvwHideColumn(lvwContratos, "apimax")
    intRes = lvwHideColumn(lvwContratos, "azumin")
    intRes = lvwHideColumn(lvwContratos, "azumax")
    intRes = lvwHideColumn(lvwContratos, "tipocalculo")
    intRes = lvwHideColumn(lvwContratos, "mespromedio")
    intRes = lvwHideColumn(lvwContratos, "preciodesde")
    intRes = lvwHideColumn(lvwContratos, "preciohasta")
    intRes = lvwHideColumn(lvwContratos, "diasprevios")
    intRes = lvwHideColumn(lvwContratos, "diasposteriores")
    intRes = lvwHideColumn(lvwContratos, "incluyeEntrega")
    intRes = lvwHideColumn(lvwContratos, "empresa")
    intRes = lvwHideColumn(lvwContratos, "cliente")
    intRes = lvwHideColumn(lvwContratos, "preciotipo")
    intRes = lvwHideColumn(lvwContratos, "tablaRangos")
    intRes = lvwHideColumn(lvwContratos, "formulaajuprecio")
    intRes = lvwHideColumn(lvwContratos, "formulaajuapi")
    intRes = lvwHideColumn(lvwContratos, "formulaajuvarios")
    intRes = lvwHideColumn(lvwContratos, "redondeo1")
    intRes = lvwHideColumn(lvwContratos, "redondeo2")
    intRes = lvwHideColumn(lvwContratos, "redondeo3")
    intRes = lvwHideColumn(lvwContratos, "redondeo4")
  
  
    ' recupero nombre de tabla actual
    strTableNameActual = strAux
  
  End If

End Function


Private Sub lvwContratos_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

  If Button = 2 Then
    PopupMenu mnuSeleccionar
  End If

End Sub

Private Sub mnuEmbarque_Click()

  Dim intRes As Integer
  Dim strFecha, strFechaAUX, strFechaDesde, strFechaHasta, strAux As String
  Dim rsPrecios As ADODB.Recordset
  
  ' valido seleccion de contrato
  ' cargo formulario y le asigno valores de contrato
  ' busco 2 fechas con cotizacion para atras
  ' busco 2 fechas con cotizacion para adelante
  ' establezco las fechas desde hasta
  ' lleno list view con rango de fechas con cotizacion
  ' muestro formulario
  ' si se selecciono alguna fecha de embarque, la asigno
  
  If Not lvwContratos.SelectedItem.Checked And lvwContratos.SelectedItem.Selected Then
    intRes = MsgBox("Para poder modificar una fecha de embarque debe seleccionar un contrato.", vbInformation + vbOKOnly, "Información")
    Exit Sub
  End If
        
  If lvwGetValue(lvwContratos, "entregaCli") = "" Then
    intRes = MsgBox("El contrato que selecciono no tiene fecha de embarque, por lo tanto no se puede modificar.", vbInformation + vbOKOnly, "Información")
    Exit Sub
  End If
        
  Load frmVentasFechaEmbarque
    
  With frmVentasFechaEmbarque
  .txtFechaEntregaCli = lvwGetValue(lvwContratos, "entregaCli")
  .txtDiasPrevios = lvwGetValue(lvwContratos, "diasprevios")
  .txtDiasPosteriores = lvwGetValue(lvwContratos, "diasposteriores")
  .txtIncluyeEntregaCli = lvwGetValue(lvwContratos, "incluyeEntrega")
  .txtMesPromedio = lvwGetValue(lvwContratos, "mespromedio")
  End With
    
  strFecha = frmVentasFechaEmbarque.txtFechaEntregaCli
  strFechaAUX = dateToIso(str(CDate(strFecha)))
        
  strSQL = "SELECT TOP 2 * FROM ViewVentasPrecios " & _
           "WHERE Fecha <= '" & strFechaAUX & "' AND " & _
           "Operacion = '" & frmVentasInfo.cboOperacion.List(frmVentasInfo.cboOperacion.ListIndex) & "' AND " & _
           "PrecioTipoID = " & lvwGetValue(lvwContratos, "precioTipo") & " AND " & _
           "ValorMin > 0 " & _
           "ORDER BY Fecha DESC"
                   
  Set rsPrecios = adoGetRS(strSQL)  ' abro recordset
  If Not rsPrecios.EOF Then   ' Chequeo que exista registro
    rsPrecios.MoveLast
    strFechaDesde = dateToIso(str(rsPrecios!fecha))
  End If
  rsPrecios.Close
        
  strFechaAUX = dateToIso(str(CDate(strFecha)))
       
  strSQL = "SELECT TOP 2 * FROM ViewVentasPrecios " & _
           "WHERE Fecha >= '" & strFechaAUX & "' AND " & _
           "Operacion = '" & frmVentasInfo.cboOperacion.List(frmVentasInfo.cboOperacion.ListIndex) & "' AND " & _
           "PrecioTipoID = " & lvwGetValue(lvwContratos, "precioTipo") & " AND " & _
           "ValorMin > 0 " & _
           "ORDER BY Fecha"
                   
  Set rsPrecios = adoGetRS(strSQL)
  If Not rsPrecios.EOF Then
    rsPrecios.MoveLast   ' pongo puntero en fecha de embarque
    strFechaHasta = dateToIso(str(rsPrecios!fecha))
  End If
  rsPrecios.Close
      
  strSQL = "SELECT * FROM ViewVentasPrecios " & _
           "WHERE Fecha BETWEEN '" & strFechaDesde & "' AND '" & strFechaHasta & "' AND " & _
           "Operacion = '" & frmVentasInfo.cboOperacion.List(frmVentasInfo.cboOperacion.ListIndex) & "' AND " & _
           "PrecioTipoID = " & lvwGetValue(lvwContratos, "precioTipo") & " AND " & _
           "ValorMin > 0 " & _
           "ORDER BY Fecha"
      
  ' guardo nombre de tabla actual
  strAux = strTableNameActual
  strTableNameActual = "ViewVentasPrecios"
      
  ' refresh info
  intRes = ListViewAppearanceChange(frmVentasFechaEmbarque.lvwPrecios)
  intRes = ListViewRefresh(frmVentasFechaEmbarque.lvwPrecios, strSQL)
          
  ' recupero nombre de tabla
  strTableNameActual = strAux
          
  ' oculto columnas
  intRes = lvwHideColumn(frmVentasFechaEmbarque.lvwPrecios, "operacion")
  intRes = lvwHideColumn(frmVentasFechaEmbarque.lvwPrecios, "preciotipoid")
          
  ' abro form
  frmVentasFechaEmbarque.Show vbModal

  If blnAceptar Then
  
    If frmVentasFechaEmbarque.lvwPrecios.SelectedItem.Selected Then
      ' le asigno la fecha seleccionada al contrato
      intRes = lvwSetValue(lvwContratos, "entregacli", frmVentasFechaEmbarque.lvwPrecios.SelectedItem)
    End If
  
  End If

End Sub

Private Sub mnuPrecios_Click()
  
  'valido seleccion de contrato
  If Not lvwContratos.SelectedItem.Checked And lvwContratos.SelectedItem.Selected Then
    intRes = MsgBox("Para poder modificar una fecha de embarque debe seleccionar un contrato.", vbInformation + vbOKOnly, "Información")
    Exit Sub
  End If
  
  'valido que contrato sea por rango
  If lvwGetValue(lvwContratos, "TipoCalculo") <> "Rango" Then
    intRes = MsgBox("El contrato seleccionado no esta definido por rango de fechas.", vbInformation + vbOKOnly, "Información")
    Exit Sub
  End If
  
  'cargo frm
  Load VentasPreciosRangoFrm
  
  'paso las fechas
  VentasPreciosRangoFrm.txtDesde = strFechaDesde
  VentasPreciosRangoFrm.txtHasta = strFechaHasta
  
  'muestro
  VentasPreciosRangoFrm.Show vbModal
  
  If blnAceptar Then
    
    strFechaDesde = VentasPreciosRangoFrm.txtDesde
    strFechaHasta = VentasPreciosRangoFrm.txtHasta
  
  End If

End Sub

Private Sub mnuBarrels_Click()
  
  Dim intCuenta, intCantidad, intRes As Integer
  Dim strCual As String
  
  ' solo cambio el valor si esta tildado el contrato
  If Not lvwContratos.SelectedItem.Checked Then
    intRes = MsgBox("Para poder asignar el volumen, el contrato debe estar seleccionado.", vbInformation + vbOKOnly, "Información")
    Exit Sub
  End If
  
  'valido que se haya seleccionada cboBase M3 o Barrels
  If cboBase.ListIndex = -1 Then
    intRes = MsgBox("Para poder asignar el volumen, debe seleccionar facturacion en base a Barrels o M3.", vbInformation + vbOKOnly, "Información")
    Exit Sub
  End If
  
  ' cargo form
  Load frmVentasBarrelsUpdate
  
  'si ya se ingresaron los volumenes a los contratos se los muestro
  If Val(lvwGetValue(lvwContratos, "volss15")) <> 0 Then
  
    ' le paso valor actual de barrels
    frmVentasBarrelsUpdate.txtVolM315 = lvwGetValue(lvwContratos, "volss15")
    frmVentasBarrelsUpdate.txtVolM315.SelLength = Len(frmVentasBarrelsUpdate.txtVolM315)
    frmVentasBarrelsUpdate.txtVolM31556 = lvwGetValue(lvwContratos, "vol1556")
    frmVentasBarrelsUpdate.txtVolM31556.SelLength = Len(frmVentasBarrelsUpdate.txtVolM31556)
    frmVentasBarrelsUpdate.txtVolBarrels = lvwGetValue(lvwContratos, "barrelss")
    frmVentasBarrelsUpdate.txtVolBarrels.SelLength = Len(frmVentasBarrelsUpdate.txtVolBarrels)
  
  'si no les paso los volumenes del embarque seleccionado
  Else
    
    'verifico que el embarque seleccionado este con check
    If lvwEntregasCli.SelectedItem.Checked Then
      
      frmVentasBarrelsUpdate.txtVolM315 = lvwGetValue(lvwEntregasCli, "ngsv")
      frmVentasBarrelsUpdate.txtVolM315.SelLength = Len(frmVentasBarrelsUpdate.txtVolM315)
      frmVentasBarrelsUpdate.txtVolM31556 = lvwGetValue(lvwEntregasCli, "ncubmeters1556")
      frmVentasBarrelsUpdate.txtVolM31556.SelLength = Len(frmVentasBarrelsUpdate.txtVolM31556)
      frmVentasBarrelsUpdate.txtVolBarrels = lvwGetValue(lvwEntregasCli, "nbarrels60")
      frmVentasBarrelsUpdate.txtVolBarrels.SelLength = Len(frmVentasBarrelsUpdate.txtVolBarrels)
      
    Else
      intRes = MsgBox("El embarque seleccionado no esta con un tilde.", vbInformation + vbOKOnly, "Información")
      Exit Sub
    End If
  
  End If
  
  ' lo muestro modal
  frmVentasBarrelsUpdate.Show vbModal
  
  ' si acepto update
  If blnAceptar Then
    intRes = lvwSetValue(lvwContratos, "volss15", frmVentasBarrelsUpdate.txtVolM315)
    intRes = lvwSetValue(lvwContratos, "vol1556", frmVentasBarrelsUpdate.txtVolM31556)
    intRes = lvwSetValue(lvwContratos, "barrelss", frmVentasBarrelsUpdate.txtVolBarrels)
  End If
  
  ' descargo form
  Unload frmVentasBarrelsUpdate

End Sub

