VERSION 5.00
Begin VB.Form frmEntregasCliInfo 
   BackColor       =   &H80000018&
   Caption         =   "Entregas Clientes"
   ClientHeight    =   5235
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   12360
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5235
   ScaleWidth      =   12360
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox txtEntregaNro 
      BackColor       =   &H80000018&
      ForeColor       =   &H00808080&
      Height          =   285
      Left            =   1845
      TabIndex        =   11
      Top             =   3375
      Width           =   3525
   End
   Begin VB.TextBox txtTitulo 
      Appearance      =   0  'Flat
      BackColor       =   &H80000018&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      ForeColor       =   &H00808080&
      Height          =   285
      Index           =   0
      Left            =   270
      TabIndex        =   81
      Text            =   "Certificados Nro"
      Top             =   3375
      Width           =   1500
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H80000018&
      Caption         =   "Datos Generales"
      Height          =   4650
      Left            =   135
      TabIndex        =   63
      Top             =   90
      Width           =   5955
      Begin VB.TextBox txtTitulo 
         Appearance      =   0  'Flat
         BackColor       =   &H80000018&
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         ForeColor       =   &H00808080&
         Height          =   285
         Index           =   2
         Left            =   3510
         TabIndex        =   83
         Text            =   "Hasta"
         Top             =   2970
         Width           =   510
      End
      Begin VB.TextBox txtCerHasta 
         BackColor       =   &H80000018&
         ForeColor       =   &H00808080&
         Height          =   285
         Left            =   4095
         TabIndex        =   10
         Top             =   2970
         Width           =   1140
      End
      Begin VB.TextBox txtTitulo 
         Appearance      =   0  'Flat
         BackColor       =   &H80000018&
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         ForeColor       =   &H00808080&
         Height          =   285
         Index           =   1
         Left            =   135
         TabIndex        =   82
         Text            =   "Certificados Desde"
         Top             =   2970
         Width           =   1500
      End
      Begin VB.TextBox txtCerDesde 
         BackColor       =   &H80000018&
         ForeColor       =   &H00808080&
         Height          =   285
         Left            =   1710
         TabIndex        =   9
         Top             =   2970
         Width           =   1140
      End
      Begin VB.TextBox txtLegajo 
         BackColor       =   &H80000018&
         ForeColor       =   &H00808080&
         Height          =   285
         Left            =   1710
         TabIndex        =   0
         Top             =   225
         Width           =   3525
      End
      Begin VB.ComboBox cboEmpresa 
         Appearance      =   0  'Flat
         BackColor       =   &H80000018&
         ForeColor       =   &H00808080&
         Height          =   315
         Left            =   1710
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   3
         Top             =   1125
         Width           =   3525
      End
      Begin VB.ComboBox cboCliente 
         Appearance      =   0  'Flat
         BackColor       =   &H80000018&
         ForeColor       =   &H00808080&
         Height          =   315
         Left            =   1710
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   4
         Top             =   1440
         Width           =   3525
      End
      Begin VB.TextBox txtCertificado 
         BackColor       =   &H80000018&
         ForeColor       =   &H00808080&
         Height          =   285
         Left            =   1710
         TabIndex        =   6
         Top             =   2070
         Width           =   3525
      End
      Begin VB.ComboBox cboInspeccion 
         Appearance      =   0  'Flat
         BackColor       =   &H80000018&
         ForeColor       =   &H00808080&
         Height          =   315
         Left            =   1710
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   5
         Top             =   1755
         Width           =   3525
      End
      Begin VB.TextBox txtTitulo 
         Appearance      =   0  'Flat
         BackColor       =   &H80000018&
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         ForeColor       =   &H00808080&
         Height          =   285
         Index           =   47
         Left            =   135
         TabIndex        =   80
         Text            =   "Legajo"
         Top             =   225
         Width           =   1500
      End
      Begin VB.TextBox txtTitulo 
         Appearance      =   0  'Flat
         BackColor       =   &H80000018&
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         ForeColor       =   &H00808080&
         Height          =   285
         Index           =   46
         Left            =   135
         TabIndex        =   79
         Text            =   "Fecha"
         Top             =   810
         Width           =   1500
      End
      Begin VB.TextBox txtTitulo 
         Appearance      =   0  'Flat
         BackColor       =   &H80000018&
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         ForeColor       =   &H00808080&
         Height          =   285
         Index           =   33
         Left            =   135
         TabIndex        =   78
         Text            =   "Empresa"
         Top             =   1140
         Width           =   1500
      End
      Begin VB.TextBox txtTitulo 
         Appearance      =   0  'Flat
         BackColor       =   &H80000018&
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         ForeColor       =   &H00808080&
         Height          =   285
         Index           =   32
         Left            =   135
         TabIndex        =   77
         Text            =   "Cliente"
         Top             =   1440
         Width           =   1500
      End
      Begin VB.TextBox txtTitulo 
         Appearance      =   0  'Flat
         BackColor       =   &H80000018&
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         ForeColor       =   &H00808080&
         Height          =   285
         Index           =   31
         Left            =   135
         TabIndex        =   76
         Text            =   "Inspección"
         Top             =   1785
         Width           =   1500
      End
      Begin VB.TextBox txtTitulo 
         Appearance      =   0  'Flat
         BackColor       =   &H80000018&
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         ForeColor       =   &H00808080&
         Height          =   285
         Index           =   30
         Left            =   135
         TabIndex        =   75
         Text            =   "FechaCertificado"
         Top             =   2070
         Width           =   1500
      End
      Begin VB.TextBox txtTitulo 
         Appearance      =   0  'Flat
         BackColor       =   &H80000018&
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         ForeColor       =   &H00808080&
         Height          =   285
         Index           =   29
         Left            =   135
         TabIndex        =   74
         Text            =   "Barco"
         Top             =   2370
         Width           =   1500
      End
      Begin VB.TextBox txtTitulo 
         Appearance      =   0  'Flat
         BackColor       =   &H80000018&
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         ForeColor       =   &H00808080&
         Height          =   285
         Index           =   28
         Left            =   135
         TabIndex        =   73
         Text            =   "Provisionado"
         Top             =   3600
         Width           =   1500
      End
      Begin VB.TextBox txtFecha 
         BackColor       =   &H80000018&
         ForeColor       =   &H00808080&
         Height          =   285
         Left            =   1710
         TabIndex        =   2
         Top             =   810
         Width           =   3525
      End
      Begin VB.ComboBox cboBarco 
         Appearance      =   0  'Flat
         BackColor       =   &H80000018&
         ForeColor       =   &H00808080&
         Height          =   315
         Left            =   1710
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   7
         Top             =   2340
         Width           =   3525
      End
      Begin VB.ComboBox cboProvicionado 
         Appearance      =   0  'Flat
         BackColor       =   &H80000018&
         ForeColor       =   &H00808080&
         Height          =   315
         ItemData        =   "frmEntregasCliInfo.frx":0000
         Left            =   1710
         List            =   "frmEntregasCliInfo.frx":000A
         Style           =   2  'Dropdown List
         TabIndex        =   12
         Top             =   3600
         Width           =   3525
      End
      Begin VB.CommandButton cmdNewEmpresa 
         Caption         =   "New"
         Height          =   280
         Left            =   5265
         TabIndex        =   72
         Top             =   1125
         Width           =   510
      End
      Begin VB.CommandButton cmdNewInspeccion 
         Caption         =   "New"
         Height          =   280
         Left            =   5265
         TabIndex        =   71
         Top             =   1755
         Width           =   510
      End
      Begin VB.CommandButton cmdNewBarco 
         Caption         =   "New"
         Height          =   280
         Left            =   5265
         TabIndex        =   70
         Top             =   2340
         Width           =   510
      End
      Begin VB.CommandButton cmdNewEntregaCliTipo 
         Caption         =   "New"
         Height          =   280
         Left            =   5265
         TabIndex        =   69
         Top             =   495
         Width           =   510
      End
      Begin VB.TextBox txtTitulo 
         Appearance      =   0  'Flat
         BackColor       =   &H80000018&
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         ForeColor       =   &H00808080&
         Height          =   285
         Index           =   27
         Left            =   135
         TabIndex        =   68
         Text            =   "TipoEntrega"
         Top             =   540
         Width           =   1500
      End
      Begin VB.ComboBox cboEntregaCliTipo 
         Appearance      =   0  'Flat
         BackColor       =   &H80000018&
         ForeColor       =   &H00808080&
         Height          =   315
         Left            =   1710
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   495
         Width           =   3525
      End
      Begin VB.CommandButton cmdNewTerminal 
         Caption         =   "New"
         Height          =   280
         Left            =   5265
         TabIndex        =   67
         Top             =   2655
         Width           =   510
      End
      Begin VB.ComboBox cboTerminal 
         Appearance      =   0  'Flat
         BackColor       =   &H80000018&
         ForeColor       =   &H00808080&
         Height          =   315
         Left            =   1710
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   8
         Top             =   2655
         Width           =   3525
      End
      Begin VB.TextBox txtTitulo 
         Appearance      =   0  'Flat
         BackColor       =   &H80000018&
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         ForeColor       =   &H00808080&
         Height          =   285
         Index           =   26
         Left            =   135
         TabIndex        =   66
         Text            =   "Terminal"
         Top             =   2670
         Width           =   1500
      End
      Begin VB.TextBox txtTitulo 
         Appearance      =   0  'Flat
         BackColor       =   &H80000018&
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         ForeColor       =   &H00808080&
         Height          =   240
         Index           =   25
         Left            =   135
         TabIndex        =   65
         Text            =   "Azufre"
         Top             =   3915
         Width           =   1500
      End
      Begin VB.TextBox txtAzufre 
         BackColor       =   &H80000018&
         ForeColor       =   &H00808080&
         Height          =   285
         Left            =   1710
         TabIndex        =   13
         Top             =   3915
         Width           =   3525
      End
      Begin VB.TextBox txtTitulo 
         Appearance      =   0  'Flat
         BackColor       =   &H80000018&
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         ForeColor       =   &H00808080&
         Height          =   250
         Index           =   24
         Left            =   135
         TabIndex        =   64
         Text            =   "OtrosAjustes"
         Top             =   4230
         Width           =   1500
      End
      Begin VB.TextBox txtOtrosAjustes 
         BackColor       =   &H80000018&
         ForeColor       =   &H00808080&
         Height          =   285
         Left            =   1710
         TabIndex        =   14
         Top             =   4230
         Width           =   3525
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H80000018&
      Caption         =   "Neto"
      Height          =   4650
      Left            =   9135
      TabIndex        =   51
      Top             =   90
      Width           =   3120
      Begin VB.TextBox txtTitulo 
         Appearance      =   0  'Flat
         BackColor       =   &H80000018&
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         ForeColor       =   &H00808080&
         Height          =   220
         Index           =   21
         Left            =   135
         TabIndex        =   62
         Text            =   "Bsw Vol"
         Top             =   3075
         Width           =   1275
      End
      Begin VB.TextBox txtNBswCBM 
         BackColor       =   &H80000018&
         ForeColor       =   &H00808080&
         Height          =   285
         Left            =   1455
         TabIndex        =   38
         Text            =   "0"
         Top             =   3060
         Width           =   1500
      End
      Begin VB.TextBox txtTitulo 
         Appearance      =   0  'Flat
         BackColor       =   &H80000018&
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         ForeColor       =   &H00808080&
         Height          =   220
         Index           =   20
         Left            =   135
         TabIndex        =   61
         Text            =   "APIGravity"
         Top             =   2805
         Width           =   1275
      End
      Begin VB.TextBox txtNAPIGravity 
         BackColor       =   &H80000018&
         ForeColor       =   &H00808080&
         Height          =   285
         Left            =   1455
         TabIndex        =   37
         Text            =   "0"
         Top             =   2790
         Width           =   1500
      End
      Begin VB.TextBox txtTitulo 
         Appearance      =   0  'Flat
         BackColor       =   &H80000018&
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         ForeColor       =   &H00808080&
         Height          =   220
         Index           =   19
         Left            =   135
         TabIndex        =   60
         Text            =   "LongTons"
         Top             =   2535
         Width           =   1275
      End
      Begin VB.TextBox txtNLongTons 
         BackColor       =   &H80000018&
         ForeColor       =   &H00808080&
         Height          =   285
         Left            =   1455
         TabIndex        =   36
         Text            =   "0"
         Top             =   2500
         Width           =   1500
      End
      Begin VB.TextBox txtTitulo 
         Appearance      =   0  'Flat
         BackColor       =   &H80000018&
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         ForeColor       =   &H00808080&
         Height          =   220
         Index           =   17
         Left            =   135
         TabIndex        =   59
         Text            =   "Gallons at 60°F"
         Top             =   2265
         Width           =   1275
      End
      Begin VB.TextBox txtNGallons60 
         BackColor       =   &H80000018&
         ForeColor       =   &H00808080&
         Height          =   285
         Left            =   1455
         TabIndex        =   35
         Text            =   "0"
         Top             =   2220
         Width           =   1500
      End
      Begin VB.TextBox txtNDensity 
         BackColor       =   &H80000018&
         ForeColor       =   &H00808080&
         Height          =   285
         Left            =   1455
         TabIndex        =   31
         Text            =   "0"
         Top             =   1080
         Width           =   1500
      End
      Begin VB.TextBox txtTitulo 
         Appearance      =   0  'Flat
         BackColor       =   &H80000018&
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         ForeColor       =   &H00808080&
         Height          =   220
         Index           =   16
         Left            =   135
         TabIndex        =   58
         Text            =   "Density"
         Top             =   1140
         Width           =   1275
      End
      Begin VB.TextBox txtNGsv 
         BackColor       =   &H80000018&
         ForeColor       =   &H00808080&
         Height          =   285
         Left            =   1455
         TabIndex        =   30
         Text            =   "0"
         Top             =   790
         Width           =   1500
      End
      Begin VB.TextBox txtTitulo 
         Appearance      =   0  'Flat
         BackColor       =   &H80000018&
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         ForeColor       =   &H00808080&
         Height          =   220
         Index           =   13
         Left            =   135
         TabIndex        =   57
         Text            =   "Gsv (m3 at 15°C)"
         Top             =   855
         Width           =   1275
      End
      Begin VB.TextBox txtNBarrels60 
         BackColor       =   &H80000018&
         ForeColor       =   &H00808080&
         Height          =   285
         Left            =   1455
         TabIndex        =   34
         Text            =   "0"
         Top             =   1930
         Width           =   1500
      End
      Begin VB.TextBox txtNCubMeters1556 
         BackColor       =   &H80000018&
         ForeColor       =   &H00808080&
         Height          =   285
         Left            =   1455
         TabIndex        =   33
         Text            =   "0"
         Top             =   1650
         Width           =   1500
      End
      Begin VB.TextBox txtTitulo 
         Appearance      =   0  'Flat
         BackColor       =   &H80000018&
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         ForeColor       =   &H00808080&
         Height          =   220
         Index           =   12
         Left            =   135
         TabIndex        =   56
         Text            =   "Barrels at 60°F"
         Top             =   1995
         Width           =   1275
      End
      Begin VB.TextBox txtTitulo 
         Appearance      =   0  'Flat
         BackColor       =   &H80000018&
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         ForeColor       =   &H00808080&
         Height          =   220
         Index           =   11
         Left            =   135
         TabIndex        =   55
         Text            =   "m3 at 15.56°C"
         Top             =   1695
         Width           =   1275
      End
      Begin VB.TextBox txtTitulo 
         Appearance      =   0  'Flat
         BackColor       =   &H80000018&
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         ForeColor       =   &H00808080&
         Height          =   220
         Index           =   9
         Left            =   135
         TabIndex        =   54
         Text            =   "MetricTong"
         Top             =   1410
         Width           =   1275
      End
      Begin VB.TextBox txtTitulo 
         Appearance      =   0  'Flat
         BackColor       =   &H80000018&
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         ForeColor       =   &H00808080&
         Height          =   220
         Index           =   8
         Left            =   135
         TabIndex        =   53
         Text            =   "Tcv"
         Top             =   585
         Width           =   1275
      End
      Begin VB.TextBox txtTitulo 
         Appearance      =   0  'Flat
         BackColor       =   &H80000018&
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         ForeColor       =   &H00808080&
         Height          =   220
         Index           =   7
         Left            =   135
         TabIndex        =   52
         Text            =   "Gov"
         Top             =   285
         Width           =   1275
      End
      Begin VB.TextBox txtNMetricTong 
         BackColor       =   &H80000018&
         ForeColor       =   &H00808080&
         Height          =   285
         Left            =   1455
         TabIndex        =   32
         Text            =   "0"
         Top             =   1360
         Width           =   1500
      End
      Begin VB.TextBox txtNGov 
         BackColor       =   &H80000018&
         ForeColor       =   &H00808080&
         Height          =   285
         Left            =   1455
         TabIndex        =   28
         Text            =   "0"
         Top             =   225
         Width           =   1500
      End
      Begin VB.TextBox txtNTcv 
         BackColor       =   &H80000018&
         ForeColor       =   &H00808080&
         Height          =   285
         Left            =   1455
         TabIndex        =   29
         Text            =   "0"
         Top             =   505
         Width           =   1500
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H80000018&
      Caption         =   "Gross"
      Height          =   4650
      Left            =   6075
      TabIndex        =   26
      Top             =   90
      Width           =   3120
      Begin VB.TextBox txtGTcv 
         BackColor       =   &H80000018&
         ForeColor       =   &H00FF0000&
         Height          =   285
         Left            =   1485
         TabIndex        =   16
         Text            =   "0"
         Top             =   495
         Width           =   1450
      End
      Begin VB.TextBox txtGGov 
         BackColor       =   &H80000018&
         ForeColor       =   &H00FF0000&
         Height          =   285
         Left            =   1485
         TabIndex        =   15
         Text            =   "0"
         Top             =   225
         Width           =   1450
      End
      Begin VB.TextBox txtGMetricTong 
         BackColor       =   &H80000018&
         ForeColor       =   &H00808080&
         Height          =   285
         Left            =   1485
         TabIndex        =   20
         Text            =   "0"
         Top             =   1330
         Width           =   1450
      End
      Begin VB.TextBox txtTitulo 
         Appearance      =   0  'Flat
         BackColor       =   &H80000018&
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   220
         Index           =   45
         Left            =   135
         TabIndex        =   50
         Text            =   "Gov"
         Top             =   270
         Width           =   1320
      End
      Begin VB.TextBox txtTitulo 
         Appearance      =   0  'Flat
         BackColor       =   &H80000018&
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   220
         Index           =   44
         Left            =   135
         TabIndex        =   49
         Text            =   "Tcv"
         Top             =   540
         Width           =   1320
      End
      Begin VB.TextBox txtTitulo 
         Appearance      =   0  'Flat
         BackColor       =   &H80000018&
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         ForeColor       =   &H00808080&
         Height          =   220
         Index           =   43
         Left            =   135
         TabIndex        =   48
         Text            =   "MetricTong"
         Top             =   1365
         Width           =   1320
      End
      Begin VB.TextBox txtTitulo 
         Appearance      =   0  'Flat
         BackColor       =   &H80000018&
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         ForeColor       =   &H00808080&
         Height          =   220
         Index           =   41
         Left            =   135
         TabIndex        =   47
         Text            =   "m3 at 15.56°C"
         Top             =   1650
         Width           =   1320
      End
      Begin VB.TextBox txtTitulo 
         Appearance      =   0  'Flat
         BackColor       =   &H80000018&
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         ForeColor       =   &H00808080&
         Height          =   220
         Index           =   40
         Left            =   135
         TabIndex        =   46
         Text            =   "Barrels at 60°F"
         Top             =   1950
         Width           =   1320
      End
      Begin VB.TextBox txtGCubMeters1556 
         BackColor       =   &H80000018&
         ForeColor       =   &H00808080&
         Height          =   285
         Left            =   1485
         TabIndex        =   21
         Text            =   "0"
         Top             =   1620
         Width           =   1450
      End
      Begin VB.TextBox txtGBarrels60 
         BackColor       =   &H80000018&
         ForeColor       =   &H00808080&
         Height          =   285
         Left            =   1485
         TabIndex        =   22
         Text            =   "0"
         Top             =   1905
         Width           =   1450
      End
      Begin VB.TextBox txtTitulo 
         Appearance      =   0  'Flat
         BackColor       =   &H80000018&
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         ForeColor       =   &H00808080&
         Height          =   220
         Index           =   39
         Left            =   135
         TabIndex        =   45
         Text            =   "Gsv (m3 at 15°C)"
         Top             =   810
         Width           =   1320
      End
      Begin VB.TextBox txtGGsv 
         BackColor       =   &H80000018&
         ForeColor       =   &H00808080&
         Height          =   285
         Left            =   1485
         TabIndex        =   19
         Text            =   "0"
         Top             =   780
         Width           =   1450
      End
      Begin VB.TextBox txtTitulo 
         Appearance      =   0  'Flat
         BackColor       =   &H80000018&
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   220
         Index           =   38
         Left            =   135
         TabIndex        =   44
         Text            =   "Density"
         Top             =   1095
         Width           =   1320
      End
      Begin VB.TextBox txtGDensity 
         BackColor       =   &H80000018&
         ForeColor       =   &H00FF0000&
         Height          =   285
         Left            =   1485
         TabIndex        =   17
         Text            =   "0"
         Top             =   1060
         Width           =   1450
      End
      Begin VB.TextBox txtGGallons60 
         BackColor       =   &H80000018&
         ForeColor       =   &H00808080&
         Height          =   285
         Left            =   1485
         TabIndex        =   23
         Text            =   "0"
         Top             =   2185
         Width           =   1450
      End
      Begin VB.TextBox txtTitulo 
         Appearance      =   0  'Flat
         BackColor       =   &H80000018&
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         ForeColor       =   &H00808080&
         Height          =   220
         Index           =   37
         Left            =   135
         TabIndex        =   43
         Text            =   "Gallons at 60°F"
         Top             =   2220
         Width           =   1320
      End
      Begin VB.TextBox txtGLongTons 
         BackColor       =   &H80000018&
         ForeColor       =   &H00808080&
         Height          =   285
         Left            =   1485
         TabIndex        =   24
         Text            =   "0"
         Top             =   2470
         Width           =   1450
      End
      Begin VB.TextBox txtTitulo 
         Appearance      =   0  'Flat
         BackColor       =   &H80000018&
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         ForeColor       =   &H00808080&
         Height          =   220
         Index           =   36
         Left            =   135
         TabIndex        =   41
         Text            =   "LongTons"
         Top             =   2490
         Width           =   1320
      End
      Begin VB.TextBox txtGAPIGravity 
         BackColor       =   &H80000018&
         ForeColor       =   &H00808080&
         Height          =   285
         Left            =   1485
         TabIndex        =   25
         Text            =   "0"
         Top             =   2760
         Width           =   1450
      End
      Begin VB.TextBox txtTitulo 
         Appearance      =   0  'Flat
         BackColor       =   &H80000018&
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         ForeColor       =   &H00808080&
         Height          =   220
         Index           =   35
         Left            =   135
         TabIndex        =   39
         Text            =   "APIGravity"
         Top             =   2805
         Width           =   1320
      End
      Begin VB.TextBox txtGBswCBM 
         BackColor       =   &H80000018&
         ForeColor       =   &H00FF0000&
         Height          =   285
         Left            =   1485
         TabIndex        =   18
         Text            =   "0"
         Top             =   3040
         Width           =   1450
      End
      Begin VB.TextBox txtTitulo 
         Appearance      =   0  'Flat
         BackColor       =   &H80000018&
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   220
         Index           =   34
         Left            =   135
         TabIndex        =   27
         Text            =   "Bsw %"
         Top             =   3060
         Width           =   1320
      End
   End
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "&Aceptar"
      Height          =   330
      Left            =   10755
      TabIndex        =   40
      Top             =   4815
      Width           =   1500
   End
   Begin VB.CommandButton cmdCancelar 
      Caption         =   "&Cancelar"
      Height          =   330
      Left            =   9045
      TabIndex        =   42
      Top             =   4815
      Width           =   1500
   End
End
Attribute VB_Name = "frmEntregasCliInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Function infoCalcGross()
  
    ' gross
  Me.txtGGsv = CCur(Me.txtGTcv)
  Me.txtGMetricTong = Round(CCur(Me.txtGGsv) * CSng(Me.txtGDensity), 3)
  Me.txtGCubMeters1556 = Round(CCur(Me.txtGGsv) * getParam("m315TOm31556"), 3)
  Me.txtGBarrels60 = Round(CCur(Me.txtGCubMeters1556) * getParam("m31556TOBarr1556"), 3)
  Me.txtGGallons60 = Round(CCur(Me.txtGBarrels60) * getParam("Barr1556TOGall1556"), 0)
  Me.txtGLongTons = Round(CCur(Me.txtGMetricTong) * getParam("MTongTOLTons"), 3)
  Me.txtGAPIGravity = Round(141.5 / (CSng(Me.txtGDensity) + 0.0005) - 131.5, 2)
  Me.txtNBswCBM = Round(CCur(Me.txtGGsv) * CCur(Me.txtGBswCBM) / 100, 3)
  
  ' gross
  Me.txtGGsv = CCur(Me.txtGTcv)
  Me.txtGMetricTong = Round(CCur(Me.txtGGsv) * CSng(Me.txtGDensity), 3)
  Me.txtGCubMeters1556 = Round(CCur(Me.txtGGsv) * getParam("m315TOm31556"), 3)
  Me.txtGBarrels60 = Round(CCur(Me.txtGCubMeters1556) * getParam("m31556TOBarr1556"), 3)
  Me.txtGGallons60 = Round(CCur(Me.txtGBarrels60) * getParam("Barr1556TOGall1556"), 0)
  Me.txtGLongTons = Round(CCur(Me.txtGMetricTong) * getParam("MTongTOLTons"), 3)
  Me.txtGAPIGravity = Round(141.5 / (CSng(Me.txtGDensity) + 0.0005) - 131.5, 2)
  Me.txtNBswCBM = Round(CCur(Me.txtGGsv) * CCur(Me.txtGBswCBM) / 100, 3)

  ' si es oleaducto pongo valores a cero
  If Me.cboEntregaCliTipo = "Oleoducto" Then
    
    Me.txtGGov = Round(0, 3)
    Me.txtGMetricTong = Round(0, 3)
    Me.txtGCubMeters1556 = Round(0, 3)
    Me.txtGBarrels60 = Round(0, 3)
    Me.txtGGallons60 = Round(0, 0)
    Me.txtGLongTons = Round(0, 3)
    Me.txtGAPIGravity = Round(0, 2)
    Me.txtNGov = Round(0, 3)
    Me.txtNTcv = Round(0, 3)
    Me.txtNMetricTong = Round(0, 3)
    Me.txtNGallons60 = Round(0, 0)
    Me.txtNLongTons = Round(0, 3)
    
  End If

End Function
  
Private Function infoCalcNeto()
  
  ' neto
  Me.txtNGsv = Round(CCur(Me.txtGGsv) - CCur(Me.txtNBswCBM), 3)
  Me.txtNMetricTong = Round(CCur(Me.txtGMetricTong) - CCur(Me.txtNBswCBM), 3)
  If Val(Me.txtNGsv) <> 0 Then
    Me.txtNDensity = Round(CCur(Me.txtNMetricTong) / CCur(Me.txtNGsv), 6)
  End If
  Me.txtNCubMeters1556 = Round(CCur(Me.txtNGsv) * getParam("m315TOm31556"), 3)
  Me.txtNBarrels60 = Round(CCur(Me.txtNCubMeters1556) * getParam("m31556TOBarr1556"), 3)
  Me.txtNGallons60 = Round(CCur(Me.txtNBarrels60) * getParam("Barr1556TOGall1556"), 0)
  Me.txtNLongTons = Round(CCur(Me.txtNMetricTong) * getParam("MTongTOLTons"), 3)
  If Val(Me.txtNDensity) <> 0 Then
    Me.txtNAPIGravity = Round(141.5 / (CSng(Me.txtNDensity) + 0.0005) - 131.5, 2)
  End If

  ' si es oleaducto pongo valores a cero
  If Me.cboEntregaCliTipo = "Oleoducto" Then
    
    Me.txtGGov = Round(0, 3)
    Me.txtGMetricTong = Round(0, 3)
    Me.txtGCubMeters1556 = Round(0, 3)
    Me.txtGBarrels60 = Round(0, 3)
    Me.txtGGallons60 = Round(0, 0)
    Me.txtGLongTons = Round(0, 3)
    Me.txtGAPIGravity = Round(0, 2)
    Me.txtNGov = Round(0, 3)
    Me.txtNTcv = Round(0, 3)
    Me.txtNMetricTong = Round(0, 3)
    Me.txtNGallons60 = Round(0, 0)
    Me.txtNLongTons = Round(0, 3)
    
  End If

End Function

Private Sub cmdAceptar_Click()
  
  If Not DataValidate(txtLegajo, "@15", True) Then Exit Sub
  If Not DataValidate(cboEntregaCliTipo, , True) Then Exit Sub
  If Not DataValidate(txtFecha, "dd/mm/yyyy", True) Then Exit Sub
  If Not DataValidate(cboEmpresa, , True) Then Exit Sub
  If Not DataValidate(cboCliente, , True) Then Exit Sub
  If Not DataValidate(cboInspeccion, , True) Then Exit Sub
  If Not DataValidate(txtCertificado, "dd/mm/yyyy", True) Then Exit Sub
  If Not DataValidate(cboBarco, , True) Then Exit Sub
  If Not DataValidate(cboTerminal, , True) Then Exit Sub
  If Not DataValidate(txtCerDesde, "dd/mm/yyyy", True) Then Exit Sub
  If Not DataValidate(txtCerHasta, "dd/mm/yyyy", True) Then Exit Sub
  If Not DataValidate(txtEntregaNro, "@200") Then Exit Sub
  If Not DataValidate(cboProvicionado, , True) Then Exit Sub
  If Not DataValidate(txtAzufre, "###.######") Then Exit Sub
  If Not DataValidate(txtOtrosAjustes, "###.##") Then Exit Sub
  ' gross
  If Not DataValidate(txtGGov, "#########.###", True) Then Exit Sub
  If Not DataValidate(txtGTcv, "#########.###", True) Then Exit Sub
  If Not DataValidate(txtGGsv, "#########.###", True) Then Exit Sub
  If Not DataValidate(txtGDensity, "##.######", True) Then Exit Sub
  If Not DataValidate(txtGMetricTong, "#########.###", True) Then Exit Sub
  If Not DataValidate(txtGCubMeters1556, "#########.###", True) Then Exit Sub
  If Not DataValidate(txtGBarrels60, "#########.###", True) Then Exit Sub
  If Not DataValidate(txtGGallons60, "#########", True) Then Exit Sub
  If Not DataValidate(txtGLongTons, "#########.###", True) Then Exit Sub
  If Not DataValidate(txtGAPIGravity, "##.##", True) Then Exit Sub
  If Not DataValidate(txtGBswCBM, "##.####", True) Then Exit Sub
  ' Neto
  If Not DataValidate(txtNGov, "#########.###", False) Then Exit Sub
  If Not DataValidate(txtNTcv, "#########.###", False) Then Exit Sub
  If Not DataValidate(txtNGsv, "#########.###", True) Then Exit Sub
  If Not DataValidate(txtNDensity, "##.######", True) Then Exit Sub
  If Not DataValidate(txtNMetricTong, "#########.###", True) Then Exit Sub
  If Not DataValidate(txtNCubMeters1556, "#########.###", True) Then Exit Sub
  If Not DataValidate(txtNBarrels60, "#########.###", True) Then Exit Sub
  If Not DataValidate(txtNGallons60, "#########", True) Then Exit Sub
  If Not DataValidate(txtNLongTons, "#########.###", True) Then Exit Sub
  If Not DataValidate(txtNAPIGravity, "##.##", True) Then Exit Sub
  If Not DataValidate(txtNBswCBM, "########.###", True) Then Exit Sub
   
  blnAceptar = True
  blnCancelar = False
  Me.Hide

End Sub

Private Sub cmdCancelar_Click()

  blnAceptar = False
  blnCancelar = True
  Unload Me

End Sub

Private Sub txtGravity_Change()

End Sub

Private Sub cmdNewBarco_Click()

  Dim strStore, strView, strDato As String

  strStore = "spBarcosInsert"
  strView = "SELECT * FROM ViewBArcos"
  strDato = ComboBoxAddItem(Me, cboBarco, "@50", strStore, strView)

End Sub


Private Sub cmdNewEmpresa_Click()

  Dim strStore, strView, strDato As String

  strStore = "spEmpresasInsert"
  strView = "SELECT * FROM ViewEmpresas"
  strDato = ComboBoxAddItem(Me, cboEmpresa, "@50", strStore, strView)

End Sub

Private Sub cmdNewEntregaCliTipo_Click()
  Dim strStore, strView, strDato As String

  strStore = "spEntregasCliTiposInsert"
  strView = "SELECT * FROM ViewEntregasCliTipos"
  strDato = ComboBoxAddItem(Me, cboEntregaCliTipo, "@50", strStore, strView)

End Sub

Private Sub cmdNewInspeccion_Click()

  Dim strStore, strView, strDato As String

  strStore = "spInspeccionesInsert"
  strView = "SELECT * FROM ViewInspecciones"
  strDato = ComboBoxAddItem(Me, cboInspeccion, "@50", strStore, strView)

End Sub

Private Sub cmdNewterminal_Click()
  Dim strStore, strView, strDato As String

  strStore = "spTerminalesInsert"
  strView = "SELECT * FROM ViewTerminales"
  strDato = ComboBoxAddItem(Me, cboTerminal, "@50", strStore, strView)

End Sub

Private Sub Form_Load()
  
  strSQL = "SELECT * FROM ViewEntregasCliTipos"
  intRes = ComboBoxRefresh(cboEntregaCliTipo, strSQL)
  
  strSQL = "SELECT * FROM ViewEmpresas"
  intRes = ComboBoxRefresh(cboEmpresa, strSQL)
  
  strSQL = "SELECT * FROM ViewClientes"
  intRes = ComboBoxRefresh(cboCliente, strSQL)

  strSQL = "SELECT * FROM ViewInspecciones"
  intRes = ComboBoxRefresh(cboInspeccion, strSQL)

  strSQL = "SELECT * FROM ViewBarcos"
  intRes = ComboBoxRefresh(cboBarco, strSQL)

  strSQL = "SELECT * FROM ViewTerminales"
  intRes = ComboBoxRefresh(cboTerminal, strSQL)

End Sub

Private Sub txtGBswCBM_LostFocus()
  intRes = infoCalcGross()
  intRes = infoCalcNeto()
End Sub

Private Sub txtGDensity_LostFocus()
  intRes = infoCalcGross()
  intRes = infoCalcNeto()
End Sub

Private Sub txtGGov_LostFocus()
  intRes = infoCalcGross()
End Sub

Private Sub txtGTcv_LostFocus()
  intRes = infoCalcGross()
  intRes = infoCalcNeto()
End Sub
