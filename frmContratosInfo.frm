VERSION 5.00
Begin VB.Form frmContratosInfo 
   BackColor       =   &H80000018&
   BorderStyle     =   5  'Sizable ToolWindow
   Caption         =   "Contratos"
   ClientHeight    =   8400
   ClientLeft      =   60
   ClientTop       =   300
   ClientWidth     =   5820
   ControlBox      =   0   'False
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8400
   ScaleWidth      =   5820
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame1 
      BackColor       =   &H80000018&
      Caption         =   "Descuento"
      Height          =   1275
      Left            =   135
      TabIndex        =   65
      Top             =   6075
      Width           =   5595
      Begin VB.TextBox txtDesMeses 
         BackColor       =   &H80000018&
         ForeColor       =   &H00808080&
         Height          =   285
         Left            =   5085
         TabIndex        =   31
         Top             =   900
         Width           =   420
      End
      Begin VB.TextBox Text38 
         Appearance      =   0  'Flat
         BackColor       =   &H80000018&
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         ForeColor       =   &H00808080&
         Height          =   285
         Left            =   3780
         TabIndex        =   71
         Text            =   "Cantidad meses"
         Top             =   945
         Width           =   1170
      End
      Begin VB.TextBox txtDesPoste 
         BackColor       =   &H80000018&
         ForeColor       =   &H00808080&
         Height          =   285
         Left            =   3060
         TabIndex        =   30
         Top             =   900
         Width           =   420
      End
      Begin VB.TextBox txtDesPrevios 
         BackColor       =   &H80000018&
         ForeColor       =   &H00808080&
         Height          =   285
         Left            =   1125
         TabIndex        =   29
         Top             =   900
         Width           =   420
      End
      Begin VB.TextBox txtDesHasta 
         BackColor       =   &H80000018&
         ForeColor       =   &H00808080&
         Height          =   285
         Left            =   4050
         TabIndex        =   28
         Top             =   585
         Width           =   1455
      End
      Begin VB.TextBox txtDesDesde 
         BackColor       =   &H80000018&
         ForeColor       =   &H00808080&
         Height          =   285
         Left            =   1170
         TabIndex        =   27
         Top             =   585
         Width           =   1455
      End
      Begin VB.TextBox txtDesFormula 
         BackColor       =   &H80000018&
         ForeColor       =   &H00808080&
         Height          =   285
         Left            =   765
         TabIndex        =   26
         Top             =   270
         Width           =   4740
      End
      Begin VB.TextBox Text32 
         Appearance      =   0  'Flat
         BackColor       =   &H80000018&
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         ForeColor       =   &H00808080&
         Height          =   285
         Left            =   2970
         TabIndex        =   70
         Text            =   "Fecha hasta"
         Top             =   630
         Width           =   1800
      End
      Begin VB.TextBox Text27 
         Appearance      =   0  'Flat
         BackColor       =   &H80000018&
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         ForeColor       =   &H00808080&
         Height          =   285
         Left            =   90
         TabIndex        =   69
         Text            =   "Fecha desde"
         Top             =   630
         Width           =   990
      End
      Begin VB.TextBox Text26 
         Appearance      =   0  'Flat
         BackColor       =   &H80000018&
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         ForeColor       =   &H00808080&
         Height          =   285
         Left            =   1755
         TabIndex        =   68
         Text            =   "Dias posteriores"
         Top             =   945
         Width           =   1170
      End
      Begin VB.TextBox Text24 
         Appearance      =   0  'Flat
         BackColor       =   &H80000018&
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         ForeColor       =   &H00808080&
         Height          =   285
         Left            =   90
         TabIndex        =   67
         Text            =   "Dias previos"
         Top             =   945
         Width           =   990
      End
      Begin VB.TextBox Text22 
         Appearance      =   0  'Flat
         BackColor       =   &H80000018&
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         ForeColor       =   &H00808080&
         Height          =   285
         Left            =   90
         TabIndex        =   66
         Text            =   "Fórmula"
         Top             =   315
         Width           =   630
      End
   End
   Begin VB.ComboBox cboRangos 
      Appearance      =   0  'Flat
      BackColor       =   &H80000018&
      ForeColor       =   &H00808080&
      Height          =   315
      ItemData        =   "frmContratosInfo.frx":0000
      Left            =   1980
      List            =   "frmContratosInfo.frx":0002
      Sorted          =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   13
      Top             =   3330
      Width           =   3060
   End
   Begin VB.CommandButton cmdRangos 
      Caption         =   "Rango"
      Height          =   300
      Left            =   5040
      TabIndex        =   64
      Top             =   3330
      Width           =   645
   End
   Begin VB.TextBox Text31 
      Appearance      =   0  'Flat
      BackColor       =   &H80000018&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      ForeColor       =   &H00808080&
      Height          =   285
      Left            =   3960
      TabIndex        =   63
      Text            =   "Redondeo"
      Top             =   4065
      Width           =   765
   End
   Begin VB.TextBox Text30 
      Appearance      =   0  'Flat
      BackColor       =   &H80000018&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      ForeColor       =   &H00808080&
      Height          =   285
      Left            =   3960
      TabIndex        =   62
      Text            =   "Redondeo"
      Top             =   4425
      Width           =   765
   End
   Begin VB.TextBox Text29 
      Appearance      =   0  'Flat
      BackColor       =   &H80000018&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      ForeColor       =   &H00808080&
      Height          =   285
      Left            =   3960
      TabIndex        =   61
      Text            =   "Redondeo"
      Top             =   4785
      Width           =   765
   End
   Begin VB.TextBox Text28 
      Appearance      =   0  'Flat
      BackColor       =   &H80000018&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      ForeColor       =   &H00808080&
      Height          =   285
      Left            =   3960
      TabIndex        =   60
      Text            =   "Redondeo"
      Top             =   3705
      Width           =   765
   End
   Begin VB.TextBox txtRedondeo2 
      BackColor       =   &H80000018&
      ForeColor       =   &H00808080&
      Height          =   285
      Left            =   4815
      TabIndex        =   17
      Top             =   4020
      Width           =   870
   End
   Begin VB.TextBox txtRedondeo1 
      BackColor       =   &H80000018&
      ForeColor       =   &H00808080&
      Height          =   285
      Left            =   4815
      TabIndex        =   15
      Top             =   3660
      Width           =   870
   End
   Begin VB.TextBox txtRedondeo3 
      BackColor       =   &H80000018&
      ForeColor       =   &H00808080&
      Height          =   285
      Left            =   4815
      TabIndex        =   19
      Top             =   4380
      Width           =   870
   End
   Begin VB.TextBox txtRedondeo4 
      BackColor       =   &H80000018&
      ForeColor       =   &H00808080&
      Height          =   285
      Left            =   4815
      TabIndex        =   21
      Top             =   4740
      Width           =   870
   End
   Begin VB.TextBox Text25 
      Appearance      =   0  'Flat
      BackColor       =   &H80000018&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      ForeColor       =   &H00808080&
      Height          =   285
      Left            =   135
      TabIndex        =   59
      Text            =   "Formula Ajuste Precio"
      Top             =   3705
      Width           =   1800
   End
   Begin VB.TextBox Text23 
      Appearance      =   0  'Flat
      BackColor       =   &H80000018&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      ForeColor       =   &H00808080&
      Height          =   285
      Left            =   135
      TabIndex        =   58
      Text            =   "Formula Ajuste API"
      Top             =   4065
      Width           =   1800
   End
   Begin VB.TextBox Text19 
      Appearance      =   0  'Flat
      BackColor       =   &H80000018&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      ForeColor       =   &H00808080&
      Height          =   285
      Left            =   135
      TabIndex        =   57
      Text            =   "Formula Ajuste Varios"
      Top             =   4380
      Width           =   1800
   End
   Begin VB.ComboBox cboAjusteVarios 
      Appearance      =   0  'Flat
      BackColor       =   &H80000018&
      ForeColor       =   &H00808080&
      Height          =   315
      ItemData        =   "frmContratosInfo.frx":0004
      Left            =   1980
      List            =   "frmContratosInfo.frx":000E
      Sorted          =   -1  'True
      TabIndex        =   18
      Text            =   "cboAjusteVarios"
      Top             =   4380
      Width           =   1905
   End
   Begin VB.ComboBox cboAjusteApi 
      Appearance      =   0  'Flat
      BackColor       =   &H80000018&
      ForeColor       =   &H00808080&
      Height          =   315
      ItemData        =   "frmContratosInfo.frx":001E
      Left            =   1980
      List            =   "frmContratosInfo.frx":0028
      Sorted          =   -1  'True
      TabIndex        =   16
      Text            =   "cboAjusteApi"
      Top             =   4020
      Width           =   1905
   End
   Begin VB.ComboBox cboAjustePrecio 
      Appearance      =   0  'Flat
      BackColor       =   &H80000018&
      ForeColor       =   &H00808080&
      Height          =   315
      ItemData        =   "frmContratosInfo.frx":0038
      Left            =   1980
      List            =   "frmContratosInfo.frx":0042
      Sorted          =   -1  'True
      TabIndex        =   14
      Text            =   "cboAjustePrecio"
      Top             =   3660
      Width           =   1905
   End
   Begin VB.TextBox Text21 
      Appearance      =   0  'Flat
      BackColor       =   &H80000018&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      ForeColor       =   &H00808080&
      Height          =   285
      Left            =   135
      TabIndex        =   56
      Text            =   "Tabla Rangos"
      Top             =   3345
      Width           =   1800
   End
   Begin VB.TextBox Text14 
      Appearance      =   0  'Flat
      BackColor       =   &H80000018&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      ForeColor       =   &H00808080&
      Height          =   285
      Left            =   135
      TabIndex        =   55
      Text            =   "TipoCalculo"
      Top             =   4740
      Width           =   1800
   End
   Begin VB.ComboBox cboTipoCalculo 
      Appearance      =   0  'Flat
      BackColor       =   &H80000018&
      ForeColor       =   &H00808080&
      Height          =   315
      ItemData        =   "frmContratosInfo.frx":0052
      Left            =   1980
      List            =   "frmContratosInfo.frx":005F
      Style           =   2  'Dropdown List
      TabIndex        =   20
      Top             =   4740
      Width           =   1905
   End
   Begin VB.TextBox Text20 
      Appearance      =   0  'Flat
      BackColor       =   &H80000018&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      ForeColor       =   &H00808080&
      Height          =   285
      Left            =   135
      TabIndex        =   54
      Text            =   "Barrels at 60°F"
      Top             =   1710
      Width           =   1800
   End
   Begin VB.TextBox txtBarrels60 
      BackColor       =   &H80000018&
      ForeColor       =   &H00808080&
      Height          =   285
      Left            =   1980
      TabIndex        =   6
      Top             =   1710
      Width           =   3700
   End
   Begin VB.TextBox Text16 
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
      ForeColor       =   &H00808080&
      Height          =   285
      Left            =   135
      TabIndex        =   53
      Text            =   "m3 at 1556°C"
      Top             =   1395
      Width           =   1800
   End
   Begin VB.TextBox txtm31556 
      BackColor       =   &H80000018&
      ForeColor       =   &H00FF0000&
      Height          =   285
      Left            =   1980
      TabIndex        =   5
      Top             =   1395
      Width           =   3700
   End
   Begin VB.TextBox Text18 
      Appearance      =   0  'Flat
      BackColor       =   &H80000018&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      ForeColor       =   &H00808080&
      Height          =   285
      Left            =   150
      TabIndex        =   52
      Text            =   "AzuMinimo"
      Top             =   2700
      Width           =   1755
   End
   Begin VB.TextBox Text17 
      Appearance      =   0  'Flat
      BackColor       =   &H80000018&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      ForeColor       =   &H00808080&
      Height          =   285
      Left            =   3480
      TabIndex        =   51
      Text            =   "AzuMaximo"
      Top             =   2700
      Width           =   855
   End
   Begin VB.TextBox txtAzuMinimo 
      BackColor       =   &H80000018&
      ForeColor       =   &H00808080&
      Height          =   285
      Left            =   1980
      TabIndex        =   10
      Top             =   2700
      Width           =   1410
   End
   Begin VB.TextBox txtAzuMaximo 
      BackColor       =   &H80000018&
      ForeColor       =   &H00808080&
      Height          =   285
      Left            =   4365
      TabIndex        =   11
      Top             =   2700
      Width           =   1320
   End
   Begin VB.CommandButton Command1 
      Caption         =   "New"
      Height          =   300
      Left            =   5130
      TabIndex        =   50
      Top             =   2025
      Width           =   555
   End
   Begin VB.CommandButton cmdNewEmpresa 
      Caption         =   "New"
      Height          =   290
      Left            =   5130
      TabIndex        =   49
      Top             =   90
      Width           =   555
   End
   Begin VB.ComboBox cboMesPromedio 
      Appearance      =   0  'Flat
      BackColor       =   &H80000018&
      Enabled         =   0   'False
      ForeColor       =   &H00808080&
      Height          =   315
      ItemData        =   "frmContratosInfo.frx":0079
      Left            =   1980
      List            =   "frmContratosInfo.frx":0086
      Style           =   2  'Dropdown List
      TabIndex        =   25
      Top             =   5730
      Width           =   3750
   End
   Begin VB.TextBox Text15 
      Appearance      =   0  'Flat
      BackColor       =   &H80000018&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      ForeColor       =   &H00808080&
      Height          =   285
      Left            =   135
      TabIndex        =   48
      Text            =   "MesPromedio"
      Top             =   5730
      Width           =   1800
   End
   Begin VB.TextBox txtDiasPosteriores 
      BackColor       =   &H80000018&
      Enabled         =   0   'False
      ForeColor       =   &H00808080&
      Height          =   285
      Left            =   4365
      TabIndex        =   23
      Top             =   5100
      Width           =   1365
   End
   Begin VB.TextBox Text10 
      Appearance      =   0  'Flat
      BackColor       =   &H80000018&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      ForeColor       =   &H00808080&
      Height          =   285
      Left            =   3195
      TabIndex        =   46
      Text            =   "DiasPosteriores"
      Top             =   5100
      Width           =   1170
   End
   Begin VB.TextBox txtObservaciones 
      BackColor       =   &H80000018&
      ForeColor       =   &H00808080&
      Height          =   330
      Left            =   135
      MultiLine       =   -1  'True
      TabIndex        =   32
      Top             =   7620
      Width           =   5595
   End
   Begin VB.ComboBox cboIncluyeEntregaCli 
      Appearance      =   0  'Flat
      BackColor       =   &H80000018&
      Enabled         =   0   'False
      ForeColor       =   &H00808080&
      Height          =   315
      ItemData        =   "frmContratosInfo.frx":00A5
      Left            =   1980
      List            =   "frmContratosInfo.frx":00AF
      Style           =   2  'Dropdown List
      TabIndex        =   24
      Top             =   5415
      Width           =   3750
   End
   Begin VB.ComboBox cboPrecioTipo 
      Appearance      =   0  'Flat
      BackColor       =   &H80000018&
      ForeColor       =   &H00808080&
      Height          =   315
      ItemData        =   "frmContratosInfo.frx":00BB
      Left            =   1980
      List            =   "frmContratosInfo.frx":00C5
      Sorted          =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   7
      Top             =   2025
      Width           =   3155
   End
   Begin VB.TextBox txtDiasPrevios 
      BackColor       =   &H80000018&
      Enabled         =   0   'False
      ForeColor       =   &H00808080&
      Height          =   285
      Left            =   1980
      TabIndex        =   22
      Top             =   5100
      Width           =   1140
   End
   Begin VB.TextBox txtDescuento 
      BackColor       =   &H80000018&
      ForeColor       =   &H00808080&
      Height          =   285
      Left            =   1980
      TabIndex        =   12
      Top             =   3015
      Width           =   3700
   End
   Begin VB.TextBox txtAPIMaximo 
      BackColor       =   &H80000018&
      ForeColor       =   &H00808080&
      Height          =   285
      Left            =   4365
      TabIndex        =   9
      Top             =   2385
      Width           =   1320
   End
   Begin VB.TextBox txtAPIMinimo 
      BackColor       =   &H80000018&
      ForeColor       =   &H00808080&
      Height          =   285
      Left            =   1980
      TabIndex        =   8
      Top             =   2385
      Width           =   1410
   End
   Begin VB.TextBox txtm315 
      BackColor       =   &H80000018&
      ForeColor       =   &H00FF0000&
      Height          =   285
      Left            =   1980
      TabIndex        =   4
      Top             =   1080
      Width           =   3700
   End
   Begin VB.TextBox txtFechaHasta 
      BackColor       =   &H80000018&
      ForeColor       =   &H00808080&
      Height          =   285
      Left            =   4365
      TabIndex        =   3
      Top             =   765
      Width           =   1320
   End
   Begin VB.TextBox txtFechaDesde 
      BackColor       =   &H80000018&
      ForeColor       =   &H00808080&
      Height          =   285
      Left            =   1980
      TabIndex        =   2
      Top             =   765
      Width           =   1365
   End
   Begin VB.ComboBox cboCliente 
      Appearance      =   0  'Flat
      BackColor       =   &H80000018&
      ForeColor       =   &H00808080&
      Height          =   315
      Left            =   1980
      Sorted          =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   405
      Width           =   3750
   End
   Begin VB.ComboBox cboEmpresa 
      Appearance      =   0  'Flat
      BackColor       =   &H80000018&
      ForeColor       =   &H00808080&
      Height          =   315
      Left            =   1980
      Sorted          =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   90
      Width           =   3165
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H80000018&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      ForeColor       =   &H00808080&
      Height          =   240
      Left            =   180
      TabIndex        =   35
      Text            =   "Observaciones"
      Top             =   7395
      Width           =   5505
   End
   Begin VB.TextBox Text13 
      Appearance      =   0  'Flat
      BackColor       =   &H80000018&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      ForeColor       =   &H00808080&
      Height          =   285
      Left            =   135
      TabIndex        =   47
      Text            =   "IncluyeEntregaCli"
      Top             =   5415
      Width           =   1800
   End
   Begin VB.TextBox Text12 
      Appearance      =   0  'Flat
      BackColor       =   &H80000018&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      ForeColor       =   &H00808080&
      Height          =   285
      Left            =   135
      TabIndex        =   41
      Text            =   "PreciosTipo"
      Top             =   2070
      Width           =   1800
   End
   Begin VB.TextBox Text11 
      Appearance      =   0  'Flat
      BackColor       =   &H80000018&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      ForeColor       =   &H00808080&
      Height          =   285
      Left            =   135
      TabIndex        =   45
      Text            =   "DiasPrevios"
      Top             =   5100
      Width           =   1800
   End
   Begin VB.TextBox Text9 
      Appearance      =   0  'Flat
      BackColor       =   &H80000018&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      ForeColor       =   &H00808080&
      Height          =   285
      Left            =   135
      TabIndex        =   44
      Text            =   "Descuento"
      Top             =   3015
      Width           =   1800
   End
   Begin VB.TextBox Text8 
      Appearance      =   0  'Flat
      BackColor       =   &H80000018&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      ForeColor       =   &H00808080&
      Height          =   285
      Left            =   3465
      TabIndex        =   43
      Text            =   "APIMaximo"
      Top             =   2385
      Width           =   900
   End
   Begin VB.TextBox Text7 
      Appearance      =   0  'Flat
      BackColor       =   &H80000018&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      ForeColor       =   &H00808080&
      Height          =   285
      Left            =   135
      TabIndex        =   42
      Text            =   "APIMinimo"
      Top             =   2385
      Width           =   1800
   End
   Begin VB.TextBox Text6 
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
      ForeColor       =   &H00808080&
      Height          =   285
      Left            =   135
      TabIndex        =   40
      Text            =   "m3 at 15°C"
      Top             =   1080
      Width           =   1800
   End
   Begin VB.TextBox Text5 
      Appearance      =   0  'Flat
      BackColor       =   &H80000018&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      ForeColor       =   &H00808080&
      Height          =   285
      Left            =   3420
      TabIndex        =   39
      Text            =   "FechaHasta"
      Top             =   765
      Width           =   990
   End
   Begin VB.TextBox Text4 
      Appearance      =   0  'Flat
      BackColor       =   &H80000018&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      ForeColor       =   &H00808080&
      Height          =   285
      Left            =   135
      TabIndex        =   38
      Text            =   "FechaDesde"
      Top             =   765
      Width           =   1800
   End
   Begin VB.TextBox Text3 
      Appearance      =   0  'Flat
      BackColor       =   &H80000018&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      ForeColor       =   &H00808080&
      Height          =   285
      Left            =   135
      TabIndex        =   37
      Text            =   "Cliente"
      Top             =   450
      Width           =   1800
   End
   Begin VB.TextBox Text2 
      Appearance      =   0  'Flat
      BackColor       =   &H80000018&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      ForeColor       =   &H00808080&
      Height          =   285
      Left            =   135
      TabIndex        =   36
      Text            =   "Empresa"
      Top             =   135
      Width           =   1800
   End
   Begin VB.CommandButton cmdCancelar 
      Caption         =   "&Cancelar"
      Height          =   285
      Left            =   2445
      TabIndex        =   34
      Top             =   8025
      Width           =   1500
   End
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "&Aceptar"
      Height          =   285
      Left            =   4200
      TabIndex        =   33
      Top             =   8025
      Width           =   1500
   End
End
Attribute VB_Name = "frmContratosInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public Function llenoRangos()

  'lleno combo con tabla Rangos, lo hago de esta manera porque el comboBox no tiene ID
  Dim rs As ADODB.Recordset
    
  strSQL = "select nombre from contratosRangos group by nombre"
  Set rs = adoGetRS(strSQL)
    
  'chequeo errores
  If Not lngAdoErrNum = -1 Then
    adoError
    Exit Function
  End If
      
  'recorro
  Me.cboRangos.Clear
  While Not rs.EOF
    Me.cboRangos.AddItem IIf(IsNull(rs!Nombre), "", rs!Nombre)
    rs.MoveNext
  Wend

End Function


Private Sub cboEmpresa_LostFocus()

  Dim intRes As Integer
  Dim strSQL As String
  Dim strEmpresa As String

  If ComboBoxNotinList(cboEmpresa) Then
    
    intRes = MsgBox("La empresa ingresada no esta en la lista, desea agregarla.", vbQuestion + vbYesNo, "Confirmacón")
    If intRes Then
    
      strSQL = "EXEC spEmpresasInsert '" & cboEmpresa.Text & "'"
      intRes = adoExecSQL(strSQL)
    
      strEmpresa = Left(cboEmpresa.Text, 15)
    
      ' refresh ComboBox
  
      strSQL = "SELECT * FROM ViewEmpresas"
      intRes = ComboBoxRefresh(cboEmpresa, strSQL)
    
      ' hubico listindex en elemento ingresado
    
      cboEmpresa.ListIndex = ComboBoxFindItem(cboEmpresa, strEmpresa)
    
    End If
  
  End If

End Sub

Private Sub cboPrecioTipo_lostfocus()

  Dim intRes As Integer
  Dim strSQL As String
  Dim strPrecioTipo As String

  If ComboBoxNotinList(cboPrecioTipo) Then
    
    intRes = MsgBox("El Tipo de Precio ingresado no esta en la lista, desea agregarlo.", vbQuestion + vbYesNo, "Confirmacón")
    If intRes Then
    
      strSQL = "EXEC spPreciosTiposInsert '" & cboPrecioTipo.Text & "'"
      intRes = adoExecSQL(strSQL)
    
      strPrecioTipo = cboPrecioTipo.Text
    
      ' refresh ComboBox
  
      strSQL = "SELECT * FROM ViewPreciostipos"
      intRes = ComboBoxRefresh(cboPrecioTipo, strSQL)
    
      ' hubico listindex en elemento ingresado
    
      cboPrecioTipo.ListIndex = ComboBoxFindItem(cboPrecioTipo, strPrecioTipo)
    
    End If
  
  End If
End Sub

Private Sub cboTipoCalculo_Click()
  
  Select Case Me.cboTipoCalculo.List(Me.cboTipoCalculo.ListIndex)
  
  Case "Dias"
    Me.txtDiasPrevios.Enabled = True
    Me.txtDiasPosteriores.Enabled = True
    Me.cboIncluyeEntregaCli.Enabled = True
    Me.cboMesPromedio.Enabled = False
  
  Case "Mensual"
    Me.txtDiasPrevios.Enabled = False
    Me.txtDiasPosteriores.Enabled = False
    Me.cboIncluyeEntregaCli.Enabled = False
    Me.cboMesPromedio.Enabled = True
  
  Case "Rango"
    Me.txtDiasPrevios.Enabled = False
    Me.txtDiasPosteriores.Enabled = False
    Me.cboIncluyeEntregaCli.Enabled = False
    Me.cboMesPromedio.Enabled = False
  
  End Select

End Sub

Private Sub cmdAceptar_Click()
  
  ' validacion de datos
  
  If Not DataValidate(cboEmpresa, , True) Then Exit Sub
  If Not DataValidate(cboCliente, , True) Then Exit Sub
  If Not DataValidate(txtFechaDesde, "dd/mm/yyyy", True) Then Exit Sub
  If Not DataValidate(txtFechaHasta, "dd/mm/yyyy", True) Then Exit Sub
  If Not DataValidate(cboPrecioTipo, , True) Then Exit Sub
  If Not DataValidate(txtm315, "########.###", True) Then Exit Sub
  If Not DataValidate(txtm31556, "########.###", True) Then Exit Sub
  If Not DataValidate(txtBarrels60, "########.###", True) Then Exit Sub
  If Not DataValidate(txtAPIMinimo, "###.###") Then Exit Sub
  If Not DataValidate(txtAPIMaximo, "###.###") Then Exit Sub
  If Not DataValidate(txtAzuMinimo, "###.###") Then Exit Sub
  If Not DataValidate(txtAzuMaximo, "###.###") Then Exit Sub
  If Not DataValidate(TxtDescuento, "@100") Then Exit Sub
  If Not DataValidate(cboRangos, , True) Then Exit Sub
  If Not DataValidate(cboTipoCalculo, , True) Then Exit Sub
  If Not DataValidate(txtDiasPrevios, "##") Then Exit Sub
  If Not DataValidate(txtDiasPosteriores, "##") Then Exit Sub
  If Not DataValidate(cboIncluyeEntregaCli, , False) Then Exit Sub
  If Not DataValidate(cboMesPromedio, , False) Then Exit Sub
  If Not DataValidate(txtObservaciones, "@250") Then Exit Sub
  If Not DataValidate(cboAjustePrecio, , False) Then Exit Sub
  If Not DataValidate(cboAjusteApi, , False) Then Exit Sub
  If Not DataValidate(cboAjusteVarios, , False) Then Exit Sub
  If Not DataValidate(cboTipoCalculo, , False) Then Exit Sub
  If Not DataValidate(txtRedondeo1, "#", False) Then Exit Sub
  If Not DataValidate(txtRedondeo2, "#", False) Then Exit Sub
  If Not DataValidate(txtRedondeo3, "#", False) Then Exit Sub
  If Not DataValidate(txtRedondeo4, "#", False) Then Exit Sub
  If Not DataValidate(txtDesFormula, "@250") Then Exit Sub
  If Not DataValidate(txtDesDesde, "dd/mm/yyyy", True) Then Exit Sub
  If Not DataValidate(txtDesHasta, "dd/mm/yyyy", True) Then Exit Sub
  If Not DataValidate(txtDesPrevios, "##") Then Exit Sub
  If Not DataValidate(txtDesPoste, "##") Then Exit Sub
  If Not DataValidate(txtDesMeses, "#", False) Then Exit Sub
  
  blnAceptar = True
  blnCancelar = False
  Me.Hide

End Sub

Private Sub cmdCancelar_Click()

  blnAceptar = False
  blnCancelar = True
  Unload Me
  
End Sub

Private Sub Combo1_Change()

End Sub



Private Sub cmdNewEmpresa_Click()

  Dim strStore, strView, strDato As String

  strStore = "spEmpresasInsert"
  strView = "SELECT * FROM ViewEmpresas"
  strDato = ComboBoxAddItem(Me, cboEmpresa, "@50", strStore, strView)

End Sub

Private Sub cmdRangos_Click()
    
  Dim strAnt As String
  Dim intI As Integer
  
  contratosRangosFRM.Show vbModal
    
  'guardo valor del comboBox
  strAnt = Me.cboRangos
  
  'lleno combo box nuevamente
  intRes = llenoRangos()
  
  'set valor de comboBox anterior, primero busco si existe, para evitar error
  For intI = 0 To Me.cboRangos.ListCount - 1
    If strAnt = Me.cboRangos.List(intI) Then
      Me.cboRangos = strAnt
      Exit For
    End If
  Next
  
End Sub

Private Sub Command1_Click()

  Dim strStore, strView, strDato As String

  strStore = "spPreciosTiposInsert"
  strView = "SELECT * FROM ViewPreciosTipos"
  strDato = ComboBoxAddItem(Me, cboPrecioTipo, "@50", strStore, strView)

End Sub

Private Sub Form_Load()
  Dim strSQL As String
  Dim rsParam As ADODB.Recordset
  
  ' lleno combo Empresas
  strSQL = "SELECT * FROM ViewEmpresas"
  intRes = ComboBoxRefresh(cboEmpresa, strSQL)

  ' lleno combo Clientes
  strSQL = "SELECT * FROM ViewClientes"
  intRes = ComboBoxRefresh(cboCliente, strSQL)

  ' lleno combo tipos de precio
  strSQL = "SELECT * FROM ViewPreciosTipos"
  intRes = ComboBoxRefresh(cboPrecioTipo, strSQL)

  ' lleno combo formula precio
  strSQL = "SELECT * FROM parametrosFormulas_View where referencia like 'ajustePrecio%'"
  intRes = ComboBoxRefresh(cboAjustePrecio, strSQL)

  ' lleno combo formula api
  strSQL = "SELECT * FROM parametrosFormulas_View where referencia like 'ajusteAPI%'"
  intRes = ComboBoxRefresh(cboAjusteApi, strSQL)
  
  ' lleno combo formula varios
  strSQL = "SELECT * FROM parametrosFormulas_View where referencia like 'ajusteVarios%'"
  intRes = ComboBoxRefresh(cboAjusteVarios, strSQL)
  
  'lleno combo con tabla Rangos, lo hago de esta manera porque el comboBox no tiene ID
  intRes = llenoRangos()
      
End Sub

Private Sub txtm315_LostFocus()
  Me.txtm31556 = Round(Val(Me.txtm315) * getParam("m315TOm31556"), 3)
  Me.txtBarrels60 = Round(Val(Me.txtm31556) * getParam("m31556TOBarr1556"), 3)
End Sub

Private Sub txtm31556_LostFocus()
  Me.txtBarrels60 = Round(Val(Me.txtm31556) * getParam("m31556TOBarr1556"), 3)
End Sub
