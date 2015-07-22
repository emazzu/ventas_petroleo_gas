VERSION 5.00
Begin VB.Form frmSubconcesionParamUpdate 
   BackColor       =   &H80000018&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Parametros PorSubconcesion"
   ClientHeight    =   5190
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5820
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5190
   ScaleWidth      =   5820
   StartUpPosition =   2  'CenterScreen
   Begin VB.ComboBox cboAjuApiStk 
      Appearance      =   0  'Flat
      BackColor       =   &H80000018&
      ForeColor       =   &H00808080&
      Height          =   315
      ItemData        =   "frmSubconcesionParamUpdate.frx":0000
      Left            =   1935
      List            =   "frmSubconcesionParamUpdate.frx":000A
      Style           =   2  'Dropdown List
      TabIndex        =   4
      Top             =   1395
      Width           =   3795
   End
   Begin VB.TextBox txtTitulo 
      Appearance      =   0  'Flat
      BackColor       =   &H80000018&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      ForeColor       =   &H00808080&
      Height          =   285
      Index           =   13
      Left            =   135
      TabIndex        =   29
      Text            =   "Ajuste API Stock"
      Top             =   1395
      Width           =   1800
   End
   Begin VB.TextBox txtTitulo 
      Appearance      =   0  'Flat
      BackColor       =   &H80000018&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      ForeColor       =   &H00808080&
      Height          =   285
      Index           =   10
      Left            =   135
      TabIndex        =   28
      Text            =   "Embarque"
      Top             =   3645
      Width           =   1800
   End
   Begin VB.TextBox txtEmbarque 
      Appearance      =   0  'Flat
      BackColor       =   &H80000018&
      BorderStyle     =   0  'None
      ForeColor       =   &H00808080&
      Height          =   285
      Left            =   1980
      TabIndex        =   11
      Top             =   3645
      Width           =   3700
   End
   Begin VB.TextBox txtAlicRetExporta 
      Appearance      =   0  'Flat
      BackColor       =   &H80000018&
      BorderStyle     =   0  'None
      ForeColor       =   &H00808080&
      Height          =   285
      Left            =   1980
      TabIndex        =   12
      Top             =   3960
      Width           =   3700
   End
   Begin VB.TextBox txtTitulo 
      Appearance      =   0  'Flat
      BackColor       =   &H80000018&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      ForeColor       =   &H00808080&
      Height          =   285
      Index           =   9
      Left            =   135
      TabIndex        =   27
      Text            =   "AlicRetExportaciones"
      Top             =   3960
      Width           =   1800
   End
   Begin VB.TextBox txtTitulo 
      Appearance      =   0  'Flat
      BackColor       =   &H80000018&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      ForeColor       =   &H00808080&
      Height          =   285
      Index           =   8
      Left            =   135
      TabIndex        =   26
      Text            =   "MercadoLocal"
      Top             =   4275
      Width           =   1800
   End
   Begin VB.TextBox txtMercadoLocal 
      Appearance      =   0  'Flat
      BackColor       =   &H80000018&
      BorderStyle     =   0  'None
      ForeColor       =   &H00808080&
      Height          =   285
      Left            =   1980
      TabIndex        =   13
      Top             =   4275
      Width           =   3700
   End
   Begin VB.TextBox TxtDescuento 
      Appearance      =   0  'Flat
      BackColor       =   &H80000018&
      BorderStyle     =   0  'None
      ForeColor       =   &H00808080&
      Height          =   285
      Left            =   1980
      TabIndex        =   6
      Top             =   2070
      Width           =   3700
   End
   Begin VB.TextBox txtTitulo 
      Appearance      =   0  'Flat
      BackColor       =   &H80000018&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      ForeColor       =   &H00808080&
      Height          =   285
      Index           =   7
      Left            =   135
      TabIndex        =   25
      Text            =   "Transporte"
      Top             =   2700
      Width           =   1800
   End
   Begin VB.TextBox txtTitulo 
      Appearance      =   0  'Flat
      BackColor       =   &H80000018&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      ForeColor       =   &H00808080&
      Height          =   285
      Index           =   6
      Left            =   135
      TabIndex        =   24
      Text            =   "Descuento"
      Top             =   2070
      Width           =   1800
   End
   Begin VB.TextBox txtTransporte 
      Appearance      =   0  'Flat
      BackColor       =   &H80000018&
      BorderStyle     =   0  'None
      ForeColor       =   &H00808080&
      Height          =   285
      Left            =   1980
      TabIndex        =   8
      Top             =   2700
      Width           =   3700
   End
   Begin VB.TextBox txtIngBrutos 
      Appearance      =   0  'Flat
      BackColor       =   &H80000018&
      BorderStyle     =   0  'None
      ForeColor       =   &H00808080&
      Height          =   285
      Left            =   1980
      TabIndex        =   7
      Top             =   2385
      Width           =   3700
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
      TabIndex        =   23
      Text            =   "IngresosBrutos"
      Top             =   2385
      Width           =   1800
   End
   Begin VB.TextBox txtAlmacen 
      Appearance      =   0  'Flat
      BackColor       =   &H80000018&
      BorderStyle     =   0  'None
      ForeColor       =   &H00808080&
      Height          =   285
      Left            =   1980
      TabIndex        =   9
      Top             =   3015
      Width           =   3700
   End
   Begin VB.TextBox txtTitulo 
      Appearance      =   0  'Flat
      BackColor       =   &H80000018&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      ForeColor       =   &H00808080&
      Height          =   285
      Index           =   4
      Left            =   135
      TabIndex        =   22
      Text            =   "Almacenamiento"
      Top             =   3015
      Width           =   1800
   End
   Begin VB.TextBox txtTitulo 
      Appearance      =   0  'Flat
      BackColor       =   &H80000018&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      ForeColor       =   &H00808080&
      Height          =   285
      Index           =   2
      Left            =   135
      TabIndex        =   21
      Text            =   "DiasAlmacenamiento"
      Top             =   3330
      Width           =   1800
   End
   Begin VB.TextBox txtDiasAlmacen 
      Appearance      =   0  'Flat
      BackColor       =   &H80000018&
      BorderStyle     =   0  'None
      ForeColor       =   &H00808080&
      Height          =   285
      Left            =   1980
      TabIndex        =   10
      Top             =   3330
      Width           =   3700
   End
   Begin VB.TextBox txtParticOil 
      Appearance      =   0  'Flat
      BackColor       =   &H80000018&
      BorderStyle     =   0  'None
      ForeColor       =   &H00808080&
      Height          =   285
      Left            =   1980
      TabIndex        =   0
      Top             =   135
      Width           =   3700
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
      TabIndex        =   20
      Text            =   "PjeSubconcesEntregas"
      Top             =   765
      Width           =   1800
   End
   Begin VB.TextBox txtTitulo 
      Appearance      =   0  'Flat
      BackColor       =   &H80000018&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      ForeColor       =   &H00808080&
      Height          =   285
      Index           =   0
      Left            =   135
      TabIndex        =   19
      Text            =   "ParticipacionOil"
      Top             =   135
      Width           =   1800
   End
   Begin VB.CommandButton cmdCancelar 
      Caption         =   "&Cancelar"
      Height          =   330
      Left            =   2445
      TabIndex        =   15
      Top             =   4725
      Width           =   1500
   End
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "&Aceptar"
      Height          =   330
      Left            =   4200
      TabIndex        =   14
      Top             =   4725
      Width           =   1500
   End
   Begin VB.TextBox txtPjeSubEnt 
      Appearance      =   0  'Flat
      BackColor       =   &H80000018&
      BorderStyle     =   0  'None
      ForeColor       =   &H00808080&
      Height          =   285
      Left            =   1980
      TabIndex        =   2
      Top             =   765
      Width           =   3700
   End
   Begin VB.TextBox txtParticGas 
      Appearance      =   0  'Flat
      BackColor       =   &H80000018&
      BorderStyle     =   0  'None
      ForeColor       =   &H00808080&
      Height          =   285
      Left            =   1980
      TabIndex        =   1
      Top             =   450
      Width           =   3700
   End
   Begin VB.TextBox txtTitulo 
      Appearance      =   0  'Flat
      BackColor       =   &H80000018&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      ForeColor       =   &H00808080&
      Height          =   285
      Index           =   3
      Left            =   135
      TabIndex        =   18
      Text            =   "ParticipacionGas"
      Top             =   450
      Width           =   1800
   End
   Begin VB.TextBox txtPjeTerEnt 
      Appearance      =   0  'Flat
      BackColor       =   &H80000018&
      BorderStyle     =   0  'None
      ForeColor       =   &H00808080&
      Height          =   285
      Left            =   1980
      TabIndex        =   3
      Top             =   1080
      Width           =   3700
   End
   Begin VB.TextBox txtTitulo 
      Appearance      =   0  'Flat
      BackColor       =   &H80000018&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      ForeColor       =   &H00808080&
      Height          =   285
      Index           =   11
      Left            =   135
      TabIndex        =   17
      Text            =   "PjeTerminalEntregas"
      Top             =   1080
      Width           =   1800
   End
   Begin VB.TextBox txtTitulo 
      Appearance      =   0  'Flat
      BackColor       =   &H80000018&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      ForeColor       =   &H00808080&
      Height          =   285
      Index           =   12
      Left            =   135
      TabIndex        =   16
      Text            =   "PrecioVenta"
      Top             =   1755
      Width           =   1800
   End
   Begin VB.TextBox txtPrecioVenta 
      Appearance      =   0  'Flat
      BackColor       =   &H80000018&
      BorderStyle     =   0  'None
      ForeColor       =   &H00808080&
      Height          =   285
      Left            =   1980
      TabIndex        =   5
      Top             =   1755
      Width           =   3700
   End
End
Attribute VB_Name = "frmSubconcesionParamUpdate"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdAceptar_Click()

  ' validacion de datos
  If Not DataValidate(txtParticOil, "###.###") Then Exit Sub
  If Not DataValidate(txtParticGas, "###.###") Then Exit Sub
  If Not DataValidate(txtPjeSubEnt, "###.######") Then Exit Sub
  If Not DataValidate(txtPrecioVenta, "###.###") Then Exit Sub
  If Not DataValidate(cboAjuApiStk, , True) Then Exit Sub
  If Not DataValidate(TxtDescuento, "###.###") Then Exit Sub
  If Not DataValidate(txtIngBrutos, "###.###") Then Exit Sub
  If Not DataValidate(txtTransporte, "###.###") Then Exit Sub
  If Not DataValidate(txtAlmacen, "###.####") Then Exit Sub
  If Not DataValidate(txtDiasAlmacen, "###") Then Exit Sub
  If Not DataValidate(txtEmbarque, "###.###") Then Exit Sub
  If Not DataValidate(txtAlicRetExporta, "###.###") Then Exit Sub
  If Not DataValidate(txtMercadoLocal, "###.####") Then Exit Sub
  
  blnAceptar = True
  blnCancelar = False
  Me.Hide

End Sub

Private Sub cmdCancelar_Click()

  blnAceptar = False
  blnCancelar = True
  Unload Me

End Sub

