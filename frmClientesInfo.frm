VERSION 5.00
Begin VB.Form frmClientesInfo 
   BackColor       =   &H00C0E0FF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Clientes"
   ClientHeight    =   4200
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6210
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4200
   ScaleWidth      =   6210
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtReferencia 
      BackColor       =   &H00C0E0FF&
      ForeColor       =   &H00808080&
      Height          =   285
      Left            =   1485
      TabIndex        =   1
      Top             =   450
      Width           =   4605
   End
   Begin VB.TextBox Text12 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0E0FF&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      ForeColor       =   &H00808080&
      Height          =   285
      Left            =   135
      Locked          =   -1  'True
      TabIndex        =   24
      Text            =   "Referencia"
      Top             =   450
      Width           =   1300
   End
   Begin VB.TextBox Text11 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0E0FF&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      ForeColor       =   &H00808080&
      Height          =   330
      Left            =   135
      Locked          =   -1  'True
      TabIndex        =   23
      Text            =   "Iva RG 3337"
      Top             =   2700
      Width           =   1300
   End
   Begin VB.ComboBox cboRg3337 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0E0FF&
      ForeColor       =   &H00808080&
      Height          =   315
      ItemData        =   "frmClientesInfo.frx":0000
      Left            =   1485
      List            =   "frmClientesInfo.frx":000A
      Style           =   2  'Dropdown List
      TabIndex        =   8
      Top             =   2700
      Width           =   4605
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0E0FF&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      ForeColor       =   &H00808080&
      Height          =   285
      Left            =   135
      Locked          =   -1  'True
      TabIndex        =   22
      Text            =   "Cuit"
      Top             =   2025
      Width           =   1300
   End
   Begin VB.CommandButton cmdCancelar 
      Caption         =   "&Cancelar"
      Height          =   330
      Left            =   4275
      TabIndex        =   21
      Top             =   3780
      Width           =   800
   End
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "&Aceptar"
      Height          =   330
      Left            =   5310
      TabIndex        =   20
      Top             =   3780
      Width           =   800
   End
   Begin VB.TextBox txtCliente 
      BackColor       =   &H00C0E0FF&
      ForeColor       =   &H00808080&
      Height          =   285
      Left            =   1485
      TabIndex        =   0
      Top             =   135
      Width           =   4605
   End
   Begin VB.ComboBox cboCondicionIva 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0E0FF&
      ForeColor       =   &H00808080&
      Height          =   315
      ItemData        =   "frmClientesInfo.frx":0016
      Left            =   1485
      List            =   "frmClientesInfo.frx":0026
      Style           =   2  'Dropdown List
      TabIndex        =   7
      Top             =   2340
      Width           =   4605
   End
   Begin VB.TextBox txtDiasVentas 
      BackColor       =   &H00C0E0FF&
      ForeColor       =   &H00808080&
      Height          =   285
      Left            =   1485
      TabIndex        =   9
      Top             =   3060
      Width           =   4605
   End
   Begin VB.ComboBox cboExportacion 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0E0FF&
      ForeColor       =   &H00808080&
      Height          =   315
      IntegralHeight  =   0   'False
      ItemData        =   "frmClientesInfo.frx":0068
      Left            =   1485
      List            =   "frmClientesInfo.frx":0072
      Style           =   2  'Dropdown List
      TabIndex        =   10
      Top             =   3375
      Width           =   4605
   End
   Begin VB.TextBox txtDomicilio 
      BackColor       =   &H00C0E0FF&
      ForeColor       =   &H00808080&
      Height          =   285
      Left            =   1485
      TabIndex        =   2
      Top             =   765
      Width           =   4605
   End
   Begin VB.TextBox txtCodigoPostal 
      BackColor       =   &H00C0E0FF&
      ForeColor       =   &H00808080&
      Height          =   285
      Left            =   1485
      TabIndex        =   3
      Top             =   1080
      Width           =   4605
   End
   Begin VB.TextBox txtLocalidad 
      BackColor       =   &H00C0E0FF&
      ForeColor       =   &H00808080&
      Height          =   285
      Left            =   1485
      TabIndex        =   4
      Top             =   1395
      Width           =   4605
   End
   Begin VB.TextBox txtPais 
      BackColor       =   &H00C0E0FF&
      ForeColor       =   &H00808080&
      Height          =   285
      Left            =   1485
      TabIndex        =   5
      Top             =   1710
      Width           =   4605
   End
   Begin VB.TextBox txtCuit 
      BackColor       =   &H00C0E0FF&
      ForeColor       =   &H00808080&
      Height          =   285
      Left            =   1485
      TabIndex        =   6
      Top             =   2025
      Width           =   4605
   End
   Begin VB.TextBox text2 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0E0FF&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      ForeColor       =   &H00808080&
      Height          =   285
      Left            =   135
      Locked          =   -1  'True
      TabIndex        =   19
      Text            =   "Cliente"
      Top             =   135
      Width           =   1300
   End
   Begin VB.TextBox Text3 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0E0FF&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      ForeColor       =   &H00808080&
      Height          =   285
      Left            =   135
      Locked          =   -1  'True
      TabIndex        =   18
      Text            =   "Domicilio"
      Top             =   765
      Width           =   1300
   End
   Begin VB.TextBox Text4 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0E0FF&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      ForeColor       =   &H00808080&
      Height          =   285
      Left            =   135
      Locked          =   -1  'True
      TabIndex        =   17
      Text            =   "CodigoPostal"
      Top             =   1080
      Width           =   1300
   End
   Begin VB.TextBox Text5 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0E0FF&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      ForeColor       =   &H00808080&
      Height          =   285
      Left            =   135
      Locked          =   -1  'True
      TabIndex        =   16
      Text            =   "Localidad"
      Top             =   1395
      Width           =   1300
   End
   Begin VB.TextBox Text6 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0E0FF&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      ForeColor       =   &H00808080&
      Height          =   285
      Left            =   135
      Locked          =   -1  'True
      TabIndex        =   15
      Text            =   "Pais"
      Top             =   1710
      Width           =   1300
   End
   Begin VB.TextBox Text7 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0E0FF&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      ForeColor       =   &H00808080&
      Height          =   285
      Left            =   -3240
      Locked          =   -1  'True
      TabIndex        =   14
      Text            =   "Cuit"
      Top             =   -3915
      Width           =   1600
   End
   Begin VB.TextBox Text8 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0E0FF&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      ForeColor       =   &H00808080&
      Height          =   330
      Left            =   135
      Locked          =   -1  'True
      TabIndex        =   13
      Text            =   "CondiciónIva"
      Top             =   2340
      Width           =   1300
   End
   Begin VB.TextBox Text9 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0E0FF&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      ForeColor       =   &H00808080&
      Height          =   285
      Left            =   135
      Locked          =   -1  'True
      TabIndex        =   12
      Text            =   "DiasVentas"
      Top             =   3060
      Width           =   1300
   End
   Begin VB.TextBox Text10 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0E0FF&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      ForeColor       =   &H00808080&
      Height          =   330
      Left            =   135
      Locked          =   -1  'True
      TabIndex        =   11
      Text            =   "Exportación"
      Top             =   3375
      Width           =   1300
   End
End
Attribute VB_Name = "frmClientesInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdAceptar_Click()

  ' validacion de datos
  
  ' validacion de datos
  
  If Not DataValidate(txtCliente, "@150", True) Then Exit Sub
  If Not DataValidate(txtDomicilio, "@100", True) Then Exit Sub
  If Not DataValidate(txtReferencia, "@50", True) Then Exit Sub
  If Not DataValidate(txtCodigoPostal, "@20", True) Then Exit Sub
  If Not DataValidate(txtLocalidad, "@50", True) Then Exit Sub
  If Not DataValidate(txtPais, "@30", True) Then Exit Sub
  If Not DataValidate(txtCuit, "@13", True) Then Exit Sub
  If Not DataValidate(cboCondicionIva, , True) Then Exit Sub
  If Not DataValidate(cboRg3337, , True) Then Exit Sub
  If Not DataValidate(txtDiasVentas, "##") Then Exit Sub
  If Not DataValidate(cboExportacion, , True) Then Exit Sub
  
  If (txtDiasVentas = "") Then
    txtDiasVentas = "0"
  End If
  
  blnAceptar = True
  blnCancelar = False
  Me.Hide

End Sub

Private Sub cmdCancelar_Click()

  blnAceptar = False
  blnCancelar = True
  Unload Me

End Sub

