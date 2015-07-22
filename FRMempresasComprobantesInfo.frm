VERSION 5.00
Begin VB.Form FRMEmpresasComprobantesInfo 
   BackColor       =   &H00C0E0FF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Comprobantes"
   ClientHeight    =   3150
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6450
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3150
   ScaleWidth      =   6450
   StartUpPosition =   2  'CenterScreen
   Begin VB.ComboBox cboTipo 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0E0FF&
      ForeColor       =   &H00808080&
      Height          =   315
      ItemData        =   "FRMempresasComprobantesInfo.frx":0000
      Left            =   1710
      List            =   "FRMempresasComprobantesInfo.frx":0010
      Style           =   2  'Dropdown List
      TabIndex        =   2
      Top             =   765
      Width           =   4605
   End
   Begin VB.TextBox txtVencimiento 
      BackColor       =   &H00C0E0FF&
      ForeColor       =   &H00808080&
      Height          =   285
      Left            =   1710
      TabIndex        =   7
      Top             =   2340
      Width           =   4605
   End
   Begin VB.TextBox txtCAI 
      BackColor       =   &H00C0E0FF&
      ForeColor       =   &H00808080&
      Height          =   285
      Left            =   1710
      TabIndex        =   6
      Top             =   2025
      Width           =   4605
   End
   Begin VB.TextBox txtFecha 
      BackColor       =   &H00C0E0FF&
      ForeColor       =   &H00808080&
      Height          =   285
      Left            =   1710
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
      TabIndex        =   18
      Text            =   "Fecha"
      Top             =   450
      Width           =   1530
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
      TabIndex        =   17
      Text            =   "Vencimiento"
      Top             =   2385
      Width           =   1530
   End
   Begin VB.CommandButton cmdCancelar 
      Caption         =   "&Cancelar"
      Height          =   330
      Left            =   4455
      TabIndex        =   16
      Top             =   2745
      Width           =   800
   End
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "&Aceptar"
      Height          =   330
      Left            =   5490
      TabIndex        =   15
      Top             =   2745
      Width           =   800
   End
   Begin VB.ComboBox cboEmpresa 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0E0FF&
      ForeColor       =   &H00808080&
      Height          =   315
      ItemData        =   "FRMempresasComprobantesInfo.frx":0020
      Left            =   1710
      List            =   "FRMempresasComprobantesInfo.frx":0030
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   135
      Width           =   4605
   End
   Begin VB.TextBox txtSucursal 
      BackColor       =   &H00C0E0FF&
      ForeColor       =   &H00808080&
      Height          =   285
      Left            =   1710
      TabIndex        =   3
      Top             =   1080
      Width           =   4605
   End
   Begin VB.TextBox txtNumeroDesde 
      BackColor       =   &H00C0E0FF&
      ForeColor       =   &H00808080&
      Height          =   285
      Left            =   1710
      TabIndex        =   4
      Top             =   1395
      Width           =   4605
   End
   Begin VB.TextBox txtNumeroHasta 
      BackColor       =   &H00C0E0FF&
      ForeColor       =   &H00808080&
      Height          =   285
      Left            =   1710
      TabIndex        =   5
      Top             =   1710
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
      TabIndex        =   14
      Text            =   "Empresa"
      Top             =   135
      Width           =   1530
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
      TabIndex        =   13
      Text            =   "Tipo"
      Top             =   765
      Width           =   1530
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
      TabIndex        =   12
      Text            =   "Sucursal"
      Top             =   1080
      Width           =   1530
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
      TabIndex        =   11
      Text            =   "Numeración Desde"
      Top             =   1395
      Width           =   1530
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
      TabIndex        =   10
      Text            =   "Numeración Hasta"
      Top             =   1710
      Width           =   1530
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
      TabIndex        =   9
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
      TabIndex        =   8
      Text            =   "CAI"
      Top             =   2025
      Width           =   1530
   End
End
Attribute VB_Name = "FRMEmpresasComprobantesInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdAceptar_Click()

  ' validacion de datos
  
  ' validacion de datos
  
  If Not DataValidate(cboEmpresa, , True) Then Exit Sub
  If Not DataValidate(txtFecha, "dd/mm/yyyy", True) Then Exit Sub
  If Not DataValidate(cboTipo, , True) Then Exit Sub
  If Not DataValidate(txtSucursal, "####", True) Then Exit Sub
  If Not DataValidate(txtNumeroDesde, "########", True) Then Exit Sub
  If Not DataValidate(txtNumeroHasta, "########", True) Then Exit Sub
  If Not DataValidate(txtCAI, "@16", True) Then Exit Sub
  If Not DataValidate(txtVencimiento, "dd/mm/yyyy", True) Then Exit Sub
  
  blnAceptar = True
  blnCancelar = False
  Me.Hide

End Sub

Private Sub cmdCancelar_Click()

  blnAceptar = False
  blnCancelar = True
  Unload Me

End Sub

Private Sub Form_Load()
  
  strSQL = "SELECT * FROM maeEmpresas_vw"
  intRes = ComboBoxRefresh(cboEmpresa, strSQL)
  
End Sub
