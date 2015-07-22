VERSION 5.00
Begin VB.Form frmSubConcesionesInfo 
   BackColor       =   &H80000018&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "SubConcesiones"
   ClientHeight    =   3750
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5805
   FillColor       =   &H00404040&
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3750
   ScaleWidth      =   5805
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdNewEmpresa 
      Caption         =   "New"
      Height          =   290
      Left            =   5130
      TabIndex        =   25
      Top             =   135
      Width           =   555
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
      Top             =   135
      Width           =   3120
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
      TabIndex        =   24
      Text            =   "Empresa"
      Top             =   135
      Width           =   1800
   End
   Begin VB.CommandButton cmdNewTerminal 
      Caption         =   "New"
      Height          =   290
      Left            =   5130
      TabIndex        =   23
      Top             =   1485
      Width           =   555
   End
   Begin VB.ComboBox cboTerminal 
      Appearance      =   0  'Flat
      BackColor       =   &H80000018&
      ForeColor       =   &H00808080&
      Height          =   315
      Left            =   1980
      Sorted          =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   4
      Top             =   1485
      Width           =   3120
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
      TabIndex        =   22
      Text            =   "Terminal"
      Top             =   1485
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
      TabIndex        =   21
      Text            =   "Provincia"
      Top             =   2115
      Width           =   1800
   End
   Begin VB.ComboBox cboProvincia 
      Appearance      =   0  'Flat
      BackColor       =   &H80000018&
      ForeColor       =   &H00808080&
      Height          =   315
      Left            =   1980
      Sorted          =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   6
      Top             =   2115
      Width           =   3120
   End
   Begin VB.CommandButton cmdNewProvincia 
      Caption         =   "New"
      Height          =   290
      Left            =   5130
      TabIndex        =   20
      Top             =   2115
      Width           =   555
   End
   Begin VB.CommandButton cmdNewCarga 
      Caption         =   "New"
      Height          =   290
      Left            =   5130
      TabIndex        =   19
      Top             =   1800
      Width           =   555
   End
   Begin VB.ComboBox cboPuntoCarga 
      Appearance      =   0  'Flat
      BackColor       =   &H80000018&
      ForeColor       =   &H00808080&
      Height          =   315
      Left            =   1980
      Sorted          =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   5
      Top             =   1800
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
      Left            =   135
      TabIndex        =   18
      Text            =   "PuntoCarga"
      Top             =   1800
      Width           =   1800
   End
   Begin VB.TextBox txtPAPCtrl 
      Appearance      =   0  'Flat
      BackColor       =   &H80000018&
      BorderStyle     =   0  'None
      ForeColor       =   &H00808080&
      Height          =   285
      Left            =   1980
      TabIndex        =   7
      Top             =   2475
      Width           =   3700
   End
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "&Aceptar"
      Height          =   330
      Left            =   4200
      TabIndex        =   9
      Top             =   3240
      Width           =   1500
   End
   Begin VB.CommandButton cmdCancelar 
      Caption         =   "&Cancelar"
      Height          =   330
      Left            =   2445
      TabIndex        =   10
      Top             =   3240
      Width           =   1500
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
      TabIndex        =   17
      Text            =   "SubConcesion"
      Top             =   495
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
      TabIndex        =   16
      Text            =   "Concesión"
      Top             =   855
      Width           =   1800
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
      TabIndex        =   15
      Text            =   "Area"
      Top             =   1170
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
      TabIndex        =   14
      Text            =   "PAPPath"
      Top             =   2790
      Width           =   1800
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
      TabIndex        =   13
      Text            =   "PAPCtrl"
      Top             =   2475
      Width           =   1800
   End
   Begin VB.ComboBox cboConcesion 
      Appearance      =   0  'Flat
      BackColor       =   &H80000018&
      ForeColor       =   &H00808080&
      Height          =   315
      Left            =   1980
      Sorted          =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   2
      Top             =   855
      Width           =   3120
   End
   Begin VB.ComboBox cboArea 
      Appearance      =   0  'Flat
      BackColor       =   &H80000018&
      ForeColor       =   &H00808080&
      Height          =   315
      Left            =   1980
      Sorted          =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   3
      Top             =   1170
      Width           =   3120
   End
   Begin VB.TextBox txtSubConcesion 
      Appearance      =   0  'Flat
      BackColor       =   &H80000018&
      BorderStyle     =   0  'None
      ForeColor       =   &H00808080&
      Height          =   285
      Left            =   1980
      TabIndex        =   1
      Top             =   495
      Width           =   3700
   End
   Begin VB.TextBox txtPAPPath 
      Appearance      =   0  'Flat
      BackColor       =   &H80000018&
      BorderStyle     =   0  'None
      ForeColor       =   &H00808080&
      Height          =   285
      Left            =   1980
      TabIndex        =   8
      Top             =   2790
      Width           =   3700
   End
   Begin VB.CommandButton cmdNewConcesion 
      Caption         =   "New"
      Height          =   290
      Left            =   5130
      TabIndex        =   12
      Top             =   855
      Width           =   555
   End
   Begin VB.CommandButton cmdNewArea 
      Caption         =   "New"
      Height          =   290
      Left            =   5130
      TabIndex        =   11
      Top             =   1170
      Width           =   555
   End
End
Attribute VB_Name = "frmSubConcesionesInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdAceptar_Click()

  ' validacion de datos
  
  If Not DataValidate(cboEmpresa, "@50", True) Then Exit Sub
  If Not DataValidate(txtSubConcesion, "@50", True) Then Exit Sub
  If Not DataValidate(cboConcesion, , True) Then Exit Sub
  If Not DataValidate(cboArea, , True) Then Exit Sub
  If Not DataValidate(cboPuntoCarga, , True) Then Exit Sub
  If Not DataValidate(cboTerminal, , True) Then Exit Sub
  If Not DataValidate(cboProvincia, , True) Then Exit Sub
  If Not DataValidate(txtPAPCtrl, "##") Then Exit Sub
  If Not DataValidate(txtPAPPath, "@100") Then Exit Sub
  
  blnAceptar = True
  blnCancelar = False
  Me.Hide

End Sub

Private Sub cmdCancelar_Click()

  blnAceptar = False
  blnCancelar = True
  Unload Me

End Sub

Private Sub cmdNewArea_Click()

  Dim strStore, strView, strDato As String

  strStore = "spAreasInsert"
  strView = "SELECT * FROM ViewAreas"
  strDato = ComboBoxAddItem(Me, cboArea, "@50", strStore, strView)

End Sub

Private Sub cmdNewCarga_Click()
  Dim strStore, strView, strDato As String

  strStore = "spCargasInsert"
  strView = "SELECT * FROM ViewCargas"
  strDato = ComboBoxAddItem(Me, cboPuntoCarga, "@50", strStore, strView)

End Sub

Private Sub cmdNewConcesion_Click()

  Dim strStore, strView, strDato As String

  strStore = "spConcesionesInsert"
  strView = "SELECT * FROM ViewConcesiones"
  strDato = ComboBoxAddItem(Me, cboConcesion, "@50", strStore, strView)

End Sub

Private Sub cmdNewEmpresa_Click()
  Dim strStore, strView, strDato As String

  strStore = "spEmpresasInsert"
  strView = "SELECT * FROM ViewEmpresas"
  strDato = ComboBoxAddItem(Me, cboEmpresa, "@50", strStore, strView)

End Sub

Private Sub cmdNewProvincia_Click()
  Dim strStore, strView, strDato As String

  strStore = "spProvinciasInsert"
  strView = "SELECT * FROM ViewProvincias"
  strDato = ComboBoxAddItem(Me, cboProvincia, "@50", strStore, strView)

End Sub

Private Sub cmdNewterminal_Click()
  Dim strStore, strView, strDato As String

  strStore = "spTerminalesInsert"
  strView = "SELECT * FROM ViewTerminales"
  strDato = ComboBoxAddItem(Me, cboTerminal, "@50", strStore, strView)

End Sub

Private Sub Form_Load()
  
  strSQL = "SELECT * FROM ViewEmpresas"
  intResul = ComboBoxRefresh(cboEmpresa, strSQL)
  
  strSQL = "SELECT * FROM ViewConcesiones"
  intResul = ComboBoxRefresh(cboConcesion, strSQL)

  strSQL = "SELECT * FROM ViewAreas"
  intResul = ComboBoxRefresh(cboArea, strSQL)

  strSQL = "SELECT * FROM ViewCargas"
  intResul = ComboBoxRefresh(cboPuntoCarga, strSQL)
  
  strSQL = "SELECT * FROM ViewProvincias"
  intResul = ComboBoxRefresh(cboProvincia, strSQL)
  
  strSQL = "SELECT * FROM ViewTerminales"
  intResul = ComboBoxRefresh(cboTerminal, strSQL)

End Sub
