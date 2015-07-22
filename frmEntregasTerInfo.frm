VERSION 5.00
Begin VB.Form frmEntregasTerInfo 
   BackColor       =   &H80000018&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Entregas Terminales"
   ClientHeight    =   3255
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5850
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3255
   ScaleWidth      =   5850
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtAPICoefAju 
      BackColor       =   &H80000018&
      ForeColor       =   &H00808080&
      Height          =   285
      Left            =   1980
      Locked          =   -1  'True
      TabIndex        =   7
      Top             =   2430
      Width           =   3750
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
      TabIndex        =   20
      Text            =   "API Coeficiente Ajuste"
      Top             =   2430
      Width           =   1800
   End
   Begin VB.TextBox txtPjeMermas 
      BackColor       =   &H80000018&
      ForeColor       =   &H00808080&
      Height          =   285
      Left            =   1980
      Locked          =   -1  'True
      TabIndex        =   6
      Top             =   2070
      Width           =   3750
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
      TabIndex        =   19
      Text            =   "(%) Mermas"
      Top             =   2070
      Width           =   1800
   End
   Begin VB.CommandButton cmdNewTerminal 
      Caption         =   "New"
      Height          =   285
      Left            =   5250
      TabIndex        =   18
      Top             =   750
      Width           =   510
   End
   Begin VB.CommandButton cmdNewTransportista 
      Caption         =   "New"
      Height          =   285
      Left            =   5250
      TabIndex        =   17
      Top             =   420
      Width           =   510
   End
   Begin VB.TextBox txtTitulo 
      Appearance      =   0  'Flat
      BackColor       =   &H80000018&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      ForeColor       =   &H00808080&
      Height          =   285
      Index           =   8
      Left            =   120
      TabIndex        =   16
      Text            =   "API"
      Top             =   1740
      Width           =   1800
   End
   Begin VB.TextBox txtTitulo 
      Appearance      =   0  'Flat
      BackColor       =   &H80000018&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      ForeColor       =   &H00808080&
      Height          =   285
      Index           =   7
      Left            =   120
      TabIndex        =   15
      Text            =   "Fecha"
      Top             =   1425
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
      Left            =   120
      TabIndex        =   14
      Text            =   "Certificado"
      Top             =   1110
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
      Left            =   120
      TabIndex        =   13
      Text            =   "Terminal"
      Top             =   750
      Width           =   1800
   End
   Begin VB.TextBox txtTitulo 
      Appearance      =   0  'Flat
      BackColor       =   &H80000018&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      ForeColor       =   &H00808080&
      Height          =   285
      Index           =   1
      Left            =   120
      TabIndex        =   12
      Text            =   "Transportista"
      Top             =   420
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
      Left            =   120
      TabIndex        =   11
      Text            =   "Empresa"
      Top             =   90
      Width           =   1800
   End
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "&Aceptar"
      Height          =   330
      Left            =   4245
      TabIndex        =   8
      Top             =   2820
      Width           =   1500
   End
   Begin VB.CommandButton cmdCancelar 
      Caption         =   "&Cancelar"
      Height          =   330
      Left            =   2490
      TabIndex        =   9
      Top             =   2820
      Width           =   1500
   End
   Begin VB.TextBox txtFecha 
      BackColor       =   &H80000018&
      ForeColor       =   &H00808080&
      Height          =   285
      Left            =   1965
      TabIndex        =   4
      Top             =   1425
      Width           =   3795
   End
   Begin VB.ComboBox cboTerminal 
      Appearance      =   0  'Flat
      BackColor       =   &H80000018&
      ForeColor       =   &H00808080&
      Height          =   315
      Left            =   1965
      Sorted          =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   2
      Top             =   750
      Width           =   3255
   End
   Begin VB.ComboBox cboTransportista 
      Appearance      =   0  'Flat
      BackColor       =   &H80000018&
      ForeColor       =   &H00808080&
      Height          =   315
      Left            =   1965
      Sorted          =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   420
      Width           =   3255
   End
   Begin VB.TextBox txtAPI 
      BackColor       =   &H80000018&
      ForeColor       =   &H00808080&
      Height          =   285
      Left            =   1965
      TabIndex        =   5
      Top             =   1740
      Width           =   3795
   End
   Begin VB.CommandButton cmdNewEmpresa 
      Caption         =   "New"
      Height          =   285
      Left            =   5250
      TabIndex        =   10
      Top             =   90
      Width           =   510
   End
   Begin VB.ComboBox cboEmpresa 
      Appearance      =   0  'Flat
      BackColor       =   &H80000018&
      ForeColor       =   &H00808080&
      Height          =   315
      Left            =   1965
      Sorted          =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   90
      Width           =   3255
   End
   Begin VB.TextBox txtCertificado 
      BackColor       =   &H80000018&
      ForeColor       =   &H00808080&
      Height          =   285
      Left            =   1965
      TabIndex        =   3
      Top             =   1110
      Width           =   3795
   End
End
Attribute VB_Name = "frmEntregasTerInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cboTerminal_Click()

'tomo el (%) de mermas de la terminal seleccioanda
If cboTerminal.ListIndex <> -1 Then
  
  Dim rs As ADODB.Recordset
  strSQL = "select pjeMermas from terminales where IDTerminal = " & cboTerminal.ItemData(cboTerminal.ListIndex)
  Set rs = adoGetRS(strSQL)
  
  If Not rs.EOF Then
    Me.txtPjeMermas = rs!PjeMermas
  End If
  
End If

End Sub

Private Sub cmdAceptar_Click()
  
  If Not DataValidate(cboEmpresa, , True) Then Exit Sub
  If Not DataValidate(cboTransportista, , True) Then Exit Sub
  If Not DataValidate(cboTerminal, , True) Then Exit Sub
  If Not DataValidate(txtCertificado, "@10", True) Then Exit Sub
  If Not DataValidate(txtFecha, "dd/mm/yyyy", True) Then Exit Sub
  If Not DataValidate(txtAPI, "###.#####", True) Then Exit Sub
  If Not DataValidate(Me.txtPjeMermas, "##.###", False) Then Exit Sub
  If Not DataValidate(Me.txtAPICoefAju, "##.###", False) Then Exit Sub
  
  blnAceptar = True
  blnCancelar = False
  Me.Hide

End Sub

Private Sub cmdCancelar_Click()

  blnAceptar = False
  blnCancelar = True
  Unload Me

End Sub

Private Sub cmdNewEmpresa_Click()
  Dim strStore, strView, strDato As String

  strStore = "spEmpresasInsert"
  strView = "SELECT * FROM ViewEmpresas"
  strDato = ComboBoxAddItem(Me, cboEmpresa, "@50", strStore, strView)
 
End Sub

Private Sub cmdNewterminal_Click()
  Dim strStore, strView, strDato As String

  strStore = "spTerminalesInsert"
  strView = "SELECT * FROM ViewTerminales"
  strDato = ComboBoxAddItem(Me, cboTerminal, "@50", strStore, strView)

End Sub

Private Sub cmdNewTransportista_Click()
  Dim strStore, strView, strDato As String

  strStore = "spTransportistasInsert"
  strView = "SELECT * FROM ViewTransportistas"
  strDato = ComboBoxAddItem(Me, cboTransportista, "@50", strStore, strView)

End Sub

Private Sub Form_Load()

  Dim strSQL As String
  Dim intResul As Integer
  
  strSQL = "SELECT * FROM ViewEmpresas"
  intResul = ComboBoxRefresh(cboEmpresa, strSQL)

  strSQL = "SELECT * FROM ViewTransportistas"
  intResul = ComboBoxRefresh(cboTransportista, strSQL)

  strSQL = "SELECT * FROM ViewTerminales"
  intResul = ComboBoxRefresh(cboTerminal, strSQL)
  
  'si esta vacio APICoeficienteAju lo tomo de parametros
  If Me.txtAPICoefAju = "" Then
    Me.txtAPICoefAju = getParam("apiCoeficienteAju")
  End If
  
End Sub


