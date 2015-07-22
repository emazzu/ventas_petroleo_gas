VERSION 5.00
Begin VB.Form frmEntregasTraInfo 
   BackColor       =   &H80000018&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Entregas Transportistas"
   ClientHeight    =   7275
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5880
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7275
   ScaleWidth      =   5880
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtTitulo 
      Appearance      =   0  'Flat
      BackColor       =   &H80000018&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      ForeColor       =   &H00808080&
      Height          =   265
      Index           =   20
      Left            =   135
      TabIndex        =   47
      Text            =   "Vol Mermas Tpte"
      Top             =   5805
      Width           =   1800
   End
   Begin VB.TextBox txtVolMermasTpte 
      BackColor       =   &H80000018&
      ForeColor       =   &H00FF0000&
      Height          =   285
      Left            =   1980
      TabIndex        =   18
      Top             =   5760
      Width           =   3760
   End
   Begin VB.TextBox txtTitulo 
      Appearance      =   0  'Flat
      BackColor       =   &H80000018&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      ForeColor       =   &H00808080&
      Height          =   265
      Index           =   19
      Left            =   135
      TabIndex        =   46
      Text            =   "Vol Mermas Otras 1"
      Top             =   6120
      Width           =   1800
   End
   Begin VB.TextBox txtVolMermasOtras1 
      BackColor       =   &H80000018&
      ForeColor       =   &H00FF0000&
      Height          =   285
      Left            =   1980
      TabIndex        =   19
      Top             =   6075
      Width           =   3760
   End
   Begin VB.TextBox txtTitulo 
      Appearance      =   0  'Flat
      BackColor       =   &H80000018&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      ForeColor       =   &H00808080&
      Height          =   265
      Index           =   18
      Left            =   135
      TabIndex        =   45
      Text            =   "Vol Mermas Otras 2"
      Top             =   5175
      Width           =   1800
   End
   Begin VB.TextBox txtVolMermasOtras2 
      BackColor       =   &H80000018&
      ForeColor       =   &H00FF0000&
      Height          =   285
      Left            =   1980
      TabIndex        =   16
      Top             =   5130
      Width           =   3760
   End
   Begin VB.TextBox txtTitulo 
      Appearance      =   0  'Flat
      BackColor       =   &H80000018&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      ForeColor       =   &H00808080&
      Height          =   265
      Index           =   17
      Left            =   135
      TabIndex        =   44
      Text            =   "Vol Agua + Sedimentos"
      Top             =   4860
      Width           =   1800
   End
   Begin VB.TextBox txtVolAguaSedim 
      BackColor       =   &H80000018&
      ForeColor       =   &H00FF0000&
      Height          =   285
      Left            =   1980
      TabIndex        =   15
      Top             =   4815
      Width           =   3760
   End
   Begin VB.TextBox txtTitulo 
      Appearance      =   0  'Flat
      BackColor       =   &H80000018&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      ForeColor       =   &H00808080&
      Height          =   270
      Index           =   16
      Left            =   135
      TabIndex        =   43
      Text            =   "Dens Seco Seco 15"
      Top             =   2025
      Width           =   1800
   End
   Begin VB.TextBox txtDensiSecoSeco15 
      BackColor       =   &H80000018&
      ForeColor       =   &H00808080&
      Height          =   285
      Left            =   1980
      TabIndex        =   6
      Top             =   1980
      Width           =   3760
   End
   Begin VB.TextBox txtNeto15 
      BackColor       =   &H80000018&
      ForeColor       =   &H00FF0000&
      Height          =   285
      Left            =   1980
      TabIndex        =   20
      Top             =   6390
      Width           =   3760
   End
   Begin VB.TextBox txtTitulo 
      Appearance      =   0  'Flat
      BackColor       =   &H80000018&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      ForeColor       =   &H00808080&
      Height          =   265
      Index           =   15
      Left            =   135
      TabIndex        =   42
      Text            =   "Vol Neto15"
      Top             =   6435
      Width           =   1800
   End
   Begin VB.TextBox txtSecoSeco15 
      BackColor       =   &H80000018&
      ForeColor       =   &H00FF0000&
      Height          =   285
      Left            =   1980
      TabIndex        =   17
      Top             =   5445
      Width           =   3760
   End
   Begin VB.TextBox txtTitulo 
      Appearance      =   0  'Flat
      BackColor       =   &H80000018&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      ForeColor       =   &H00808080&
      Height          =   265
      Index           =   14
      Left            =   135
      TabIndex        =   41
      Text            =   "Vol SecoSeco15"
      Top             =   5490
      Width           =   1800
   End
   Begin VB.TextBox txtTitulo 
      Appearance      =   0  'Flat
      BackColor       =   &H80000018&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      ForeColor       =   &H00808080&
      Height          =   265
      Index           =   13
      Left            =   135
      TabIndex        =   40
      Text            =   "MermasOtras2 %"
      Top             =   3600
      Width           =   1800
   End
   Begin VB.TextBox txtMermasOtras2 
      BackColor       =   &H80000018&
      ForeColor       =   &H00808080&
      Height          =   285
      Left            =   1980
      TabIndex        =   11
      Top             =   3555
      Width           =   3760
   End
   Begin VB.TextBox txtMermasOtras1 
      BackColor       =   &H80000018&
      ForeColor       =   &H00808080&
      Height          =   285
      Left            =   1980
      TabIndex        =   13
      Top             =   4185
      Width           =   3760
   End
   Begin VB.TextBox txtTitulo 
      Appearance      =   0  'Flat
      BackColor       =   &H80000018&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      ForeColor       =   &H00808080&
      Height          =   265
      Index           =   12
      Left            =   135
      TabIndex        =   39
      Text            =   "MermasOtras1 %"
      Top             =   4230
      Width           =   1800
   End
   Begin VB.TextBox txtAgua 
      BackColor       =   &H80000018&
      ForeColor       =   &H00808080&
      Height          =   285
      Left            =   1980
      TabIndex        =   9
      Top             =   2925
      Width           =   3760
   End
   Begin VB.TextBox txtFecha 
      BackColor       =   &H80000018&
      ForeColor       =   &H00808080&
      Height          =   285
      Left            =   1980
      TabIndex        =   3
      Top             =   990
      Width           =   3760
   End
   Begin VB.TextBox txtActa 
      BackColor       =   &H80000018&
      ForeColor       =   &H00808080&
      Height          =   285
      Left            =   1980
      TabIndex        =   2
      Top             =   720
      Width           =   3760
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
      Width           =   3255
   End
   Begin VB.CommandButton cmdNewEmpresa 
      Caption         =   "New"
      Height          =   285
      Left            =   5220
      TabIndex        =   38
      Top             =   90
      Width           =   510
   End
   Begin VB.TextBox txtAPI 
      BackColor       =   &H80000018&
      ForeColor       =   &H00FF0000&
      Height          =   285
      Left            =   1980
      TabIndex        =   7
      Top             =   2295
      Width           =   3760
   End
   Begin VB.ComboBox cboTransportista 
      Appearance      =   0  'Flat
      BackColor       =   &H80000018&
      ForeColor       =   &H00808080&
      Height          =   315
      Left            =   1980
      Sorted          =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   405
      Width           =   3255
   End
   Begin VB.ComboBox cboPuntoCarga 
      Appearance      =   0  'Flat
      BackColor       =   &H80000018&
      ForeColor       =   &H00808080&
      Height          =   315
      Left            =   1980
      Sorted          =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   4
      Top             =   1305
      Width           =   3255
   End
   Begin VB.TextBox txtSedimento 
      BackColor       =   &H80000018&
      ForeColor       =   &H00808080&
      Height          =   285
      Left            =   1980
      TabIndex        =   10
      Top             =   3240
      Width           =   3760
   End
   Begin VB.TextBox txtVolHidratado15 
      BackColor       =   &H80000018&
      ForeColor       =   &H00808080&
      Height          =   285
      Left            =   1980
      TabIndex        =   14
      Top             =   4500
      Width           =   3760
   End
   Begin VB.CommandButton cmdCancelar 
      Caption         =   "&Cancelar"
      Height          =   330
      Left            =   2475
      TabIndex        =   22
      Top             =   6795
      Width           =   1500
   End
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "&Aceptar"
      Height          =   330
      Left            =   4230
      TabIndex        =   21
      Top             =   6780
      Width           =   1500
   End
   Begin VB.TextBox txtSalinidad 
      BackColor       =   &H80000018&
      ForeColor       =   &H00808080&
      Height          =   285
      Left            =   1980
      TabIndex        =   8
      Top             =   2610
      Width           =   3760
   End
   Begin VB.TextBox txtTitulo 
      Appearance      =   0  'Flat
      BackColor       =   &H80000018&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      ForeColor       =   &H00808080&
      Height          =   265
      Index           =   0
      Left            =   135
      TabIndex        =   37
      Text            =   "Empresa"
      Top             =   135
      Width           =   1800
   End
   Begin VB.TextBox txtTitulo 
      Appearance      =   0  'Flat
      BackColor       =   &H80000018&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      ForeColor       =   &H00808080&
      Height          =   265
      Index           =   1
      Left            =   135
      TabIndex        =   36
      Text            =   "Transportista"
      Top             =   450
      Width           =   1800
   End
   Begin VB.TextBox txtTitulo 
      Appearance      =   0  'Flat
      BackColor       =   &H80000018&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      ForeColor       =   &H00808080&
      Height          =   265
      Index           =   2
      Left            =   135
      TabIndex        =   35
      Text            =   "Acta"
      Top             =   720
      Width           =   1800
   End
   Begin VB.TextBox txtTitulo 
      Appearance      =   0  'Flat
      BackColor       =   &H80000018&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      ForeColor       =   &H00808080&
      Height          =   265
      Index           =   3
      Left            =   135
      TabIndex        =   34
      Text            =   "Fecha"
      Top             =   1035
      Width           =   1800
   End
   Begin VB.TextBox txtTitulo 
      Appearance      =   0  'Flat
      BackColor       =   &H80000018&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      ForeColor       =   &H00808080&
      Height          =   265
      Index           =   4
      Left            =   135
      TabIndex        =   33
      Text            =   "PuntoCarga"
      Top             =   1350
      Width           =   1800
   End
   Begin VB.TextBox txtTitulo 
      Appearance      =   0  'Flat
      BackColor       =   &H80000018&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      ForeColor       =   &H00808080&
      Height          =   265
      Index           =   5
      Left            =   135
      TabIndex        =   32
      Text            =   "TerminalAlmacenaje"
      Top             =   1665
      Width           =   1800
   End
   Begin VB.TextBox txtTitulo 
      Appearance      =   0  'Flat
      BackColor       =   &H80000018&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      ForeColor       =   &H00808080&
      Height          =   265
      Index           =   6
      Left            =   135
      TabIndex        =   31
      Text            =   "Agua %"
      Top             =   2970
      Width           =   1800
   End
   Begin VB.TextBox txtTitulo 
      Appearance      =   0  'Flat
      BackColor       =   &H80000018&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      ForeColor       =   &H00808080&
      Height          =   265
      Index           =   7
      Left            =   135
      TabIndex        =   30
      Text            =   "Sedimento %"
      Top             =   3285
      Width           =   1800
   End
   Begin VB.TextBox txtTitulo 
      Appearance      =   0  'Flat
      BackColor       =   &H80000018&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      ForeColor       =   &H00808080&
      Height          =   265
      Index           =   8
      Left            =   135
      TabIndex        =   29
      Text            =   "API"
      Top             =   2340
      Width           =   1800
   End
   Begin VB.TextBox txtTitulo 
      Appearance      =   0  'Flat
      BackColor       =   &H80000018&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      ForeColor       =   &H00808080&
      Height          =   265
      Index           =   9
      Left            =   135
      TabIndex        =   28
      Text            =   "Salinidad"
      Top             =   2655
      Width           =   1800
   End
   Begin VB.TextBox txtTitulo 
      Appearance      =   0  'Flat
      BackColor       =   &H80000018&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      ForeColor       =   &H00808080&
      Height          =   265
      Index           =   10
      Left            =   135
      TabIndex        =   27
      Text            =   "VolHidratado15"
      Top             =   4545
      Width           =   1800
   End
   Begin VB.TextBox txtTitulo 
      Appearance      =   0  'Flat
      BackColor       =   &H80000018&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      ForeColor       =   &H00808080&
      Height          =   265
      Index           =   11
      Left            =   135
      TabIndex        =   26
      Text            =   "MermasTpte %"
      Top             =   3915
      Width           =   1800
   End
   Begin VB.ComboBox cboTerminal 
      Appearance      =   0  'Flat
      BackColor       =   &H80000018&
      ForeColor       =   &H00808080&
      Height          =   315
      Left            =   1980
      Sorted          =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   5
      Top             =   1620
      Width           =   3255
   End
   Begin VB.TextBox txtMermasTpte 
      BackColor       =   &H80000018&
      ForeColor       =   &H00808080&
      Height          =   285
      Left            =   1980
      TabIndex        =   12
      Top             =   3870
      Width           =   3760
   End
   Begin VB.CommandButton cmdNewTransportista 
      Caption         =   "New"
      Height          =   285
      Left            =   5220
      TabIndex        =   25
      Top             =   405
      Width           =   510
   End
   Begin VB.CommandButton cmdNewPuntoCarga 
      Caption         =   "New"
      Height          =   285
      Left            =   5220
      TabIndex        =   24
      Top             =   1305
      Width           =   510
   End
   Begin VB.CommandButton cmdNewPuntoDevolucion 
      Caption         =   "New"
      Height          =   285
      Left            =   5220
      TabIndex        =   23
      Top             =   1620
      Width           =   510
   End
End
Attribute VB_Name = "frmEntregasTraInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim oControl As Object
  
Function calculos()
  Dim PAgua, PSedimentos, PMermastpte, PMermasotras1, PMermasotras2 As Double
  Dim PDensiToAPI1, PDensiToAPI2, PDensiToAPI3 As Double
  Dim VSecoseco15, VNeto15, VDensiSecoSeco15, VAPI As Double
  
  ' tomo porcentajes de controles y se los paso a variables del tipo adecuado
  VHidratado15 = Val(Me.txtVolHidratado15)
  VDensiSecoSeco15 = Val(Me.txtDensiSecoSeco15)
  PAgua = Val(Me.txtAgua)
  PSedimentos = Val(Me.txtSedimento)
  PMermasotras2 = Val(Me.txtmermasOtras2)
  PMermastpte = Val(Me.txtMermasTpte)
  PMermasotras1 = Val(Me.txtMermasOtras1)
    
  ' tomo parametros segun puntoCarga
  PDensiToAPI1 = 0
  PDensiToAPI2 = 0
  PDensiToAPI3 = 0
  If cboPuntoCarga.ListIndex >= 0 Then
    
    Dim rs As ADODB.Recordset
    strSQL = "select * from cargas where idcarga = " & cboPuntoCarga.ItemData(cboPuntoCarga.ListIndex)
    Set rs = adoGetRS(strSQL)
    If Not rs.EOF Then
      PDensiToAPI1 = rs!densiToAPI1
      PDensiToAPI2 = rs!densiToAPI2
      PDensiToAPI3 = rs!densiToAPI3
    End If
    
  End If
  
  ' calculo volumenes
  If VDensiSecoSeco15 + PDensiToAPI2 <> 0 Then
    VAPI = Round(PDensiToAPI1 / (VDensiSecoSeco15 + PDensiToAPI2) - PDensiToAPI3, 2)
  End If
  Me.txtAPI = Format(VAPI, "##0.000")
  
  Dim VAguaSedim, VMermasOtras2, VMermastpte, VMermasOtras1 As Double
  
  VAguaSedim = Round(VHidratado15 * (PAgua + PSedimentos) / 100, 3)
  Me.txtVolAguaSedim = Format(VAguaSedim, "###0.000")
  
  VMermasOtras2 = Round((VHidratado15 - VAguaSedim) * PMermasotras2 / 100, 3)
  Me.txtVolMermasOtras2 = Format(VMermasOtras2, "###0.000")
  
  VSecoseco15 = Round(VHidratado15 - VAguaSedim - VMermasOtras2, 3)
  Me.txtSecoSeco15 = Format(VSecoseco15, "#######0.000")
  
  VMermastpte = Round(VSecoseco15 * PMermastpte / 100, 3)
  Me.txtVolMermasTpte = Format(VMermastpte, "###0.000")
  
  VMermasOtras1 = Round(VSecoseco15 * PMermasotras1 / 100, 3)
  Me.txtVolMermasOtras1 = Format(VMermasOtras1, "###0.000")
  
  VNeto15 = Round(VSecoseco15 - VMermastpte - VMermasOtras1, 3)
  Me.txtNeto15 = Format(VNeto15, "#######0.000")

End Function

Private Sub cboPuntoCarga_LostFocus()
  intRes = calculos()
End Sub

Private Sub cmdAceptar_Click()
  
  If Not DataValidate(cboEmpresa, , True) Then Exit Sub
  If Not DataValidate(cboTransportista, , True) Then Exit Sub
  If Not DataValidate(txtActa, "@10", True) Then Exit Sub
  If Not DataValidate(txtFecha, "dd/mm/yyyy", True) Then Exit Sub
  If Not DataValidate(cboPuntoCarga, , True) Then Exit Sub
  If Not DataValidate(cboTerminal, , True) Then Exit Sub
  If Not DataValidate(txtAgua, "##.######", True) Then Exit Sub
  If Not DataValidate(txtSedimento, "##.######", True) Then Exit Sub
  If Not DataValidate(txtmermasOtras2, "##.###") Then Exit Sub
  If Not DataValidate(txtDensiSecoSeco15, "#.######", True) Then Exit Sub
  If Not DataValidate(txtAPI, "###.###", True) Then Exit Sub
  If Not DataValidate(txtSalinidad, "####.###", True) Then Exit Sub
  If Not DataValidate(txtVolHidratado15, "#########.###", True) Then Exit Sub
  If Not DataValidate(txtMermasTpte, "##.####") Then Exit Sub
  If Not DataValidate(txtMermasOtras1, "##.###") Then Exit Sub
  If Not DataValidate(txtSecoSeco15, "########.###") Then Exit Sub
  If Not DataValidate(txtNeto15, "########.###") Then Exit Sub
  
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

Private Sub cmdNewPuntoCarga_Click()
  Dim strStore, strView, strDato As String
  
  strStore = "spCargasInsert"
  strView = "SELECT * FROM ViewCargas"
  strDato = ComboBoxAddItem(Me, cboPuntoCarga, "@50", strStore, strView)

End Sub

Private Sub cmdNewPuntoDevolucion_Click()
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

  strSQL = "SELECT * FROM ViewCargas"
  intResul = ComboBoxRefresh(cboPuntoCarga, strSQL)

  strSQL = "SELECT * FROM ViewTerminales"
  intResul = ComboBoxRefresh(cboTerminal, strSQL)

End Sub

Private Sub txtAgua_LostFocus()
  intRes = calculos()
End Sub


Private Sub txtDensiSecoSeco15_LostFocus()
  intRes = calculos()
End Sub

Private Sub txtMermasOtras1_LostFocus()
  intRes = calculos()
End Sub

Private Sub txtmermasOtras2_LostFocus()
  intRes = calculos()
End Sub

Private Sub txtMermasTpte_LostFocus()
  intRes = calculos()
End Sub

Private Sub txtSedimento_LostFocus()
  intRes = calculos
End Sub

Private Sub txtVolHidratado15_LostFocus()
  intRes = calculos()
End Sub
