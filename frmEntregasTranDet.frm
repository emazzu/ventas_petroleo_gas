VERSION 5.00
Begin VB.Form frmEntregasTraDet 
   BackColor       =   &H80000018&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Entregas Transportistas Detalle"
   ClientHeight    =   3555
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5880
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3555
   ScaleWidth      =   5880
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtmermasOtras2 
      BackColor       =   &H80000018&
      ForeColor       =   &H00808080&
      Height          =   330
      Left            =   2025
      TabIndex        =   3
      Top             =   1170
      Width           =   3700
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
      TabIndex        =   18
      Text            =   "MermasOtras2"
      Top             =   1215
      Width           =   1800
   End
   Begin VB.TextBox txtEmpresaID 
      Height          =   285
      Left            =   225
      TabIndex        =   17
      Top             =   3060
      Visible         =   0   'False
      Width           =   960
   End
   Begin VB.TextBox txtVolumenNeto15 
      BackColor       =   &H80000018&
      ForeColor       =   &H00808080&
      Height          =   330
      Left            =   2025
      TabIndex        =   7
      Top             =   2610
      Width           =   3700
   End
   Begin VB.ComboBox cboConcesion 
      Appearance      =   0  'Flat
      BackColor       =   &H80000018&
      ForeColor       =   &H00808080&
      Height          =   315
      Left            =   2025
      Sorted          =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   135
      Width           =   3705
   End
   Begin VB.TextBox txtTitulo 
      Appearance      =   0  'Flat
      BackColor       =   &H80000018&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      ForeColor       =   &H00808080&
      Height          =   285
      Index           =   11
      Left            =   180
      TabIndex        =   16
      Text            =   "VolumenNeto15"
      Top             =   2655
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
      Left            =   180
      TabIndex        =   15
      Text            =   "MermasOtras1"
      Top             =   2295
      Width           =   1800
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
      TabIndex        =   14
      Text            =   "MermasTpte"
      Top             =   1935
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
      Left            =   180
      TabIndex        =   13
      Text            =   "VolumenSecoSeco15"
      Top             =   1575
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
      Left            =   180
      TabIndex        =   12
      Text            =   "VolumenAguaSedimento"
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
      Index           =   6
      Left            =   180
      TabIndex        =   11
      Text            =   "VolumenHidratado15"
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
      Index           =   5
      Left            =   180
      TabIndex        =   10
      Text            =   "Concesión"
      Top             =   135
      Width           =   1800
   End
   Begin VB.TextBox txtMermasTpte 
      BackColor       =   &H80000018&
      ForeColor       =   &H00808080&
      Height          =   330
      Left            =   2025
      TabIndex        =   5
      Top             =   1890
      Width           =   3700
   End
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "&Aceptar"
      Height          =   330
      Left            =   4230
      TabIndex        =   8
      Top             =   3060
      Width           =   1500
   End
   Begin VB.CommandButton cmdCancelar 
      Caption         =   "&Cancelar"
      Height          =   330
      Left            =   2475
      TabIndex        =   9
      Top             =   3060
      Width           =   1500
   End
   Begin VB.TextBox txtMermasOtras1 
      BackColor       =   &H80000018&
      ForeColor       =   &H00808080&
      Height          =   330
      Left            =   2025
      TabIndex        =   6
      Top             =   2250
      Width           =   3700
   End
   Begin VB.TextBox txtVolumenAguaSedimento 
      BackColor       =   &H80000018&
      ForeColor       =   &H00808080&
      Height          =   330
      Left            =   2025
      TabIndex        =   2
      Top             =   810
      Width           =   3700
   End
   Begin VB.TextBox txtVolumenSecoSeco15 
      BackColor       =   &H80000018&
      ForeColor       =   &H00808080&
      Height          =   330
      Left            =   2025
      TabIndex        =   4
      Top             =   1530
      Width           =   3700
   End
   Begin VB.TextBox txtVolumenHidratado15 
      BackColor       =   &H80000018&
      ForeColor       =   &H00808080&
      Height          =   330
      Left            =   2025
      TabIndex        =   1
      Top             =   450
      Width           =   3700
   End
End
Attribute VB_Name = "frmEntregasTraDet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdAceptar_Click()
  
  If Not DataValidate(Me.cboConcesion, , True) Then Exit Sub
  If Not DataValidate(Me.txtVolumenHidratado15, "########.###", True) Then Exit Sub
  If Not DataValidate(Me.txtVolumenAguaSedimento, "######.###", True) Then Exit Sub
  If Not DataValidate(Me.txtmermasOtras2, "######.###", True) Then Exit Sub
  If Not DataValidate(Me.txtVolumenSecoSeco15, "########.###", True) Then Exit Sub
  If Not DataValidate(Me.txtMermasTpte, "######.####", True) Then Exit Sub
  If Not DataValidate(Me.txtMermasOtras1, "######.###", True) Then Exit Sub
  If Not DataValidate(Me.txtVolumenNeto15, "########.###", True) Then Exit Sub
  
  blnAceptar = True
  blnCancelar = False
  Me.Hide

End Sub

Private Sub cmdCancelar_Click()

  blnAceptar = False
  blnCancelar = True
  Unload Me

End Sub

Private Sub txtEmpresaID_Change()

  strSQL = "SELECT * FROM ViewConcesionesxEmpresa WHERE empresaID = " & Me.txtEmpresaID
  intRes = ComboBoxRefresh(cboConcesion, strSQL)

End Sub

Private Sub txtVolumenHidratado15_LostFocus()
  Dim PAgua, PSedimentos As Double
  Dim VHidratado15, VAguasedimentos, VSecoseco15, VMermasOtras2 As Currency
  Dim PMermastpte, PMermasotras, VMermastpte, VMermasOtras1, VNeto15 As Currency

  ' tomo porcentajes de listview y se los paso a variables del tipo adecuado
  PAgua = CDbl(lvwGetValue(frmEntregasTra.lvwDatos, "agua"))
  PSedimentos = CDbl(lvwGetValue(frmEntregasTra.lvwDatos, "sedimento"))
  PMermasotras2 = CCur(lvwGetValue(frmEntregasTra.lvwDatos, "mermasotras2"))
  PMermastpte = CCur(lvwGetValue(frmEntregasTra.lvwDatos, "mermastpte"))
  PMermasotras1 = CCur(lvwGetValue(frmEntregasTra.lvwDatos, "mermasotras1"))
  
  ' calculo volumenes
  VHidratado15 = Round(CCur(Me.txtVolumenHidratado15), 3)
  VAguasedimentos = Round(VHidratado15 * (PAgua + PSedimentos) / 100, 3)
  VMermasOtras2 = Round((CCur(Me.txtVolumenHidratado15) * (1 - ((PAgua + PSedimentos) / 100))) * PMermasotras2 / 100, 3)
  VSecoseco15 = Round(VHidratado15 - VAguasedimentos - VMermasOtras2, 3)
  VMermastpte = Round(VSecoseco15 * PMermastpte / 100, 3)
  VMermasotras = Round(VSecoseco15 * PMermasotras / 100, 3)
  VNeto15 = Round(VSecoseco15 - (VMermastpte + VMermasotras), 3)
  
  Me.txtVolumenHidratado15 = Format(VHidratado15, "#######0.000")
  Me.txtVolumenAguaSedimento = Format(VAguasedimentos, "#######0.000")
  Me.txtmermasOtras2 = Format(VMermasOtras2, "#######0.000")
  Me.txtVolumenSecoSeco15 = Format(VSecoseco15, "#######0.000")
  Me.txtMermasTpte = Format(VMermastpte, "#######0.0000")
  Me.txtMermasOtras1 = Format(Round(VMermasOtras1, 3), "#######0.000")
  Me.txtVolumenNeto15 = Format(VNeto15, "#######0.000")

End Sub
