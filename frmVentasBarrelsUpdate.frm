VERSION 5.00
Begin VB.Form frmVentasBarrelsUpdate 
   BackColor       =   &H00C0E0FF&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Asignando volumen"
   ClientHeight    =   1725
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   3105
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1725
   ScaleWidth      =   3105
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtVolBarrels 
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H00808080&
      Height          =   300
      Left            =   1395
      TabIndex        =   2
      Top             =   855
      Width           =   1575
   End
   Begin VB.TextBox txtVolM31556 
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H00808080&
      Height          =   300
      Left            =   1395
      TabIndex        =   1
      Top             =   495
      Width           =   1575
   End
   Begin VB.TextBox Text2 
      BackColor       =   &H00FFFFFF&
      Enabled         =   0   'False
      ForeColor       =   &H00808080&
      Height          =   300
      Left            =   135
      Locked          =   -1  'True
      TabIndex        =   7
      Text            =   "m3 15.56"
      Top             =   495
      Width           =   1230
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00FFFFFF&
      Enabled         =   0   'False
      ForeColor       =   &H00808080&
      Height          =   300
      Left            =   135
      Locked          =   -1  'True
      TabIndex        =   6
      Text            =   "m3 15"
      Top             =   135
      Width           =   1230
   End
   Begin VB.CommandButton cmdCancelar 
      Caption         =   "&Cancelar"
      Height          =   285
      Left            =   2160
      TabIndex        =   4
      Top             =   1305
      Width           =   795
   End
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "&Aceptar"
      Height          =   285
      Left            =   1200
      TabIndex        =   3
      Top             =   1305
      Width           =   795
   End
   Begin VB.TextBox txtTitulo 
      BackColor       =   &H00FFFFFF&
      Enabled         =   0   'False
      ForeColor       =   &H00808080&
      Height          =   300
      Left            =   135
      Locked          =   -1  'True
      TabIndex        =   5
      Text            =   "Barrels"
      Top             =   855
      Width           =   1230
   End
   Begin VB.TextBox txtVolM315 
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H00808080&
      Height          =   300
      Left            =   1395
      TabIndex        =   0
      Top             =   135
      Width           =   1575
   End
End
Attribute VB_Name = "frmVentasBarrelsUpdate"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmdAceptar_Click()

  blnAceptar = True
  blnCancelar = False
  Me.Hide

End Sub

Private Sub cmdCancelar_Click()
  
  blnAceptar = False
  blnCancelar = True
  Me.Hide

End Sub

Private Sub txtDato1_KeyPress(KeyAscii As Integer)

  If KeyAscii = 27 Then       ' escapar
    blnAceptar = False
    blnCancelar = True
    Me.Hide
    Exit Sub
  End If

End Sub

Private Sub txtVolBarrels_LostFocus()

  'si se selecciono barrels calculo vol 15 y vol 1556
  If frmVentasInfo.cboBase.List(frmVentasInfo.cboBase.ListIndex) = "Barrels" Then
    
    txtVolM31556 = Round(CSng(Val(txtVolBarrels)) / CSng(getParam("m31556TObarr1556")), 3)
    txtVolM315 = Round(CSng(Val(txtVolM31556)) / CSng(getParam("m315TOm31556")), 3)
  
  End If

End Sub

Private Sub txtVolM315_LostFocus()

  'si se selecciono m3 calculo Vol 1556 o Barriles
  If frmVentasInfo.cboBase.List(frmVentasInfo.cboBase.ListIndex) = "M3" Then
    
    txtVolM31556 = Round(CSng(Val(txtVolM315)) * CSng(getParam("m315TOm31556")), 3)
    txtVolBarrels = Round(CSng(Val(txtVolM31556)) * CSng(getParam("m31556TObarr1556")), 3)
    
  End If

End Sub
