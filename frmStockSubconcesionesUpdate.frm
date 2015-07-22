VERSION 5.00
Begin VB.Form frmStockSubconcesionesUpdate 
   BackColor       =   &H80000018&
   BorderStyle     =   0  'None
   Caption         =   "Stock Subconcesiones "
   ClientHeight    =   525
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   8925
   LinkTopic       =   "Form2"
   ScaleHeight     =   525
   ScaleWidth      =   8925
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtDato1 
      BackColor       =   &H00C0FFFF&
      ForeColor       =   &H00808080&
      Height          =   330
      Left            =   5220
      TabIndex        =   2
      Top             =   90
      Width           =   1725
   End
   Begin VB.TextBox Text2 
      BackColor       =   &H00C0FFFF&
      Enabled         =   0   'False
      ForeColor       =   &H00808080&
      Height          =   330
      Left            =   3690
      TabIndex        =   1
      Text            =   " Ingresos / Egresos"
      Top             =   90
      Width           =   1500
   End
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "&Aceptar"
      Height          =   300
      Left            =   6960
      TabIndex        =   3
      Top             =   90
      Width           =   930
   End
   Begin VB.CommandButton cmdCancelar 
      Caption         =   "&Cancelar"
      Height          =   300
      Left            =   7905
      TabIndex        =   4
      Top             =   90
      Width           =   885
   End
   Begin VB.TextBox txtSubconcesion 
      BackColor       =   &H00C0FFFF&
      Enabled         =   0   'False
      ForeColor       =   &H00808080&
      Height          =   330
      Left            =   90
      TabIndex        =   0
      Top             =   90
      Width           =   3570
   End
End
Attribute VB_Name = "frmStockSubconcesionesUpdate"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdAceptar_Click()
  Dim intRes As Integer

  If Not DataValidate(Me.txtDato1, "########.###-", True) Then Exit Sub

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

