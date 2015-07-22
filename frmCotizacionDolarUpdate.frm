VERSION 5.00
Begin VB.Form frmCotizacionDolarUpdate 
   BackColor       =   &H80000018&
   BorderStyle     =   0  'None
   Caption         =   "Cotizacion Dolar Update"
   ClientHeight    =   480
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6735
   LinkTopic       =   "Form1"
   ScaleHeight     =   480
   ScaleWidth      =   6735
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdCancelar 
      Caption         =   "&Cancelar"
      Height          =   300
      Left            =   5745
      TabIndex        =   5
      Top             =   90
      Width           =   885
   End
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "&Aceptar"
      Height          =   300
      Left            =   4860
      TabIndex        =   4
      Top             =   90
      Width           =   885
   End
   Begin VB.TextBox Text3 
      BackColor       =   &H00C0FFFF&
      Enabled         =   0   'False
      ForeColor       =   &H00808080&
      Height          =   300
      Left            =   2445
      TabIndex        =   3
      Text            =   "Vendedor"
      Top             =   90
      Width           =   1050
   End
   Begin VB.TextBox Text2 
      BackColor       =   &H00C0FFFF&
      Enabled         =   0   'False
      ForeColor       =   &H00808080&
      Height          =   300
      Left            =   90
      TabIndex        =   2
      Text            =   "Comprador"
      Top             =   90
      Width           =   1005
   End
   Begin VB.TextBox txtDato2 
      BackColor       =   &H00C0FFFF&
      ForeColor       =   &H00808080&
      Height          =   300
      Left            =   3510
      TabIndex        =   1
      Top             =   90
      Width           =   1300
   End
   Begin VB.TextBox txtDato1 
      BackColor       =   &H00C0FFFF&
      ForeColor       =   &H00808080&
      Height          =   300
      Left            =   1110
      TabIndex        =   0
      Top             =   90
      Width           =   1300
   End
End
Attribute VB_Name = "frmCotizacionDolarUpdate"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmdAceptar_Click()
  Dim intRes As Integer

  If Not DataValidate(Me.txtDato1, "###.###", True) Then Exit Sub
  If Not DataValidate(Me.txtDato2, "###.###", True) Then Exit Sub

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

Private Sub txtDato2_KeyPress(KeyAscii As Integer)

  If KeyAscii = 27 Then       ' escapar
    blnAceptar = False
    blnCancelar = True
    Me.Hide
    Exit Sub
  End If

End Sub

