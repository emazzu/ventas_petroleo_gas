VERSION 5.00
Begin VB.Form frmTerminalParamUpdate 
   BackColor       =   &H80000018&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   1770
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3615
   LinkTopic       =   "Form1"
   ScaleHeight     =   1770
   ScaleWidth      =   3615
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtDato2 
      BackColor       =   &H00C0FFFF&
      ForeColor       =   &H00808080&
      Height          =   300
      Left            =   90
      TabIndex        =   3
      Top             =   1020
      Width           =   3360
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFFF&
      Enabled         =   0   'False
      ForeColor       =   &H00808080&
      Height          =   300
      Left            =   90
      TabIndex        =   2
      Text            =   "Tipo Oil"
      Top             =   720
      Width           =   3345
   End
   Begin VB.TextBox txtDato1 
      BackColor       =   &H00C0FFFF&
      ForeColor       =   &H00808080&
      Height          =   300
      Left            =   90
      TabIndex        =   1
      Top             =   420
      Width           =   3360
   End
   Begin VB.TextBox Text2 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFFF&
      Enabled         =   0   'False
      ForeColor       =   &H00808080&
      Height          =   300
      Left            =   90
      TabIndex        =   0
      Text            =   " Valor"
      Top             =   120
      Width           =   3345
   End
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "&Aceptar"
      Height          =   300
      Left            =   1650
      TabIndex        =   4
      Top             =   1365
      Width           =   885
   End
   Begin VB.CommandButton cmdCancelar 
      Caption         =   "&Cancelar"
      Height          =   300
      Left            =   2580
      TabIndex        =   5
      Top             =   1365
      Width           =   885
   End
End
Attribute VB_Name = "frmTerminalParamUpdate"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdAceptar_Click()
  Dim intRes As Integer

  If Not DataValidate(Me.txtDato1, "####.#######", True) Then Exit Sub
  If Not DataValidate(Me.txtDato2, "@50", True) Then Exit Sub

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


