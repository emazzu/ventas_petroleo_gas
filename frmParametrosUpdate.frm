VERSION 5.00
Begin VB.Form frmParametrosUpdate 
   BackColor       =   &H80000018&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Parametros"
   ClientHeight    =   2130
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   12675
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2130
   ScaleWidth      =   12675
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtDato2 
      BackColor       =   &H00C0FFFF&
      ForeColor       =   &H00808080&
      Height          =   300
      Left            =   90
      TabIndex        =   3
      Top             =   1350
      Width           =   12495
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFFF&
      Enabled         =   0   'False
      ForeColor       =   &H00808080&
      Height          =   300
      Left            =   90
      TabIndex        =   2
      Text            =   "Valor"
      Top             =   1020
      Width           =   12480
   End
   Begin VB.CommandButton cmdCancelar 
      Caption         =   "&Cancelar"
      Height          =   300
      Left            =   11715
      TabIndex        =   5
      Top             =   1755
      Width           =   885
   End
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "&Aceptar"
      Height          =   300
      Left            =   10755
      TabIndex        =   4
      Top             =   1755
      Width           =   885
   End
   Begin VB.TextBox Text2 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFFF&
      Enabled         =   0   'False
      ForeColor       =   &H00808080&
      Height          =   300
      Left            =   90
      TabIndex        =   0
      Text            =   "Fórmula"
      Top             =   75
      Width           =   12480
   End
   Begin VB.TextBox txtDato1 
      BackColor       =   &H00C0FFFF&
      ForeColor       =   &H00808080&
      Height          =   600
      Left            =   90
      MaxLength       =   500
      MultiLine       =   -1  'True
      TabIndex        =   1
      Top             =   405
      Width           =   12495
   End
End
Attribute VB_Name = "frmParametrosUpdate"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmdAceptar_Click()
  
  Dim strT As String
  
  If Not DataValidate(Me.txtDato1, "@500", True) Then Exit Sub
  If Not DataValidate(Me.txtDato2, "#####.######", True) Then Exit Sub
  
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

