VERSION 5.00
Begin VB.Form frmAddTipoItem 
   BackColor       =   &H00C0E0FF&
   Caption         =   "agregando Tipo de Item"
   ClientHeight    =   1065
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   3765
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   ScaleHeight     =   1065
   ScaleWidth      =   3765
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C0E0FF&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      ForeColor       =   &H00808080&
      Height          =   195
      Left            =   2190
      Locked          =   -1  'True
      TabIndex        =   5
      Text            =   "Referencia"
      Top             =   90
      Width           =   1485
   End
   Begin VB.TextBox text2 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C0E0FF&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      ForeColor       =   &H00808080&
      Height          =   195
      Left            =   90
      Locked          =   -1  'True
      TabIndex        =   4
      Text            =   "Tipo de Item"
      Top             =   90
      Width           =   2025
   End
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "&Aceptar"
      Height          =   300
      Left            =   2895
      TabIndex        =   3
      Top             =   690
      Width           =   800
   End
   Begin VB.CommandButton cmdCancelar 
      Caption         =   "&Cancelar"
      Height          =   300
      Left            =   1890
      TabIndex        =   2
      Top             =   690
      Width           =   800
   End
   Begin VB.TextBox txtTipoItemCorto 
      BackColor       =   &H00C0E0FF&
      ForeColor       =   &H00808080&
      Height          =   285
      Left            =   2160
      TabIndex        =   1
      Top             =   300
      Width           =   1515
   End
   Begin VB.TextBox txtTipoItem 
      BackColor       =   &H00C0E0FF&
      ForeColor       =   &H00808080&
      Height          =   285
      Left            =   60
      TabIndex        =   0
      Top             =   300
      Width           =   2055
   End
End
Attribute VB_Name = "frmAddTipoItem"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdAceptar_Click()

  ' validacion de datos
  If Not DataValidate(txtTipoItem, "@25", True) Then Exit Sub
  If Not DataValidate(txtTipoItemCorto, "@10", True) Then Exit Sub
  
  blnAceptar = True
  blnCancelar = False
  Me.Hide

End Sub

Private Sub cmdCancelar_Click()

  blnAceptar = False
  blnCancelar = True
  Unload Me

End Sub


