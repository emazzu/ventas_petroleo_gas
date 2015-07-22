VERSION 5.00
Begin VB.Form frmAddCondiciones 
   BackColor       =   &H00C0E0FF&
   Caption         =   "Agregando Condiciones de Pago"
   ClientHeight    =   2700
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4260
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2700
   ScaleWidth      =   4260
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdCancelar 
      Caption         =   "&Cancelar"
      Height          =   285
      Left            =   2430
      TabIndex        =   3
      Top             =   2340
      Width           =   795
   End
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "&Aceptar"
      Height          =   285
      Left            =   3420
      TabIndex        =   2
      Top             =   2340
      Width           =   795
   End
   Begin VB.TextBox txtIdentificacion 
      BackColor       =   &H00C0E0FF&
      ForeColor       =   &H00808080&
      Height          =   285
      Left            =   45
      TabIndex        =   0
      Top             =   315
      Width           =   4155
   End
   Begin VB.TextBox txtDetalle 
      BackColor       =   &H00C0E0FF&
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808080&
      Height          =   1320
      Left            =   45
      MultiLine       =   -1  'True
      TabIndex        =   1
      Top             =   900
      Width           =   4155
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
      TabIndex        =   5
      Text            =   "Identificación"
      Top             =   90
      Width           =   4125
   End
   Begin VB.TextBox Text3 
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
      Text            =   "Detalle"
      Top             =   675
      Width           =   4125
   End
End
Attribute VB_Name = "frmAddCondiciones"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdAceptar_Click()

  If Not DataValidate(txtIdentificacion, "@50", True) Then Exit Sub
  If Not DataValidate(txtDetalle, "@250", True) Then Exit Sub

  blnAceptar = True
  blnCancelar = False
  Me.Hide

End Sub

Private Sub cmdCancelar_Click()

  blnAceptar = False
  blnCancelar = True
  Unload Me

End Sub

