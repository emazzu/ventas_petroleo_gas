VERSION 5.00
Begin VB.Form unidadFRM 
   BackColor       =   &H00C0E0FF&
   Caption         =   "Agregando Unidad"
   ClientHeight    =   1020
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   2205
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   ScaleHeight     =   1020
   ScaleWidth      =   2205
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtUnidad 
      BackColor       =   &H00C0E0FF&
      ForeColor       =   &H00808080&
      Height          =   285
      Left            =   60
      TabIndex        =   0
      Top             =   300
      Width           =   2055
   End
   Begin VB.CommandButton cmdCancelar 
      Caption         =   "&Cancelar"
      Height          =   270
      Left            =   330
      TabIndex        =   3
      Top             =   690
      Width           =   800
   End
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "&Aceptar"
      Height          =   270
      Left            =   1335
      TabIndex        =   2
      Top             =   690
      Width           =   800
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
      TabIndex        =   1
      Text            =   "Unidad"
      Top             =   90
      Width           =   2025
   End
End
Attribute VB_Name = "unidadFRM"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdAceptar_Click()

  ' validacion de datos
  If Not DataValidate(txtUnidad, "@25", True) Then Exit Sub
  
  blnAceptar = True
  blnCancelar = False
  Me.Hide

End Sub

Private Sub cmdCancelar_Click()

  blnAceptar = False
  blnCancelar = True
  Unload Me

End Sub
