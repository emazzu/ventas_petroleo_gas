VERSION 5.00
Begin VB.Form ventasDetalleConsultaFrm 
   BackColor       =   &H00C0E0FF&
   Caption         =   "Ventas Detalle Consulta"
   ClientHeight    =   7680
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7245
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7680
   ScaleWidth      =   7245
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtConsulta 
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
      Height          =   7170
      Left            =   45
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   1
      Top             =   45
      Width           =   7125
   End
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "&Aceptar"
      Height          =   330
      Left            =   5805
      TabIndex        =   0
      Top             =   7290
      Width           =   1365
   End
End
Attribute VB_Name = "ventasDetalleConsultaFrm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdAceptar_Click()

  Unload Me

End Sub

