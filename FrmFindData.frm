VERSION 5.00
Begin VB.Form frmFindData 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Buscar..."
   ClientHeight    =   1170
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5175
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1170
   ScaleWidth      =   5175
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtQue 
      ForeColor       =   &H00808080&
      Height          =   285
      Left            =   180
      MaxLength       =   100
      TabIndex        =   0
      Top             =   360
      Width           =   4830
   End
   Begin VB.CommandButton cmdCancelar 
      Caption         =   "&Cancelar"
      Height          =   315
      Left            =   2115
      TabIndex        =   2
      Top             =   765
      Width           =   1365
   End
   Begin VB.CommandButton cmdBuscar 
      Caption         =   "&Buscar"
      Height          =   315
      Left            =   3645
      TabIndex        =   1
      Top             =   765
      Width           =   1365
   End
   Begin VB.Label Label1 
      Caption         =   "Ingrese información a buscar"
      Height          =   240
      Left            =   225
      TabIndex        =   3
      Top             =   135
      Width           =   4785
   End
End
Attribute VB_Name = "frmFindData"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmdBuscar_Click()

  If Not DataValidate(Me.txtQue, "@100", True) Then Exit Sub

  blnAceptar = True
  blnCancelar = False
  Me.Hide
  
End Sub

Private Sub cmdCancelar_Click()
  
  blnAceptar = False
  blnCancelar = True
  Me.Hide

End Sub

