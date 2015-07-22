VERSION 5.00
Begin VB.Form frmSortData 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Ordenar..."
   ClientHeight    =   3210
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   2655
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3210
   ScaleWidth      =   2655
   StartUpPosition =   2  'CenterScreen
   Begin VB.ComboBox cboForma 
      Height          =   315
      ItemData        =   "frmSortData.frx":0000
      Left            =   180
      List            =   "frmSortData.frx":000A
      Style           =   2  'Dropdown List
      TabIndex        =   3
      Top             =   2250
      Width           =   2300
   End
   Begin VB.CommandButton cmdCancelar 
      Caption         =   "&Cancelar"
      Height          =   315
      Left            =   180
      TabIndex        =   5
      Top             =   2745
      Width           =   1000
   End
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "&Aceptar"
      Height          =   315
      Left            =   1485
      TabIndex        =   4
      Top             =   2745
      Width           =   990
   End
   Begin VB.ComboBox cboTres 
      Height          =   315
      Left            =   180
      Style           =   2  'Dropdown List
      TabIndex        =   2
      Top             =   1620
      Width           =   2300
   End
   Begin VB.ComboBox cboDos 
      Height          =   315
      Left            =   180
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   990
      Width           =   2300
   End
   Begin VB.ComboBox cboUno 
      Height          =   315
      Left            =   180
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   360
      Width           =   2300
   End
   Begin VB.Label Label5 
      Caption         =   "Tipo de orden"
      Height          =   240
      Left            =   225
      TabIndex        =   9
      Top             =   2025
      Width           =   1680
   End
   Begin VB.Label Label3 
      Caption         =   "Tercer Orden"
      Height          =   240
      Left            =   225
      TabIndex        =   8
      Top             =   1395
      Width           =   1680
   End
   Begin VB.Label Label2 
      Caption         =   "Segundo Orden"
      Height          =   240
      Left            =   225
      TabIndex        =   7
      Top             =   765
      Width           =   1680
   End
   Begin VB.Label Label1 
      Caption         =   "Primer Orden"
      Height          =   240
      Left            =   225
      TabIndex        =   6
      Top             =   135
      Width           =   1680
   End
End
Attribute VB_Name = "frmSortData"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cboDos_Change()

  If Not Me.cboDos = "" Then Me.cboForma.ListIndex = 0

End Sub

Private Sub cboTres_Click()

  If Not Me.cboTres = "" Then Me.cboForma.ListIndex = 0

End Sub

Private Sub cboUno_Click()

  If Not Me.cboUno = "" Then Me.cboForma.ListIndex = 0

End Sub

Private Sub cmdAceptar_Click()

  If Me.cboUno = "" And cboDos = "" And cboTres = "" Then Exit Sub

  blnAceptar = True
  blnCancelar = False
  Me.Hide
  
End Sub

Private Sub cmdCancelar_Click()
  
  blnAceptar = False
  blnCancelar = True
  Me.Hide

End Sub
