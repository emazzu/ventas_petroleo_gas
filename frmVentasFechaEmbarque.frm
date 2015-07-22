VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form frmVentasFechaEmbarque 
   BackColor       =   &H00C0E0FF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Selección Fecha Embarque"
   ClientHeight    =   4125
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3180
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4125
   ScaleWidth      =   3180
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdCancelar 
      Caption         =   "&Cancelar"
      Height          =   330
      Left            =   900
      TabIndex        =   7
      Top             =   3690
      Width           =   1000
   End
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "&Aceptar"
      Height          =   330
      Left            =   2070
      TabIndex        =   6
      Top             =   3690
      Width           =   1000
   End
   Begin VB.TextBox txtFechaEntregaCli 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0E0FF&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      ForeColor       =   &H00808080&
      Height          =   285
      Left            =   1800
      TabIndex        =   0
      Top             =   135
      Width           =   1300
   End
   Begin VB.TextBox txtTitulo 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0E0FF&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      ForeColor       =   &H00808080&
      Height          =   285
      Index           =   6
      Left            =   135
      TabIndex        =   12
      Text            =   "FechaEntregaCli"
      Top             =   135
      Width           =   1600
   End
   Begin MSComctlLib.ListView lvwPrecios 
      Height          =   1635
      Left            =   90
      TabIndex        =   5
      Top             =   1980
      Width           =   2985
      _ExtentX        =   5265
      _ExtentY        =   2884
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   0
   End
   Begin VB.TextBox txtDiasPrevios 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0E0FF&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      ForeColor       =   &H00808080&
      Height          =   285
      Left            =   1800
      TabIndex        =   1
      Top             =   495
      Width           =   1300
   End
   Begin VB.TextBox txtDiasPosteriores 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0E0FF&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      ForeColor       =   &H00808080&
      Height          =   285
      Left            =   1800
      TabIndex        =   2
      Top             =   855
      Width           =   1300
   End
   Begin VB.TextBox txtIncluyeEntregaCli 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0E0FF&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      ForeColor       =   &H00808080&
      Height          =   285
      Left            =   1800
      TabIndex        =   3
      Top             =   1215
      Width           =   1300
   End
   Begin VB.TextBox txtMesPromedio 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0E0FF&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      ForeColor       =   &H00808080&
      Height          =   285
      Left            =   1800
      TabIndex        =   4
      Top             =   1575
      Width           =   1300
   End
   Begin VB.TextBox txtTitulo 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0E0FF&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      ForeColor       =   &H00808080&
      Height          =   285
      Index           =   5
      Left            =   135
      TabIndex        =   11
      Text            =   "MesPromedio"
      Top             =   1575
      Width           =   1600
   End
   Begin VB.TextBox txtTitulo 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0E0FF&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      ForeColor       =   &H00808080&
      Height          =   285
      Index           =   4
      Left            =   135
      TabIndex        =   10
      Text            =   "IncluyeEntregaCli"
      Top             =   1215
      Width           =   1600
   End
   Begin VB.TextBox txtTitulo 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0E0FF&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      ForeColor       =   &H00808080&
      Height          =   285
      Index           =   3
      Left            =   135
      TabIndex        =   9
      Text            =   "DiasPosteriores"
      Top             =   855
      Width           =   1600
   End
   Begin VB.TextBox txtTitulo 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0E0FF&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      ForeColor       =   &H00808080&
      Height          =   285
      Index           =   2
      Left            =   135
      TabIndex        =   8
      Text            =   "DiasPrevios"
      Top             =   495
      Width           =   1600
   End
End
Attribute VB_Name = "frmVentasFechaEmbarque"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdAceptar_Click()

  blnAceptar = True
  blnCancelar = False
  Me.Hide

End Sub

Private Sub cmdCancelar_Click()

  blnAceptar = False
  blnCancelar = True
  Unload Me
  
End Sub

