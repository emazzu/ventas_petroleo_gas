VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmEntregasTerDet 
   BackColor       =   &H00C0FFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Entregas Terminales Detalle"
   ClientHeight    =   3240
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10470
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3240
   ScaleWidth      =   10470
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtEmpresaID 
      Height          =   285
      Left            =   180
      TabIndex        =   3
      Top             =   2835
      Visible         =   0   'False
      Width           =   960
   End
   Begin VB.CommandButton cmdCancelar 
      Caption         =   "&Cancelar"
      Height          =   330
      Left            =   7080
      TabIndex        =   2
      Top             =   2790
      Width           =   1500
   End
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "&Aceptar"
      Height          =   330
      Left            =   8865
      TabIndex        =   1
      Top             =   2790
      Width           =   1500
   End
   Begin MSComctlLib.ListView lvwCon 
      Height          =   2595
      Left            =   150
      TabIndex        =   0
      Top             =   120
      Width           =   10200
      _ExtentX        =   17992
      _ExtentY        =   4577
      MultiSelect     =   -1  'True
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      Checkboxes      =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   0
   End
End
Attribute VB_Name = "frmEntregasTerDet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdAceptar_Click()
  Dim intInd, blnSelecciono As Integer
    
  ' chequeo que se haya seleccionado algo
  blnSelecciono = False
  For intInd = 1 To Me.lvwCon.ListItems.Count
    
    If Me.lvwCon.ListItems(intInd).Checked = True Then
      blnSelecciono = True
      Exit For
    End If
  
  Next
  
  If Not blnSelecciono Then
    intRes = MsgBox("No hay algún item seleccionado.", vbApplicationModal + vbOKOnly + vbInformation, "Informacion")
    Exit Sub
  End If
  
  blnAceptar = True
  blnCancelar = False
  Me.Hide

End Sub

Private Sub cmdCancelar_Click()

  blnAceptar = False
  blnCancelar = True
  Unload Me

End Sub

Private Sub txtEmpresaID_Change()
  
  ' para leer INI
  strTableNameActual = "entregasTerDetCon"
  
  strSQL = "SELECT * FROM ViewEntregasTerDetCon WHERE empresaID = " & Me.txtEmpresaID & " AND " & _
           "TerminalID = " & lvwGetValue(frmEntregasTer.lvwDatos, "terminalID") & " AND " & _
           "EntregaTerID = 0 ORDER BY Fecha"
  intRes = ListViewAppearanceChange(lvwCon)
  intRes = ListViewRefresh(lvwCon, strSQL)
  intRes = lvwHideColumn(lvwCon, "entregaTraID")
  intRes = lvwHideColumn(lvwCon, "entregaTerID")
  intRes = lvwHideColumn(lvwCon, "concesionID")
  intRes = lvwHideColumn(lvwCon, "terminalID")
  intRes = lvwHideColumn(lvwCon, "empresaID")

End Sub
