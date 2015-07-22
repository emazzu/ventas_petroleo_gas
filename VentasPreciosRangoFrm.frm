VERSION 5.00
Begin VB.Form VentasPreciosRangoFrm 
   BackColor       =   &H00C0E0FF&
   Caption         =   "Seleccionar Rango de Fechas"
   ClientHeight    =   1380
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   3240
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1380
   ScaleWidth      =   3240
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtDesde 
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H00808080&
      Height          =   300
      Left            =   1485
      TabIndex        =   0
      Top             =   180
      Width           =   1575
   End
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "&Aceptar"
      Height          =   285
      Left            =   2235
      TabIndex        =   2
      Top             =   990
      Width           =   795
   End
   Begin VB.CommandButton cmdCancelar 
      Caption         =   "&Cancelar"
      Height          =   285
      Left            =   1305
      TabIndex        =   3
      Top             =   990
      Width           =   795
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00FFFFFF&
      Enabled         =   0   'False
      ForeColor       =   &H00808080&
      Height          =   300
      Left            =   225
      Locked          =   -1  'True
      TabIndex        =   5
      Text            =   "Desde"
      Top             =   180
      Width           =   1230
   End
   Begin VB.TextBox Text2 
      BackColor       =   &H00FFFFFF&
      Enabled         =   0   'False
      ForeColor       =   &H00808080&
      Height          =   300
      Left            =   225
      Locked          =   -1  'True
      TabIndex        =   4
      Text            =   "Hasta"
      Top             =   540
      Width           =   1230
   End
   Begin VB.TextBox txtHasta 
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H00808080&
      Height          =   300
      Left            =   1485
      TabIndex        =   1
      Top             =   540
      Width           =   1575
   End
End
Attribute VB_Name = "VentasPreciosRangoFrm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdAceptar_Click()
  Dim rs As ADODB.Recordset

  'valido info
  If Not DataValidate(Me.txtDesde, "dd/mm/yyyy", True) Then Exit Sub
  If Not DataValidate(Me.txtHasta, "dd/mm/yyyy", True) Then Exit Sub

  'validando rango
  strSQL = "SELECT * FROM ViewVentasPrecios " & _
           "WHERE Fecha BETWEEN '" & dateToIso(Me.txtDesde) & "' AND '" & dateToIso(Me.txtHasta) & "' AND " & _
           "Operacion = '" & frmVentasInfo.cboOperacion.List(frmVentasInfo.cboOperacion.ListIndex) & "' AND " & _
           "PrecioTipoID = " & lvwGetValue(frmVentasInfo.lvwContratos, "preciotipo") & " AND " & _
           "ValorMin > 0 " & _
           "ORDER BY Fecha"
  Set rs = adoGetRS(strSQL)
  
  If rs.EOF Then
    intRes = MsgBox("El rango de fechas seleccionado no tiene informacion.", vbInformation + vbOKOnly, "Información")
    Me.txtDesde = ""
    Me.txtHasta = ""
    Exit Sub
  End If
  
  blnAceptar = True
  blnCancelar = False
  Me.Hide

End Sub

Private Sub cmdCancelar_Click()
  
  blnAceptar = False
  blnCancelar = True
  Me.Hide

End Sub

