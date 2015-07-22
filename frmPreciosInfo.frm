VERSION 5.00
Begin VB.Form frmPreciosInfo 
   BackColor       =   &H80000018&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Precios"
   ClientHeight    =   2280
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5370
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2280
   ScaleWidth      =   5370
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "&Aceptar"
      Height          =   330
      Left            =   3735
      TabIndex        =   5
      Top             =   1845
      Width           =   1500
   End
   Begin VB.CommandButton cmdCancelar 
      Caption         =   "&Cancelar"
      Height          =   330
      Left            =   1980
      TabIndex        =   6
      Top             =   1845
      Width           =   1500
   End
   Begin VB.TextBox txtTitulo 
      Appearance      =   0  'Flat
      BackColor       =   &H80000018&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      ForeColor       =   &H00808080&
      Height          =   285
      Index           =   0
      Left            =   135
      TabIndex        =   12
      Text            =   "Fecha"
      Top             =   810
      Width           =   1400
   End
   Begin VB.TextBox txtTitulo 
      Appearance      =   0  'Flat
      BackColor       =   &H80000018&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      ForeColor       =   &H00808080&
      Height          =   285
      Index           =   3
      Left            =   135
      TabIndex        =   11
      Text            =   "ValorMínimo"
      Top             =   1125
      Width           =   1400
   End
   Begin VB.TextBox txtTitulo 
      Appearance      =   0  'Flat
      BackColor       =   &H80000018&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      ForeColor       =   &H00808080&
      Height          =   285
      Index           =   4
      Left            =   135
      TabIndex        =   10
      Text            =   "ValorMáximo"
      Top             =   1440
      Width           =   1400
   End
   Begin VB.TextBox txtTitulo 
      Appearance      =   0  'Flat
      BackColor       =   &H80000018&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      ForeColor       =   &H00808080&
      Height          =   285
      Index           =   2
      Left            =   135
      TabIndex        =   9
      Text            =   "TipoPrecio"
      Top             =   450
      Width           =   1400
   End
   Begin VB.TextBox txtTitulo 
      Appearance      =   0  'Flat
      BackColor       =   &H80000018&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      ForeColor       =   &H00808080&
      Height          =   285
      Index           =   1
      Left            =   135
      TabIndex        =   8
      Text            =   "Operación"
      Top             =   135
      Width           =   1400
   End
   Begin VB.TextBox txtFecha 
      Appearance      =   0  'Flat
      BackColor       =   &H80000018&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      ForeColor       =   &H00808080&
      Height          =   285
      Left            =   1575
      TabIndex        =   2
      Top             =   810
      Width           =   3660
   End
   Begin VB.TextBox txtValorMinimo 
      Appearance      =   0  'Flat
      BackColor       =   &H80000018&
      BorderStyle     =   0  'None
      ForeColor       =   &H00808080&
      Height          =   285
      Left            =   1575
      TabIndex        =   3
      Top             =   1125
      Width           =   3660
   End
   Begin VB.TextBox txtValorMaximo 
      Appearance      =   0  'Flat
      BackColor       =   &H80000018&
      BorderStyle     =   0  'None
      ForeColor       =   &H00808080&
      Height          =   285
      Left            =   1575
      TabIndex        =   4
      Top             =   1440
      Width           =   3660
   End
   Begin VB.ComboBox cboPrecioTipo 
      Appearance      =   0  'Flat
      BackColor       =   &H80000018&
      ForeColor       =   &H00808080&
      Height          =   315
      ItemData        =   "frmPreciosInfo.frx":0000
      Left            =   1575
      List            =   "frmPreciosInfo.frx":000A
      Sorted          =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   450
      Width           =   3120
   End
   Begin VB.ComboBox cboOperacion 
      Appearance      =   0  'Flat
      BackColor       =   &H80000018&
      ForeColor       =   &H00808080&
      Height          =   315
      ItemData        =   "frmPreciosInfo.frx":001A
      Left            =   1575
      List            =   "frmPreciosInfo.frx":0024
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   135
      Width           =   3700
   End
   Begin VB.CommandButton Command1 
      Caption         =   "New"
      Height          =   285
      Left            =   4725
      TabIndex        =   7
      Top             =   450
      Width           =   520
   End
End
Attribute VB_Name = "frmPreciosInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cboOperacion_Click()
  intRes = PrecioNuevo()
End Sub

Private Sub cboPrecioTipo_Click()
  intRes = PrecioNuevo()
End Sub

Private Sub cmdAceptar_Click()

  ' validacion de datos
  If Not DataValidate(txtFecha, "dd/mm/yyyy", True) Then Exit Sub
  If Not DataValidate(cboOperacion, , True) Then Exit Sub
  If Not DataValidate(cboPrecioTipo, , True) Then Exit Sub
  If Not DataValidate(txtValorMinimo, "###.###") Then Exit Sub
  If Not DataValidate(txtValorMaximo, "###.###") Then Exit Sub
  
  'graba
  strSQL = "EXEC spPreciosInsert " & _
  "'" & dateToIso(Me.txtFecha) & "'," & _
  Me.cboOperacion.List(Me.cboOperacion.ListIndex) & "," & _
  Me.cboPrecioTipo.ItemData(Me.cboPrecioTipo.ListIndex) & "," & _
  Val(Me.txtValorMinimo) & "," & _
  Val(Me.txtValorMaximo)
  
  intRes = adoExecSQL(strSQL)
  
  'blanqueo
  Me.txtValorMinimo = "0"
  Me.txtValorMaximo = "0"
  Me.txtValorMinimo.SelLength = Len(Me.txtValorMinimo)
  Me.txtValorMaximo.SelLength = Len(Me.txtValorMaximo)
  
  'busco nueva fecha para precio
  intRes = PrecioNuevo()
    
  'set variables
  blnAceptar = True
  blnCancelar = False

End Sub

Private Sub cmdCancelar_Click()

  blnAceptar = False
  blnCancelar = True
  Unload Me

End Sub

Private Sub Command1_Click()
  
  Dim strStore, strView, strDato As String

  strStore = "spPreciosTiposInsert"
  strView = "SELECT * FROM ViewPreciosTipos"
  strDato = ComboBoxAddItem(Me, cboPrecioTipo, "@50", strStore, strView)

End Sub

Private Function PrecioNuevo()
  Dim dtmFecha As Date
  Dim rsPrecio As ADODB.Recordset

  If Me.cboOperacion.ListIndex > -1 And Me.cboPrecioTipo.ListIndex > -1 Then

    ' tomo la ultima fecha segun Oil/Gas y Producto
    strSQL = "SELECT MAX(Fecha) AS Fecha FROM ViewPrecios WHERE " & _
              "Operacion = '" & Me.cboOperacion.List(Me.cboOperacion.ListIndex) & "' AND " & _
              "TipoPrecio = '" & Me.cboPrecioTipo.List(Me.cboPrecioTipo.ListIndex) & "'"
    Set rsPrecio = adoGetRS(strSQL)
  
    If Not rsPrecio.EOF And Not IsNull(rsPrecio!fecha) Then
      Me.txtFecha = rsPrecio!fecha + 1
      Me.txtFecha.Enabled = False
      Me.txtValorMinimo.SetFocus
    Else
      intRes = MsgBox("No hay precios para los datos seleccionados, ingrese la fecha usted mismo.", vbApplicationModal + vbOKOnly + vbInformation, "Informacion")
      Me.txtFecha.Text = ""
      Me.txtFecha.Enabled = True
      Me.txtFecha.SetFocus
    End If
    rsPrecio.Close
  
  End If

End Function

Private Sub Form_Load()
  
  ' lleno combo tipos de precio
  strSQL = "SELECT * FROM ViewPreciosTipos"
  a = ComboBoxRefresh(cboPrecioTipo, strSQL)

  'blanqueo
  Me.txtValorMinimo = "0"
  Me.txtValorMaximo = "0"
  Me.txtValorMinimo.SelLength = Len(Me.txtValorMinimo)
  Me.txtValorMaximo.SelLength = Len(Me.txtValorMaximo)

End Sub
