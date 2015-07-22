VERSION 5.00
Begin VB.Form frmCotizacionDolarInfo 
   BackColor       =   &H80000018&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Cotizacion Dolar"
   ClientHeight    =   1905
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4785
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1905
   ScaleWidth      =   4785
   StartUpPosition =   2  'CenterScreen
   Begin VB.ComboBox cboBanco 
      Appearance      =   0  'Flat
      BackColor       =   &H80000018&
      ForeColor       =   &H00808080&
      Height          =   315
      ItemData        =   "frmCotizacionDolarInfo.frx":0000
      Left            =   1530
      List            =   "frmCotizacionDolarInfo.frx":000A
      Sorted          =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   9
      Top             =   90
      Width           =   3120
   End
   Begin VB.TextBox txtVendedor 
      Appearance      =   0  'Flat
      BackColor       =   &H80000018&
      BorderStyle     =   0  'None
      ForeColor       =   &H00808080&
      Height          =   285
      Left            =   1530
      TabIndex        =   8
      Top             =   1080
      Width           =   3120
   End
   Begin VB.TextBox txtComprador 
      Appearance      =   0  'Flat
      BackColor       =   &H80000018&
      BorderStyle     =   0  'None
      ForeColor       =   &H00808080&
      Height          =   285
      Left            =   1530
      TabIndex        =   7
      Top             =   765
      Width           =   3120
   End
   Begin VB.TextBox txtFecha 
      Appearance      =   0  'Flat
      BackColor       =   &H80000018&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      ForeColor       =   &H00808080&
      Height          =   285
      Left            =   1530
      TabIndex        =   6
      Top             =   450
      Width           =   3120
   End
   Begin VB.TextBox txtTitulo 
      Appearance      =   0  'Flat
      BackColor       =   &H80000018&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      ForeColor       =   &H00808080&
      Height          =   285
      Index           =   2
      Left            =   90
      TabIndex        =   5
      Text            =   "Banco"
      Top             =   90
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
      Left            =   90
      TabIndex        =   4
      Text            =   "Vendedor"
      Top             =   1080
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
      Left            =   90
      TabIndex        =   3
      Text            =   "Comprador"
      Top             =   765
      Width           =   1400
   End
   Begin VB.TextBox txtTitulo 
      Appearance      =   0  'Flat
      BackColor       =   &H80000018&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      ForeColor       =   &H00808080&
      Height          =   285
      Index           =   0
      Left            =   90
      TabIndex        =   2
      Text            =   "Fecha"
      Top             =   450
      Width           =   1400
   End
   Begin VB.CommandButton cmdCancelar 
      Caption         =   "&Cancelar"
      Height          =   330
      Left            =   1890
      TabIndex        =   1
      Top             =   1485
      Width           =   1320
   End
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "&Aceptar"
      Height          =   330
      Left            =   3375
      TabIndex        =   0
      Top             =   1485
      Width           =   1320
   End
End
Attribute VB_Name = "frmCotizacionDolarInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cboBanco_Click()
  intRes = BancoNuevo()
End Sub

Private Sub cmdAceptar_Click()
  
  ' validacion de datos
  If Not DataValidate(cboBanco, , True) Then Exit Sub
  If Not DataValidate(txtFecha, "dd/mm/yyyy", True) Then Exit Sub
  If Not DataValidate(txtComprador, "###.###") Then Exit Sub
  If Not DataValidate(txtVendedor, "###.###") Then Exit Sub
  
  'graba
  strSQL = "EXEC spCotizacionDolarInsert " & _
  Me.cboBanco.ItemData(Me.cboBanco.ListIndex) & "," & _
  "'" & dateToIso(Me.txtFecha) & "'," & _
  Val(Me.txtComprador) & "," & _
  Val(Me.txtVendedor)
  
  intRes = adoExecSQL(strSQL)
  
  'blanqueo
  Me.txtComprador = "0"
  Me.txtVendedor = "0"
  Me.txtComprador.SelLength = Len(Me.txtComprador)
  Me.txtVendedor.SelLength = Len(Me.txtVendedor)
  
  'busco nueva fecha para precio
  intRes = BancoNuevo()
    
  'set variables
  blnAceptar = True
  blnCancelar = False

End Sub

Private Sub cmdCancelar_Click()

  blnAceptar = False
  blnCancelar = True
  Unload Me

End Sub

Private Function BancoNuevo()
  Dim dtmFecha As Date
  Dim rsCot As ADODB.Recordset

  If Me.cboBanco.ListIndex > -1 Then

    'tomo la ultima fecha segun
    strSQL = "SELECT MAX(Fecha) AS Fecha FROM cotizacionDolar_vw WHERE " & _
              "Banco = '" & Me.cboBanco & "'"
    Set rsCot = adoGetRS(strSQL)
  
    If Not rsCot.EOF And Not IsNull(rsCot!fecha) Then
      Me.txtFecha = rsCot!fecha + 1
      Me.txtFecha.Enabled = False
      Me.txtComprador.SetFocus
    Else
      intRes = MsgBox("No hay info para los datos seleccionados, ingrese la fecha usted mismo.", vbApplicationModal + vbOKOnly + vbInformation, "Informacion")
      Me.txtFecha.Text = ""
      Me.txtFecha.Enabled = True
      Me.txtFecha.SetFocus
    End If
    rsCot.Close
  
  End If

End Function

Private Sub Form_Load()
  
  ' lleno combo tipos de precio
  strSQL = "SELECT * FROM bancos"
  a = ComboBoxRefresh(cboBanco, strSQL)

  'blanqueo
  Me.txtComprador = "0"
  Me.txtVendedor = "0"
  Me.txtComprador.SelLength = Len(Me.txtComprador)
  Me.txtVendedor.SelLength = Len(Me.txtVendedor)

End Sub


