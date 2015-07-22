VERSION 5.00
Begin VB.Form frmDebitoCreditoUpdate 
   BackColor       =   &H00C0E0FF&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Agregando Detalle"
   ClientHeight    =   3825
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   5595
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3825
   ScaleWidth      =   5595
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CheckBox chkIva 
      BackColor       =   &H00C0E0FF&
      Height          =   240
      Left            =   360
      TabIndex        =   4
      Top             =   2610
      Width           =   195
   End
   Begin VB.TextBox Text5 
      Alignment       =   2  'Center
      BackColor       =   &H00C0E0FF&
      Enabled         =   0   'False
      ForeColor       =   &H00808080&
      Height          =   315
      Left            =   270
      TabIndex        =   13
      Text            =   "Iva"
      Top             =   2190
      Width           =   405
   End
   Begin VB.TextBox txtCantidad 
      BackColor       =   &H00C0E0FF&
      ForeColor       =   &H00808080&
      Height          =   300
      Left            =   270
      TabIndex        =   1
      Top             =   1845
      Width           =   1635
   End
   Begin VB.TextBox txtImporte 
      BackColor       =   &H00C0E0FF&
      ForeColor       =   &H00808080&
      Height          =   300
      Left            =   3660
      Locked          =   -1  'True
      TabIndex        =   3
      Text            =   "0.00"
      Top             =   1845
      Width           =   1605
   End
   Begin VB.TextBox txtPrecio 
      BackColor       =   &H00C0E0FF&
      ForeColor       =   &H00808080&
      Height          =   300
      Left            =   1950
      TabIndex        =   2
      Text            =   "0.000000000000"
      Top             =   1845
      Width           =   1665
   End
   Begin VB.TextBox Text4 
      BackColor       =   &H00C0E0FF&
      Enabled         =   0   'False
      ForeColor       =   &H00808080&
      Height          =   300
      Left            =   1950
      TabIndex        =   10
      Text            =   "Precio"
      Top             =   1530
      Width           =   1680
   End
   Begin VB.TextBox Text3 
      BackColor       =   &H00C0E0FF&
      Enabled         =   0   'False
      ForeColor       =   &H00808080&
      Height          =   300
      Left            =   3660
      TabIndex        =   11
      Text            =   "Importe"
      Top             =   1530
      Width           =   1605
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00C0E0FF&
      Enabled         =   0   'False
      ForeColor       =   &H00808080&
      Height          =   300
      Left            =   270
      TabIndex        =   9
      Text            =   "Cantidad"
      Top             =   1530
      Width           =   1650
   End
   Begin VB.CommandButton cmdCancelar 
      Caption         =   "&Cancelar"
      Height          =   300
      Left            =   4590
      TabIndex        =   8
      Top             =   3420
      Width           =   885
   End
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "&Aceptar"
      Height          =   300
      Left            =   3615
      TabIndex        =   7
      Top             =   3420
      Width           =   885
   End
   Begin VB.TextBox Text2 
      BackColor       =   &H00C0E0FF&
      Enabled         =   0   'False
      ForeColor       =   &H00808080&
      Height          =   300
      Left            =   270
      TabIndex        =   12
      Text            =   "Concepto"
      Top             =   270
      Width           =   4980
   End
   Begin VB.TextBox txtConcepto 
      BackColor       =   &H00C0E0FF&
      ForeColor       =   &H00808080&
      Height          =   885
      Left            =   270
      MaxLength       =   300
      MultiLine       =   -1  'True
      TabIndex        =   0
      Top             =   585
      Width           =   4980
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0E0FF&
      Height          =   3285
      Left            =   90
      TabIndex        =   14
      Top             =   45
      Width           =   5385
      Begin VB.CommandButton cmdUnidadNew 
         Caption         =   "New"
         Height          =   280
         Left            =   2970
         TabIndex        =   18
         Top             =   2880
         Width           =   2235
      End
      Begin VB.ComboBox cboUnidad 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0E0FF&
         ForeColor       =   &H00808080&
         Height          =   315
         ItemData        =   "frmDebitoCreditoUpdate.frx":0000
         Left            =   2970
         List            =   "frmDebitoCreditoUpdate.frx":000A
         Style           =   2  'Dropdown List
         TabIndex        =   6
         Top             =   2490
         Width           =   2235
      End
      Begin VB.TextBox Text7 
         BackColor       =   &H00C0E0FF&
         Enabled         =   0   'False
         ForeColor       =   &H00808080&
         Height          =   315
         Left            =   2970
         TabIndex        =   17
         Text            =   "Unidad"
         Top             =   2130
         Width           =   2235
      End
      Begin VB.CommandButton cmdTipoItemNew 
         Caption         =   "New"
         Height          =   280
         Left            =   600
         TabIndex        =   16
         Top             =   2880
         Width           =   2295
      End
      Begin VB.TextBox Text6 
         BackColor       =   &H00C0E0FF&
         Enabled         =   0   'False
         ForeColor       =   &H00808080&
         Height          =   315
         Left            =   600
         TabIndex        =   15
         Text            =   "Tipo Item"
         Top             =   2130
         Width           =   2325
      End
      Begin VB.ComboBox cboTipoItem 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0E0FF&
         ForeColor       =   &H00808080&
         Height          =   315
         ItemData        =   "frmDebitoCreditoUpdate.frx":001B
         Left            =   600
         List            =   "frmDebitoCreditoUpdate.frx":0025
         Style           =   2  'Dropdown List
         TabIndex        =   5
         Top             =   2490
         Width           =   2325
      End
   End
End
Attribute VB_Name = "frmDebitoCreditoUpdate"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdAceptar_Click()

  If Not DataValidate(txtConcepto, "@300", True) Then Exit Sub
  If Not DataValidate(txtCantidad, "########.###-", True) Then Exit Sub
  If Not DataValidate(txtPrecio, "########.############", True) Then Exit Sub
  If Not DataValidate(txtImporte, "########.###-", True) Then Exit Sub
  If Not DataValidate(cboTipoItem, , True) Then Exit Sub
  If Not DataValidate(cboUnidad, , True) Then Exit Sub

  blnAceptar = True
  blnCancelar = False
  Me.Hide

End Sub

Private Sub cmdCancelar_Click()
  
  blnAceptar = False
  blnCancelar = True
  Me.Hide

End Sub

Private Sub cmdTipoItemNew_Click()
  Dim strAux As String
        
  ' carga formulario
  Load frmAddTipoItem
    
  ' muestra formulario
  frmAddTipoItem.Show vbModal
    
  ' si hace clicn en aceptar
  If blnAceptar Then
    
    With frmAddTipoItem
    
    strSQL = "EXEC ventaTipoItem_sp '" & .txtTipoItem & "','" & .txtTipoItemCorto & "'"
    intRes = adoExecSQL(strSQL)
    
    End With
    
    ' refresh ComboBox
    strSQL = "SELECT * FROM ventasTipoItem_vw"
    intRes = ComboBoxRefresh(cboTipoItem, strSQL)
    
    ' hubico listindex en elemento agregado
    cboTipoItem.ListIndex = ComboBoxFindItem(cboTipoItem, frmAddTipoItem.txtTipoItem)
    
    ' descarga formulario
    Unload frmAddTipoItem
      
  End If

End Sub

Private Sub Command1_Click()

End Sub

Private Sub cmdUnidadNew_Click()
  Dim strAux As String
        
  ' carga formulario
  Load unidadFRM
    
  ' muestra formulario
  unidadFRM.Show vbModal
    
  ' si hace clicn en aceptar
  If blnAceptar Then
    
    With unidadFRM
    
    strSQL = "EXEC ventasUnidad_sp '" & .txtUnidad & "'"
    intRes = adoExecSQL(strSQL)
    
    End With
    
    ' refresh ComboBox
    strSQL = "SELECT * FROM ventasUnidades"
    intRes = ComboBoxRefresh(cboUnidad, strSQL)
    
    ' hubico listindex en elemento agregado
    cboUnidad.ListIndex = ComboBoxFindItem(cboUnidad, unidadFRM.txtUnidad)
    
    ' descarga formulario
    Unload unidadFRM
      
  End If

End Sub

Private Sub Form_Activate()

  'si es por diferencia calcula automaticamente
  If frmDebitoCredito.cboOperacion = "Dcg" Or frmDebitoCredito.cboOperacion = "Dco" Or frmDebitoCredito.cboOperacion = "Dcv" Then
    
    'le paso valores al detalle desde el comprobante origen
    Me.txtConcepto = "Fact. Nro: " & frmDebitoCredito.txtComprobante
    Me.txtCantidad = CDbl(frmDebitoCredito.txtSubtotalOrigen)
    Me.txtPrecio = Format(((CDbl(frmDebitoCredito.txtTipoCambio) * CDbl(frmDebitoCredito.txtAjuste)) / 100) - Val(frmDebitoCredito.txtCotizaOrigen), "########0.000000000000")
    Me.txtImporte = Format(CDbl(Me.txtCantidad) * CDbl(Me.txtPrecio), "########0.00")
        
    'si el precio es negativo es un CREDITO, pero con importes positivos, sino debito
    If Val(Me.txtPrecio) < 0 Then
      frmDebitoCredito.cboComprobante = "Credito"
      Me.txtPrecio = CDbl(Me.txtPrecio) * -1
      Me.txtImporte = Format(CDbl(Me.txtImporte) * -1, "########0.00")
    Else
      frmDebitoCredito.cboComprobante = "Debito"
    End If
        
    'fuerzo operacion a lo que corresponda
    Select Case frmDebitoCredito.txtOperacionOrigen
    
    Case "Oil"
      frmDebitoCredito.cboOperacion = "Dco"
    Case "Gas"
      frmDebitoCredito.cboOperacion = "Dcg"
    Case "Var"
      frmDebitoCredito.cboOperacion = "Dcv"
    
    End Select
        
    'fuerzo unidad a pesos
    Me.cboUnidad = "Pesos"
        
    'fuerzo tipo de item a diferencia de cambio
    Me.cboTipoItem = "DIFPRE"
        
  End If

End Sub

Private Sub Form_Load()
  
  'set combo tipo de venta
  strSQL = "SELECT * FROM ventasTipoItem_vw"
  intRes = ComboBoxRefresh(cboTipoItem, strSQL)
  
  'set combo unidad
  strSQL = "SELECT * FROM ventasUnidades"
  intRes = ComboBoxRefresh(Me.cboUnidad, strSQL)
  
End Sub

Private Sub txtCantidad_LostFocus()
  
  txtImporte = Format(CDbl(txtCantidad) * CDbl(txtPrecio), "########0.00")

End Sub

Private Sub txtConcepto_KeyPress(KeyAscii As Integer)
  
  'evito la comilla simple ya que SQL lo utiliza para string
  If KeyAscii = 39 Then
    KeyAscii = 0
  End If
  
End Sub

Private Sub txtPrecio_LostFocus()

  txtImporte = Format(CDbl(txtCantidad) * CDbl(txtPrecio), "########0.00")

End Sub
