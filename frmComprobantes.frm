VERSION 5.00
Object = "{F856EC8B-F03C-4515-BDC6-64CBD617566A}#6.0#0"; "fpspr60.ocx"
Begin VB.Form frmComprobantes 
   BackColor       =   &H80000018&
   Caption         =   "Comprobantes"
   ClientHeight    =   4290
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   7080
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4290
   ScaleWidth      =   7080
   StartUpPosition =   2  'CenterScreen
   Begin FPSpreadADO.fpSpread spd 
      Height          =   1860
      Left            =   180
      TabIndex        =   0
      Top             =   360
      Width           =   6765
      _Version        =   393216
      _ExtentX        =   11933
      _ExtentY        =   3281
      _StockProps     =   64
      DisplayRowHeaders=   0   'False
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      OperationMode   =   2
      SelectBlockOptions=   0
      SpreadDesigner  =   "frmComprobantes.frx":0000
   End
End
Attribute VB_Name = "frmComprobantes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Dim Frm As Form

Private Sub Form_Activate()
  
  'set 3 decimales
  Me.spd.Row = -1
  Me.spd.Col = 6
  Me.spd.TypeCurrencyDecPlaces = 3
  Me.spd.TypeCurrencyShowSymbol = False
  
  'busca comprobante origen y pone puntero en comprobante seleccionado con anterioridad
  intRes = Me.spd.SearchCol(3, 1, -1, frmDebitoCredito.txtComprobante, SearchFlagsCaseSensitive)
  If intRes <> -1 Then
    Me.spd.SetActiveCell 1, intRes
  End If
  
End Sub

Private Sub Form_Load()
  
  Dim rs As ADODB.Recordset
  
  'referencio formulario desde donde se lo llamo
  Set Frm = Screen.ActiveForm
  
  'tomo las facturas para la empresa y cliente seleccionada
  strSQL = "select * from ND_NC_origen_vw where empresaID = " & _
    Frm.cboEmpresa.ItemData(Frm.cboEmpresa.ListIndex) & " and clienteID = " & _
    Frm.cboCliente.ItemData(Frm.cboCliente.ListIndex) & " order by fecha desc"
  Set rs = adoGetRS(strSQL)
    
  'ajusta columnas
  Me.spd.DAutoSizeCols = DAutoSizeColsMax
    
  'color zona fuera grilla
  Me.spd.GrayAreaBackColor = RGB(255, 255, 255)
  
  'no muestra encabezado de fila
  Me.spd.RowHeadersShow = False
    
  'seleccion uan fila por vez
  Me.spd.OperationMode = OperationModeRow
      
  'paso datos
  Set Me.spd.DataSource = rs
      
  'set limites a grilla
  Me.spd.MaxRows = rs.RecordCount
  Me.spd.MaxCols = rs.Fields.Count
    
  'que no se puedan modificar datos
  Me.spd.Col = -1
  Me.spd.Row = -1
  Me.spd.Lock = True
  Me.spd.Protect = True
  
  'oculto columnas
  Me.spd.Col = 8
  Me.spd.ColHidden = True
  Me.spd.Col = 9
  Me.spd.ColHidden = True
  
End Sub

Private Sub Form_Resize()
  
  Me.spd.Top = Me.ScaleTop
  Me.spd.Left = Me.ScaleLeft
  Me.spd.Width = Me.ScaleWidth
  Me.spd.Height = Me.ScaleHeight
  
End Sub

Private Sub Form_Unload(Cancel As Integer)
  
  'lo utilizo para saber si selecciono comprobante
  blnAceptar = True
  blnCancelar = False
  
End Sub

Private Sub spd_DblClick(ByVal Col As Long, ByVal Row As Long)
  
  Dim varComprobante, varOperacion, varSubtotal, varCoti As Variant
  
  'solo si hace click arriba de cotizacion
  If Col = 5 Or Col = 6 Then
    
    'tomo valores
    Me.spd.GetText 3, Row, varComprobante
    Me.spd.GetText 2, Row, varOperacion
    Me.spd.GetText Col, Row, varSubtotal
    Me.spd.GetText 7, Row, varCoti
    
    'los guardo para poder calcular
    Frm.txtComprobante = varComprobante
    Frm.txtOperacionOrigen = varOperacion
    Frm.txtSubtotalOrigen = varSubtotal
    Frm.txtCotizaOrigen = varCoti
              
    'lo utilizo para saber si selecciono comprobante
    blnAceptar = True
    blnCancelar = False
              
    'descargo
    Unload Me
    
  End If
  
End Sub
