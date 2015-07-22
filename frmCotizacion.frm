VERSION 5.00
Object = "{F856EC8B-F03C-4515-BDC6-64CBD617566A}#6.0#0"; "fpSpr60.ocx"
Begin VB.Form frmCotizacion 
   Caption         =   "Cotización"
   ClientHeight    =   3990
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   3585
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3990
   ScaleWidth      =   3585
   StartUpPosition =   2  'CenterScreen
   Begin FPSpreadADO.fpSpread spd 
      Height          =   1950
      Left            =   90
      TabIndex        =   0
      Top             =   90
      Width           =   3390
      _Version        =   393216
      _ExtentX        =   5980
      _ExtentY        =   3440
      _StockProps     =   64
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      SpreadDesigner  =   "frmCotizacion.frx":0000
   End
End
Attribute VB_Name = "frmCotizacion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Activate()
  
  'set 3 decimales
  Me.spd.Row = -1
  Me.spd.Col = 4
  Me.spd.TypeCurrencyDecPlaces = 3
  Me.spd.TypeCurrencyShowSymbol = False
  
  Me.spd.Row = -1
  Me.spd.Col = 5
  Me.spd.TypeCurrencyDecPlaces = 3
  Me.spd.TypeCurrencyShowSymbol = False
  
  'busca cotizacion origen y pone puntero en fila seleccionado con anterioridad
  intRes = Me.spd.SearchCol(1, 1, -1, frmDebitoCredito.txtIDcotizacionDolar, SearchFlagsValue)
  If intRes <> -1 Then
    Me.spd.SetActiveCell 1, intRes
  End If

End Sub

Private Sub Form_Load()
  Dim rs As ADODB.Recordset
    
  'tomo las facturas para la empresa y cliente seleccionada
  strSQL = "select top 1000 * from cotizacionDolar_vw order by fecha desc"
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
  Me.spd.Col = 1
  Me.spd.ColHidden = True
  
End Sub

Private Sub Form_Resize()
  
  Me.spd.Top = Me.ScaleTop
  Me.spd.Left = Me.ScaleLeft
  Me.spd.Width = Me.ScaleWidth
  Me.spd.Height = Me.ScaleHeight
  

End Sub

Private Sub spd_DblClick(ByVal Col As Long, ByVal Row As Long)
  Dim varID, varCoti As Variant
  
  'solo si hace click arriba de cotizacion
  If Col = 4 Or Col = 5 Then
    
    Me.spd.GetText Col, Row, varCoti
    Me.spd.GetText 1, Row, varID
    frmDebitoCredito.txtTipoCambio = varCoti
    frmDebitoCredito.txtIDcotizacionDolar = varID
    frmDebitoCredito.txtAjuste = 100
    
    Unload Me
    
  End If
  
End Sub
