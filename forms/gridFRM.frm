VERSION 5.00
Object = "{F856EC8B-F03C-4515-BDC6-64CBD617566A}#6.0#0"; "fpSpr60.ocx"
Begin VB.Form gridFRM 
   ClientHeight    =   4080
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6930
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   4080
   ScaleWidth      =   6930
   Begin FPSpreadADO.fpSpread spdGrid 
      Height          =   3765
      Left            =   60
      TabIndex        =   0
      Top             =   45
      Width           =   6570
      _Version        =   393216
      _ExtentX        =   11589
      _ExtentY        =   6641
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
      SpreadDesigner  =   "gridFRM.frx":0000
   End
End
Attribute VB_Name = "gridFRM"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private m_strDataSource As String
Private m_strDataWhere As String
Private m_strDataOrder As String
Private m_blnDataRefresh As Boolean
Private m_lngDataMaximo As Long
Private m_lngDataTopMax As Long
Private m_fldDataFields As ADODB.Fields
Private m_strDataStoreProcedure As String
Private m_strDataNoMuestraEnGrilla As String
Private m_strDataNoMuestraEnEdit As String
Private m_strDataSoloLecturaEnEdit As String
Private m_strDataObligatorioEnEdit As String
Private m_strDataCantDecimales As String
Private m_strDataComboBox As String
Private m_strDataValorDefault As String
Private m_strDataImportRelaciones As String
Private m_strDataAnchoColumna As String
Private m_strDataFiltro() As String
Private m_strDataFiltroAdd As String
Private m_strDataFiltroRemove As String
Private m_intDataFiltroItems As Integer

'muestro info en barra de estado
'
Function setInfoBarraEstado()
  
  MainMDI.staBarra.Panels(1).Text = " filas: " & Trim(Me.DataTopMax) & " "
  MainMDI.staBarra.Panels(2).Text = " filtro: " & Me.DataWhere & " "
  MainMDI.staBarra.Panels(3).Text = " orden: " & Me.DataOrder & " "
  MainMDI.staBarra.Panels(4).Text = " fila: " & str(Me.spdGrid.ActiveRow) & ", Col: " & str(Me.spdGrid.ActiveCol) & " "
  MainMDI.staBarra.Panels(1).AutoSize = sbrContents
  MainMDI.staBarra.Panels(2).AutoSize = sbrContents
  MainMDI.staBarra.Panels(3).AutoSize = sbrContents
  MainMDI.staBarra.Panels(4).AutoSize = sbrContents

End Function

'habilito-deshabilito botones en barra de herramienta
'
Function SetMenuButton()

  If Me.spdGrid.MaxRows = 0 Then
    MainMDI.tlbMenu.Buttons("editar").Enabled = False
    MainMDI.tlbMenu.Buttons("eliminar").Enabled = False
    MainMDI.tlbMenu.Buttons("rapido").Enabled = False
    MainMDI.tlbMenu.Buttons("buscar").Enabled = False
    MainMDI.tlbMenu.Buttons("rapido").Enabled = False
    MainMDI.tlbMenu.Buttons("avanzado").Enabled = False
    MainMDI.tlbMenu.Buttons("atras").Enabled = True
    MainMDI.tlbMenu.Buttons("excel").Enabled = True
  Else
    MainMDI.tlbMenu.Buttons("editar").Enabled = True
    MainMDI.tlbMenu.Buttons("eliminar").Enabled = True
    MainMDI.tlbMenu.Buttons("buscar").Enabled = True
    MainMDI.tlbMenu.Buttons("avanzado").Enabled = True
    MainMDI.tlbMenu.Buttons("atras").Enabled = True
    MainMDI.tlbMenu.Buttons("excel").Enabled = True
    MainMDI.tlbMenu.Buttons("actualizar").Enabled = True
  
    If Me.spdGrid.MaxRows > 0 Then
      MainMDI.tlbMenu.Buttons("rapido").Enabled = True
    End If
  
  End If

End Function

Public Property Let DataSource(ByVal strDataSource As String)
  m_strDataSource = strDataSource
End Property

Public Property Set DataFields(ByVal fldDataFields As ADODB.Fields)
  Set m_fldDataFields = fldDataFields
End Property

Public Property Let DataMaximo(ByVal lngDataMaximo As Long)
  m_lngDataMaximo = lngDataMaximo
  
End Property

Public Property Let DataTopMax(ByVal lngDataTopMax As Long)
  m_lngDataTopMax = lngDataTopMax
  
  Me.spdGrid.MaxRows = lngDataTopMax

End Property

Public Property Let DataWhere(ByVal strDataWhere As String)
  m_strDataWhere = strDataWhere

End Property

Public Property Let DataOrder(ByVal strDataOrder As String)
  m_strDataOrder = strDataOrder
  
End Property

Public Property Let DataStoreProcedure(ByVal strDataStoreProcedure As String)
  m_strDataStoreProcedure = strDataStoreProcedure
End Property

Public Property Let DataNoMuestraEnGrilla(ByVal strDataNoMuestraEnGrilla As String)
  m_strDataNoMuestraEnGrilla = strDataNoMuestraEnGrilla
End Property

Public Property Let DataNoMuestraEnEdit(ByVal strDataNoMuestraEnEdit As String)
  m_strDataNoMuestraEnEdit = strDataNoMuestraEnEdit
End Property

Public Property Let DataSoloLecturaEnEdit(ByVal strDataSoloLecturaEnEdit As String)
  m_strDataSoloLecturaEnEdit = strDataSoloLecturaEnEdit
End Property

Public Property Let DataObligatorioEnEdit(ByVal strDataObligatorioEnEdit As String)
  m_strDataObligatorioEnEdit = strDataObligatorioEnEdit
End Property

Public Property Let DataCantDecimales(ByVal str As String)
  m_strDataCantDecimales = str
End Property

Public Property Let DataComboBox(ByVal strDataComboBox As String)
  m_strDataComboBox = strDataComboBox
End Property

Public Property Let DataAnchoColumna(ByVal str As String)
  m_strDataAnchoColumna = str
End Property

Public Property Let DataRefresh(ByVal blnDataRefresh As Boolean)
  m_blnDataRefresh = blnDataRefresh
  
  Dim rs As ADODB.Recordset
  Dim strDataSource, strDataTop, strDataWhere, strDataOrder As String
   
  'si establece la propiedad en true actualiza rs
  If blnDataRefresh Then
    
    'cambio puntero mouse a reloj arena
    Screen.MousePointer = vbHourglass
    
    'formando el Query completo
    strDataSource = ""
    If m_strDataSource <> "" Then
      strDataSource = m_strDataSource
    End If
    
    'si tiene definido un maximo de filas cuando no hay filtro lo aplico
    If m_strDataWhere = "" Then
      strDataSource = Left(strDataSource, 6) & " top " & str(m_lngDataMaximo) & " " & Mid(strDataSource, 7)
    End If
    
    ' si tiene where lo agrego al query
    strDataWhere = ""
    If m_strDataWhere <> "" Then
      strDataWhere = " where " & m_strDataWhere
    End If
    
    'si tiene orden lo agrego al query
    strDataOrder = ""
    If m_strDataOrder <> "" Then
      strDataOrder = " order by " & m_strDataOrder
    End If
    
    'abro rs
    Set rs = adoGetRS(strDataSource & strDataWhere & strDataOrder)
        
    'chequeo error
    If Not lngAdoErrNum = -1 Then
      adoError
      Exit Property
    End If
    
    'set propiedad fields - estructura de la tabla/vista
    Set Me.DataFields = rs.Fields
    
    'le paso el recordset a la grilla
    Set Me.spdGrid.DataSource = rs
    
    'set propiedad dataTopMax
    Me.DataTopMax = rs.RecordCount
    Me.spdGrid.MaxRows = rs.RecordCount

    'defino variables
    Dim strNombreCol  As String
    Dim intCantDecimales As Integer
    Dim varCantDecimales As Variant
    Dim varAnchoColumna As Variant

    'si existen datos en el recordset
    If rs.RecordCount > 0 Then
      
      'set celda activa
      Me.spdGrid.Row = 1
      Me.spdGrid.Col = 1
      Me.spdGrid.SetActiveCell 1, 1
      
      'con nombre de columna y cantidad de decimales
      varCantDecimales = separateText(Me.DataCantDecimales)
      
      'recorro columnas
      For intRes = 1 To Me.spdGrid.MaxCols
        
        'le asigno un ID a cada columna para acceder directamente
        Me.spdGrid.Row = 0
        Me.spdGrid.Col = intRes
        strNombreCol = "[" & LCase(Me.spdGrid.Text) & "]"
        Me.spdGrid.ColID = strNombreCol
        
        'cambio decimales en tipos de datos numericos
        If spdGrid.CellType = CellTypeNumber Then
          spdGrid.Row = -1
          spdGrid.Col = intRes
          'busca si en columna actual se personalizaron los decimales
          If IsArray(varCantDecimales) Then intCantDecimales = Val(arrayGetValue(varCantDecimales, strNombreCol))
          Else: intCantDecimales = 2
          spdGrid.TypeNumberDecPlaces = intCantDecimales
        End If
            
        'cambio decimales en tipos de datos currency
        If spdGrid.CellType = CellTypeCurrency Then
          spdGrid.Row = -1
          spdGrid.Col = intRes
          'busca si en columna actual se personalizaron los decimales
          If IsArray(varCantDecimales) Then intCantDecimales = Val(arrayGetValue(varCantDecimales, strNombreCol))
          Else: intCantDecimales = 2
          spdGrid.TypeCurrencyDecPlaces = intCantDecimales
        End If
          
      Next
          
      'set blockMode a true para poder modificar parametros en bloque
      Me.spdGrid.BlockMode = True
      
      'lock de celdas para que no sean modificadas
      Me.spdGrid.Col = 1
      Me.spdGrid.Col2 = -1
      Me.spdGrid.Row = 1
      Me.spdGrid.Row2 = -1
      Me.spdGrid.Lock = True
      Me.spdGrid.Protect = True
      
      'set a grilla completa color default por si quedo alguna
      'fila seleccionada con anterioridad y no es la primera
      'esto soluciona el problema que queden 2 filas seleccionadas
      Me.spdGrid.Row2 = -1
      Me.spdGrid.Col2 = -1
      Me.spdGrid.BackColor = RGB(245, 245, 240)
      Me.spdGrid.ForeColor = RGB(60, 60, 60)
      
      'marco fila actual como seleccionada cambiando el color de fila completa
      Me.spdGrid.Row2 = 1
      Me.spdGrid.Col2 = -1
      Me.spdGrid.BackColor = RGB(255, 255, 250)
      Me.spdGrid.ForeColor = RGB(4, 130, 255)
      
    Else
      
      Me.spdGrid.SetActiveCell 0, 0
      
    End If
      
    'leo ancho de columnas del ini y lo guardo en propiedad
    Me.DataAnchoColumna = ReadIni(Me.Caption, "anchoColumna")
    
    'array con el nombre de columna y ancho
    varAnchoColumna = separateText(Me.DataAnchoColumna)
      
    'cambio ancho a las columnas
    If IsArray(varAnchoColumna) Then
        
      'recorro array y cambio ancho
      For intRes = 1 To UBound(varAnchoColumna, 1) - 1 Step 2
        If Me.spdGrid.GetColFromID("[" & LCase(varAnchoColumna(intRes)) & "]") <> -1 Then
          Me.spdGrid.ColWidth(Me.spdGrid.GetColFromID("[" & LCase(varAnchoColumna(intRes)) & "]")) = varAnchoColumna(intRes + 1)
        End If
      Next
      
    End If
    
    'escondo columnas
      
    'set habilito o no opciones de menu
    intRes = SetMenuButton()
    
    'show info en barra de estado
    intRes = setInfoBarraEstado()
    
     'recupero puntero mouse
    Screen.MousePointer = vbDefault

  End If

End Property

Public Property Let DataFiltroAdd(str As String)
  
  m_intDataFiltroItems = m_intDataFiltroItems + 1
  ReDim Preserve m_strDataFiltro(m_intDataFiltroItems)
  m_strDataFiltro(m_intDataFiltroItems) = str

  'aplico filtro completo a propiedad DataWhere
  Dim strWhere As String
  strWhere = ""
  For intRes = 1 To Me.DataFiltroItems
    strWhere = strWhere & Me.DataFiltro(intRes) & " and "
  Next
     
  'elimino el ultimo end
  If Me.DataFiltroItems > 0 Then
    strWhere = Left(strWhere, Len(strWhere) - 5)
  End If
  
  Me.DataWhere = strWhere

End Property

Public Property Let DataFiltroRemove(bln As Boolean)
  
  If m_intDataFiltroItems > 0 Then
    
    m_intDataFiltroItems = m_intDataFiltroItems - 1
    ReDim Preserve m_strDataFiltro(m_intDataFiltroItems)
    
    'aplico filtro completo a propiedad DataWhere
    Dim strWhere As String
    strWhere = ""
    For intRes = 1 To Me.DataFiltroItems
      strWhere = strWhere & Me.DataFiltro(intRes) & " and "
    Next
      
    'elimino el ultimo end
    If Me.DataFiltroItems > 0 Then
      strWhere = Left(strWhere, Len(strWhere) - 5)
    End If
    
    Me.DataWhere = strWhere
    
  End If
  
End Property

Public Property Get DataFiltroItems() As Integer
  DataFiltroItems = m_intDataFiltroItems
End Property

Public Property Get DataFiltro(intIndice As Integer) As String
  DataFiltro = m_strDataFiltro(intIndice)
End Property

Public Property Get DataSource() As String
  DataSource = m_strDataSource
End Property

Public Property Get DataFields() As ADODB.Fields
  Set DataFields = m_fldDataFields
End Property

Public Property Get DataMaximo() As Long
  DataMaximo = m_lngDataMaximo
End Property

Public Property Get DataTopMax() As Long
  DataTopMax = m_lngDataTopMax
End Property

Public Property Get DataWhere() As String
  DataWhere = m_strDataWhere
End Property

Public Property Get DataOrder() As String
  DataOrder = m_strDataOrder
End Property

Public Property Get DataStoreProcedure() As String
  DataStoreProcedure = m_strDataStoreProcedure
End Property

Public Property Get DataNoMuestraEnGrilla() As String
  DataNoMuestraEnGrilla = m_strDataNoMuestraEnGrilla
End Property

Public Property Get DataNoMuestraEnEdit() As String
  DataNoMuestraEnEdit = m_strDataNoMuestraEnEdit
End Property

Public Property Get DataSoloLecturaEnEdit() As String
  DataSoloLecturaEnEdit = m_strDataSoloLecturaEnEdit
End Property

Public Property Get DataObligatorioEnEdit() As String
  DataObligatorioEnEdit = m_strDataObligatorioEnEdit
End Property

Public Property Get DataCantDecimales() As String
  DataCantDecimales = m_strDataCantDecimales
End Property

Public Property Get DataComboBox() As String
  DataComboBox = m_strDataComboBox
End Property

Public Property Get DataAnchoColumna() As String
  DataAnchoColumna = m_strDataAnchoColumna
End Property

Public Property Get DataRefresh() As Boolean
  DataRefresh = m_blnDataRefresh
End Property

Public Property Get DataImportRelaciones() As String
  DataImportRelaciones = m_strDataImportRelaciones
End Property

Public Property Let DataImportRelaciones(ByVal strDataImportRelaciones As String)
  m_strDataImportRelaciones = strDataImportRelaciones
End Property

Public Property Get DataValorDefault() As String
  DataValorDefault = m_strDataValorDefault
End Property

Public Property Let DataValorDefault(ByVal strDataValorDefault As String)
  m_strDataValorDefault = strDataValorDefault
End Property

Private Sub Form_Activate()
  
  'set botones menu invisibles
  MainMDI.tlbMenu.Buttons("insertar").Enabled = True
  MainMDI.tlbMenu.Buttons("editar").Enabled = True
  MainMDI.tlbMenu.Buttons("eliminar").Enabled = True
  MainMDI.tlbMenu.Buttons("buscar").Enabled = True
  MainMDI.tlbMenu.Buttons("rapido").Enabled = True
  MainMDI.tlbMenu.Buttons("avanzado").Enabled = True
  MainMDI.tlbMenu.Buttons("atras").Enabled = True
  MainMDI.tlbMenu.Buttons("excel").Enabled = True
  MainMDI.tlbMenu.Buttons("actualizar").Enabled = True
  
  'set habilito-deshabilito botones en barra de herramienta
  intRes = SetMenuButton()
  
  'show info en barra de estado
  intRes = setInfoBarraEstado()
  
End Sub

Private Sub Form_Load()
  
  ' personalizo grilla
  intRes = spdCustomize(spdGrid)
  
  'ajusto grilla a frm
  spdGrid.Top = 0
  spdGrid.Left = 0
  If (Me.Height - 400) >= 0 Then
    spdGrid.Height = Me.Height - 400
  End If
  spdGrid.Width = Me.Width - 110
  
  'prueba
  intDataFiltroItems = 0
  
End Sub

Private Sub Form_Resize()
  
  'cuando minimizo y ajusto la grilla al tamaño el -500 en Me.Height y el -110 en Me.Width
  'lo hago porque sino queda sin verse las barra de desplazamiento en la grilla pero cuando
  'minimizo el 500 y 110 a veces es mas grande que el tamaño del frm por eso los iif
  spdGrid.Top = 0
  spdGrid.Left = 0
  spdGrid.Height = Me.Height - IIf(Me.Height < 500, Me.Height, 500)
  spdGrid.Width = Me.Width - IIf(Me.Width < 110, Me.Width, 110)
  
End Sub

Private Sub Form_Unload(Cancel As Integer)
  Dim intRes As Integer
  
  'grabando en INI
  
  'ubicacion de FRM
  strValor = "left;" & Me.Left & ";top;" & Me.Top & ";width;" & Me.Width & ";height;" & Me.Height
  intRes = WriteIni(Me.Caption, "ubicacion", strValor)
  
  'ancho de columnas
  strValor = ""
  For intRes = 1 To Me.spdGrid.MaxCols
    Me.spdGrid.GetText intRes, 0, strnombre
    strValor = strValor & strnombre & ";" & Me.spdGrid.ColWidth(intRes) & ";"
  Next
  intRes = WriteIni(Me.Caption, "anchoColumna", strValor)
  
  'where - order - cantidad de filas mostradas
  Dim strWhere As String
  strWhere = ""
  For intRes = 1 To Me.DataFiltroItems
    strWhere = strWhere & Me.DataFiltro(intRes) & ","
  Next
  
  'elimino la ultima coma
  If strWhere <> "" Then
    strWhere = Left(strWhere, Len(strWhere) - 1)
  End If

  strValor = "datawhere;" & strWhere & ";dataorder;" & Me.DataOrder & ";datatopmax;" & Me.DataTopMax
  intRes = WriteIni(Me.Caption, "data", strValor)
  
  'set botones menu invisibles
  MainMDI.tlbMenu.Buttons("insertar").Enabled = False
  MainMDI.tlbMenu.Buttons("editar").Enabled = False
  MainMDI.tlbMenu.Buttons("eliminar").Enabled = False
  MainMDI.tlbMenu.Buttons("buscar").Enabled = False
  MainMDI.tlbMenu.Buttons("rapido").Enabled = False
  MainMDI.tlbMenu.Buttons("avanzado").Enabled = False
  MainMDI.tlbMenu.Buttons("atras").Enabled = False
  MainMDI.tlbMenu.Buttons("excel").Enabled = False
  MainMDI.tlbMenu.Buttons("actualizar").Enabled = False
  
  'elimino info de barra de estado
  MainMDI.staBarra.Panels(1).Text = " filas: "
  MainMDI.staBarra.Panels(2).Text = " filtro: "
  MainMDI.staBarra.Panels(3).Text = " orden: "
  MainMDI.staBarra.Panels(4).Text = " fila: "
  MainMDI.staBarra.Panels(1).AutoSize = sbrContents
  MainMDI.staBarra.Panels(2).AutoSize = sbrContents
  MainMDI.staBarra.Panels(3).AutoSize = sbrContents
  MainMDI.staBarra.Panels(4).AutoSize = sbrContents
  
End Sub

Private Sub spdGrid_AfterUserSort(ByVal Col As Long)
  Dim strRowName As Variant
  
  ' recupero puntero mouse
  Screen.MousePointer = vbDefault

  'marco fila actual como seleccionada cambiando el color de fila completa
  Me.spdGrid.Row2 = Me.spdGrid.ActiveRow
  Me.spdGrid.Col = 1
  Me.spdGrid.Col2 = -1
  Me.spdGrid.BackColor = RGB(255, 255, 250)
  Me.spdGrid.ForeColor = RGB(4, 130, 255)

End Sub

Private Sub spdGrid_BeforeUserSort(ByVal Col As Long, ByVal State As Long, DefaultAction As Long)
  Dim strRowName As Variant
  Dim strType As String
  
  ' cambio puntero mouse
  Screen.MousePointer = vbHourglass
  
  ' muestra orden en la barra de herramientas
  intRes = Me.spdGrid.GetText(Col, 0, strRowName)
  strType = IIf(State = 2, "asc", "desc")
  Me.DataOrder = "[" & strRowName & "] " & strType
  
  'show info en barra de estado
  intRes = setInfoBarraEstado()
  
  'set color default en grilla en fila para abandonar
  Me.spdGrid.Row = Me.spdGrid.ActiveRow
  Me.spdGrid.Col = Me.spdGrid.ActiveCol
  Me.spdGrid.Col2 = -1
  Me.spdGrid.BackColor = RGB(245, 245, 245)
  Me.spdGrid.ForeColor = RGB(60, 60, 60)
  
End Sub

Private Sub spdGrid_LeaveCell(ByVal Col As Long, ByVal Row As Long, ByVal NewCol As Long, ByVal NewRow As Long, Cancel As Boolean)
  
  'cuando la grilla pierde el foco, pasa por aca, y pone a newCol y newRow en -1
  'y me caga cuando muestro la ubicacion de celda en pantalla por eso este IF
  If NewCol <> -1 And NewRow <> -1 Then
  
    'set color default de grilla a celda para abandonar
    Me.spdGrid.Row = Row
    Me.spdGrid.Col = Col
    Me.spdGrid.Col2 = -1
    Me.spdGrid.BackColor = RGB(245, 245, 245)
    Me.spdGrid.ForeColor = RGB(60, 60, 60)
  
    MainMDI.staBarra.Panels(4).Text = " fila: " & str(NewRow) & ", Col: " & str(NewCol) & " "
    MainMDI.staBarra.Panels(4).AutoSize = sbrContents
  
    'set color de seleccion a celda nueva
    Me.spdGrid.Row = NewRow
    Me.spdGrid.Row2 = NewRow
    Me.spdGrid.Col2 = -1
    Me.spdGrid.BackColor = RGB(255, 255, 250)
    Me.spdGrid.ForeColor = RGB(4, 130, 255)
    
  End If
  
End Sub
