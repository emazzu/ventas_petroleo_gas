VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{F856EC8B-F03C-4515-BDC6-64CBD617566A}#6.0#0"; "fpSpr60.ocx"
Begin VB.Form ProcesosFRM 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   3090
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7110
   FillColor       =   &H00404040&
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   3090
   ScaleWidth      =   7110
   Begin VB.TextBox txtMsjOnLine 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808080&
      Height          =   300
      Left            =   0
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      TabIndex        =   5
      Top             =   2160
      Width           =   7110
   End
   Begin VB.TextBox txtMsjAdicional 
      Alignment       =   2  'Center
      ForeColor       =   &H00404040&
      Height          =   705
      Left            =   0
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      TabIndex        =   4
      Top             =   1440
      Width           =   7110
   End
   Begin MSComctlLib.ProgressBar prgAvance 
      Height          =   405
      Left            =   0
      TabIndex        =   3
      Top             =   2520
      Width           =   7110
      _ExtentX        =   12541
      _ExtentY        =   714
      _Version        =   393216
      Appearance      =   1
      Scrolling       =   1
   End
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "&Aceptar"
      Height          =   330
      Left            =   4140
      TabIndex        =   1
      Top             =   405
      Width           =   1230
   End
   Begin VB.CommandButton cmdCancelar 
      Caption         =   "&Cancelar"
      Height          =   330
      Left            =   5400
      TabIndex        =   0
      Top             =   405
      Width           =   1230
   End
   Begin FPSpreadADO.fpSpread spdParam 
      Height          =   1365
      Left            =   0
      TabIndex        =   2
      Top             =   60
      Visible         =   0   'False
      Width           =   2805
      _Version        =   393216
      _ExtentX        =   4948
      _ExtentY        =   2408
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
      GrayAreaBackColor=   16777215
      SpreadDesigner  =   "ProcesosFRM.frx":0000
      ClipboardOptions=   3
   End
End
Attribute VB_Name = "ProcesosFRM"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim m_strDataTitulo As String         'titulo del reporte para caption
Dim m_strDataIDProceso As String      'referencia por medio de ID a los parametros del proceso
Dim m_strDataMsjAdicional As String   'mensaje adicional
Dim m_rsDataParam As ADODB.Recordset  'parametros
Dim m_strDataComboBox As String       'si el parametro es un combo o no
Dim m_strDataWhere As String          'string con where
Dim m_strDataWhereArr As Variant      'array con cada campo del where por separado
Dim m_blnDataParameters As Boolean    'si tiene parametros

Private Sub cmdAceptar_Click()
  Dim rs As ADODB.Recordset
  Dim intFila, intParamCant As Integer
  Dim strParcial, strWhere As String
  Dim varNombreParam, varNombreCol, varOperadorLogico, varValor As Variant
  Dim strArr() As String
  
  'deshabilita boton cancelar
  Me.cmdCancelar.Enabled = False
  
  'recorro grilla
  strWhere = ""
  intParamCant = 0
  For intFila = 1 To spdParam.MaxRows
      
    'tomo nombre Parametro, nombre columna, operadorLogico, valor y orden
    intRes = spdParam.GetText(1, intFila, varNombreParam)
    intRes = spdParam.GetText(4, intFila, varNombreCol)
    intRes = spdParam.GetText(5, intFila, varOperadorLogico)
    intRes = spdParam.GetText(2, intFila, varValor)
    intRes = spdParam.GetText(3, intFila, varOrden)
          
    'puntero a fila columna
    spdParam.Row = intFila
    spdParam.Col = 2
      
    'si no se selecciono nada no incluyo en where
    If varValor <> "" Then
      
      'lleno array con propiedad where por separado, nombre parametro - valoro esto lo hago
      'para poder tomar los parametros por separado desde la funcion reportesFRMOtros
      intParamCant = intParamCant + 2
      ReDim Preserve strArr(intParamCant)
      strArr(intParamCant - 1) = varNombreParam
      strArr(intParamCant - 0) = varValor
      
      'determino tipo de Dato y tipo de celda
      Select Case spdParam.CellType
      
      Case CellTypeDate
        strParcial = "[" & varNombreCol & "]" & varOperadorLogico & "'" & dateToIso(varValor) & "'"
      
      Case CellTypeComboBox
        
        'chequeo si se selecciono algo o si se selecciono la opcion <Todos>
        If Me.spdParam.Text <> "" And varValor <> "<Todos>" Then
          strParcial = "[" & varNombreCol & "]" & varOperadorLogico & "'" & varValor & "'"
        End If
            
      Case CellTypeEdit
        strParcial = "[" & varNombreCol & "]" & varOperadorLogico & "'" & varValor & "'"
      
      Case CellTypeNumber
        strParcial = "[" & varNombreCol & "]" & varOperadorLogico & varValor
      
      Case CellTypeCheckBox
        strParcial = "[" & varNombreCol & "]" & varOperadorLogico & varValor
      
      End Select
      
      'voy armando el where final
      If strWhere = "" Then
        strWhere = strParcial
      Else
        strWhere = strWhere & " and " & strParcial
      End If
      
    End If  'valor <> ""
          
  Next
  
  'le paso el where y order a la propiedad
  Me.DataWhere = strWhere
  Me.DataWhereArr = strArr
  
  'llamo a funcion que ejecuta proceso
  intRes = procesosEXEC(Me, Me.DataIDProceso, Me.DataWhere, Me.DataWhereArr)
  
  'habilita boton cancelar
  Me.cmdCancelar.Enabled = True
  
End Sub

Private Sub cmdCancelar_Click()
  Unload Me
End Sub

Private Sub Form_Activate()

  'set botones menu invisibles
  MainMDI.tlbMenu.Buttons("insertar").Enabled = False
  MainMDI.tlbMenu.Buttons("editar").Enabled = False
  MainMDI.tlbMenu.Buttons("eliminar").Enabled = False
  MainMDI.tlbMenu.Buttons("buscar").Enabled = False
  MainMDI.tlbMenu.Buttons("rapido").Enabled = False
  MainMDI.tlbMenu.Buttons("avanzado").Enabled = False
  MainMDI.tlbMenu.Buttons("borrar").Enabled = False
  MainMDI.tlbMenu.Buttons("excel").Enabled = False
    MainMDI.tlbMenu.Buttons("actualiza").Enabled = False

  Me.txtMsjAdicional = "Sin Mensaje Adicional"
  If Me.DataMsjAdicional <> "" Then
    Me.txtMsjAdicional = Me.DataMsjAdicional
  End If
  
  Me.prgAvance.Value = 100
  
End Sub

Private Sub Form_Load()
  
  'personalizo grilla
  intRes = spdCustomize(spdParam)
    
  'no permite ordenar filas
  spdParam.UserColAction = UserColActionDefault
  
  'agrego que con ENTER baje a la proxima fila
  spdParam.EditEnterAction = EditEnterActionDown
  
  'cambio colo fondo a frm
  Me.BackColor = spdParam.GrayAreaBackColor
    
End Sub

Property Let DataTitulo(ByVal strTitulo As String)
  Dim lngHeight As Long
  
  'guardo propiedad
  m_strDataTitulo = strTitulo
  
  'pongo titulo en caption
  Me.Caption = m_strDataTitulo
  
  'busco si rpt tiene parametros
  Dim rs As ADODB.Recordset
  strSQL = "select * from Parametros where idReferencia = '" & Me.DataIDProceso & "' order by idOrden"
  Set rs = adoGetRS(strSQL)
    
  'si encontro parametros los guardo
  If Not rs.EOF Then
    Set Me.DataParam = rs
    Me.DataComboBox = rs!ComboBox
    Me.DataParameters = True
  Else
    Me.DataParameters = False
  End If
  
  'para guardar altura de grilla
  lngHeight = 0
  
  'personalizo frm segun parametros
  If Me.DataParameters Then
    
    'visible grilla parametros
    spdParam.Visible = True
    
    'veo la altura de cada fila x la cantidad de filas
    Dim lngAlturaGrilla As Long
    
    'convierto altura grilla a Twips
    lngAlturaGrilla = 0
    spdParam.RowHeightToTwips 1, 20 + (spdParam.RowHeight(1) * rs.RecordCount), lngAlturaGrilla
    
    'ubico grilla
    spdParam.Top = 0
    spdParam.Left = 0
    spdParam.Width = Me.ScaleWidth
    spdParam.Height = lngAlturaGrilla
    
    'set tamaño de la grilla
    spdParam.MaxRows = rs.RecordCount
    spdParam.MaxCols = 5
    
    'poniendo titulos
    spdParam.SetText 1, 0, "Parametro"
    spdParam.SetText 2, 0, "Valor"
    spdParam.SetText 3, 0, "Ordenar"
    
    'set alto y ancho
    Dim sngWidth As Single
    'convierte de twips a Width
    spdParam.TwipsToColWidth (Me.Width / 2) - 40, sngWidth
    
    spdParam.RowHeight(0) = 15
    spdParam.ColWidth(1) = sngWidth         'nombre parametro
    spdParam.ColWidth(2) = sngWidth         'dato a ingresar o seleccionar en caso de combo
    spdParam.ColWidth(3) = 0                'para poder seleccionar un orden NO SE USA
    spdParam.ColWidth(4) = 0                'nombre de campo de recordset
    spdParam.ColWidth(5) = 0                'operadsor logico a utilizar
    
    'recorro rs
    Dim intFila As Integer
    intFila = 0
    rs.MoveFirst
    While Not rs.EOF
      
      'cuento fila
      intFila = intFila + 1
      
      'puntero en fila
      spdParam.Row = intFila
      
      'pongo nombre de parametro en columna 1
      spdParam.Col = 1
      spdParam.Text = rs!nombreParametro
      spdParam.CellType = CellTypeStaticText
      
      'puntero en columna 2 tipo de dato segun parametro
      spdParam.Col = 2
      
      'determino tipo de Dato y tipo de celda
      Select Case LCase(Trim(rs!tipo))
      
      Case "fecha"
        spdParam.CellType = CellTypeDate
      
      Case "texto"
        ' text
        If rs!ComboBox = "" Then
          spdParam.CellType = CellTypeEdit
        'comboBox
        Else
          'set ComboBox si es una columna que se muestra
          ' en un comboBox tambien armo una fila de tipo comboBox
          Me.spdParam.CellType = CellTypeComboBox
          Me.spdParam.TypeComboBoxEditable = True
          
          'lleno combo con datos
          intRes = spdDataToCbo(spdParam, rs!nombreParametro, rs!ComboBox)
          
          'agrego la opcion Todos si esta definido en menuParametros
          If rs!opcionTodos Then
            Me.spdParam.TypeComboBoxList = "<Todos>" & Chr(9) & Me.spdParam.TypeComboBoxList
          End If
        
        End If
      
      Case "entero"
        spdParam.CellType = CellTypeNumber
        spdParam.TypeHAlign = TypeHAlignLeft
        spdParam.TypeNumberDecPlaces = 0
      
      Case "decimal"
        spdParam.CellType = CellTypeNumber
        spdParam.TypeHAlign = TypeHAlignLeft
        spdParam.TypeNumberDecPlaces = 3
      
      Case "sino"
        spdParam.CellType = CellTypeCheckBox
      
      End Select
      
      'lleno combo con parametros para establecer un orden
      spdParam.Col = 3
      Me.spdParam.CellType = CellTypeComboBox
      Me.spdParam.TypeComboBoxEditable = False
      Me.spdParam.TypeComboBoxList = "ascendente" & vbTab & "descendente" & vbTab & "ninguno"
      
      'pongo nombre del Parametro o Recordset, si nombre
      'recordset esta vacio pongo nombre parametro
      spdParam.Col = 4
      If rs!nombreRecordset <> "" Then
        spdParam.Text = rs!nombreRecordset
      Else
        spdParam.Text = rs!nombreParametro
      End If
      
      'pongo operador logico a utilizar
      spdParam.Col = 5
      spdParam.Text = rs!OperadorLogico
      
      'puntero proximo registro
      rs.MoveNext
      
    Wend
    
  'crystal sin parametros
  Else
    
    'oculto grilla parametros y aceptar
    spdParam.Visible = False
    
    'ubico crystal
'    rptVisor.Top = 0
'    rptVisor.Left = 0
'    rptVisor.Width = Me.ScaleWidth
'    rptVisor.Height = Me.ScaleHeight
    
  End If
  
  'cierro rs
  rs.Close
  
  'ajusto boton cancelar
  cmdCancelar.Left = Me.ScaleWidth - cmdCancelar.Width - 25
  cmdCancelar.Top = lngAlturaGrilla + 100
    
  'ajusto boton aceptar
  cmdAceptar.Left = Me.ScaleWidth - cmdCancelar.Width - cmdAceptar.Width - 300
  cmdAceptar.Top = lngAlturaGrilla + 100
  
  'ajusto txt mensaje adicional
  Me.txtMsjAdicional.Top = Me.cmdAceptar.Top + Me.cmdAceptar.Height + 100
  
  'ajusto txt mensaje OnLine
  Me.txtMsjOnLine.Top = Me.txtMsjAdicional.Top + Me.txtMsjAdicional.Height + 100
  
  'ajusto barra progreso
  Me.prgAvance.Top = Me.txtMsjOnLine.Top + Me.txtMsjOnLine.Height + 100
    
  'cambio tamaño al formulario
  Me.BorderStyle = vbFixedSingle
  Me.Height = Me.prgAvance.Top + Me.prgAvance.Height + 500
  Me.BorderStyle = vbSizable
  
End Property

Property Let DataComboBox(ByVal str As String)
  m_strDataComboBox = str
End Property

Property Set DataParam(ByVal rs As ADODB.Recordset)
  Set m_rsDataParam = rs
End Property

Property Let DataParameters(ByVal bln As Boolean)
  m_blnDataParameters = bln
End Property

Property Let DataWhere(ByVal str As String)
  m_strDataWhere = str
End Property

Property Let DataMsjAdicional(ByVal str As String)
  m_strDataMsjAdicional = str
End Property

Property Get DataMsjAdicional() As String
  DataMsjAdicional = m_strDataMsjAdicional
End Property

Property Let DataIDProceso(ByVal str As String)
  m_strDataIDProceso = str
End Property

Property Get DataIDProceso() As String
  DataIDProceso = m_strDataIDProceso
End Property

Property Let DataRefresh(ByVal bln As Boolean)
  m_blnDataRefresh = bln
End Property

Property Let DataWhereArr(ByVal arr As Variant)
  m_strDataWhereArr = arr
End Property

Property Get DataWhereArr() As Variant
  DataWhereArr = m_strDataWhereArr
End Property

Property Get DataTitulo() As String
  DataTitulo = m_strDataTitulo
End Property

Property Get DataParameters() As Boolean
  DataParameters = m_blnDataParameters
End Property

Property Get DataComboBox() As String
  DataComboBox = m_strDataComboBox
End Property

Property Get DataParam() As ADODB.Recordset
  Set DataParam = m_rsDataParam
End Property

Property Get DataWhere() As String
  DataWhere = m_strDataWhere
End Property

Property Get DataRefresh() As Boolean
  DataRefresh = m_blnDataRefresh
End Property

Private Sub Form_Unload(Cancel As Integer)
  
  ' graba en ini ubicacion de frm
  strValor = "left;" & Me.Left & ";top;" & Me.Top & ";width;" & Me.Width & ";height;" & Me.Height
  intRes = WriteIni(Me.Caption, "ubicacion", strValor)

End Sub
