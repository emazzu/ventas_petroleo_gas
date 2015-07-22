VERSION 5.00
Object = "{C4847593-972C-11D0-9567-00A0C9273C2A}#8.0#0"; "crviewer.dll"
Object = "{F856EC8B-F03C-4515-BDC6-64CBD617566A}#6.0#0"; "fpSpr60.ocx"
Begin VB.Form ReportesFRM 
   BackColor       =   &H00FFFFFF&
   ClientHeight    =   8490
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   10185
   LinkTopic       =   "Form2"
   MDIChild        =   -1  'True
   ScaleHeight     =   8490
   ScaleWidth      =   10185
   Begin VB.CommandButton cmdCancelar 
      Caption         =   "&Cancelar"
      Height          =   330
      Left            =   6840
      TabIndex        =   5
      Top             =   90
      Visible         =   0   'False
      Width           =   1230
   End
   Begin VB.CommandButton cmdBajar 
      Caption         =   "&Bajar fila"
      Height          =   330
      Left            =   5535
      TabIndex        =   4
      Top             =   90
      Visible         =   0   'False
      Width           =   1230
   End
   Begin VB.CommandButton cmdSubir 
      Caption         =   "&Subir fila"
      Height          =   330
      Left            =   4230
      TabIndex        =   3
      Top             =   90
      Visible         =   0   'False
      Width           =   1230
   End
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "&Aceptar"
      Height          =   330
      Left            =   2970
      TabIndex        =   2
      Top             =   90
      Visible         =   0   'False
      Width           =   1230
   End
   Begin FPSpreadADO.fpSpread spdParam 
      Height          =   1365
      Left            =   90
      TabIndex        =   1
      Top             =   90
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
      SpreadDesigner  =   "ReportesFRM.frx":0000
      ClipboardOptions=   3
   End
   Begin CRVIEWERLibCtl.CRViewer rptVisor 
      Height          =   3930
      Left            =   90
      TabIndex        =   0
      Top             =   1530
      Width           =   8025
      DisplayGroupTree=   0   'False
      DisplayToolbar  =   -1  'True
      EnableGroupTree =   0   'False
      EnableNavigationControls=   -1  'True
      EnableStopButton=   -1  'True
      EnablePrintButton=   -1  'True
      EnableZoomControl=   -1  'True
      EnableCloseButton=   -1  'True
      EnableProgressControl=   -1  'True
      EnableSearchControl=   -1  'True
      EnableRefreshButton=   -1  'True
      EnableDrillDown =   -1  'True
      EnableAnimationControl=   -1  'True
      EnableSelectExpertButton=   0   'False
      EnableToolbar   =   -1  'True
      DisplayBorder   =   -1  'True
      DisplayTabs     =   0   'False
      DisplayBackgroundEdge=   -1  'True
      SelectionFormula=   ""
      EnablePopupMenu =   -1  'True
      EnableExportButton=   -1  'True
      EnableSearchExpertButton=   0   'False
      EnableHelpButton=   0   'False
   End
End
Attribute VB_Name = "ReportesFRM"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim m_strDataReportName As String     'nombre reporte con path completo
Dim m_strDataIDReporte As String      'ID del Reporte
Dim m_strDataTitulo As String         'titulo del reporte para caption
Dim m_rsDataParam As ADODB.Recordset  'parametros
Dim m_strDataComboBox As String       'si el parametro es un combo o no
Dim m_strDataSource As String         'recordset asociado
Dim m_strDataFormula As String        'datos para formulas
Dim m_strDataSourceSubRpt As String   'subreport asociados al report principal
Dim m_strDataWhereShow As String      'string con where a mostrar en el reporte
Dim m_strDataWhere As String          'string con where
Dim m_strDataWhereArr As Variant      'array con cada campo del where por separado
Dim m_strDataOrderBy As String        'orden
Dim m_blnDataParameters As Boolean    'si tiene parametros
Dim m_DataRefresh As Boolean          'fuerza refresh
Dim blnPorUnicaVez As Boolean         'refresh primera vez

Private Sub cmdAceptar_Click()
  Dim rs As ADODB.Recordset
  Dim intFila, intParamCant As Integer
  Dim strParcial, strParcialShow, strWhere, strWhereShow, strORDER As String
  Dim varNombreParam, varNombreCol, varOperadorLogico, varValor, varOrden, NoUsarEnWhere, varTipoParametro As Variant
  Dim strArr() As String
  
  'recorro grilla
  strWhere = ""
  strWhereShow = ""
  strORDER = ""
  intParamCant = 0
  For intFila = 1 To spdParam.MaxRows
      
    'tomo nombre Parametro, nombre columna, operadorLogico, valor y orden
    intRes = spdParam.GetText(1, intFila, varNombreParam)
    intRes = spdParam.GetText(4, intFila, varNombreCol)
    intRes = spdParam.GetText(5, intFila, varOperadorLogico)
    intRes = spdParam.GetText(6, intFila, varNoUsarEnWhere)
    intRes = spdParam.GetText(2, intFila, varValor)
    intRes = spdParam.GetText(3, intFila, varOrden)
    intRes = spdParam.GetText(7, intFila, varTipoParametro)
          
    'puntero a fila columna
    spdParam.Row = intFila
    spdParam.Col = 2
      
    'lleno array con propiedad where por separado, nombre parametro - valor ,esto lo hago
    'para poder tomar los parametros por separado desde la funcion reportesFRMOtros
    intParamCant = intParamCant + 2
    ReDim Preserve strArr(intParamCant)
    strArr(intParamCant - 1) = varNombreParam
    
    'si tipo de parametro es periodo lo convierto al formato correcto yyyy/dd
    'el tipo periodo es el unico caso que se convierte
    If varTipoParametro = "periodo" Then
      strArr(intParamCant - 0) = dateToPeriodo(varValor)
    ElseIf varTipoParametro = "fecha" Then
      strArr(intParamCant - 0) = dateToIso(varValor)
        Else
          strArr(intParamCant - 0) = varValor
        End If
        
    'si no se ingreso el parametro no lo incluyo en el string where
    If varValor <> "" Then
      
      'determino tipo de Dato y tipo de celda
      Select Case spdParam.CellType
      
      Case CellTypeDate
        If varTipoParametro = "periodo" Then
          strParcial = "[" & varNombreCol & "]" & varOperadorLogico & "'" & dateToPeriodo(varValor) & "'"
        Else
          strParcial = "[" & varNombreCol & "]" & varOperadorLogico & "'" & dateToIso(varValor) & "'"
        End If
        strParcialShow = "[" & varNombreCol & "]" & varOperadorLogico & "'" & varValor & "'"
        
      Case CellTypeComboBox
        'chequeo si se selecciono algo o si se selecciono la opcion <Todos>
        If Me.spdParam.Text <> "" And varValor <> "<Todos>" Then
          strParcial = "[" & varNombreCol & "]" & varOperadorLogico & "'" & varValor & "'"
          strParcialShow = "[" & varNombreCol & "]" & varOperadorLogico & "'" & varValor & "'"
        End If
            
      Case CellTypeEdit
        strParcial = "[" & varNombreCol & "]" & varOperadorLogico & "'" & varValor & "'"
        strParcialShow = "[" & varNombreCol & "]" & varOperadorLogico & "'" & varValor & "'"
      
      Case CellTypeNumber
        strParcial = "[" & varNombreCol & "]" & varOperadorLogico & varValor
        strParcialShow = "[" & varNombreCol & "]" & varOperadorLogico & varValor
      
      Case CellTypeCheckBox
        strParcial = "[" & varNombreCol & "]" & varOperadorLogico & varValor
        strParcialShow = "[" & varNombreCol & "]" & varOperadorLogico & varValor
      
      End Select
      
      'solo si se definio que se incluya en el where
      If Not varNoUsarEnWhere Then
        
        'voy armando el where final
        If strWhere = "" Then
          strWhere = strParcial
        Else
          strWhere = strWhere & " and " & strParcial
        End If
        
      End If
      
      'voy armando el where para mostrar en el reporte
      If strWhereShow = "" Then
        strWhereShow = strParcialShow
      Else
        strWhereShow = strWhereShow & " y " & strParcialShow
      End If
      
    End If  'valor <> ""
          
    'voy armando el order final
    If varOrden <> "" And varOrden <> "ninguno" Then
      If strORDER = "" Then
        strORDER = "[" & varNombreCol & "] " & IIf(varOrden = "ascendente", "asc", "desc")
      Else
        strORDER = strORDER & ", [" & varNombreCol & "] " & IIf(varOrden = "ascendente", "asc", "desc")
      End If
    End If
      
  Next
  
  'le paso el where y order a la propiedad
  Me.DataWhere = strWhere
  Me.DataWhereShow = strWhereShow
  Me.DataWhereArr = strArr
  Me.DataOrderBy = strORDER
  
  'activo visor
  Me.DataRefresh = True
  
End Sub

Private Sub cmdBajar_Click()

  'verifico que no este hubicado en fila 1
  If spdParam.ActiveRow < spdParam.MaxRows Then
    
    'agrego fila temporal al final de la grilla
    spdParam.MaxRows = spdParam.MaxRows + 1
    spdParam.InsertRows spdParam.MaxRows, 1
    
    'guardo la fila en fila temporal el valor de la fila posterior a la que va a ser movida
    spdParam.MoveRowRange spdParam.ActiveRow + 1, spdParam.ActiveRow + 1, spdParam.MaxRows
    
    'muevo fila seleccionada a la posicion posterior
    spdParam.MoveRowRange spdParam.ActiveRow, spdParam.ActiveRow, spdParam.ActiveRow + 1
    
    'muevo fila guardada en posicion temporal a la ubicacion correcta
    spdParam.MoveRowRange spdParam.MaxRows, spdParam.MaxRows, spdParam.ActiveRow
    
    'elimino fila temporal
    spdParam.DeleteRows spdParam.MaxRows, 1
    spdParam.MaxRows = spdParam.MaxRows - 1
    
    'pongo puntero a fila movida
    spdParam.SetActiveCell 1, spdParam.ActiveRow + 1
    
  End If

End Sub

Private Sub cmdCancelar_Click()
  Unload Me
End Sub

Private Sub cmdSubir_Click()
  
  'verifico que no este hubicado en fila 1
  If spdParam.ActiveRow > 1 Then
    
    'agrego fila temporal al final de la grilla
    spdParam.MaxRows = spdParam.MaxRows + 1
    spdParam.InsertRows spdParam.MaxRows, 1
    
    'guardo la fila en fila temporal el valor de la fila anterior a la que va a ser movida
    spdParam.MoveRowRange spdParam.ActiveRow - 1, spdParam.ActiveRow - 1, spdParam.MaxRows
    
    'muevo fila seleccionada a la posicion anterior
    spdParam.MoveRowRange spdParam.ActiveRow, spdParam.ActiveRow, spdParam.ActiveRow - 1
    
    'muevo fila guardada en posicion temporal a la ubicacion correcta
    spdParam.MoveRowRange spdParam.MaxRows, spdParam.MaxRows, spdParam.ActiveRow
    
    'elimino fila temporal
    spdParam.DeleteRows spdParam.MaxRows, 1
    spdParam.MaxRows = spdParam.MaxRows - 1
    
    'pongo puntero a fila movida
    spdParam.SetActiveCell 1, spdParam.ActiveRow - 1
    
  End If
  
End Sub

Private Sub Form_Activate()
  
  'set botones menu deshabilitados
  MainMDI.tlbMenu.Buttons("insertar").Enabled = False
  MainMDI.tlbMenu.Buttons("editar").Enabled = False
  MainMDI.tlbMenu.Buttons("eliminar").Enabled = False
  MainMDI.tlbMenu.Buttons("buscar").Enabled = False
  MainMDI.tlbMenu.Buttons("rapido").Enabled = False
  MainMDI.tlbMenu.Buttons("avanzado").Enabled = False
  MainMDI.tlbMenu.Buttons("atras").Enabled = False
  MainMDI.tlbMenu.Buttons("excel").Enabled = False
    MainMDI.tlbMenu.Buttons("actualizar").Enabled = False
  
  'si no tiene parametros y es primera vez hago el refresh
  If Not Me.DataParameters And Not blnPorUnicaVez Then
    Me.DataRefresh = True
    blnPorUnicaVez = True
  End If
  
End Sub

Private Sub Form_Load()
  
  'personalizo grilla
  intRes = spdCustomize(spdParam)
    
  'no permite ordenar filas
  spdParam.UserColAction = UserColActionDefault
  
  'agrego que con ENTER baje a la proxima fila
  spdParam.EditEnterAction = EditEnterActionDown
  
  'elimino barra de desplazamiento
  spdParam.ScrollBars = ScrollBarsNone
  
  'set dato a ingresar que se reemplace
  spdParam.EditModeReplace = True
  
  'cambio colo fondo a frm
  Me.BackColor = spdParam.GrayAreaBackColor
    
End Sub

Private Sub Form_Resize()
  Dim lngGrilla, lngBotonAceptar As Long
  
  'reporte con parametros
  If Me.DataParameters Then
    
    'ajusto ancho grilla
    spdParam.Width = Me.ScaleWidth
    
    'ajusto columnas grilla
    Dim sngWidth As Single
    spdParam.TwipsToColWidth (Me.Width / 3) - 40, sngWidth
    spdParam.ColWidth(1) = sngWidth
    spdParam.ColWidth(2) = sngWidth
    spdParam.ColWidth(3) = sngWidth
    
    'ajusto boton cancelar
    cmdCancelar.Left = Me.ScaleWidth - cmdCancelar.Width
    cmdCancelar.Top = Me.spdParam.Height
    
    'ajusto boton bajar fila
    cmdBajar.Left = Me.ScaleWidth - cmdCancelar.Width - cmdBajar.Width - 100
    cmdBajar.Top = Me.spdParam.Height
    
    'ajusto boton subir fila
    cmdSubir.Left = Me.ScaleWidth - cmdCancelar.Width - cmdBajar.Width - cmdSubir.Width - 200
    cmdSubir.Top = Me.spdParam.Height
    
    'ajusto boton aceptar
    cmdAceptar.Left = Me.ScaleWidth - cmdCancelar.Width - cmdBajar.Width - cmdSubir.Width - cmdAceptar.Width - 300
    cmdAceptar.Top = Me.spdParam.Height
    
    'guardo altura de grilla y aceptar
    lngGrilla = spdParam.Height
    lngBotonAceptar = cmdAceptar.Height
    
  'report sin parametros
  Else
    lngGrilla = 0
    lngBotonAceptar = 0
  End If
  
  'ajusto crystal
  rptVisor.Width = Me.ScaleWidth
  
  'para que no de error y que el height del crystal no sea negativo
  If Me.ScaleHeight - lngGrilla - lngBotonAceptar > 50 Then
    rptVisor.Height = Me.ScaleHeight - lngGrilla - lngBotonAceptar
  End If
  
End Sub

Property Let DataTitulo(ByVal valueT As String)
  m_strDataTitulo = valueT
  
  'pongo titulo al form
  Me.Caption = m_strDataTitulo
  
  'busco si rpt tiene parametros
  Dim rs As ADODB.Recordset
  strSQL = "select * from menuParametros where idReferencia = '" & Me.DataIDReporte & "' order by idOrden"
  Set rs = adoGetRS(strSQL)
    
  'si encontro parametros los guardo
  If Not rs.EOF Then
    Set Me.DataParam = rs
    Me.DataComboBox = rs!ComboBox
    Me.DataParameters = True
  Else
    Me.DataParameters = False
  End If
  
  'personalizo frm segun parametros
  If Me.DataParameters Then
    
    'visible grilla parametros y aceptar
    spdParam.Visible = True
    cmdCancelar.Visible = True
    cmdBajar.Visible = True
    cmdSubir.Visible = True
    cmdAceptar.Visible = True
    
    'veo la altura de cada fila x la cantidad de filas
    Dim lngAlturaGrilla As Long
    
    'convierto altura grilla a Twips
    spdParam.RowHeightToTwips 1, 20 + (spdParam.RowHeight(1) * rs.RecordCount), lngAlturaGrilla
    
    'ubico grilla
    spdParam.Top = 0
    spdParam.Left = 0
    spdParam.Width = Me.ScaleWidth
    spdParam.Height = lngAlturaGrilla
    
    'ajusto boton cancelar
    cmdCancelar.Left = Me.ScaleWidth - cmdCancelar.Width
    cmdCancelar.Top = Me.spdParam.Height
    
    'ajusto boton bajar fila
    cmdBajar.Left = Me.ScaleWidth - cmdCancelar.Width - cmdBajar.Width - 100
    cmdBajar.Top = Me.spdParam.Height
    
    'ajusto boton subir fila
    cmdSubir.Left = Me.ScaleWidth - cmdCancelar.Width - cmdBajar.Width - cmdSubir.Width - 200
    cmdSubir.Top = Me.spdParam.Height
    
    'ajusto boton aceptar
    cmdAceptar.Left = Me.ScaleWidth - cmdCancelar.Width - cmdBajar.Width - cmdSubir.Width - cmdAceptar.Width - 300
    cmdAceptar.Top = Me.spdParam.Height
    
    'hubico crystal
    rptVisor.Top = spdParam.Height + cmdAceptar.Height
    rptVisor.Left = 0
    rptVisor.Width = Me.ScaleWidth
    rptVisor.Height = Me.ScaleHeight - spdParam.Height - cmdAceptar.Height
    
    'set tamaño de la grilla
    spdParam.MaxRows = rs.RecordCount
    spdParam.MaxCols = 7
    
    'poniendo titulos
    spdParam.SetText 1, 0, "Parametro"
    spdParam.SetText 2, 0, "Valor"
    spdParam.SetText 3, 0, "Ordenar"
    
    'set alto y ancho
    Dim sngWidth As Single
    'convierte de twips a Width
    spdParam.TwipsToColWidth (Me.Width / 3) - 40, sngWidth
    
    spdParam.RowHeight(0) = 15
    spdParam.ColWidth(1) = sngWidth         'nombre parametro
    spdParam.ColWidth(2) = sngWidth         'dato a ingresar o seleccionar en caso de combo
    spdParam.ColWidth(3) = sngWidth         'para poder seleccionar un orden
    spdParam.ColWidth(4) = 0                'nombre de campo de recordset
    spdParam.ColWidth(5) = 0                'operadsor logico a utilizar
    spdParam.ColWidth(6) = 0                'No se utiliza en Where, no pertenece al recordset
    spdParam.ColWidth(7) = 0                'tipo de celda proviene de la definicion de parametros
    
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
      
      Case "periodo"
        spdParam.CellType = CellTypeDate
        spdParam.TypeDateFormat = TypeDateFormatDDMMYY
      
      Case "fecha"
        spdParam.CellType = CellTypeDate
        spdParam.TypeDateFormat = TypeDateFormatDDMMYY
        
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
      
      'No se utiliza en where, no pertenece al recordset
      spdParam.Col = 6
      spdParam.Text = rs!NoUsarEnWhere
      
      'tipo de parametro definido en tabla parametros, fecha, periodo, decimal, entero, texto
      spdParam.Col = 7
      spdParam.Text = rs!tipo
      
      'puntero proximo registro
      rs.MoveNext
      
    Wend
    
    'set foco en grilla
    spdParam.SetFocus
    
    'set posicion en  fila 1, columna 2
    spdParam.SetActiveCell 2, 1
    
  'crystal sin parametros
  Else
    
    'oculto grilla parametros y aceptar
    spdParam.Visible = False
    cmdAceptar.Visible = False
    
    'ubico crystal
    rptVisor.Top = 0
    rptVisor.Left = 0
    rptVisor.Width = Me.ScaleWidth
    rptVisor.Height = Me.ScaleHeight
    
  End If
  
  'cierro rs
  rs.Close
 
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

Property Let DataReportName(ByVal str As String)
  m_strDataReportName = str
End Property

Property Let DataSource(ByVal str As String)
  m_strDataSource = str
End Property

Property Let DataFormula(ByVal str As String)
  m_strDataFormula = str
End Property

Property Let DataSourceSubRpt(ByVal str As String)
  m_strDataSourceSubRpt = str
End Property

Property Let DataWhere(ByVal str As String)
  m_strDataWhere = str
End Property

Property Let DataWhereShow(ByVal str As String)
  m_strDataWhereShow = str
End Property

Property Let DataOrderBy(ByVal str As String)
  m_strDataOrderBy = str
End Property

Property Let DataRefresh(ByVal bln As Boolean)
  m_blnDataRefresh = bln
  
  'si refresh
  If bln Then
    
    'mouse reloj
    Screen.MousePointer = vbHourglass
    
    'armo recordset
    strSQL = Me.DataSource
    
    'si tiene where lo agrego
    If Me.DataWhere <> "" Then
      strSQL = strSQL & " where " & Me.DataWhere
    End If
    
    'si tiene orderBy lo agrego
    If Me.DataOrderBy <> "" Then
      strSQL = strSQL & " order by " & Me.DataOrderBy
    End If
    
    'reclaro report
    Dim rptReport As CRAXDRT.Report
'    Dim rptSubRep As CRAXDRT.Report
    Dim appRpt As New CRAXDRT.Application
    Dim arrSubRpt, arrFormula As Variant
    Dim rs As ADODB.Recordset
    Dim intCuenta As Integer
    
    'abro report
    Set rptReport = appRpt.OpenReport(Me.DataReportName)
    
    'si DataSource contiene algo, tomo recordset
    If Me.DataSource <> "" Then
          
      'abro rs y se lo paso al Report
      Set rs = adoGetRS(strSQL)
      
      If Not lngAdoErrNum = -1 Then
        adoError
        Exit Property
      End If
      
      rptReport.Database.SetDataSource rs
      Set rs = Nothing
      
    End If
    
    'si tiene formulas
    If Me.DataFormula <> "" Then
      
      'separo nombre de formulas y dato asociado
      arrFormula = separateText(Me.DataFormula)
      
      'recorro array y le paso a cada formula el valor
      For intCuenta = 1 To UBound(arrFormula) - 1 Step 2
        
        rptReport.FormulaFields.GetItemByName(arrFormula(intCuenta)).Text = arrFormula(intCuenta + 1)
        
      Next
    
    End If
       
    'le paso la informacion filtrada y el orden establecido si existe la formula parametrosFiltrados en el reporte
    'a la formula del crystal hay que pasarle el texto encerrado entre ' comullas simples sino da error
    For intCuenta = 1 To rptReport.FormulaFields.Count
      If rptReport.FormulaFields(intCuenta).Name = "{@parametrosFiltrados}" Then
        rptReport.FormulaFields.GetItemByName("parametrosFiltrados").Text = "'" & _
        IIf(Me.DataWhereShow <> "", "Filtro: " & Replace(Me.DataWhereShow, "'", ""), "") & IIf(Me.DataOrderBy <> "", ", Orden: " & Me.DataOrderBy, "") & "'"
      End If
    Next
      
    'agregar caracteristicas adicionales al reporte
    intRes = reportesFRMOtros(Me.DataIDReporte, rptReport, Me.DataWhere, Me.DataWhereArr)
           
    'el visor toma el objeto rptReport y muestra
    rptVisor.ReportSource = rptReport
    Me.rptVisor.ViewReport
    
    'mouse default
    Screen.MousePointer = vbDefault
    
  End If
  
End Property

Property Let DataIDReporte(ByVal str As String)
  m_strDataIDReporte = str
End Property

Property Get DataIDReporte() As String
  DataIDReporte = m_strDataIDReporte
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

Property Get DataReportName() As String
  DataReportName = m_strDataReportName
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

Property Get DataSource() As String
  DataSource = m_strDataSource
End Property

Property Get DataFormula() As String
  DataFormula = m_strDataFormula
End Property

Property Get DataSourceSubRpt() As String
  DataSourceSubRpt = m_strDataSourceSubRpt
End Property

Property Get DataWhere() As String
  DataWhere = m_strDataWhere
End Property

Property Get DataWhereShow() As String
  DataWhereShow = m_strDataWhereShow
End Property

Property Get DataOrderBy() As String
  DataOrderBy = m_strDataOrderBy
End Property

Property Get DataRefresh() As Boolean
  DataRefresh = m_blnDataRefresh
End Property

Private Sub Form_Unload(Cancel As Integer)
  
  ' graba en ini ubicacion de frm
  strValor = "left;" & Me.Left & ";top;" & Me.Top & ";width;" & Me.Width & ";height;" & Me.Height
  intRes = WriteIni(Me.Caption, "ubicacion", strValor)

End Sub
