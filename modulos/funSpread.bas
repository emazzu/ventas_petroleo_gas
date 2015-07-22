Attribute VB_Name = "funSpread"

'
' CAMBIA APARIENCIA A GRILLA HORIZONTAL
'
Function spdCustomize(ByRef spd As fpSpread) As Boolean

  'varios
  spd.Appearance = Appearance3D
  spd.AutoSize = False
  spd.CursorStyle = CursorStyleArrow            ' stilo cursor
  spd.CursorType = CursorTypeDefault            ' tipo cursor
  spd.UserColAction = UserColActionSort         ' posibilidad de ordenar columnas
  spd.DAutoSizeCols = DAutoSizeColsMax          ' tamaño de columnas automaticas
  spd.BorderStyle = BorderStyleNone             ' estilo borde grilla
  spd.ColHeadersAutoText = DispBlank            ' encabezado texto a mostrar
  spd.FontSize = 9                              'tamaño letra
  spd.RowHeight(0) = 20                         'altura fila de titulos
  spd.MoveActiveOnFocus = False                 '
  spd.Protect = False                            'exporta a excel sin proteccion
  
  'encabezado
  spd.ColHeadersShow = True                     ' muestra encabezado columnas
  spd.RowHeadersShow = False                    ' muestra encabezado de filas
  spd.ShadowText = RGB(0, 0, 0)                 'texto de encabezados negro
  
  'lineas horizontales y verticales
  spd.BackColorStyle = BackColorStyleUnderGrid  'estilo
  spd.GridShowHoriz = True                      'muestra grilla horizontal
  spd.GridShowVert = True                       'muestra grilla vertical
  spd.GridColor = RGB(200, 200, 200)            'color muy suave
  spd.NoBorder = True                           'sin borde fin zona de datos
  
  'barra de desplazamiento
  spd.ScrollBars = ScrollBarsBoth               'ambas
  spd.ScrollBarExtMode = True                   'cuando las necetite
  spd.VScrollSpecial = False                     'barra especial
      
  'color
  intRes = spd.SetOddEvenRowColor(RGB(245, 245, 245), RGB(60, 60, 60), RGB(245, 245, 245), RGB(60, 60, 60))
  spd.SelBackColor = RGB(248, 251, 179)         'fondo del area seleccionada amarillo
  spd.SelForeColor = RGB(4, 130, 255)           'texto del area seleccionada azul
  spd.GrayAreaBackColor = RGB(240, 240, 255)    'fondo fuera de la grilla
  
  'modo virtual
  spd.VirtualMode = False                        ' ajusta rows al tamaño del recordset
  'spd.VirtualRows = 300                         ' rows a leer del virtual buffer
  'spd.VirtualScrollBuffer = True                ' scroll vertical lee de tantas rows del buffer
  
  'setea para mostrar tooltip en las celdas donde no se ve toda la info
  spd.TextTip = TextTipFixed
  spd.TextTipDelay = 250
  
End Function

'
'ARMA GRILLA VERTICAL EN BASE A LA GRILLA
'HORIZONTAL PARA PODER EDITAR LOS DATOS
'FUNCIONA CON UNA TABLA LLAMADA menuOpciones
'Y DIBUJA TODO AUTOMATICAMENTYE AJUSTANDO FRM
'
Function spdEdit(ByVal frmBase As Form, ByRef frmEdit As Form, ByVal strOperacion As String, Optional blnAjustable As Boolean) As Boolean
  Dim a As fpSpread
  Dim varDefault As Variant
   
  'separo en un array los valores por default
  'ejemplo: [telefono],4361-0000,[pais],Argentina
  varDefault = separateText(frmBase.DataValorDefault)
  
  'cambio caption del form edit
  frmEdit.Caption = frmBase.Caption & " - " & UCase(strOperacion)
  
  'set apariencia de borde
  frmEdit.spdEdit.Appearance = Appearance3D
      
  'set no muestra encabezado fila y columna
  frmEdit.spdEdit.ColHeadersShow = False
  frmEdit.spdEdit.RowHeadersShow = False
  
  'barra de desplazamiento vertical solo cuando la necesita
  frmEdit.spdEdit.ScrollBars = ScrollBarsVertical
  frmEdit.spdEdit.ScrollBarExtMode = True
      
  'set muestro grilla horizontal vertical y color de fondo debajo de grilla
  frmEdit.spdEdit.BackColorStyle = BackColorStyleUnderGrid
  frmEdit.spdEdit.GridShowHoriz = True
  frmEdit.spdEdit.GridShowVert = True
          
  'set enter fila siguiente
  frmEdit.spdEdit.EditEnterAction = EditEnterActionDown
          
  'set cuando ingreso valor arriba de otro lo reemplaza
  frmEdit.spdEdit.EditModeReplace = True
          
  'set lineas de fondo
  frmEdit.spdEdit.GridColor = RGB(192, 192, 192)
      
  'set area fuera de la grilla blanco
  frmEdit.spdEdit.GrayAreaBackColor = RGB(255, 255, 255)
      
  'estilo borde grilla no
  frmEdit.spdEdit.BorderStyle = BorderStyleFixedSingle
      
  frmEdit.spdEdit.NoBorder = True
      
  'maximo de columnas
  frmEdit.spdEdit.MaxCols = 3
      
  'set ancho columna 1
  'Dim z As Double
  'Dim s As fpSpread
  
  frmEdit.spdEdit.ColWidth(1) = (Screen.Height * 0.001649)
  
  'set ancho columna 2
  frmEdit.spdEdit.ColWidth(2) = (Screen.Height * 0.00217)
      
  'set oculto columna 3 para indice de combo
  frmEdit.spdEdit.Col = 3
  frmEdit.spdEdit.ColHidden = True
      
  'dim variables cuenta columnas y toma titulo
  Dim intCol, intUltimaFilaAgregada As Integer
  Dim sngAlturaTotalFilas As Single
  Dim varTitulo, varDato As Variant
                
  intUltimaFilaAgregada = 0
  sngAlturaTotalFilas = 0
                
  'recorro columnas para armar grilla
  For intCol = 1 To frmBase.spdGrid.MaxCols
        
    'set columna 1 donde se encuentran los titulos de las columnas
    frmEdit.spdEdit.Col = 1
        
    'get nombre de columna de grilla base
    intRes = frmBase.spdGrid.GetText(intCol, 0, varTitulo)
        
    'busco columna que no este definida como no se muestran en edit
    If InStr(LCase(frmBase.DataNoMuestraEnEdit), "[" & LCase(varTitulo) & "]") = 0 Then
        
      intUltimaFilaAgregada = intUltimaFilaAgregada + 1
        
      'inserto fila
      frmEdit.spdEdit.InsertRows intUltimaFilaAgregada, 1
          
      'seteo maximo fila
      frmEdit.spdEdit.MaxRows = intUltimaFilaAgregada
        
      'set nombre de columna en grilla edit
      frmEdit.spdEdit.SetText 1, intUltimaFilaAgregada, varTitulo
       
      'set dato en grilla edit
      frmEdit.spdEdit.SetText 2, intUltimaFilaAgregada, varDato
        
      'set puntero en fila y columna actual para cambiar propiedades
      frmEdit.spdEdit.Col = 1
      frmEdit.spdEdit.Row = intUltimaFilaAgregada
        
      'set altura de fila porcentaje de la altura maxima de screen
      frmEdit.spdEdit.RowHeight(intUltimaFilaAgregada) = (Screen.Height * 0.0012)
        
      'set alineacion vertical
      frmEdit.spdEdit.TypeVAlign = TypeVAlignCenter

      frmEdit.spdEdit.RowHeight(intUltimaFilaAgregada) = 12
      
      'set color fondo gris para titulos
      frmEdit.spdEdit.BackColor = RGB(240, 240, 240)
          
      'set color texto gris para titulos
      frmEdit.spdEdit.ForeColor = RGB(131, 131, 131)
                
      'tamaño de letra para titulos
      frmEdit.spdEdit.FontSize = 10
      
      'set titulos estaticos
      frmEdit.spdEdit.CellType = CellTypeStaticText
      
      'set columna 2 donde se encuentran los datos a actualizar
      frmEdit.spdEdit.Col = 2
          
      'set color de fondo para ingreso de datos
      frmEdit.spdEdit.BackColor = RGB(255, 255, 255)
          
      'set color de texto para ingreso de datos
      frmEdit.spdEdit.ForeColor = RGB(0, 0, 0)
          
      'tamaño de letra para ingreso de datos
      frmEdit.spdEdit.FontSize = 9
          
      'case tipo de celda para formatear y alinear
      Select Case frmBase.DataFields(varTitulo).Type
          
      'Bit
      Case conBit
      
        frmEdit.spdEdit.CellType = CellTypeCheckBox
        frmEdit.spdEdit.Value = 0
          
      'enteros
      Case conSmallInt, conInt, conTinyInt
      
        frmEdit.spdEdit.CellType = CellTypeNumber
        frmEdit.spdEdit.TypeNumberDecPlaces = 0
        frmEdit.spdEdit.TypeHAlign = TypeHAlignRight
        frmEdit.spdEdit.Value = 0
                    
      'decimal
      Case conMoney, conSmallMoney, conReal, conFloat, conNumeric, conDecimal
            
        frmEdit.spdEdit.CellType = CellTypeNumber
        frmEdit.spdEdit.TypeNumberDecPlaces = 3
        frmEdit.spdEdit.TypeHAlign = TypeHAlignRight
        frmEdit.spdEdit.Value = 0
                    
      'fecha
      Case conSmallDateTime, conDateTime
      
        frmEdit.spdEdit.CellType = CellTypeDate
        frmEdit.spdEdit.Text = "06/17/00"
                                        
      'string
      Case conChar, conNchar, conVarchar, conText, conNVarchar, conText
        
        'set ComboBox si es una columna que se muestra en un comboBox
        'tambien armo una fila de tipo comboBox en la columan 3 para
        'guardar los index de cada combo
        If InStr(LCase(frmBase.DataComboBox), "[" & LCase(varTitulo) & "]") <> 0 Then
          frmEdit.spdEdit.CellType = CellTypeComboBox
          frmEdit.spdEdit.TypeComboBoxEditable = False
          frmEdit.spdEdit.Col = 3
          frmEdit.spdEdit.CellType = CellTypeComboBox
          frmEdit.spdEdit.TypeComboBoxEditable = False
          frmEdit.spdEdit.Col = 2
          'lleno combo con datos
          intRes = spdDataToCbo(frmEdit.spdEdit, varTitulo, frmBase.DataComboBox)
        Else
          'sino se muestra solo un texto
          frmEdit.spdEdit.CellType = CellTypeEdit
        End If
                                       
      End Select
         
      'set valores por default, pueden ser fijos o pueden venir en un select
      If InStr(LCase(frmBase.DataValorDefault), "[" & LCase(varTitulo) & "]") <> 0 Then
        frmEdit.spdEdit.Text = arrayGetValue(varDefault, "[" & LCase(varTitulo) & "]")
      End If
          
      'si operacion es un U de update o D de delete busco el valor de la
      'columna en grilla horizontal luego lo guardo en grilla vertical
      If strOperacion = "editar" Or strOperacion = "eliminar" Or strOperacion = "consultar" Then
        intRes = frmBase.spdGrid.GetText(intCol, frmBase.spdGrid.ActiveRow, varDato)
        intRes = frmEdit.spdEdit.SetText(2, intCol, varDato)
      End If
         
      'set color texto negro columna 2
      frmEdit.spdEdit.ForeColor = RGB(0, 0, 0)
          
      'set alineacion vertical columna 2
      frmEdit.spdEdit.TypeVAlign = TypeVAlignCenter
          
      'set lock y Backcolor cuando columna es No Permite Edit
      If InStr(LCase(frmBase.DataSoloLecturaEnEdit), "[" & LCase(varTitulo) & "]") <> 0 Then
        frmEdit.spdEdit.LockBackColor = RGB(240, 240, 240)
        frmEdit.spdEdit.Lock = True
      End If
    
      'se Backcolor cuando columna es obligatoria
      If InStr(LCase(frmBase.DataObligatorioEnEdit), "[" & LCase(varTitulo) & "]") <> 0 Then
        frmEdit.spdEdit.BackColor = RGB(225, 241, 255)
      End If
    
      'set lock si operacion es Consulta
      If strOperacion = "consultar" Then
        frmEdit.spdEdit.Lock = True
      End If
    
      'sumo altura de fila actual para determinar altura total de la grilla
      sngAlturaTotalFilas = sngAlturaTotalFilas + frmEdit.spdEdit.RowHeight(frmEdit.spdEdit.MaxRows) + 0.28
    
    End If
    
  Next
      
  'si es formulario ajustable
  If blnAjustable Then
    
    'ancho columna 1 con titulos ajusta al mas ancho automaticamente
    'si el campo mas ancho es menor que 25 por default es 25
    Dim dblAnchoMaximo As Double
    dblAnchoMaximo = frmEdit.spdEdit.MaxTextColWidth(1)
  
    If dblAnchoMaximo < 25 Then
      dblAnchoMaximo = 25
    End If
  
    frmEdit.spdEdit.ColWidth(1) = dblAnchoMaximo
      
    'ancho columna 2 en donde se ingresan los datos igual
    'al ancho de titulos para que quede una grilla pareja
    frmEdit.spdEdit.ColWidth(2) = dblAnchoMaximo
      
    'ancho columna 3 para indice de combo oculta
    frmEdit.spdEdit.Col = 3
    frmEdit.spdEdit.ColHidden = True
          
    'ancho de grilla dinamico es la suma del ancho de la
    'columna 1 + 2 pero primero debo convertir a twips
    Dim lngAnchoGrilla As Long
    frmEdit.spdEdit.ColWidthToTwips (frmEdit.spdEdit.ColWidth(1) + frmEdit.spdEdit.ColWidth(2)), lngAnchoGrilla
    frmEdit.spdEdit.Width = lngAnchoGrilla + 100
        
    'alto de grilla dinamico es la suma de la altura de
    'todas las filas, pero primero debo convertir a twips
    Dim lngAltoGrilla As Long
    frmEdit.spdEdit.RowHeightToTwips 1, sngAlturaTotalFilas, lngAltoGrilla
    frmEdit.spdEdit.Height = lngAltoGrilla + 300
    
    'ancho formulario
    frmEdit.Width = frmEdit.spdEdit.Width + 300
    
    'alto del formulario
    frmEdit.Height = frmEdit.spdEdit.Height + 1000
    
    'ubico grilla en form
    frmEdit.spdEdit.Left = 100
    frmEdit.spdEdit.Top = 100
    
    'cambio tamaño a botones aceptar
    frmEdit.cmdAceptar.Width = lngAnchoGrilla / 2 + 30
    frmEdit.cmdAceptar.Height = 300
  
    'cambio tamaño a botones cancelar
    frmEdit.cmdCancelar.Width = lngAnchoGrilla / 2 + 30
    frmEdit.cmdCancelar.Height = 300
      
    'cambio ubicacion de botones aceptar
    frmEdit.cmdAceptar.Left = 100
    frmEdit.cmdAceptar.Top = frmEdit.spdEdit.Height + 200
    
    'cambio ubicacion de botones cancelar
    frmEdit.cmdCancelar.Left = frmEdit.cmdAceptar.Left + frmEdit.cmdAceptar.Width + 30
    frmEdit.cmdCancelar.Top = frmEdit.spdEdit.Height + 200
  
  End If
  
  'selda activa fila 1 columna 2
  frmEdit.spdEdit.SetActiveCell 2, 1
  
  'muestra form
  frmEdit.Show vbModal

End Function

'
'ESTA FUNCION SE UTILIZA PARA CUANDO TENGO UNA GRILLA PERSONALIZADA, O SEA NO AUTOMATICA
'SE LE PASA LA GRILLA HORIZONTAL BASE Y EL NOMBRE DE LA GRILLA PERSONALIZADA
'PINTA LOS OBLIGATORIOS, BLOQUEA LO QUE SE HAYA DEFINIDO ASI Y SE ENCARGA DE TRANSFERIRLE
'LOS DATOS EN CASO DE EDICION Y ELIMINACION
'
Function spdEditSet(ByVal frmBase As Form, ByRef spdEdit As fpSpread, ByVal strOperacion As String) As Boolean
  Dim varTitulo, varDato As Variant
  Dim lngEncontroFila As Long
                
  'set columna 2 donde se encuentran los datos
  spdEdit.Col = 2
                
  'recorro columnas
  For intCol = 1 To frmBase.spdGrid.MaxCols
    
    'tomo nombre de columna de grilla base
    intRes = frmBase.spdGrid.GetText(intCol, 0, varTitulo)
        
    'busco en grilla edit la posicion del nombre de la columna de grilla base
    lngEncontroFila = spdEdit.SearchCol(1, 1, frmBase.spdGrid.MaxCols, varTitulo, SearchFlagsNone)
    
    'si encontro nombre de columna
    If lngEncontroFila <> -1 Then
    
      'si editar o eliminar tomo valor grilla horizontal y pongo en grilla vertical
      If (strOperacion = "editar" Or strOperacion = "eliminar" Or strOperacion = "consultar") Then
      
        intRes = frmBase.spdGrid.GetText(intCol, frmBase.spdGrid.ActiveRow, varDato)
        spdEdit.SetText 2, lngEncontroFila, varDato
    
      End If
      
      'set puntero a fila actual
      spdEdit.Row = lngEncontroFila
    
      'set lock y Backcolor cuando columna es No Permite Edit
      If InStr(LCase(frmBase.DataSoloLecturaEnEdit), "[" & LCase(varTitulo) & "]") <> 0 Then
        spdEdit.LockBackColor = RGB(240, 240, 240)
        spdEdit.Lock = True
      End If
    
      'se Backcolor cuando columna es obligatoria
      If InStr(LCase(frmBase.DataObligatorioEnEdit), "[" & LCase(varTitulo) & "]") <> 0 Then
        spdEdit.BackColor = RGB(225, 241, 255)
      End If
    
      'set lock si operacion es Consulta
      If strOperacion = "consultar" Then
        spdEdit.Lock = True
      End If
    
    End If
   
  Next
  
End Function

'
'ESTA FUNCION SE UTILIZA PARA CUANDO TENGO UNA GRILLA PERSONALIZADA, O SEA NO AUTOMATICA
'SE DEVUELVEN LOS VALORES SEPARADOS POR COMA EN UN STRING PASADO COMO ARGUMENTO PARA LUEGO
'SE LOS PUEDA DAR A UN STORE PROCEDURE Y ESTE ACTUALICE UNA TABLA
'
Function spdEditGet(ByVal spdEdit As fpSpread, ByRef str As String) As Boolean
  Dim intFila As Integer
  Dim varTitulo, varDato As Variant
  Dim lngEncontroFila As Long
                
  'recorro columnas
  For intFila = 1 To spdEdit.MaxRows
    
    'tomo nombre de columna de grilla base
    intRes = spdEdit.GetText(1, intFila, varTitulo)
        
    'set fila 2 datos
    spdEdit.GetText 2, intFila, varDato
        
    'set fila columna actual
    spdEdit.Row = intFila
    spdEdit.Col = 2
        
    If spdEdit.CellType <> CellTypeStaticText Then
        
      'pongo coma
      If str <> "" Then
        str = str & ","
      End If
        
      'armo string segun tipo de celda
      Select Case spdEdit.CellType
      
      Case CellTypeCheckBox, CellTypeNumber
        str = str & Val(varDato)
       
      Case CellTypeDate
        str = str & "'" & dateToIso(varDato) & "'"
      
      Case CellTypeEdit
        str = str & "'" & varDato & "'"
      
      Case CellTypeComboBox
        'si es un comboBox
         
        Dim intCantidadItem As Integer
        Dim varIndice As Variant
        Dim blnIndNumerico As Boolean
         
        'puntero columna 3 en donde se encuentra el indice asociado
        spdEdit.Col = 3
            
        'chequeo si tengo que devolver dato numerico o texto
        'si es numerico es porque el comboBox tiene indice asociado
        blnIndNumerico = False
        For a = intCantidadItem To spdEdit.TypeComboBoxCount
          spdEdit.TypeComboBoxCurSel = intCantidadItem
          If spdEdit.CellType = CellTypeNumber Then
            If spdEdit.Text <> 0 Then
              blnIndNumerico = True
            End If
          End If
        Next
            
        'puntero en donde se encuentra el texto del combo
        spdEdit.Col = 2
            
        'si se selecciono algun item
        If spdEdit.TypeComboBoxCurSel <> -1 Then
            
          'si hay que velolver numerico
          If blnIndNumerico Then
              
            Dim intIndice As Integer
              
            'puntero a columna 3 para asignarle la posicion
            'del item seleccionado al comboBox de la columna
            '3 que mantiene el identificador numerico del texto
            spdEdit.Col = 2
            intIndice = spdEdit.TypeComboBoxCurSel
            spdEdit.Col = 3
            spdEdit.TypeComboBoxCurSel = intIndice
              
            intRes = spdEdit.GetText(3, intFila, varDato)
            str = str & varDato
              
          'devuelvo texto
          Else
            str = str & "'" & varDato & "'"
          End If
            
        'no se selecciono ningun item
        Else
            
          'si hay que devolver numero devuelvo -1 sino blanco
          str = str & IIf(blnIndNumerico, "-1", "''")
            
        End If
      
      End Select
        
    End If
        
  Next
  
End Function


'
'PASA VALORES DE GRILLA VERTICAL (EDIT) A GRILLA HORIZONTAL (BASE)
'
Function spdEditToBase(ByVal frmBase As gridFRM, ByRef varNombreClave As Variant, varValorClave As Variant) As Boolean
  
  'dim variables cuenta columnas y toma titulo
  Dim rs As ADODB.Recordset
  Dim fld As ADODB.Field
  Dim varDato As Variant
                
  'busco registro
  strSQL = frmBase.DataSource & " where " & varNombreClave & " = " & varValorClave
  Set rs = adoGetRS(strSQL)
  
  'chequeo que haya traido algo
  If Not rs.EOF Then
  
    'recorro nombres de columnas
    For Each fld In rs.Fields
        
      'busco segun nombreColumna en grilla base para saber ubicacion de columna
      intRes = frmBase.spdGrid.SearchRow(0, 0, -1, fld.Name, SearchFlagsCaseSensitive)
          
      'si encontro celda reemplazo su valor
      If Not intRes = -1 Then
        
        'si el dato es BIT tengo que hacer esto porque la grilla no toma un true o false solo 0 o -1
        If fld.Type = conBit Then
          varDato = fld.Value * -1
        Else
          varDato = fld.Value
        End If
        
        frmBase.spdGrid.SetText intRes, frmBase.spdGrid.ActiveRow, varDato
      
      End If
    
    Next
    
  End If
  
End Function

'
'ESTA FUNCION SE UTILIZA PARA CUANDO TENGO UNA GRILLA PERSONALIZADA, O SEA NO AUTOMATICA
'SE LE PASA LA GRILLA HORIZONTAL BASE Y EL NOMBRE DE LA GRILLA PERSONALIZADA
'Y SE ENCARGA DE LIMPIAR LOS DATOS DE PANTALLA
'
Function spdEditClear(ByVal frmBase As Form, ByRef spdEdit As fpSpread, ByVal strOperacion As String) As Boolean
  Dim varTitulo, varDato As Variant
  Dim lngEncontroFila As Long
                
  'set columna 2 donde se encuentran los datos
  spdEdit.Col = 2
                
  'recorro columnas
  For intRow = 1 To spdEdit.MaxRows
      
    spdEdit.SetText 2, intRow, ""
      
  Next
  
End Function


'
'LLENA UN COMBO DE UNA GRILLA VERTICAL CON VARIAS POSIBILIDADES
'FUNCIONA CON UNA TABLA LLAMADA menuOpciones columna ComboBox
'
'1. pozo;select pozo, id from pozos     ' devuelve id
'2. equipo;select equipo from equipos   ' devuelve texto
'3. UNO,0,DOS,0,TRES,0,CUATRO,0         ' devuelve texto
'4. UNO,1,DOS,2,TRES,3,CUATRO,4         ' devuelve id
'
Function spdDataToCbo(ByRef spd As fpSpread, ByVal strColumna As String, ByVal strDataSelect As String) As Boolean
  Dim intInd  As Integer
  Dim strList, strItemData As String
  Dim rs As ADODB.Recordset
  Dim strArrayDatosSeparados As Variant
      
  'separo columnas de select
  strArrayDatosSeparados = separateText(strDataSelect, ";")
    
  'busca ubicacion del nombre de la columna y en el siguiente esta el select o lista a mano
  'ejemplo: select Idxxx, Nombrexxxx from xxxxxx o Uno,1,Dos,2,Tres,3
  For intInd = 1 To UBound(strArrayDatosSeparados) - 1 Step 2
    If LCase(strArrayDatosSeparados(intInd)) = "[" & LCase(strColumna) & "]" Then
      Exit For
    End If
  Next
    
  'chequeo que haya encontrado la palabra select
  If InStr(LCase(strArrayDatosSeparados(intInd + 1)), "select") <> 0 Then
  
    'abro recordset
    Set rs = adoGetRS(strArrayDatosSeparados(intInd + 1))
  
    'busco cantidad de columnas del rs para
    'saber si viene solo texto o texto e indice
    Select Case rs.Fields.Count
    
    Case 1  ' solo texto
    
      'recorro recordset
      strList = ""
      strItemData = ""
      While Not rs.EOF
        'armo descripciones e indices
        strList = strList & rs(0) & vbTab
        strItemData = strItemData & "0" & vbTab
        
        'avanzo puntero
        rs.MoveNext
      
      Wend
      
    Case 2      ' texto e indice
           
      'recorro recordset
      strList = ""
      strItemData = ""
      While Not rs.EOF
        
        'busco en cual columna viene el texto y en cual el indice
        Select Case rs(0).Type
        Case conChar, conNchar, conVarchar, conText, conNVarchar, conText
          strList = strList & rs(0) & vbTab
          strItemData = strItemData & rs(1) & vbTab
        Case Else
          strList = strList & rs(1) & vbTab
          strItemData = strItemData & rs(0) & vbTab
        End Select
      
        'avanzo puntero
        rs.MoveNext
      
      Wend
      rs.Close
    
    End Select
    
  'lista de datos a escrita a mano: Uno,1,Dos,2,Tres,3 o Uno,0,Dos,0,Tres,0,Cuatro,0
  Else
    
    'separo los elementos de la lista
    Dim strArraySeparoLista As Variant
    strArraySeparoLista = separateText(strArrayDatosSeparados(intInd + 1), ",")
      
    'recorro array
    strList = ""
    strItemData = ""
    For intInd = 1 To UBound(strArraySeparoLista) - 1 Step 2
        
      strList = strList & strArraySeparoLista(intInd) & vbTab
      strItemData = strItemData & strArraySeparoLista(intInd + 1) & vbTab
      
    Next
    
  End If
    
  'agrego la lista al comboBox de la grilla, columna 2, fila actual
  spd.Col = 2
  spd.TypeComboBoxList = strList
    
  'agrego ItemData al combobox de la grilla columna 3 invisible, fila actual
  spd.Col = 3
  spd.TypeComboBoxList = strItemData
    
  'set vuelvo a columna 2
  spd.Col = 2
    
End Function

'
'PASA VALORES DE GRILLA VERTICAL A GRILLA
'HORIZONTAL UNA VEZ EDITADOS LOS DATOS
'
Function spdEditSetToSpdBase(ByVal frmBase As Form, ByRef frmEdit As Form, ByVal strOperacion As String) As String
  
  'dim variables cuenta columnas y toma titulo
  Dim intRow As Integer
  Dim varTitulo, varDato As Variant
  Dim strResulParcial  As String
                
  'recorro filas para armar grilla
  strResulParcial = ""
  For intRow = 1 To frmEdit.spdEdit.MaxRows
        
    'tomo nombre de columna de grilla edit
    intRes = frmEdit.spdEdit.GetText(1, intRow, varTitulo)
        
    'busco columna no esta dentro de las no permiten edicion
    'If InStr(LCase(frmBase.DataSoloLecturaEnEdit), "[" & LCase(varTitulo) & "]") = 0 Then
        
      'tomo dato de columna de grilla edit
      intRes = frmEdit.spdEdit.GetText(2, intRow, varDato)
        
      'case tipo de dato para formatear y alinear
      Select Case frmBase.DataFields(varTitulo).Type
          
      'Bit, enteros, decimales
      Case conBit, conSmallInt, conInt, conTinyInt, conMoney, conSmallMoney, conReal, conFloat, conNumeric, conDecimal
          
        strResulParcial = varDato
          
      'fecha
      Case conSmallDateTime, conDateTime
      
        strResulParcial = "'" & dateToIso(varDato) & "'"
      
      'string
      Case conChar, conNchar, conVarchar, conText, conNVarchar, conText
      
        'puntero de grilla en fila y columna
        frmEdit.spdEdit.Row = intRow
        frmEdit.spdEdit.Col = 2
        
        'si es un comboBox
        If frmEdit.spdEdit.CellType = CellTypeComboBox Then
        
          Dim intCantidadItem As Integer
          Dim varIndice As Variant
          Dim blnIndNumerico As Boolean
        
          'puntero columna 3 en donde se encuentra el indice asociado
          frmEdit.spdEdit.Col = 3
            
          'chequeo si tengo que devolver dato numerico o texto
          'si es numerico es porque el comboBox tiene indice asociado
          blnIndNumerico = False
          For a = intCantidadItem To frmEdit.spdEdit.TypeComboBoxCount
            frmEdit.spdEdit.TypeComboBoxCurSel = intCantidadItem
            If frmEdit.spdEdit.Text <> 0 Then
              blnIndNumerico = True
            End If
          Next
            
          'puntero en donde se encuentra el texto del combo
          frmEdit.spdEdit.Col = 2
            
          'si se selecciono algun item
          If frmEdit.spdEdit.TypeComboBoxCurSel <> -1 Then
          
            'si hay que velolver numerico
            If blnIndNumerico Then
            
              Dim intIndice As Integer
              
              'puntero a columna 3 para asignarle la posicion
              'del item seleccionado al comboBox de la columna
              '3 que mantiene el identificador numerico del texto
              frmEdit.spdEdit.Col = 2
              intIndice = frmEdit.spdEdit.TypeComboBoxCurSel
              frmEdit.spdEdit.Col = 3
              frmEdit.spdEdit.TypeComboBoxCurSel = intIndice
            
              intRes = frmEdit.spdEdit.GetText(3, intRow, varDato)
              strResulParcial = varDato
            
            'devuelvo texto
            Else
              strResulParcial = "'" & varDato & "'"
            End If
            
          'no se selecciono ningun item
          Else
          
            'si hay que devolver numero devuelvo -1 sino blanco
            strResulParcial = IIf(blnIndNumerico, "-1", "''")
            
          End If
        
        'si es texto comun
        Else
          strResulParcial = "'" & varDato & "'"
        End If
        
      End Select
    
      'armando el sql final
      strResulFinal = strResulFinal & strResulParcial & ","
    
    'End If
    
  Next
  
  'set string final
  spdEditSetToSpdBase = "'" & strOperacion & "'," & Left(strResulFinal, Len(strResulFinal) - 1)

End Function

'
' VALIDA FILAS DE LA GRILLA TODAS SON OBLIGATORIAS EXCEPTO LAS QUE ESTAN EN LA
' COLUMNA NoObligatorio TRABAJA CON TABLA menuOpciones COLUMNA NOOBLIGATORIO
'
Function spdValidateData(ByVal frmBase As Form, ByVal spd As fpSpread) As Boolean
  
  spdValidateData = True
  
  'dim variables cuenta columnas y toma titulo
  Dim intFila As Integer
  Dim varTitulo, varDato As Variant
  Dim strLeyenda As String
                
  'recorro filas para armar grilla
  strLeyenda = ""
  For intFila = 1 To spd.MaxRows
        
    'get nombre de columna de grilla edit
    intRes = spd.GetText(1, intFila, varTitulo)
        
    'set fila columna
    spd.Row = intFila
    spd.Col = 2
    
    'si es static la paso por algo, en la mayoria de los casos
    'se va a dar en grillas perzonalizadas en filas titulo con Span
    If Not spd.CellType = CellTypeStaticText Then
        
      'busco columna, si es obligatoria
      If InStr(LCase(frmBase.DataObligatorioEnEdit), "[" & LCase(varTitulo) & "]") <> 0 Then
        
        'get dato de columna de grilla edit
        intRes = spd.GetText(2, intFila, varDato)
                
        'case tipo de celda para formatear y alinear
        Select Case frmBase.DataFields(varTitulo).Type
          
        'Bit, enteros, decimales
        Case conBit
      
        'nose valida porque por default es 0, false
      
        'enteros y decimales
        Case conSmallInt, conInt, conTinyInt, conMoney, conSmallMoney, conReal, conFloat, conNumeric, conDecimal
          
        ' nose valida porque por default es 0
          
        'fecha
        Case conSmallDateTime, conDateTime
      
          If varDato = "" Then
            strLeyenda = strLeyenda & varTitulo & vbCrLf
            spdValidateData = False
          End If
      
        'string
        Case conChar, conNchar, conVarchar, conText, conNVarchar, conNtext
        
          If varDato = "" Then
            strLeyenda = strLeyenda & varTitulo & vbCrLf
            spdValidateData = False
          End If
        
        End Select
    
      End If
      
    End If
    
  Next
  
  If strLeyenda <> "" Then
    intRes = MsgBox(strLeyenda & vbCrLf & "Información obligatoria.", vbApplicationModal + vbCritical + vbOKOnly, frmBase.Caption)
  End If
  
End Function

'
' DEVUELVE EL VALOR DE FILA ACTUAL Y NOMBRE DE COLUMNA PASADO COMO ARGUMENTO
'
Function spdGetValue(spdGrilla As fpSpread, strNombreColumna As String) As Variant
  Dim intPosicion As Integer
  Dim intColAnt As Integer
  
  'valor default
  spdGetValue = ""
  
  'busco posicion de la columna
  intPosicion = spdGrilla.SearchRow(0, 0, -1, strNombreColumna, 0)
  
  'si encontro devuelvo su valor
  If intPosicion <> -1 Then
    
    'guardo columna actual
    intColAnt = spdGrilla.Col
    
    'puntero a columna encontrada
    spdGrilla.Col = intPosicion
  
    'case tipo de celda para formatear y alinear
    Select Case spdGrilla.CellType
          
    'fecha
    Case CellTypeDate
      
       spdGetValue = "'" & dateToIso(spdGrilla.Text) & "'"
      
    'string
    Case CellTypeComboBox, CellTypeEdit
        
       spdGetValue = "'" & spdGrilla.Text & "'"
        
    'numero
    Case CellTypeCheckBox, CellTypeCurrency, CellTypeNumber
       
       spdGetValue = spdGrilla.Text
    
    End Select
  
    'recupero columna anterior
    spdGrilla.Col = intColAnt
  
  End If
  
End Function

'
' GENERA ARCHIVO PARA IMPORTAR TRABAJA CON FORMULARIO gridFRM, importExportFRM Y
'  CON TABLA MENUOPCIONES CON LA DEFINICION DE CAMPOS OBLIGATORIOS Y AUTOMATICOS
'
Function spdGeneraArchivoParaImportar(gridFRM) As Boolean

  'verifico existencia de archivo
  If Dir(App.Path & "\Importar\" & gridFRM.Caption & "_estructura.xls") <> "" Then
    intRes = MsgBox("El archivo: " & App.Path & "\Importar\" & gridFRM.Caption & "_estructura.Xls" & Chr(13) & Chr(13) & "Ya existe. Sobreescribe ?", vbApplicationModal + vbQuestion + vbYesNo)
    If intRes = 7 Then 'click boton NO, cancela generacion estructura de importacion
      Exit Function
    End If
  End If
         
  'cambio puntero mouse
  Screen.MousePointer = vbHourglass
        
  'cargo auxiliar
  Load importExportFRM
    
  'determino cantidad de filas y columnas para grilla a generar
  importExportFRM.spdAuxiliar.MaxRows = 2
  importExportFRM.spdAuxiliar.MaxCols = 0
          
  'recorro grilla activa
  Dim intCuenta As Integer
  Dim varNombreColumna As Variant
  Dim strValor As String
  Dim strObligatoria As String
  Dim strAutomatico As String
  For intCol = 1 To gridFRM.spdGrid.MaxCols
        
    'tomo nombre de columna
    gridFRM.spdGrid.GetText intCol, 0, varNombreColumna
        
    'busco columna que no este definida como NoMuestraEnEdit
    If InStr(LCase(gridFRM.DataNoMuestraEnEdit), "[" & LCase(varNombreColumna) & "]") = 0 Then
                
      'agrego columna a la grilla a generar
      importExportFRM.spdAuxiliar.MaxCols = importExportFRM.spdAuxiliar.MaxCols + 1
                
      'puntero a fila y columna
      importExportFRM.spdAuxiliar.Row = 1
      importExportFRM.spdAuxiliar.Col = importExportFRM.spdAuxiliar.MaxCols
                
      'busco columna que este definida como obligatorio
      If InStr(LCase(gridFRM.DataObligatorioEnEdit), "[" & LCase(varNombreColumna) & "]") <> 0 Then
        strObligatoria = "OBLIGATORIO"
        importExportFRM.spdAuxiliar.BackColor = RGB(225, 241, 255)
        importExportFRM.spdAuxiliar.Row = 2
        importExportFRM.spdAuxiliar.BackColor = RGB(225, 241, 255)
        importExportFRM.spdAuxiliar.Row = 1
      Else
        strObligatoria = ""
      End If
        
      'busco columna que este definida como automatico
      If InStr(LCase(gridFRM.DataSoloLecturaEnEdit), "[" & LCase(varNombreColumna) & "]") <> 0 Then
        strAutomatico = "AUTOMATICO"
        importExportFRM.spdAuxiliar.BackColor = RGB(225, 225, 225)
        importExportFRM.spdAuxiliar.Row = 2
        importExportFRM.spdAuxiliar.BackColor = RGB(225, 225, 225)
        importExportFRM.spdAuxiliar.Row = 1
      Else
        strAutomatico = ""
      End If
                
      'consulto tipo de celda para formatear y alinear
      Select Case gridFRM.DataFields(varNombreColumna).Type
          
      'Bit
      Case conBit
        strValor = strValor & "(0/1)"
        
      'enteros
      Case conSmallInt, conInt, conTinyInt
        strValor = "(entero numerico)"
                    
      'decimal
      Case conMoney, conSmallMoney, conReal, conFloat, conNumeric, conDecimal
        strValor = "(decimal numerico)"
            
      'fecha
      Case conSmallDateTime, conDateTime
        strValor = "(fecha dd/mm/yyyy)"
        
      'string
      Case conChar, conNchar, conVarchar, conText, conNVarchar, conText
        strValor = "(texto) " & gridFRM.DataFields(varNombreColumna).DefinedSize
        
      End Select
        
      'le paso el valor a la grilla fila columna
      importExportFRM.spdAuxiliar.SetText intCol, 1, varNombreColumna
      importExportFRM.spdAuxiliar.FontBold = True
      importExportFRM.spdAuxiliar.SetText intCol, 2, strValor & " " & strObligatoria & " " & strAutomatico
        
    End If  'noMuestraEnEdit
      
  Next  'cuenta columnas
                    
  'saca proteccion para excel
  importExportFRM.spdAuxiliar.Protect = False
                  
  'genero el excel con estructura
  intRes = importExportFRM.spdAuxiliar.ExportToExcel(App.Path & "\Importar\" & gridFRM.Caption & "_estructura", "default", "")
          
  'cierro auxiliar
  Unload importExportFRM
          
  'recupero puntero mouse
  Screen.MousePointer = vbDefault
        
  If intRes = -1 Then 'ok
    intRes = MsgBox("Se generó archivo para importar para: " & gridFRM.Caption & ", en:" & Chr(13) & Chr(13) & App.Path & "\Import\" & gridFRM.Caption & "_estructura.Xls", vbApplicationModal + vbInformation + vbOKOnly)
  Else ' intRes = 0
    intRes = MsgBox("La estructura no se pudo generar." & Chr(13) & Chr(13) & "Verificar que el archivo no se encuentre abierto.", vbApplicationModal + vbCritical + vbOKOnly)
  End If

End Function
'
'FUNCION EXPORTA DATOS DE UNA GRILLA A UN EXCEL
'-1:error, 0:no hay valores a exportar, 1:todo bien
'
Function spdExportarAExcel(gridFRM As gridFRM) As Integer

  If gridFRM.spdGrid.MaxRows = 0 Then
    spdExportarAExcel = 0
    Exit Function
  End If
        
  'mouse reloj
  If Dir(App.Path & "\iconos\_24x24_excel.ico") <> "" Then
    Screen.MouseIcon = LoadPicture(App.Path & "\iconos\_24x24_excel.ico")
    Screen.MousePointer = vbCustom
  Else
    Screen.MousePointer = vbHourglass
  End If
    
  'inserto 1 linea arriba de todo para ponerle los titulos de las columnas
  gridFRM.spdGrid.MaxRows = gridFRM.spdGrid.MaxRows + 1
  gridFRM.spdGrid.InsertRows 1, 1
  
  'cambio altura
  gridFRM.spdGrid.RowHeight(1) = 15
  
  For intRes = 1 To gridFRM.spdGrid.MaxCols
    gridFRM.spdGrid.GetText intRes, 0, titulo
    gridFRM.spdGrid.Row = 1
    gridFRM.spdGrid.Row2 = 1
    gridFRM.spdGrid.Col = intRes
    gridFRM.spdGrid.Col2 = intRes
    gridFRM.spdGrid.CellType = CellTypeEdit
    gridFRM.spdGrid.BackColor = RGB(205, 238, 254)
    gridFRM.spdGrid.SetText intRes, 1, titulo
  Next
  
  'saca proteccion para excel
  gridFRM.spdGrid.Protect = False
  
  'exporto
  intRes = gridFRM.spdGrid.ExportToExcel(App.Path & "\exportar\" & gridFRM.Caption, "", "")
 
  'elimino fila 1 que servia para exportar los nombres de columnas
  gridFRM.spdGrid.DeleteRows 1, 1
  gridFRM.spdGrid.MaxRows = gridFRM.spdGrid.MaxRows - 1
    
  'mouse default
  Screen.MousePointer = vbDefault
    
  'exporto OK
  If intRes = -1 Then
    spdExportarAExcel = 1
  Else
    spdExportarAExcel = -1
  End If

End Function


