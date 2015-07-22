Attribute VB_Name = "modGenerals"

' abre formulario segun opcion del menu

Public Function frmToShow(ByVal frmMenu As Form, ByRef frmActivo As Form, ByRef frmShow As Form, Optional ByVal blnAdjust As Boolean)
  Dim intLVW, intPositionLVW As Integer
  Dim lvw As ListView
  
  ' vacio informacion de la barra de estado panel 1 filtro
  MDIMenu.staGeneral.Panels(1).Text = "Filtro: ninguno"
  MDIMenu.staGeneral.Panels(1).AutoSize = sbrContents
  MDIMenu.staGeneral.Panels(1).Alignment = sbrCenter
  
  ' vacio informacion de la barra de estado panel 2 Orden
  MDIMenu.staGeneral.Panels(2).Text = "Orden: ninguno"
  MDIMenu.staGeneral.Panels(2).AutoSize = sbrContents
  MDIMenu.staGeneral.Panels(2).Alignment = sbrCenter
  
  ' vacio informacion de la barra de estado panel 3 Info General
  MDIMenu.staGeneral.Panels(3).Text = "Info: "
  MDIMenu.staGeneral.Panels(3).AutoSize = sbrContents
  MDIMenu.staGeneral.Panels(3).Alignment = sbrCenter

  Unload frmActivo                                  ' descargo form actual
  Load frmShow                                      ' cargo form a mostrar
  frmShow.Top = frmMenu.Top                         ' seteo ubicacion arriba
  frmShow.Left = frmMenu.Left + frmMenu.Width + 15  ' seteo ubicacion izquierda
  frmShow.Width = MDIMenu.Width - frmMenu.Width     ' Seteo Ancho segun mdi
  frmShow.Height = MDIMenu.Height                   ' Seteo Alto segun mdi
  'frmShow.Height = conFrmHeight                    ' Seteo Alto antes
  'frmShow.Width = conFrmWidth                      ' Seteo Ancho antes
  frmShow.BorderStyle = conFrmBorderStyle           ' Seteo Estilo Borde
  
  ' cuento la cantidad de listview en el form
  intLVW = 0
  For intIndice = 0 To frmShow.Controls.Count - 1
    If TypeName(frmShow.Controls(intIndice)) = "ListView" Then
        intLVW = intLVW + 1
        intPositionLVW = intIndice
        Set lvw = frmShow.Controls(intIndice)
    End If
  Next
  
  For intIndice = 0 To frmShow.Controls.Count - 1
    
    Select Case TypeName(frmShow.Controls(intIndice))
      
    Case Is = "ToolBar"
        
      frmShow.Controls(intIndice).Align = conTlbAlign
      frmShow.Controls(intIndice).Appearance = conTlbAppearance
      frmShow.Controls(intIndice).Height = conTlbHeight
      frmShow.Controls(intIndice).Width = MDIMenu.Width - frmMenu.Width - 150
      frmShow.Controls(intIndice).Top = conTlbTop
      frmShow.Controls(intIndice).Left = conTlbLeft
      
    Case Is = "Label"
        
      frmShow.Controls(intIndice).Height = conLblHeight
      'frmShow.Controls(intIndice).Width = conLblWidth
      frmShow.Controls(intIndice).Width = MDIMenu.Width - frmMenu.Width - 150
      frmShow.Controls(intIndice).Top = conLblTop
      frmShow.Controls(intIndice).Left = conLblLeft
      frmShow.Controls(intIndice).BackColor = conLblBackColor
      frmShow.Controls(intIndice).ForeColor = conLblForeColor
      frmShow.Controls(intIndice).Font = conLblFont
      frmShow.Controls(intIndice).FontBold = conLblFontBold
      frmShow.Controls(intIndice).FontSize = conLblFontSize
      frmShow.Controls(intIndice).Alignment = conLblAlignment
      frmShow.Controls(intIndice).BackStyle = conLblBackStyle
      frmShow.Controls(intIndice).BorderStyle = conLblBorderStyle

    Case Is = "ListView"
      
      ' ancho, para todos los ListView
      frmShow.Controls(intIndice).Width = MDIMenu.Width - frmMenu.Width - 150
      
      ' si el form tiene solo un LVW lo acomoda a la izquiqerda arriba
      If intLVW = 1 Then
        frmShow.Controls(intIndice).Top = conLvwTop
        frmShow.Controls(intIndice).Left = conLvwLeft
        frmShow.Controls(intIndice).Height = MDIMenu.Height - (conTlbHeight + conLblHeight + 700)
      Else
      
      End If
        
      ' largo, si hay uno solo o para el ultimo ListView del Form lo estira al final de mdi
      ' los 700 incluye borde del formulario MDI y altura de barra de estado
      'If intIndice = intPositionLVW Then
      '  frmShow.Controls(intIndice).Height = MDIMenu.Height - (conTlbHeight + conLblHeight + 700)
      'End If
      
      ' codigo anterior todavia no lo borro
      'If blnAdjust Then
      
        'frmShow.Controls(intIndice).Height = conLvwHeight
        'frmShow.Controls(intIndice).Width = conLvwWidth
        'frmShow.Controls(intIndice).Height = MDIMenu.Height       ' segun mdi
        'frmShow.Controls(intIndice).Width = MDIMenu.Width         ' segun mdi
        'frmShow.Controls(intIndice).Top = conLvwTop
        'frmShow.Controls(intIndice).Left = conLvwLeft
        
      'End If
      
    Case Is = "CrystalReport"
      
    End Select
      
  Next
    
  ' pongo visible el Form y Seteo Frm como Frm Activo
    
  frmShow.Show
  Set frmActivo = frmShow
  
  End Function

' cambia apariencia de un ListView

Public Function ListViewAppearanceChange(ByRef lvw As ListView)
  Dim intColumna As Integer

  lvw.BackColor = conListView_BackColor     ' color fondo
  lvw.ForeColor = conListView_ForeColor     ' color letra
  lvw.View = lvwReport                      ' forma en que se ven los datos
  lvw.LabelEdit = lvwManual                 ' editar datos o no
  lvw.HideSelection = False                 ' cuando no es control activo deja en gris atenuado la seleccion
  lvw.MultiSelect = True                    ' seleccion multiple
  lvw.AllowColumnReorder = True             ' cambiar orden a columnas
  lvw.FullRowSelect = True                  ' seleccion fila completa
  lvw.Sorted = False                        ' ordenar columnas
  lvw.Gridlines = False                     ' grilla

End Function


' llena una ListView segun Query
  
Public Function ListViewRefresh(ByRef lvw As ListView, ByVal strQuery As String, Optional ByRef strStructure As Variant)
  Dim intCuenta, intAligment, intColumna  As Integer
  Dim rsCount As Long
  Dim fldName, strTYPE, strFmat, strWidth, strWAUX, strAux() As String
  Dim strViewFmat, strViewFmatAUX As String
  Dim li As ListItem
  Dim rs As New ADODB.Recordset
  Dim fld As ADODB.Field
  
  ' leo ini formato y ancho columna
  arrFormat = keyIniToArray(strTableNameActual, "format")
  arrWidth = keyIniToArray(strTableNameActual, "width")
  
  ' obtengo recordset con titulos para ListView
  Set rs = adoOpenRecordset(strQuery)

  ' elimina el contenido del listview y datos y titulos
  lvw.ListItems.Clear
  lvw.ColumnHeaders.Clear

  ' crea las columnas con los titulos
  intColumna = 0
  For Each fld In rs.Fields

    ' evalua tipo de campo
    
    Select Case fld.Type
    
    Case conDateTime, conSmallDateTime
      intAligment = lvwColumnLeft
      strTYPE = "date"
      strFmat = "dd/mm/yyyy"
      strViewFmat = "dd/mm/yyyy"
    
    Case conDecimal, conFloat, conMoney, conNumeric, conReal, conSmallMoney
      intAligment = lvwColumnRight
      strTYPE = "numeric"
      strFmat = "#########.###-"
      strViewFmat = "########0.000"
    
    Case conInt, conSmallInt, conTinyInt
      intAligment = lvwColumnRight
      strTYPE = "numeric"
      strFmat = "#######"
      strViewFmat = "######0"

    Case conChar, conNchar, conText, conNtext, conVarchar, conNVarchar
      intAligment = lvwColumnLeft
      strTYPE = "string"
      strFmat = "@" & Trim(fld.DefinedSize)
      strViewFmat = ""
    
    Case adBoolean
      intAligment = lvwColumnRight
      strTYPE = "boolean"
      strFmat = ""

    Case Else
      intAligment = -1

    End Select

    ' busca ancho de columna en ini, si no encuentra deja standard
    strWidth = "2000"
    strWAUX = findColumn(arrWidth, fld.Name)
    If strWAUX <> "" Then
      strWidth = strWAUX
    End If

    ' si el field es correcto, creo una columna
    ' pero si es la primer columna, debe alinearse a la izquierda
    
    If intAligment <> -1 Then
      If lvw.ColumnHeaders.Count = 0 Then intAligment = lvwColumnLeft
      lvw.ColumnHeaders.Add , , fld.Name, strWidth, intAligment
    End If

    ' guardo nombre columna, tipo de dato, y formato
    ' este array es devuelto por la funcion y se utiliza
    ' para filtrar informacion de listview
    ReDim Preserve strAux(1 To 3, intColumna)
    strAux(1, intColumna) = fld.Name
    strAux(2, intColumna) = strTYPE
    strAux(3, intColumna) = strFmat
    
    ' paso a la columna siguiente
    intColumna = intColumna + 1

  Next

  ' lleno array con estructura de la tabla
  strStructure = strAux

  rsCount = 0
  
  Do Until rs.EOF

    rsCount = rsCount + 1
    
    ' agrego el objeto listitem principal (la columna)
    fldName = lvw.ColumnHeaders(1).Text
    Set li = lvw.ListItems.Add(, , rs.Fields(fldName) & "")
    
    ' agrego todos los objetos listSubItem siguientes
    For i = 2 To lvw.ColumnHeaders.Count
      
      ' toma el nombre de la columna
      fldName = lvw.ColumnHeaders(i)
    
      ' tomo formato estandard
      Select Case rs.Fields(fldName).Type
    
      Case conDateTime, conSmallDateTime
        strViewFmat = "dd/mm/yyyy"
    
      Case conDecimal, conFloat, conMoney, conNumeric, conReal, conSmallMoney
        strViewFmat = "########0.000"
    
      Case conInt, conSmallInt, conTinyInt
        strViewFmat = "######0"

      Case conChar, conNchar, conText, conNtext, conVarchar, conNVarchar
        strViewFmat = ""
    
      Case adBoolean
        strViewFmat = ""

      End Select
    
      ' busca formato en ini, si no encuentra deja standard
      strViewFmatAUX = findColumn(arrFormat, rs.Fields(fldName).Name)
      If strViewFmatAUX <> "" Then strViewFmat = strViewFmatAUX
    
      ' agrega dato al listView
      li.ListSubItems.Add , , Format(rs.Fields(fldName), strViewFmat)
    
    Next
    
    ' cheque el maximo de registros a msotrar
    
    If rsCount = MaxRecords Then Exit Do
    rs.MoveNext
  
  Loop

  ' ajusto columnas
  'intRes = lvwAdjustColumn(LVW, True)

End Function

' ajusta las columnas de un ListView

Public Function lvwAdjustColumn(lvw As ListView, Optional blnForHeaders As Boolean, Optional varColumn As Variant)

  Dim lngRow, lngColumnDesde, lngColumnHasta As Long, lngCol As Long
  Dim sngWidth, sngMaxWidth As Single
  Dim stdSaveFont As StdFont
  Dim intSaveScaleMode  As Integer
  Dim strCellText As String

  If lvw.ListItems.Count = 0 Then Exit Function
  
  ' guarda la fuente utilizada por el formulario padre
  ' y forzar la fuente ListView, Necesito hacer esta operacion
  ' para poder utilizar el metodo textwidth del formulario

  Set stdSaveFont = lvw.Parent.Font
  Set lvw.Parent.Font = lvw.Font
  
  ' forzar intscalemode = vbtwips para el padre
  
  intSaveScaleMode = lvw.Parent.ScaleMode
  lvw.Parent.ScaleMode = vbTwips

  ' seteo rango de columnas
  lngColumnDesde = 1
  lngColumnHasta = lvw.ColumnHeaders.Count

  If Not IsNull(varColumn) Then
    
    If TypeName(varColumn) = "String" Then
      
      For intIndice = 1 To lvw.ColumnHeaders.Count   ' recorro columnas
        If Format(Mid(Trim(lvw.ColumnHeaders(intIndice)), 1, Len(varColumn)), "<") = Format(varColumn, "<") Then
          If intIndice = 1 Then
            lngColumnDesde = 1
            lngColumnHasta = 1
          Else
            lngColumnDesde = intIndice
            lngColumnHasta = intIndice
          End If
        End If
      Next
    
    End If
    
  End If

  ' ajusta columnas
  For lngCol = lngColumnDesde To lngColumnHasta
  
    ' las columnas con ancho cero las deja iguales porque son acultas seguramente
    If lvw.ColumnHeaders(lngCol).Width <> 0 Then
  
    sngMaxWidth = 0
    If blnForHeaders Then
      sngMaxWidth = lvw.Parent.TextWidth(lvw.ColumnHeaders(lngCol).Text) + 200
    End If

    For lngRow = 1 To lvw.ListItems.Count
      ' recupera la cedena de texto de listitems o listsubitems
      If lngCol = 1 Then
        strCellText = lvw.ListItems(lngRow).Text
      Else
        strCellText = lvw.ListItems(lngRow).ListSubItems(lngCol - 1).Text
      End If
      
      ' calcular su anchura teniendo en cuenta campos de texto que ocupra
      ' varias lineas
      
      sngWidth = lvw.Parent.TextWidth(strCellText) + 200
      
      ' actualizar sngmaxwidth si se localiza una cadena mas larga
      
      If sngWidth > sngMaxWidth Then sngMaxWidth = sngWidth

    Next
    
    ' modifica la anchura de la columna
    
    lvw.ColumnHeaders(lngCol).Width = sngMaxWidth
    
    End If
    ' width <> 0
    
  Next
  
  ' restaurar las propiedades del formulario padre
  
  Set lvw.Parent.Font = stdSaveFont
  lvw.Parent.ScaleMode = intSaveScaleMode
    
End Function

' devuelve el valor de un ListView
' correspondiente a un nombre de columna

Public Function lvwGetValue(ByVal lvw As ListView, ByVal var As Variant)
  Dim intIndice As Integer

  lvwGetValue = "Error"
  If lvw.SelectedItem Is Nothing Then Exit Function     ' chequeo que exista seleccion
  
  If TypeName(var) = "String" Then

    ' comparo el nombre de columna, con el nombre pasado com argumento
    ' si paso el nombre de la columna en forma incompleta, tambien
    ' lo encuentra y si hay 2 nombres de columna iguales, devuelve el primero

    For intIndice = 1 To lvw.ColumnHeaders.Count   ' recorro columnas

      If Format(Mid(Trim(lvw.ColumnHeaders(intIndice)), 1, Len(var)), "<") = Format(var, "<") Then
        If intIndice = 1 Then
          lvwGetValue = lvw.SelectedItem
        Else
          lvwGetValue = lvw.SelectedItem.SubItems(intIndice - 1)
        End If
        Exit Function
      End If
      
    Next
    
  End If

  ' tambien devuelve el valor correspondiente a un numero de columna

  If TypeName(var) = "Integer" Then
        
    For intIndice = 1 To lvw.ColumnHeaders.Count  ' recorro columnas
        
      If intIndice = var Then
        If intIndice = 1 Then
          lvwGetValue = lvw.SelectedItem
        Else
          lvwGetValue = lvw.SelectedItem.SubItems(intIndice - 1)
        End If
        Exit Function
      End If
      
    Next

  End If
    
End Function

' cambia el valor de una columna de un ListView
' correspondiente a un nombre o numero de columna

Public Function lvwSetValue(ByRef lvw As ListView, ByVal varColumn As Variant, ByVal varValue As Variant)
  Dim intIndice As Integer

  If lvw.SelectedItem Is Nothing Then Exit Function     ' chequeo que exista seleccion

  If TypeName(varColumn) = "String" Then

    ' comparo el nombre de columna, con el nombre pasado com argumento
    ' si paso el nombre de la columna en forma incompleta, tambien
    ' lo encuentra y si hay 2 nombres de columna iguales, asigna al primero

    For intIndice = 1 To lvw.ColumnHeaders.Count   ' recorro columnas

      If Format(Mid(Trim(lvw.ColumnHeaders(intIndice)), 1, Len(varColumn)), "<") = Format(varColumn, "<") Then
        If intIndice = 1 Then
          ' lvw.SelectedItem = varValue
        Else
          lvw.SelectedItem.SubItems(intIndice - 1) = varValue
        End If
        Exit Function
      End If
      
    Next
    
  End If

  ' tambien devuelve el valor correspondiente a un numero de columna

  If TypeName(varColumn) = "Integer" Then
        
    For intIndice = 1 To lvw.ColumnHeaders.Count  ' recorro columnas
        
      If intIndice = varColumn Then
         If intIndice = 1 Then
           'lvw.SelectedItem = varValue
         Else
           lvw.SelectedItem.SubItems(intIndice - 1) = varValue
         End If
         Exit Function
       End If
      
    Next

  End If

End Function


' oculta una columna x nombre o por numeo de columna

Public Function lvwHideColumn(ByVal lvw As ListView, ByVal varColumn As Variant)
  
  Dim intIndice As Integer

  If TypeName(varColumn) = "String" Then

    ' comparo el nombre de columna, con el nombre pasado com argumento
    ' si paso el nombre de la columna en forma incompleta, tambien
    ' lo encuentra y si hay 2 nombres de columna iguales, oculta el primero

    For intIndice = 1 To lvw.ColumnHeaders.Count   ' recorro columnas

      If Format(Mid(Trim(lvw.ColumnHeaders(intIndice)), 1, Len(varColumn)), "<") = Format(varColumn, "<") Then
        lvw.ColumnHeaders(intIndice).Width = 0
        Exit Function
      End If
      
    Next
    
  End If

  ' tambien Oculta la columna correspondiente a un numero

  If TypeName(varColumn) = "Integer" Then
        
    For intIndice = 1 To lvw.ColumnHeaders.Count  ' recorro columnas
        
      If intIndice = varColumn Then
        lvw.ColumnHeaders(intIndice).Width = 0
        Exit Function
      End If
      
    Next

  End If
  
End Function

'
' ordenar columna actual haciendo click arriba
'
Public Function lvwSortColumnActual(ByRef lvw As ListView) As Integer
  Dim intIndice, intRow, intRowCbo, intColUno, intColDos, intColTres As Integer
  Dim strLeyenda As String
  Dim lvwColumn As ColumnHeader
End Function


'
' ordenar columnas en un listview
'
Public Function lvwSortColumn(ByRef lvw As ListView) As Integer
  Dim intIndice, intRow, intRowCbo, intColUno, intColDos, intColTres As Integer
  Dim strLeyenda As String
  Dim lvwColumn As ColumnHeader

  ' cargo formulario
  Load frmSortData
      
  ' recorro array y lleno combos con las columnas
  For intCuenta = 0 To UBound(strStruc, 2)
    frmSortData.cboUno.AddItem strStruc(1, intCuenta)
    frmSortData.cboDos.AddItem strStruc(1, intCuenta)
    frmSortData.cboTres.AddItem strStruc(1, intCuenta)
  Next
      
  ' muestro formulario para seleccionar orden
  frmSortData.Show vbModal
 
  ' si acepto
  If blnAceptar Then
  
    ' cambio puntero mouse
    Screen.MousePointer = vbHourglass
  
    ' inicializo
    intColUno = 0
    intColDos = 0
    intColTres = 0
  
    ' identifico columnas que se seleccionaron, hasta un maximo de 3
    For intRow = 1 To lvw.ColumnHeaders.Count
      
      If frmSortData.cboUno.ListIndex <> -1 Then
        If Format(frmSortData.cboUno.List(frmSortData.cboUno.ListIndex), "<") = Format(lvw.ColumnHeaders(intRow).Text, "<") Then
          intColUno = intRow - 1
        End If
      End If
    
      If frmSortData.cboDos.ListIndex <> -1 Then
        If Format(frmSortData.cboDos.List(frmSortData.cboDos.ListIndex), "<") = Format(lvw.ColumnHeaders(intRow).Text, "<") Then
          intColDos = intRow - 1
        End If
      End If
      
      If frmSortData.cboTres.ListIndex <> -1 Then
        If Format(frmSortData.cboTres.List(frmSortData.cboTres.ListIndex), "<") = Format(lvw.ColumnHeaders(intRow).Text, "<") Then
          intColTres = intRow - 1
        End If
      End If
      
    Next
  
    ' agrego una columna al final para poder ordenar
    lvw.ColumnHeaders.Add , , "", 0, lvwColumnLeft
    
    ' recorro filas de LVW
    For intRow = 1 To lvw.ListItems.Count
      
      ' vacio info de fila nueva para hacer un match de las columnas seleccionadas
      lvw.ListItems(intRow).SubItems(lvw.ColumnHeaders.Count - 1) = ""
    
      ' agrega columna 1
      If intColUno <> 0 Then
        lvw.ListItems(intRow).SubItems(lvw.ColumnHeaders.Count - 1) = lvw.ListItems(intRow).SubItems(lvw.ColumnHeaders.Count - 1) & Trim(lvw.ListItems(intRow).SubItems(intColUno))
        strLeyenda = lvw.ColumnHeaders(intColUno + 1)
      End If
    
      ' agrega columna 2
      If intColDos <> 0 Then
        lvw.ListItems(intRow).SubItems(lvw.ColumnHeaders.Count - 1) = lvw.ListItems(intRow).SubItems(lvw.ColumnHeaders.Count - 1) & Trim(lvw.ListItems(intRow).SubItems(intColDos))
        strLeyenda = strLeyenda & "+" & lvw.ColumnHeaders(intColDos + 1)
      End If
    
      ' agrega columna 3
      If intColTres <> 0 Then
        lvw.ListItems(intRow).SubItems(lvw.ColumnHeaders.Count - 1) = lvw.ListItems(intRow).SubItems(lvw.ColumnHeaders.Count - 1) & Trim(lvw.ListItems(intRow).SubItems(intColTres))
        strLeyenda = strLeyenda & "+" & lvw.ColumnHeaders(intColTres + 1)
      End If
    
    Next
  
    ' ordeno
    lvw.Sorted = True
    lvw.SortKey = lvw.ColumnHeaders.Count - 1
    
    If frmSortData.cboForma = "Ascendente" Then
      lvw.SortOrder = lvwAscending
      strLeyenda = strLeyenda & " Asc"
    Else
      lvw.SortOrder = lvwDescending
      strLeyenda = strLeyenda & " Desc"
    End If
  
    ' deshabilito orden
    lvw.Sorted = False
    
    ' elimino ultima columna generada para ordenar
    lvw.ColumnHeaders.Remove lvw.ColumnHeaders.Count
    
    ' muestro eb barra de estado orden
    MDIMenu.staGeneral.Panels(2).Text = "Orden: " & strLeyenda
    MDIMenu.staGeneral.Panels(2).AutoSize = sbrContents
    MDIMenu.staGeneral.Panels(2).Alignment = sbrCenter
    
    ' vuelvo puntero mouse
    Screen.MousePointer = vbDefault
    
  End If

  ' descargo form
  Unload frmSortData
  
End Function

'
' Suma una columna de un objeto
'
Public Function objSumColumn(ByVal obj As Object, ByVal varColumn As Variant) As Currency
  Dim intCol, intItem As Integer
  Dim sngSuma As Currency

  ' cambio puntero mouse
  Screen.MousePointer = vbHourglass
  
  objSumColumn = 0
  
  Select Case TypeName(obj)
  
  Case "ListView"

    If TypeName(varColumn) = "String" Then

      ' comparo el nombre de columna, con el nombre pasado con argumento
      ' si paso el nombre de la columna en forma incompleta, tambien
      ' lo encuentra y si hay 2 nombres de columna iguales, oculta el primero
      For intCol = 1 To obj.ColumnHeaders.Count   ' recorro columnas
        
        If Format(Mid(Trim(obj.ColumnHeaders(intCol)), 1, Len(varColumn)), "<") = Format(varColumn, "<") Then
          
          ' recorro items
          sngSuma = 0
          For intItem = 1 To obj.ListItems.Count
            
            ' si son numericos los sumo
            If IsNumeric(obj.ListItems(intItem).SubItems(intCol - 1)) Then
              sngSuma = sngSuma + (obj.ListItems(intItem).SubItems(intCol - 1))
            End If
          
          Next
        
          objSumColumn = sngSuma
        
        End If
      
      Next
    
    End If

    If TypeName(varColumn) = "Integer" Then

      ' si la columna esta dentro del rango de columnas sumo
      If varColumn >= 0 And varColumn <= obj.ColumnHeaders.Count - 1 Then
      
        ' recorro items
        sngSuma = 0
        For intItem = 1 To obj.ListItems.Count
            
          ' si son numericos los sumo
          If IsNumeric(obj.ListItems(intItem).SubItems(varColumn)) Then
            sngSuma = sngSuma + obj.ListItems(intItem).SubItems(varColumn)
          End If
          
        Next
        
        objSumColumn = sngSuma
    
      End If
    
    End If

  End Select
  
  ' vuelvo puntero mouse
  Screen.MousePointer = vbDefault

End Function

'
' Busca informacion en un objeto X
'
Public Function FindData(ByVal obj As Object)
  Dim intCuenta, intRes, intIndice, intEncontro As Integer
  Dim lit As ListItem
  Dim blnEncontro As Boolean
  
  frmFindData.Show vbModal
  
  If blnAceptar Then
  
    ' cambio puntero mouse
    Screen.MousePointer = vbHourglass
  
    Select Case TypeName(obj)
  
    Case "ListView"
    
      ' elimina selecciones existentes
  
      For intCuenta = 1 To obj.ListItems.Count
        obj.ListItems(intCuenta).Selected = False
      Next
      
      ' busca por primera vez el texto ingresado
      
      intIndice = 1
      blnEncontro = True
      Set lit = obj.FindItem(frmFindData.txtQue, 1, intIndice)
      
      ' si no encontro muestro mensaje, cierro frm y salgo
      If lit Is Nothing Then
        
        MDIMenu.staGeneral.Panels(3).Text = "Info: No se encontraron elementos."
        MDIMenu.staGeneral.Panels(3).AutoSize = sbrContents
        MDIMenu.staGeneral.Panels(3).Alignment = sbrCenter
        blnEncontro = False
      
      End If
     
      If blnEncontro Then
           
        ' encontro el primero, lo selecciono y
        ' sumo 1 al indice para proxima busqueda
        ' mientras encunetre algo, sigo buscando
        ' hasta que se terminen los elementos y
        ' los voy seleccionando a cada uno
        ' cuando finalizo pongo foco en lvw
     
        lit.Selected = True
        intIndice = lit.Index + 1
      
        ' recorro todo el listview y busco
        intEncontrados = 0
        While intIndice <= obj.ListItems.Count
          Set lit = obj.FindItem(frmFindData.txtQue, 1, intIndice)
          If Not (lit Is Nothing) Then
            lit.Selected = True
            intIndice = lit.Index + 1
            intEncontro = intEncontro + 1
          Else
            intIndice = obj.ListItems.Count + 1
          End If
        Wend
  
        ' muestro mensaje de cantidad encontrada
        MDIMenu.staGeneral.Panels(3).Text = "Info: Se encontraron " & Trim(str(intEncontro + 1)) & " elementos."
        MDIMenu.staGeneral.Panels(3).AutoSize = sbrContents
        MDIMenu.staGeneral.Panels(3).Alignment = sbrCenter
  
      End If
  
      obj.SetFocus
    
    End Select

    ' vuelvo puntero mouse
    Screen.MousePointer = vbDefault

  End If

  Unload frmFindData

End Function

'
' filtra informacion en un objeto X
'
Public Function FilterData(ByVal obj As Object) As String
  Dim intCuenta, intRes, intIndice, intEncontro As Integer
  Dim lit As ListItem
  Dim blnEncontro As Boolean
 
  Select Case TypeName(obj)
  
  Case "ListView"
    
    ' cargo formulario
    Load frmFilterData
      
    ' recorro array y lleno combo con las columnas
    For intCuenta = 0 To UBound(strStruc, 2)
      frmFilterData.cboColumna.AddItem strStruc(1, intCuenta)
    Next
      
    ' muestro formulario para filtro
    frmFilterData.Show vbModal
  
    ' devuelvo un string con la condicion
    For intCuenta = 0 To frmFilterData.lstCondicion.ListCount - 1
      FilterData = FilterData & " " & frmFilterData.lstCondicion.List(intCuenta)
    Next
  
    ' descargo form
    Unload frmFilterData
    
    ' muestro mensaje con filtro
    MDIMenu.staGeneral.Panels(1).Text = "Filtro: " & FilterData
    MDIMenu.staGeneral.Panels(1).AutoSize = sbrContents
    MDIMenu.staGeneral.Panels(1).Alignment = sbrCenter
    
  End Select

End Function

'
' Ordenar informacion en un objeto X
'
Public Function SortData(ByVal obj As Object) As String
  Dim intCuenta, intRes, intIndice, intEncontro As Integer
  Dim strFilterData As String
  Dim lit As ListItem
  Dim blnEncontro As Boolean
 
  Select Case TypeName(obj)
  
  Case "ListView"
    
    ' cargo formulario
    Load frmSortData
      
    ' recorro array y lleno combo con las columnas
    For intCuenta = 0 To UBound(strStruc, 2)
      frmSortData.cboUno.AddItem strStruc(1, intCuenta)
      frmSortData.cboDos.AddItem strStruc(1, intCuenta)
      frmSortData.cboTres.AddItem strStruc(1, intCuenta)
    Next
      
    ' muestro formulario para filtro
    frmSortData.Show vbModal
  
    ' devuelvo un string con la condicion
    For intCuenta = 0 To frmFilterData.lstCondicion.ListCount - 1
      strFilterData = strFilterData & " " & frmFilterData.lstCondicion.List(intCuenta)
    Next
  
    ' descargo form
    Unload frmFilterData
    
    ' muestro mensaje con filtro
    MDIMenu.staGeneral.Panels(1).Text = "Filtro: " & strFilterData
    MDIMenu.staGeneral.Panels(1).AutoSize = sbrContents
    MDIMenu.staGeneral.Panels(1).Alignment = sbrCenter
    
  End Select

End Function

'
' llena un ComboBox segun Query
  
Public Function ComboBoxRefresh(ByRef cbo As ComboBox, ByVal strQuery As String)

  Dim rs As New ADODB.Recordset

  ' abro recordset
  
  Set rs = adoOpenRecordset(strQuery)

  ' elimina el contenido del conbobox
  
  cbo.Clear

  ' recorro recordset y lleno Combo Box
  
  While Not rs.EOF

    ' chequeo en que posicion esta el ID
    ' ejemplo: IDCLiente, Descripcion o
    '          Descripcion, IDCliente

    If Left(rs(1).Name, 2) = "ID" Then
      cbo.AddItem rs(0)
      cbo.ItemData(cbo.NewIndex) = rs(1)
    Else
      cbo.AddItem rs(1)
      cbo.ItemData(cbo.NewIndex) = rs(0)
    End If
    
    rs.MoveNext
    
  Wend

  rs.Close

End Function

' busca en un ComboBox un Item segun Description
  
Public Function ComboBoxFindItem(ByRef cbo As ComboBox, ByVal strDescription As String)
  
  ComboBoxFindItem = -1
  For intCuenta = 0 To cbo.ListCount - 1
  
    If cbo.List(intCuenta) = strDescription Then
      ComboBoxFindItem = intCuenta
      Exit For
    End If
  
  Next

End Function

' averigua si un elemento seleccionado de un ComboBox
' es parte de la lista o se ingreso a mano

Public Function ComboBoxNotinList(ByVal cbo As ComboBox) As Boolean

  ComboBoxNotinList = False
  If cbo.Text <> "" And cbo.ListIndex = -1 Then
    ComboBoxNotinList = True
  End If

End Function

' agrega un elemento a un Combo Box

Public Function ComboBoxAddItem(ByVal Frm As Form, ByVal cbo As ComboBox, ByVal strFmat As String, ByVal strStore As String, ByVal strView As String) As String
  
  ComboBoxAddItem = ""
  
  Load frmAdd
  frmAdd.txtDato.MaxLength = IIf(Not IsNull(strFmat), Right(strFmat, Len(strFmat) - 1), 50)
  frmAdd.Left = Frm.Left + Frm.Controls(cbo.Name).Left + 10
  frmAdd.Top = Frm.Top + Frm.Controls(cbo.Name).Top + 310
  frmAdd.Width = Frm.Controls(cbo.Name).Width
  frmAdd.Height = Frm.Controls(cbo.Name).Height
  frmAdd.Show vbModal

  If blnAceptar Then
    
    intRes = adoExecuteQuery("EXEC " & strStore & " '" & frmAdd.txtDato.Text & "'")
    intRes = ComboBoxRefresh(cbo, strView)
    cbo.ListIndex = ComboBoxFindItem(cbo, frmAdd.txtDato.Text)
    ComboBoxAddItem = frmAdd.txtDato.Text
    
  End If
  
  Unload frmAdd

End Function


' validacion de datos

Public Function DataValidate(ByRef obj As Object, Optional ByVal strFmat As String, Optional ByVal blnObligatorio As Boolean)
  
  Dim intResultado, intLenMaxima, intCaracter, a  As Integer
  Dim intLenEnterosFmat, intLenDecimalesFmat As Integer
  Dim intLenEnterosValor, intLenDecimalesValor As Integer
  Dim strValidos As String

  DataValidate = vbError

  ' valido segun tipo de objeto que recibo

  Select Case TypeName(obj)
  
  Case "TextBox"
  
    ' si el dato no es obligatorio y esta vacio EXIT
  
    If Not blnObligatorio And obj.Text = "" Then
      DataValidate = True
      Exit Function
    End If
  
    ' valido dentro de TextBox que tipo de dato es
    
    Select Case Left(strFmat, 1)
  
    ' # numerico, . con decimales, - acepta negativos
  
    Case "#", "."
       
      ' cadena standard de validacion
       
      strValidos = "0123456789"
       
      ' si formato acepta decimales agrego el . como caracter valido
      ' tambien calculo la longitud maxima de los enteros y los decimales
      
      If InStr(1, strFmat, ".") Then
        strValidos = strValidos & "."
        intLenEnterosFmat = Len(Left(strFmat, InStr(1, strFmat, ".") - 1))
        intLenDecimalesFmat = Len(Right(strFmat, Len(strFmat) - InStr(1, strFmat, ".")))
      Else
        intLenEnterosFmat = Len(strFmat)
        intLenDecimalesFmat = 0
      End If
      
      ' si formato acepta negativos agrego el - como caracter valido
      ' tambien resto 1 a longitud maxima de decimales por el signo -
       
      If InStr(1, strFmat, "-") <> 0 Then
        strValidos = strValidos & "-"
        intLenDecimalesFmat = IIf(intLenDecimalesFmat > 0, intLenDecimalesFmat - 1, 0)
      End If
       
      ' calculo longitud maxima de caracteres
      ' enteros y decimales del valor ingresado
      
      If InStr(1, obj.Text, ".") Then
        intLenEnterosValor = Len(Left(obj.Text, InStr(1, obj.Text, ".") - 1))
        intLenDecimalesValor = Len(Right(obj.Text, Len(obj.Text) - InStr(1, obj.Text, ".")))
      Else
        intLenEnterosValor = Len(obj.Text)
        intLenDecimalesValor = 0
      End If
       
      ' si el valor tiene un signo - resto 1 a la cantidad entera
      
      If InStr(1, obj.Text, "-") Then
        intLenEnterosValor = IIf(intLenEnterosValor > 0, intLenEnterosValor - 1, 0)
      End If
       
      If Len(obj.Text) <> 0 Then
       
        ' valido datos segun strValidos
      
        For a = 1 To Len(obj.Text)
      
          ' si encuentra algun valor invalido, ya sean caracteres
          ' o sobrepasa las longitudes maximas del formato
          ' pone foco en control con error y pinta el valor del
          ' control pone texto de error, devuelve un error, y escapa
      
          If InStr(1, strValidos, Mid(obj.Text, a, 1)) = 0 Or _
            intLenEnterosValor > intLenEnterosFmat Or _
            intLenDecimalesValor > intLenDecimalesFmat Or _
            Not IsNumeric(obj.Text) Then
            obj.SetFocus
            obj.SelStart = 0
            obj.SelLength = 999
            obj.ToolTipText = Right(obj.Name, Len(obj.Name) - 3) & ": Deber ser un valor con formato, " & strFmat
            DataValidate = vbError
            Exit Function
          End If
      
        Next
    
      Else
            
        obj.SetFocus
        obj.SelStart = 0
        obj.SelLength = 999
        obj.ToolTipText = Right(obj.Name, Len(obj.Name) - 3) & ": Deber ser un valor con formato, " & strFmat
        DataValidate = vbError
        Exit Function
            
      End If
    
      ' numerico validado
    
      DataValidate = True
    
    ' fecha, d dia, m mes, y año
    
    Case "d", "M", "y"
    
      ' valido longitud
      
      If Len(obj.Text) <> 10 Or Not IsDate(obj.Text) Then
        obj.SetFocus
        obj.SelStart = 0
        obj.SelLength = 999
        obj.ToolTipText = Right(obj.Name, Len(obj.Name) - 3) & ": Deber ser un valor con formato, " & strFmat
        DataValidate = vbError
        Exit Function
      End If
    
      ' fecha validada
      
      DataValidate = True
    
    ' hora h hora, m minuto
    
    Case "h", "m"
    
      ' valido datos segun strValidos
      
      If Len(obj.Text) <> 5 Or Not IsDate(obj.Text) Then
      
        ' si encuentra algun valor invalido
        ' pone foco en control con error
        ' pinta el valor del control
        ' pone texto de error
        ' devuelve un error y escapa
      
        obj.SetFocus
        obj.SelStart = 0
        obj.SelLength = 999
        obj.ToolTipText = Right(obj.Name, Len(obj.Name) - 3) & ": Deber ser un valor con formato, " & strFmat
        DataValidate = vbError
        Exit Function
        
      End If
    
      ' hora validada
      
      DataValidate = True
    
    ' @ alfanumerico
    
    Case "@"
    
      ' determino longitud a validar
       
      intLenMaxima = Val(Mid(strFmat, 2, Len(strFmat)))  ' longitud validar
      
      ' valido longitud
      
      If (Len(obj.Text) > intLenMaxima) Or (Len(obj.Text) = 0 And blnObligatorio) Then
        obj.SetFocus
        obj.SelStart = 0
        obj.SelLength = intLenMaxima
        obj.ToolTipText = Right(obj.Name, Len(obj.Name) - 3) & ": Deber ser un valor con una longitud máxima de " & intLenMaxima & " caracteres."
        DataValidate = vbError
        Exit Function
      End If
      
      ' alfanumerico validado
    
      DataValidate = True
    
    End Select
    
  Case "ComboBox"
  
    ' si el dato no es obligatorio y esta vacio EXIT
  
    If Not blnObligatorio And obj.ListIndex = -1 Then
      DataValidate = True
      Exit Function
    End If
  
    If obj.ListIndex = -1 Then
      obj.SetFocus
      obj.ToolTipText = Right(obj.Name, Len(obj.Name) - 3) & ": Debe seleccionar un dato de la lista."
      DataValidate = vbError
      Exit Function
    End If
  
    ' dato validado
    
    DataValidate = True
  
  End Select

End Function

' convierte un string con formato dd/mm/yyyy a ISO yyyymmdd

Public Function dateToIso(ByVal str As String) As String

dateToIso = ""
If IsDate(str) Then
  dateToIso = Mid(str, 7, 4) & Mid(str, 4, 2) & Mid(str, 1, 2)
End If

End Function


' devuelve el primer dia del mes de acuerdo a la fecha recibida

Public Function dateToFirstDay(ByVal str As String) As String

  dateToFirstDay = "01" & Mid(str, 3, 8)

End Function

' devuelve el ultimo dia del mes de acuerdo a la fecha recibida

Public Function dateToLastDay(ByVal str As String) As String

    If IsDate("31" & Mid(str, 3, 8)) Then
      dateToLastDay = "31" & Mid(str, 3, 8)
      ElseIf IsDate("30" & Mid(str, 3, 8)) Then
        dateToLastDay = "30" & Mid(str, 3, 8)
        ElseIf IsDate("29" & Mid(str, 3, 8)) Then
          dateToLastDay = "29" & Mid(str, 3, 8)
            ElseIf IsDate("28" & Mid(str, 3, 8)) Then
              dateToLastDay = "28" & Mid(str, 3, 8)
            End If

End Function

'
' Devuelve Parametros
'
Public Function getParam(ByVal strParam As String) As Single
  Dim rsParam As ADODB.Recordset
  
  getParam = -999
  
  strSQL = "SELECT * FROM " & conDBParam & " WHERE Referencia = '" & Format(strParam, "<") & "'"
  Set rsParam = adoOpenRecordset(strSQL)

  If Not rsParam.EOF Then
    getParam = rsParam!Valor
  End If
  rsParam.Close

End Function

'
' Exportar informacion en un objeto x
'
Public Function ExportData(ByVal obj As Object) As Integer
  Dim intCuenta, intRes, intInd1, intInd2, intRow, intCol, intEncontro As Integer
  Dim lit As ListItem
  Dim myBook As Workbook
  Dim xlsApp As excel.Application
  
 
  Select Case TypeName(obj)
  
  Case "ListView"
    
    ' cargo formulario
    Load frmExportData
      
    ' recorro array y lleno lista con columnas
    For intCuenta = 0 To UBound(strStruc, 2)
      ' no agrego las columnas en donde encuentra un id
      If InStr(1, Format(strStruc(1, intCuenta), "<"), "id") = 0 Then
        frmExportData.lstSelect.AddItem strStruc(1, intCuenta)
      End If
    Next
      
    ' muestro formulario para filtro
    frmExportData.Show vbModal
  
    ' si acepto exportacion
    If blnAceptar Then
      
      ' cambio puntero mouse
      Screen.MousePointer = vbHourglass
   
      ' proceso de exportacion
      
      ' crea un libro de excel y referencio
      Set xlsApp = CreateObject("Excel.Application")
      Set myBook = xlsApp.Application.Workbooks.Add
      
      ' contador filas y columnas en excel
      intRow = 1
      intCol = 0
      
      ' recorro columnas del ListView
      For intInd1 = 1 To obj.ColumnHeaders.Count
      
        ' recorro columnas seleccionadas en lstSelecting
        For intInd2 = 0 To frmExportData.lstSelecting.ListCount - 1
          
          ' si encuentro columna seleccionada la agrego en excel
          If obj.ColumnHeaders(intInd1) = frmExportData.lstSelecting.List(intInd2) Then
            
            ' guardo titulo de columna
            intCol = intCol + 1
            myBook.Worksheets(1).Cells(intRow, intCol).Value = obj.ColumnHeaders(intInd1)
              
            ' mensaje procesando columna nnnn
            MDIMenu.staGeneral.Panels(3).Text = "Info: " & "exporting column " & Trim(intCol) & " of " & Trim(frmExportData.lstSelecting.ListCount)
            
            ' guardo informacion de cada columna
            For a = 1 To obj.ListItems.Count
              ' si es columna 1, se trata diferente al resto
              If intInd1 = 1 Then
                myBook.Worksheets(1).Cells(a + 1, intCol).Value = obj.ListItems(a)
              Else
                myBook.Worksheets(1).Cells(a + 1, intCol).Value = obj.ListItems(a).SubItems(intInd1 - 1)
              End If
            Next
            
          End If
        
        Next
      
      Next
      
      ' guardo archivo exportado
      myBook.SaveAs frmExportData.comNombre.FileName
      intRes = MsgBox("Exportation Ok", vbInformation + vbOKOnly, "Informacion")
      myBook.Close
   
      ' vuelvo puntero mouse
      Screen.MousePointer = vbDefault
    
    End If
    
    ' descargo form
    Unload frmExportData
    
    ' borro mensaje de progres0
    MDIMenu.staGeneral.Panels(3).Text = "Info:"

  End Select

End Function

