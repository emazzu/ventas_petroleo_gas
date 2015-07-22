VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmContratos 
   BorderStyle     =   0  'None
   Caption         =   "Contratos"
   ClientHeight    =   7140
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   13170
   ControlBox      =   0   'False
   ForeColor       =   &H00008000&
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   7140
   ScaleWidth      =   13170
   ShowInTaskbar   =   0   'False
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   45
      Top             =   6255
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   8
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmContratos.frx":0000
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmContratos.frx":08DA
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmContratos.frx":11B4
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmContratos.frx":1A8E
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmContratos.frx":3420
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmContratos.frx":3CFA
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmContratos.frx":45D4
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmContratos.frx":5F66
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ListView lvwContratos 
      Height          =   5325
      Left            =   0
      TabIndex        =   1
      Top             =   855
      Width           =   9240
      _ExtentX        =   16298
      _ExtentY        =   9393
      View            =   3
      Sorted          =   -1  'True
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      _Version        =   393217
      ForeColor       =   8421504
      BackColor       =   16777215
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   0
   End
   Begin MSComctlLib.Toolbar tlbOperaciones 
      Height          =   570
      Left            =   0
      TabIndex        =   2
      Top             =   0
      Width           =   11340
      _ExtentX        =   20003
      _ExtentY        =   1005
      ButtonWidth     =   2355
      ButtonHeight    =   1005
      Style           =   1
      TextAlignment   =   1
      ImageList       =   "ImageList1"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   8
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "&Agrega"
            Key             =   "agregar"
            ImageIndex      =   1
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "&Modifica"
            Key             =   "modificar"
            ImageIndex      =   2
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "&Elimina"
            Key             =   "eliminar"
            ImageIndex      =   3
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "&Busca"
            Key             =   "buscar"
            ImageIndex      =   5
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "&Filtra"
            Key             =   "filtrar"
            ImageIndex      =   4
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "&Ordena"
            Key             =   "ordenar"
            ImageIndex      =   6
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "&Exporta"
            Key             =   "exportar"
            ImageIndex      =   7
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "&Ajusta"
            Key             =   "ajustar"
            ImageIndex      =   8
         EndProperty
      EndProperty
   End
   Begin VB.Label lblTitulo 
      Alignment       =   2  'Center
      BackColor       =   &H00808000&
      Caption         =   "Contratos"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   0
      TabIndex        =   0
      Top             =   585
      Width           =   9225
   End
   Begin VB.Menu mnuContratos 
      Caption         =   "Contratos"
      Visible         =   0   'False
      Begin VB.Menu mnuAgregar 
         Caption         =   "Agregar"
      End
      Begin VB.Menu mnuModificar 
         Caption         =   "Modificar"
      End
      Begin VB.Menu mnuEliminar 
         Caption         =   "Eliminar"
      End
      Begin VB.Menu mnuBuscar 
         Caption         =   "Buscar"
      End
      Begin VB.Menu mnuFiltrar 
         Caption         =   "Filtrar"
      End
   End
End
Attribute VB_Name = "frmContratos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim strOperacion As String

Private Sub Form_Load()
  Dim strRangoIso As isoLastPeriod
  
  'muestro solo ultimo periodo ingresado
  strRangoIso = adoLastPeriod("ViewContratos", "fechadesde")
  
  strSQL = "SELECT * FROM ViewContratos where fechaDesde between " & strRangoIso.strDesde & " and " & strRangoIso.strHasta
  
  intRes = ListViewAppearanceChange(lvwContratos)
  intRes = ListViewRefresh(lvwContratos, strSQL, strStruc)
  intRes = lvwHideColumn(lvwContratos, "contra")

End Sub


Private Sub tlbOperaciones_ButtonClick(ByVal Button As MSComctlLib.Button)
  Dim pp As ListItem
  
  blnRefresh = False
  
  Select Case Button.Key
    
  Case Is = "agregar"
  
      ' cargo formulario
    
      Load frmContratosInfo
    
      ' muestro formulario
            
      frmContratosInfo.Show vbModal
      
      ' si hizo click en Aceptar, en el formulario frmContratosInfo
      ' se pone en true la variable global blnAceptar y
      ' armo string y ejecuto funcion de INSERT
      
      If blnAceptar Then
      
        With frmContratosInfo
        strSQL = "EXEC spContratosInsert " & _
        .cboEmpresa.ItemData(.cboEmpresa.ListIndex) & "," & _
        .cboCliente.ItemData(.cboCliente.ListIndex) & "," & _
        "'" & dateToIso(.txtFechaDesde) & "'," & _
        "'" & dateToIso(.txtFechaHasta) & "'," & _
        Val(.txtm315) & "," & Val(.txtm31556) & "," & Val(.txtBarrels60) & "," & _
        "'" & .TxtDescuento & "'," & _
        Val(.txtAPIMinimo) & "," & Val(.txtAPIMaximo) & "," & Val(.txtAzuMinimo) & "," & Val(.txtAzuMaximo) & "," & _
        "'" & .cboRangos.List(.cboRangos.ListIndex) & "'," & _
        "'" & .cboAjustePrecio & "','" & .cboAjusteApi & "'," & _
        "'" & .cboAjusteVarios & "'," & _
        Val(.cboPrecioTipo.ItemData(.cboPrecioTipo.ListIndex)) & "," & _
        "'" & .cboTipoCalculo.List(.cboTipoCalculo.ListIndex) & "'," & _
        Val(.txtDiasPrevios) & "," & _
        Val(.txtDiasPosteriores) & "," & _
        "'" & .cboIncluyeEntregaCli.List(.cboIncluyeEntregaCli.ListIndex) & "'," & _
        "'" & .cboMesPromedio.List(.cboMesPromedio.ListIndex) & "'," & _
        "'" & .txtObservaciones & "'," & _
        Val(.txtRedondeo1) & "," & Val(.txtRedondeo2) & "," & Val(.txtRedondeo3) & "," & Val(.txtRedondeo4) & "," & _
        "'" & .txtDesFormula & "'," & _
        "'" & dateToIso(.txtDesDesde) & "'," & _
        "'" & dateToIso(.txtDesHasta) & "'," & _
        Val(.txtDesPrevios) & "," & _
        Val(.txtDesPoste) & "," & _
        Val(.txtDesMeses)
        
        End With
        
        'exec
        a = adoExecSQL(strSQL)
        
        'chequeo errores
        If Not lngAdoErrNum = -1 Then
          adoError
          Exit Sub
        End If
        
        blnRefresh = True
        
      End If
      
      ' descargo formulario
      
      Unload frmContratosInfo
  
  Case Is = "modificar"
  
    If lvwContratos.SelectedItem > 0 Then
    
      ' cargo formulario
    
      Load frmContratosInfo
    
      ' paso los valores del list view al formulario
      With frmContratosInfo
      .cboEmpresa.ListIndex = ComboBoxFindItem(.cboEmpresa, lvwGetValue(lvwContratos, "empresa"))
      .cboCliente.ListIndex = ComboBoxFindItem(.cboCliente, lvwGetValue(lvwContratos, "cliente"))
      .txtFechaDesde = lvwGetValue(lvwContratos, "fechadesde")
      .txtFechaHasta = lvwGetValue(lvwContratos, "fechahasta")
      .txtm315 = lvwGetValue(lvwContratos, "m315")
      .txtm31556 = lvwGetValue(lvwContratos, "m31556")
      .txtBarrels60 = lvwGetValue(lvwContratos, "barrels60")
      .cboPrecioTipo.ListIndex = ComboBoxFindItem(.cboPrecioTipo, lvwGetValue(lvwContratos, "preciotipo"))
      .txtAPIMinimo = lvwGetValue(lvwContratos, "apimin")
      .txtAPIMaximo = lvwGetValue(lvwContratos, "apimax")
      .txtAzuMinimo = lvwGetValue(lvwContratos, "azumin")
      .txtAzuMaximo = lvwGetValue(lvwContratos, "azumax")
      .cboRangos.ListIndex = ComboBoxFindItem(.cboRangos, lvwGetValue(lvwContratos, "tabla Rangos"))
      .cboAjustePrecio.ListIndex = ComboBoxFindItem(.cboAjustePrecio, lvwGetValue(lvwContratos, "formulaajuprecio"))
      .cboAjusteApi.ListIndex = ComboBoxFindItem(.cboAjusteApi, lvwGetValue(lvwContratos, "formulaajuapi"))
      .cboAjusteVarios.ListIndex = ComboBoxFindItem(.cboAjusteVarios, lvwGetValue(lvwContratos, "formulaajuvarios"))
      .TxtDescuento = lvwGetValue(lvwContratos, "descuento")
      .cboTipoCalculo.ListIndex = ComboBoxFindItem(.cboTipoCalculo, lvwGetValue(lvwContratos, "tipocalculo"))
      .txtDiasPrevios = lvwGetValue(lvwContratos, "diasprev")
      .txtDiasPosteriores = lvwGetValue(lvwContratos, "diaspost")
      .cboIncluyeEntregaCli.ListIndex = ComboBoxFindItem(.cboIncluyeEntregaCli, lvwGetValue(lvwContratos, "entregaCli"))
      .cboMesPromedio.ListIndex = ComboBoxFindItem(.cboMesPromedio, lvwGetValue(lvwContratos, "mespromedio"))
      .txtObservaciones = lvwGetValue(lvwContratos, "observaciones")
      .txtRedondeo1 = lvwGetValue(lvwContratos, "redon1")
      .txtRedondeo2 = lvwGetValue(lvwContratos, "redon2")
      .txtRedondeo3 = lvwGetValue(lvwContratos, "redon3")
      .txtRedondeo4 = lvwGetValue(lvwContratos, "redon4")
      .txtDesFormula = lvwGetValue(lvwContratos, "Desc. fórmula")
      .txtDesDesde = lvwGetValue(lvwContratos, "Desc. fecha desde")
      .txtDesHasta = lvwGetValue(lvwContratos, "Desc. fecha hasta")
      .txtDesPrevios = lvwGetValue(lvwContratos, "Desc. previos")
      .txtDesPoste = lvwGetValue(lvwContratos, "Desc. poste")
      .txtDesMeses = lvwGetValue(lvwContratos, "Desc. meses")
      
      End With
      
      frmContratosInfo.Show vbModal
      
      
      If blnAceptar Then
      
        ' si hizo click en Aceptar, genero string y ejecuto
        ' funcion de UPDATE, el primer argumento enviado es
        ' el campo clave por el cual aplica el WHERE
        
        With frmContratosInfo
        strSQL = "EXEC spContratosUpdate " & _
        Me.lvwContratos.SelectedItem & "," & _
        .cboEmpresa.ItemData(.cboEmpresa.ListIndex) & "," & _
        .cboCliente.ItemData(.cboCliente.ListIndex) & "," & _
        "'" & dateToIso(.txtFechaDesde) & "'," & _
        "'" & dateToIso(.txtFechaHasta) & "'," & _
        Val(.txtm315) & "," & Val(.txtm31556) & "," & Val(.txtBarrels60) & "," & "'" & .TxtDescuento & "'," & _
        Val(.txtAPIMinimo) & "," & Val(.txtAPIMaximo) & "," & Val(.txtAzuMinimo) & "," & Val(.txtAzuMaximo) & "," & _
        "'" & .cboRangos.List(.cboRangos.ListIndex) & "'," & _
        "'" & .cboAjustePrecio & "','" & .cboAjusteApi & "','" & .cboAjusteVarios & "'," & _
        Val(.cboPrecioTipo.ItemData(.cboPrecioTipo.ListIndex)) & "," & _
        "'" & .cboTipoCalculo.List(.cboTipoCalculo.ListIndex) & "'," & _
        Val(.txtDiasPrevios) & "," & _
        Val(.txtDiasPosteriores) & "," & _
        "'" & .cboIncluyeEntregaCli.List(.cboIncluyeEntregaCli.ListIndex) & "'," & _
        "'" & .cboMesPromedio.List(.cboMesPromedio.ListIndex) & "'," & _
        "'" & .txtObservaciones & "'," & _
        Val(.txtRedondeo1) & "," & Val(.txtRedondeo2) & "," & Val(.txtRedondeo3) & "," & Val(.txtRedondeo4) & "," & _
        "'" & .txtDesFormula & "'," & _
        "'" & dateToIso(.txtDesDesde) & "'," & _
        "'" & dateToIso(.txtDesHasta) & "'," & _
        Val(.txtDesPrevios) & "," & _
        Val(.txtDesPoste) & "," & _
        Val(.txtDesMeses)
        
        End With
                
        'exec
        a = adoExecSQL(strSQL)
        
        'chequeo errores
        If Not lngAdoErrNum = -1 Then
          adoError
          Exit Sub
        End If
        
        blnRefresh = True
        
      End If
      
      ' descargo formulario
      
      Unload frmContratosInfo
      
    Else
      a = MsgBox("No hay ningun item seleccionado.", vbApplicationModal + vbOKOnly + vbInformation, "Informacion")
    End If
  
    Case Is = "eliminar"
  
      If lvwContratos Is Nothing Then
        a = MsgBox("No hay ningun item seleccionado.", vbApplicationModal + vbOKOnly + vbInformation, "Informacion")
        Exit Sub
      End If
      
      intRes = MsgBox("Esta seguro que desea eliminar el elemento seleccionado.", vbQuestion + vbYesNo, "Confirmacón")
      
      If intRes = vbYes Then
        
        'exec
        strSQL = "EXEC spContratosDelete " & Me.lvwContratos.SelectedItem
        a = adoExecSQL(strSQL)
        
        'chequeo errores
        If Not lngAdoErrNum = -1 Then
          adoError
          Exit Sub
        End If
        
        blnRefresh = True
        
      End If
 
    Case Is = "buscar"
      intRes = FindData(lvwContratos)
  
    Case Is = "filtrar"
      strWhere = FilterData(lvwContratos)
      If blnAceptar Then blnRefresh = True
  
    Case Is = "ordenar"
      intRes = lvwSortColumn(lvwContratos)
  
    Case Is = "exportar"
      intRes = ExportData(lvwContratos)
  
    Case Is = "ajustar"
      ' ajusta y envia a INI
      intRes = lvwAdjustColumn(lvwContratos, True)
      intRes = lvwWidthToKeyIni(lvwContratos, strTableNameActual)
  
  End Select

  ' refresh
  If blnRefresh Then
  
    If strWhere = "" Then
      
      'muestro solo ultimo periodo ingresado
      Dim strRangoIso As isoLastPeriod
      strRangoIso = adoLastPeriod("ViewContratos", "fechadesde")
      strWhere = "fechaDesde between " & strRangoIso.strDesde & " and " & strRangoIso.strHasta
    
    End If
    
    strSQL = "SELECT * FROM ViewContratos" & _
             IIf(Not strWhere = "", " WHERE " & strWhere, "") & " "
    intRes = ListViewAppearanceChange(lvwContratos)
    intRes = ListViewRefresh(lvwContratos, strSQL)
    intRes = lvwHideColumn(lvwContratos, "contra")
  
  End If

End Sub




