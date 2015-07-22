VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmEntregasCli 
   BorderStyle     =   0  'None
   Caption         =   "Entregas Clientes"
   ClientHeight    =   8130
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   13290
   ForeColor       =   &H80000010&
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   8130
   ScaleWidth      =   13290
   ShowInTaskbar   =   0   'False
   Begin MSComctlLib.ListView lvwDatos 
      Height          =   5325
      Left            =   0
      TabIndex        =   0
      Top             =   900
      Width           =   9240
      _ExtentX        =   16298
      _ExtentY        =   9393
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   0
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   45
      Top             =   6210
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
            Picture         =   "frmEntregasCli.frx":0000
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmEntregasCli.frx":08DA
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmEntregasCli.frx":11B4
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmEntregasCli.frx":1A8E
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmEntregasCli.frx":3420
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmEntregasCli.frx":3CFA
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmEntregasCli.frx":45D4
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmEntregasCli.frx":5F66
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar tlbOperaciones 
      Height          =   570
      Left            =   0
      TabIndex        =   2
      Top             =   0
      Width           =   12645
      _ExtentX        =   22304
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
      Appearance      =   0  'Flat
      BackColor       =   &H00808080&
      Caption         =   "Entregas Clientes"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   315
      Left            =   0
      TabIndex        =   1
      Top             =   585
      Width           =   9195
   End
End
Attribute VB_Name = "frmEntregasCli"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Form_Load()
  Dim strRangoIso As isoLastPeriod
  
  'muestro solo ultimo periodo ingresado
  strRangoIso = adoLastPeriod("ViewEntregasCli", "fecha")
  
  strSQL = "SELECT * FROM ViewEntregasCli where fecha between " & strRangoIso.strDesde & " and " & strRangoIso.strHasta & " " & _
           "ORDER BY empresa, fecha DESC"

  intRes = ListViewAppearanceChange(lvwDatos)
  intRes = ListViewRefresh(lvwDatos, strSQL, strStruc)
  intRes = lvwHideColumn(lvwDatos, "IDentregaCli")
  intRes = lvwHideColumn(lvwDatos, "status")

End Sub


Private Sub tlbOperaciones_ButtonClick(ByVal Button As MSComctlLib.Button)
  Dim strSQL As String
  Dim intRes As Integer
  
  blnRefresh = False
  
  Select Case Button.Key
    
  Case Is = "agregar"
  
      ' cargo formulario
    
      Load frmEntregasCliInfo
    
      ' muestro formulario
            
      frmEntregasCliInfo.Show vbModal
      
      ' si hizo click en Aceptar, en el formulario frmContratosInfo
      ' se pone en true la variable global blnAceptar y
      ' armo string y ejecuto funcion de INSERT
      
      If blnAceptar Then
      
        With frmEntregasCliInfo
        strSQL = "EXEC spEntregasCliInsert " & _
        "'" & .txtLegajo & "'," & _
        Val(.cboEntregaCliTipo.ItemData(.cboEntregaCliTipo.ListIndex)) & "," & _
        "'" & dateToIso(.txtFecha) & "'," & _
        Val(.cboEmpresa.ItemData(.cboEmpresa.ListIndex)) & "," & _
        Val(.cboCliente.ItemData(.cboCliente.ListIndex)) & "," & _
        Val(.cboInspeccion.ItemData(.cboInspeccion.ListIndex)) & "," & _
        "'" & dateToIso(.txtCertificado) & "'," & _
        Val(.cboBarco.ItemData(.cboBarco.ListIndex)) & "," & _
        Val(.cboTerminal.ItemData(.cboTerminal.ListIndex)) & "," & _
        "'" & dateToIso(.txtCerDesde) & "'," & _
        "'" & dateToIso(.txtCerHasta) & "'," & _
        "'" & .txtEntregaNro & "'," & _
        Val(.txtAzufre) & "," & Val(.txtOtrosAjustes) & "," & _
        "'" & .cboProvicionado.List(.cboProvicionado.ListIndex) & "'," & _
        Val(.txtNGov) & "," & Val(.txtNTcv) & "," & Val(.txtNGsv) & "," & _
        Val(.txtNDensity) & "," & Val(.txtNMetricTong) & "," & _
        Val(.txtNCubMeters1556) & "," & Val(.txtNBarrels60) & "," & _
        Val(.txtNGallons60) & "," & Val(.txtNLongTons) & "," & _
        Val(.txtNAPIGravity) & "," & Val(.txtNBswCBM) & "," & _
        Val(.txtGGov) & "," & Val(.txtGTcv) & "," & Val(.txtGGsv) & "," & _
        Val(.txtGDensity) & "," & Val(.txtGMetricTong) & "," & _
        Val(.txtGCubMeters1556) & "," & Val(.txtGBarrels60) & "," & _
        Val(.txtGGallons60) & "," & Val(.txtGLongTons) & "," & _
        Val(.txtGAPIGravity) & "," & Val(.txtGBswCBM)
        End With
        
        a = adoExecSQL(strSQL)
        blnRefresh = True
      
      End If
      
      ' descargo formulario
      
      Unload frmEntregasCliInfo
  
  Case Is = "modificar"
  
    If lvwDatos Is Nothing Then
      intRes = MsgBox("No hay ningún item seleccionado.", vbApplicationModal + vbOKOnly + vbInformation, "Informacion")
      Exit Sub
    End If
    
    If lvwGetValue(lvwDatos, "status") <> 0 Then
      intRes = MsgBox("No es posible modificarlo, primero debera borrar el stock por terminales.", vbApplicationModal + vbOKOnly + vbInformation, "Informacion")
      Exit Sub
    End If
    
    ' cargo formulario
    
    Load frmEntregasCliInfo
    
    ' paso los valores del list view al formulario
      
    With frmEntregasCliInfo
    .txtLegajo = lvwGetValue(lvwDatos, "legajo")
    .cboEntregaCliTipo.ListIndex = ComboBoxFindItem(.cboEntregaCliTipo, lvwGetValue(lvwDatos, "tipoentrega"))
    .txtFecha = lvwGetValue(lvwDatos, "fecha")
    .cboEmpresa.ListIndex = ComboBoxFindItem(.cboEmpresa, lvwGetValue(lvwDatos, "empresa"))
    .cboCliente.ListIndex = ComboBoxFindItem(.cboCliente, lvwGetValue(lvwDatos, "cliente"))
    .cboInspeccion.ListIndex = ComboBoxFindItem(.cboInspeccion, lvwGetValue(lvwDatos, "inspeccion"))
    .txtCertificado = lvwGetValue(lvwDatos, "certificado")
    .cboBarco.ListIndex = ComboBoxFindItem(.cboBarco, lvwGetValue(lvwDatos, "barco"))
    .cboTerminal.ListIndex = ComboBoxFindItem(.cboTerminal, lvwGetValue(lvwDatos, "terminal"))
    .txtCerDesde = lvwGetValue(lvwDatos, "CertifDesde")
    .txtCerHasta = lvwGetValue(lvwDatos, "CertifHasta")
    .txtEntregaNro = lvwGetValue(lvwDatos, "entreganro")
    .txtAzufre = lvwGetValue(lvwDatos, "azufre")
    .txtOtrosAjustes = lvwGetValue(lvwDatos, "otrosaju")
    .cboProvicionado.ListIndex = ComboBoxFindItem(.cboProvicionado, lvwGetValue(lvwDatos, "provisionado"))
    .txtGGov = lvwGetValue(lvwDatos, "ggov")
    .txtGTcv = lvwGetValue(lvwDatos, "gtcv")
    .txtGGsv = lvwGetValue(lvwDatos, "ggsv")
    .txtGDensity = lvwGetValue(lvwDatos, "gdensity")
    .txtGMetricTong = lvwGetValue(lvwDatos, "gmetric")
    .txtGCubMeters1556 = lvwGetValue(lvwDatos, "gcubmeters")
    .txtGBarrels60 = lvwGetValue(lvwDatos, "gbarrels")
    .txtGGallons60 = lvwGetValue(lvwDatos, "ggallons")
    .txtGLongTons = lvwGetValue(lvwDatos, "glongtons")
    .txtGAPIGravity = lvwGetValue(lvwDatos, "gapigravity")
    .txtGBswCBM = lvwGetValue(lvwDatos, "gbswcbm")
    .txtNGov = lvwGetValue(lvwDatos, "ngov")
    .txtNTcv = lvwGetValue(lvwDatos, "ntcv")
    .txtNGsv = lvwGetValue(lvwDatos, "ngsv")
    .txtNDensity = lvwGetValue(lvwDatos, "ndensity")
    .txtNMetricTong = lvwGetValue(lvwDatos, "nmetric")
    .txtNCubMeters1556 = lvwGetValue(lvwDatos, "ncubmeters")
    .txtNBarrels60 = lvwGetValue(lvwDatos, "nbarrels")
    .txtNGallons60 = lvwGetValue(lvwDatos, "ngallons")
    .txtNLongTons = lvwGetValue(lvwDatos, "nlongtons")
    .txtNAPIGravity = lvwGetValue(lvwDatos, "napigravity")
    .txtNBswCBM = lvwGetValue(lvwDatos, "nbswcbm")
    .Show vbModal
    End With
      
    If blnAceptar Then
      
      ' si hizo click en Aceptar, genero string y ejecuto
      ' funcion de UPDATE, el primer argumento enviado es
      ' el campo clave por el cual aplica el WHERE
        
      With frmEntregasCliInfo
      strSQL = "EXEC spEntregasCliUpdate " & _
      lvwGetValue(lvwDatos, "IDentregaCli") & "," & _
      "'" & .txtLegajo & "'," & _
      Val(.cboEntregaCliTipo.ItemData(.cboEntregaCliTipo.ListIndex)) & "," & _
      "'" & dateToIso(.txtFecha) & "'," & _
      Val(.cboEmpresa.ItemData(.cboEmpresa.ListIndex)) & "," & _
      Val(.cboCliente.ItemData(.cboCliente.ListIndex)) & "," & _
      Val(.cboInspeccion.ItemData(.cboInspeccion.ListIndex)) & "," & _
      "'" & dateToIso(.txtCertificado) & "'," & _
      Val(.cboBarco.ItemData(.cboBarco.ListIndex)) & "," & _
      Val(.cboTerminal.ItemData(.cboTerminal.ListIndex)) & "," & _
      "'" & dateToIso(.txtCerDesde) & "','" & dateToIso(.txtCerHasta) & "'," & _
      "'" & .txtEntregaNro & "'," & _
      "'" & .cboProvicionado.List(.cboProvicionado.ListIndex) & "'," & _
      Val(.txtAzufre) & "," & Val(.txtOtrosAjustes) & "," & _
      Val(.txtNGov) & "," & Val(.txtNTcv) & "," & Val(.txtNGsv) & "," & _
      Val(.txtNDensity) & "," & Val(.txtNMetricTong) & "," & _
      Val(.txtNCubMeters1556) & "," & Val(.txtNBarrels60) & "," & _
      Val(.txtNGallons60) & "," & Val(.txtNLongTons) & "," & _
      Val(.txtNAPIGravity) & "," & Val(.txtNBswCBM) & "," & _
      Val(.txtGGov) & "," & Val(.txtGTcv) & "," & Val(.txtGGsv) & "," & _
      Val(.txtGDensity) & "," & Val(.txtGMetricTong) & "," & _
      Val(.txtGCubMeters1556) & "," & Val(.txtGBarrels60) & "," & _
      Val(.txtGGallons60) & "," & Val(.txtGLongTons) & "," & _
      Val(.txtGAPIGravity) & "," & Val(.txtGBswCBM)
      End With
        
      a = adoExecSQL(strSQL)
      blnRefresh = True
 
    End If
      
    ' descargo formulario
      
    Unload frmEntregasCliInfo
      
    Case Is = "eliminar"
  
      If lvwDatos Is Nothing Then
        intRes = MsgBox("No hay ningún item seleccionado.", vbApplicationModal + vbOKOnly + vbInformation, "Informacion")
        Exit Sub
      End If
     
      If lvwGetValue(lvwDatos, "status") <> 0 Then
        intRes = MsgBox("No es posible eliminarlo, primero debera borrar el stock por terminales.", vbApplicationModal + vbOKOnly + vbInformation, "Informacion")
        Exit Sub
      End If
      
      intRes = MsgBox("Esta seguro que desea eliminar el elemento seleccionado.", vbQuestion + vbYesNo, "Confirmacón")
      
      If intRes = vbYes Then
        
        strSQL = "EXEC spEntregasCliDelete " & lvwGetValue(lvwDatos, "IDentregaCli")
    
        a = adoExecSQL(strSQL)
        blnRefresh = True
      
      End If
  
    Case Is = "buscar"
      intRes = FindData(lvwDatos)
  
    Case Is = "filtrar"
      strWhere = FilterData(lvwDatos)
      If blnAceptar Then blnRefresh = True
    
    Case Is = "ordenar"
      intRes = lvwSortColumn(lvwDatos)
  
    Case Is = "exportar"
      intRes = ExportData(lvwDatos)
  
    Case Is = "ajustar"
      ' ajusta y envia a INI
      intRes = lvwAdjustColumn(lvwDatos, True)
      intRes = lvwWidthToKeyIni(lvwDatos, strTableNameActual)
  
  End Select

  If blnRefresh Then
    
    If strWhere = "" Then
      
      'muestro solo ultimo periodo ingresado
      Dim strRangoIso As isoLastPeriod
      strRangoIso = adoLastPeriod("ViewEntregasCli", "fecha")
      strWhere = "fecha between " & strRangoIso.strDesde & " and " & strRangoIso.strHasta
    
    End If
    
    strSQL = "SELECT * FROM ViewEntregasCli" & _
             IIf(Not strWhere = "", " WHERE " & strWhere, "") & " " & _
             "ORDER BY empresa, fecha DESC"
    intRes = ListViewRefresh(lvwDatos, strSQL)
    intRes = lvwHideColumn(lvwDatos, "IDEntregaCli")
    intRes = lvwHideColumn(lvwDatos, "status")

  End If

End Sub

