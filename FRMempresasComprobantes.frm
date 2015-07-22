VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmEmpresasComprobantes 
   BorderStyle     =   0  'None
   Caption         =   "Comprobantes"
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
            Picture         =   "FRMempresasComprobantes.frx":0000
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FRMempresasComprobantes.frx":08DA
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FRMempresasComprobantes.frx":11B4
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FRMempresasComprobantes.frx":1A8E
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FRMempresasComprobantes.frx":3420
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FRMempresasComprobantes.frx":3CFA
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FRMempresasComprobantes.frx":45D4
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FRMempresasComprobantes.frx":5F66
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ListView lvwComprobantes 
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
      Caption         =   "Comprobantes"
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
Attribute VB_Name = "frmEmpresasComprobantes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim strOperacion As String

Private Sub Form_Load()
  Dim strRangoIso As isoLastPeriod
  
  'muestro solo ultimo periodo ingresado
  strRangoIso = adoLastPeriod("EmpresasComprobantes_vw", "id")
  
  strSQL = "SELECT * FROM EmpresasComprobantes_vw"
  
  intRes = ListViewAppearanceChange(lvwComprobantes)
  intRes = ListViewRefresh(lvwComprobantes, strSQL, strStruc)
  intRes = lvwHideColumn(lvwComprobantes, "id")
  
End Sub


Private Sub tlbOperaciones_ButtonClick(ByVal Button As MSComctlLib.Button)
  Dim pp As ListItem
  
  blnRefresh = False
  
  Select Case Button.Key
    
  Case Is = "agregar"
  
      ' cargo formulario
    
      Load FRMEmpresasComprobantesInfo
    
      ' muestro formulario
            
      FRMEmpresasComprobantesInfo.Show vbModal
      
      ' si hizo click en Aceptar, en el formulario frmContratosInfo
      ' se pone en true la variable global blnAceptar y
      ' armo string y ejecuto funcion de INSERT
      
      If blnAceptar Then
        
        With FRMEmpresasComprobantesInfo
        
        strSQL = "EXEC EmpresasComprobantes_INS_sp " & _
        .cboEmpresa.ItemData(.cboEmpresa.ListIndex) & "," & _
        "'" & dateToIso(.txtFecha) & "'," & _
        "'" & .cboTipo.List(.cboTipo.ListIndex) & "'," & _
        Val(.txtSucursal) & "," & _
        Val(.txtNumeroDesde) & "," & _
        Val(.txtNumeroHasta) & "," & _
        "'" & .txtCAI & "'," & _
        "'" & dateToIso(.txtVencimiento) & "'"
        
        End With
  
        a = adoExecSQL(strSQL)
        blnRefresh = True
        
      End If
      
      ' descargo formulario
      
      Unload FRMEmpresasComprobantesInfo
  
  Case Is = "modificar"
  
    If lvwComprobantes.SelectedItem > 0 Then
    
      ' cargo formulario
    
      Load FRMEmpresasComprobantesInfo
    
      ' paso los valores del list view al formulario
      With FRMEmpresasComprobantesInfo
      
      .cboEmpresa.ListIndex = ComboBoxFindItem(.cboEmpresa, lvwGetValue(lvwComprobantes, "empresa"))
      .txtFecha = lvwGetValue(lvwComprobantes, "fecha")
      .cboTipo.ListIndex = ComboBoxFindItem(.cboTipo, lvwGetValue(lvwComprobantes, "tipo"))
      .txtSucursal = lvwGetValue(lvwComprobantes, "sucursal")
      .txtNumeroDesde = lvwGetValue(lvwComprobantes, "numeración desde")
      .txtNumeroHasta = lvwGetValue(lvwComprobantes, "numeración hasta")
      .txtCAI = lvwGetValue(lvwComprobantes, "CAI")
      .txtVencimiento = lvwGetValue(lvwComprobantes, "Vencimiento")
      
      
      End With
      
      FRMEmpresasComprobantesInfo.Show vbModal
      
      
      If blnAceptar Then
      
        ' si hizo click en Aceptar, genero string y ejecuto
        ' funcion de UPDATE, el primer argumento enviado es
        ' el campo clave por el cual aplica el WHERE
        
        With FRMEmpresasComprobantesInfo
        
        strSQL = "EXEC empresasComprobantes_EDI_sp " & _
        Me.lvwComprobantes.SelectedItem & "," & _
        .cboEmpresa.ItemData(.cboEmpresa.ListIndex) & "," & _
        "'" & dateToIso(.txtFecha) & "'," & _
        "'" & .cboTipo.List(.cboTipo.ListIndex) & "'," & _
        Val(.txtSucursal) & "," & _
        Val(.txtNumeroDesde) & "," & _
        Val(.txtNumeroHasta) & "," & _
        "'" & .txtCAI & "'," & _
        "'" & dateToIso(.txtVencimiento) & "'"
        End With
        
        a = adoExecSQL(strSQL)
        blnRefresh = True
        
      End If
      
      ' descargo formulario
      
      Unload FRMEmpresasComprobantesInfo
      
    Else
      a = MsgBox("No hay ningun item seleccionado.", vbApplicationModal + vbOKOnly + vbInformation, "Informacion")
    End If
  
    Case Is = "eliminar"
  
      If lvwComprobantes Is Nothing Then
      
        a = MsgBox("No hay ningun item seleccionado.", vbApplicationModal + vbOKOnly + vbInformation, "Informacion")
        Exit Sub
      
      End If
      
      intRes = MsgBox("Esta seguro que desea eliminar el elemento seleccionado.", vbQuestion + vbYesNo, "Confirmacón")
      
      If intRes = vbYes Then
        
        strSQL = "EXEC empresasComprobantes_ELI_sp " & Me.lvwComprobantes.SelectedItem
        
        a = adoExecSQL(strSQL)
        blnRefresh = True
      
      End If
 
    Case Is = "buscar"
      intRes = FindData(lvwComprobantes)
  
    Case Is = "filtrar"
      strWhere = FilterData(lvwComprobantes)
      If blnAceptar Then blnRefresh = True
  
    Case Is = "ordenar"
      intRes = lvwSortColumn(lvwComprobantes)
  
    Case Is = "exportar"
      intRes = ExportData(lvwComprobantes)
  
    Case Is = "ajustar"
      ' ajusta y envia a INI
      intRes = lvwAdjustColumn(lvwComprobantes, True)
      intRes = lvwWidthToKeyIni(lvwComprobantes, strTableNameActual)
  
  End Select
  
  'refresh
  If blnRefresh Then
      
    strSQL = "SELECT * FROM EmpresasComprobantes_vw" & _
    IIf(Not strWhere = "", " WHERE " & strWhere, "")
    intRes = ListViewRefresh(lvwComprobantes, strSQL)
    intRes = lvwHideColumn(lvwComprobantes, "ID")

  End If

End Sub




