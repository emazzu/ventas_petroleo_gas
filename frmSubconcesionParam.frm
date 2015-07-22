VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form frmSubconcesionParam 
   BorderStyle     =   0  'None
   ClientHeight    =   8550
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   10035
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   8550
   ScaleWidth      =   10035
   ShowInTaskbar   =   0   'False
   Begin MSComctlLib.ListView lvwDatos 
      Height          =   5055
      Left            =   0
      TabIndex        =   0
      Top             =   855
      Width           =   9240
      _ExtentX        =   16298
      _ExtentY        =   8916
      View            =   3
      SortOrder       =   -1  'True
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      _Version        =   393217
      ForeColor       =   8421504
      BackColor       =   16777215
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   0
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   45
      Top             =   5985
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
            Picture         =   "frmSubconcesionParam.frx":0000
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSubconcesionParam.frx":08DA
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSubconcesionParam.frx":11B4
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSubconcesionParam.frx":1A8E
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSubconcesionParam.frx":3420
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSubconcesionParam.frx":3CFA
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSubconcesionParam.frx":45D4
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSubconcesionParam.frx":5F66
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar tlbOperaciones 
      Height          =   570
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   12735
      _ExtentX        =   22463
      _ExtentY        =   1005
      ButtonWidth     =   2461
      ButtonHeight    =   1005
      Style           =   1
      TextAlignment   =   1
      ImageList       =   "ImageList1"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   8
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "&Agregar"
            Key             =   "agregar"
            ImageIndex      =   1
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "&Modificar"
            Key             =   "modificar"
            ImageIndex      =   2
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "&Eliminar"
            Key             =   "eliminar"
            ImageIndex      =   3
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "&Buscar"
            Key             =   "buscar"
            ImageIndex      =   5
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "&Filtrar"
            Key             =   "filtrar"
            ImageIndex      =   4
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "&Ordenar"
            Key             =   "ordenar"
            ImageIndex      =   6
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "&Exportar"
            Key             =   "exportar"
            ImageIndex      =   7
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "&Ajustar"
            Key             =   "ajustar"
            ImageIndex      =   8
         EndProperty
      EndProperty
   End
   Begin VB.Label lblTitulo 
      Alignment       =   2  'Center
      BackColor       =   &H00808000&
      Caption         =   "Parametros PorSubConcesion"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   0
      TabIndex        =   2
      Top             =   585
      Width           =   9225
   End
End
Attribute VB_Name = "frmSubconcesionParam"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()

  strSQL = "SELECT * FROM SubConcesionesParam_View"
  intRes = ListViewAppearanceChange(lvwDatos)
  intRes = ListViewRefresh(lvwDatos, strSQL, strStruc)
  intRes = lvwHideColumn(lvwDatos, "IDSubconcesion")
  intRes = lvwHideColumn(lvwDatos, "empresaID")
  intRes = lvwHideColumn(lvwDatos, "concesionID")
  intRes = lvwHideColumn(lvwDatos, "pappath")
  intRes = lvwHideColumn(lvwDatos, "papcontrol")

End Sub

Private Sub tlbOperaciones_ButtonClick(ByVal Button As MSComctlLib.Button)

  blnRefresh = False
    
  Select Case Button.Key
  
  Case Is = "modificar"
  
    If (lvwDatos Is Nothing) Then
      intRes = MsgBox("No hay ningun item seleccionado.", vbApplicationModal + vbOKOnly + vbInformation, "Informacion")
      Exit Sub
    End If
  
    ' cargo formulario
    Load frmSubconcesionParamUpdate
    
    ' paso los valores del list view al formulario
    With frmSubconcesionParamUpdate
    .txtParticOil = lvwGetValue(lvwDatos, "particoil")
    .txtParticGas = lvwGetValue(lvwDatos, "particgas")
    .txtPjeSubEnt = lvwGetValue(lvwDatos, "pjeSubEnt")
    .txtPjeTerEnt = lvwGetValue(lvwDatos, "pjeSubTerm")
    .cboAjuApiStk.ListIndex = ComboBoxFindItem(.cboAjuApiStk, lvwGetValue(lvwDatos, "ajustaapistk"))
    .txtPrecioVenta = lvwGetValue(lvwDatos, "precioVenta")
    .TxtDescuento = lvwGetValue(lvwDatos, "Descuento")
    .txtIngBrutos = lvwGetValue(lvwDatos, "ingBrutos")
    .txtTransporte = lvwGetValue(lvwDatos, "transporte")
    .txtAlmacen = lvwGetValue(lvwDatos, "almacen")
    .txtDiasAlmacen = lvwGetValue(lvwDatos, "diasAlmacen")
    .txtEmbarque = lvwGetValue(lvwDatos, "embarque")
    .txtAlicRetExporta = lvwGetValue(lvwDatos, "alicRetExporta")
    .txtMercadoLocal = lvwGetValue(lvwDatos, "mercadoLocal")
    .Show vbModal
    End With
      
    If blnAceptar Then
      
      ' si hizo click en Aceptar, genero string y ejecuto
      ' funcion de UPDATE, el primer argumento enviado es
      ' el campo clave por el cual aplica el WHERE
        
      With frmSubconcesionParamUpdate
      strSQL = "EXEC SubconcesionesParamUpdateSP " & _
      lvwGetValue(lvwDatos, "empresaID") & "," & _
      lvwGetValue(lvwDatos, "IDsubconce") & "," & _
      Val(.txtParticOil) & "," & _
      Val(.txtParticGas) & "," & _
      Val(.txtPjeSubEnt) & "," & _
      Val(.txtPjeTerEnt) & "," & _
      "'" & .cboAjuApiStk & "'," & _
      Val(.txtPrecioVenta) & "," & _
      Val(.TxtDescuento) & "," & _
      Val(.txtIngBrutos) & "," & _
      Val(.txtTransporte) & "," & _
      Val(.txtAlmacen) & "," & _
      Val(.txtDiasAlmacen) & "," & _
      Val(.txtEmbarque) & "," & _
      Val(.txtAlicRetExporta) & "," & _
      Val(.txtMercadoLocal)
      End With

      intRes = adoExecSQL(strSQL)
      blnRefresh = True
      
    End If
      
   ' descargo formulario
    Unload frmSubconcesionParamUpdate
      
  Case Is = "eliminar"
  
  Case Is = "ordenar"
    
    intRes = lvwSortColumn(lvwDatos)
  
  Case Is = "filtrar"
      
      strWhere = FilterData(lvwDatos)
      If blnAceptar Then blnRefresh = True
      
    Case Is = "buscar"
      
      intRes = FindData(lvwDatos)
  
  Case Is = "exportar"
    
    intRes = ExportData(lvwDatos)
  
  Case Is = "ajustar"
     
    ' ajusta y envia a INI
    intRes = lvwAdjustColumn(lvwDatos, True)
    intRes = lvwWidthToKeyIni(lvwDatos, strTableNameActual)
  
  End Select

  If blnRefresh = True Then
      
    strSQL = "SELECT * FROM SubConcesionesParam_View" & _
            IIf(Not strWhere = "", " WHERE " & strWhere, "")
    intRes = ListViewRefresh(lvwDatos, strSQL)
    intRes = lvwHideColumn(lvwDatos, "IDSubconcesion")
    intRes = lvwHideColumn(lvwDatos, "empresaID")
    intRes = lvwHideColumn(lvwDatos, "concesionID")
    intRes = lvwHideColumn(lvwDatos, "pappath")
    intRes = lvwHideColumn(lvwDatos, "papcontrol")

  End If

End Sub
