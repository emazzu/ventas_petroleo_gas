VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomctl.ocx"
Begin VB.Form frmCotizacionDolar 
   BorderStyle     =   0  'None
   ClientHeight    =   6765
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   11205
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   6765
   ScaleWidth      =   11205
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
            Picture         =   "frmCotizacionDolar.frx":0000
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCotizacionDolar.frx":08DA
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCotizacionDolar.frx":11B4
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCotizacionDolar.frx":1A8E
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCotizacionDolar.frx":3420
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCotizacionDolar.frx":3CFA
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCotizacionDolar.frx":45D4
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCotizacionDolar.frx":5F66
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar tlbOperaciones 
      Height          =   570
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   12060
      _ExtentX        =   21273
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
      Caption         =   "Cotizacion Dolar"
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
Attribute VB_Name = "frmCotizacionDolar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Form_Load()
  Dim strRangoIso As isoLastPeriod
  
  'muestro solo ultimo periodo ingresado
  strRangoIso = adoLastPeriod("cotizacionDolar_vw", "fecha")
  
  strSQL = "SELECT * FROM cotizacionDolar_vw where fecha between " & strRangoIso.strDesde & " and " & strRangoIso.strHasta & " " & _
           "ORDER BY fecha DESC"

  intRes = ListViewAppearanceChange(lvwDatos)
  intRes = ListViewRefresh(lvwDatos, strSQL, strStruc)
  intRes = lvwHideColumn(lvwDatos, "IDcotizacion")

  ' ordeno por fecha

  lvwDatos.Sorted = True
  lvwDatos.SortKey = 1
  lvwDatos.SortOrder = lvwDescending

End Sub


Private Sub tlbOperaciones_ButtonClick(ByVal Button As MSComctlLib.Button)
  Dim strSQL As String
  Dim intRes As Integer
  
  blnRefresh = False
      
  Select Case Button.Key
      
  Case Is = "agregar"
      
      ' cargo formulario
      Load frmCotizacionDolarInfo
      
      ' muestro formulario
      frmCotizacionDolarInfo.Show vbModal
      
      ' descargo formulario
      Unload frmCotizacionDolarInfo
      
  Case Is = "modificar"
      
    If lvwDatos.SelectedItem > 0 Then
      
      'cargo formulario
      Load frmCotizacionDolarUpdate
      
      'paso los valores del list view al formulario
      With frmCotizacionDolarUpdate
      'le paso los datos
      .txtDato1 = lvwGetValue(lvwDatos, "comprador")
      .txtDato2 = lvwGetValue(lvwDatos, "vendedor")
      'pinto campo para edicion
      frmCotizacionDolarUpdate.txtDato1.SelLength = Len(frmCotizacionDolarUpdate.txtDato1)
      frmCotizacionDolarUpdate.txtDato2.SelLength = Len(frmCotizacionDolarUpdate.txtDato2)
      .Show vbModal
      End With
      
      If blnAceptar Then
        
        ' si hizo click en Aceptar, genero string y ejecuto
        ' funcion de UPDATE, el primer argumento enviado es
        ' el campo clave por el cual aplica el WHERE
        
        With frmCotizacionDolarUpdate
        strSQL = "EXEC spCotizacionDolarUpdate " & _
        lvwGetValue(lvwDatos, "IDcotizacion") & "," & _
        Val(.txtDato1) & "," & _
        Val(.txtDato2)
        End With
        
        intRes = adoExecSQL(strSQL)
        blnRefresh = True
        
      End If
      
      'descargo formulario
      'Unload frmPreciosUpdate
      
    Else
      a = MsgBox("No hay ningun item seleccionado.", vbApplicationModal + vbOKOnly + vbInformation, "Informacion")
    End If
  
    Case Is = "eliminar"
  
      If lvwDatos.SelectedItem > 0 Then
      
        intRes = MsgBox("Esta seguro que desea eliminar el elemento seleccionado.", vbQuestion + vbYesNo, "Confirmacón")
      
        If intRes = vbYes Then
        
          strSQL = "EXEC spCotizacionDolarDelete " & Me.lvwDatos.SelectedItem
          intRes = adoExecSQL(strSQL)
          blnRefresh = True
      
        End If
      
      Else
        a = MsgBox("No hay ningun item seleccionado.", vbApplicationModal + vbOKOnly + vbInformation, "Informacion")
      End If
  
    Case Is = "filtrar"
      strWhere = FilterData(lvwDatos)
      If blnAceptar Then blnRefresh = True

    Case Is = "buscar"
      intRes = FindData(lvwDatos)
  
    Case Is = "ordenar"
      intRes = lvwSortColumn(lvwDatos)
  
    Case Is = "exportar"
     
      intRes = ExportData(lvwDatos)
  
    Case Is = "ajustar"
     
      ' ajusta y envia a INI
      intRes = lvwAdjustColumn(lvwDatos, True)
      intRes = lvwWidthToKeyIni(lvwDatos, strTableNameActual)
  
  End Select

  ' hago un refresh de los datos para ver los cambios
  If blnRefresh Then
    
    If strWhere = "" Then
      
      'muestro solo ultimo periodo ingresado
      Dim strRangoIso As isoLastPeriod
      strRangoIso = adoLastPeriod("cotizacionDolar_vw", "fecha")
      strWhere = "fecha between " & strRangoIso.strDesde & " and " & strRangoIso.strHasta
    
    End If
     
    strSQL = "SELECT * FROM cotizacionDolar_vw" & _
              IIf(Not strWhere = "", " WHERE " & strWhere, "")
    intRes = ListViewRefresh(lvwDatos, strSQL)
    intRes = lvwHideColumn(lvwDatos, "IDcotizacion")
  
  End If

End Sub


