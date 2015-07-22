VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmEntregasTra 
   BorderStyle     =   0  'None
   Caption         =   "Entregas Transportistas"
   ClientHeight    =   11445
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   12255
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   11445
   ScaleWidth      =   12255
   ShowInTaskbar   =   0   'False
   Begin MSComctlLib.ListView lvwCon 
      Height          =   1590
      Left            =   0
      TabIndex        =   2
      Top             =   4455
      Width           =   9675
      _ExtentX        =   17066
      _ExtentY        =   2805
      View            =   3
      LabelEdit       =   1
      Sorted          =   -1  'True
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      _Version        =   393217
      ForeColor       =   8421504
      BackColor       =   16777215
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   0
   End
   Begin MSComctlLib.ListView lvwSub 
      Height          =   4065
      Left            =   0
      TabIndex        =   3
      Top             =   6030
      Width           =   9675
      _ExtentX        =   17066
      _ExtentY        =   7170
      View            =   3
      LabelEdit       =   1
      Sorted          =   -1  'True
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      _Version        =   393217
      ForeColor       =   8421504
      BackColor       =   16777215
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   0
   End
   Begin MSComctlLib.ListView lvwDatos 
      Height          =   3615
      Left            =   0
      TabIndex        =   1
      Top             =   880
      Width           =   9675
      _ExtentX        =   17066
      _ExtentY        =   6376
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
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
      TabIndex        =   4
      Top             =   0
      Width           =   11865
      _ExtentX        =   20929
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
            Style           =   5
            BeginProperty ButtonMenus {66833FEC-8583-11D1-B16A-00C0F0283628} 
               NumButtonMenus  =   1
               BeginProperty ButtonMenu1 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "concesionAgregar"
                  Text            =   "Concesion"
               EndProperty
            EndProperty
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
            Style           =   5
            BeginProperty ButtonMenus {66833FEC-8583-11D1-B16A-00C0F0283628} 
               NumButtonMenus  =   1
               BeginProperty ButtonMenu1 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "concesionEliminar"
                  Text            =   "Concesion"
               EndProperty
            EndProperty
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
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   90
      Top             =   10665
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
            Picture         =   "frmEntregasTra.frx":0000
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmEntregasTra.frx":08DA
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmEntregasTra.frx":11B4
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmEntregasTra.frx":1A8E
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmEntregasTra.frx":3420
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmEntregasTra.frx":3CFA
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmEntregasTra.frx":45D4
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmEntregasTra.frx":5F66
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Label lblTitulo 
      Alignment       =   2  'Center
      BackColor       =   &H00808000&
      Caption         =   "Entregas Transportistas"
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
      Width           =   9675
   End
End
Attribute VB_Name = "frmEntregasTra"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Public Function entRefresh()
  Dim strRangoIso As isoLastPeriod
  Dim strAux As String
  
  'muestro solo ultimo periodo ingresado
  strRangoIso = adoLastPeriod("viewEntregasTra", "fecha")
  
  strSQL = "SELECT * FROM ViewEntregasTra where fecha between " & strRangoIso.strDesde & " and " & strRangoIso.strHasta & " " & _
           "ORDER BY IDentregatra desc"

  intRes = ListViewAppearanceChange(lvwDatos)
  intRes = ListViewRefresh(lvwDatos, strSQL, strStruc)
  intRes = lvwHideColumn(lvwDatos, "IDEnt")
  intRes = lvwHideColumn(lvwDatos, "empresaID")
  intRes = lvwHideColumn(lvwDatos, "status")
  
  ' refresh concesiones
  intRes = ConRefresh()
  
End Function

Public Function ConRefresh()
  Dim strCon, strAux As String

  If lvwDatos.SelectedItem Is Nothing Then        ' arma condicion para Concesion
    strCon = "0"
  Else
    strCon = lvwGetValue(lvwDatos, "IDentregaTra")
  End If

    ' guardo nombre tabla actual
  strAux = strTableNameActual
  strTableNameActual = "entregasTraDetCon"

  strSQL = "SELECT * FROM ViewEntregasTraDetCon WHERE EntregaTraID = " & strCon
  intRes = ListViewAppearanceChange(lvwCon)
  intRes = ListViewRefresh(lvwCon, strSQL)
  intRes = lvwHideColumn(lvwCon, "entregaTraID")
  intRes = lvwHideColumn(lvwCon, "ConcesionID")
  intRes = lvwHideColumn(lvwCon, "entregaTerID")

  ' recupero tabla actual
  strTableNameActual = strAux

  ' refresh data
  intRes = SubRefresh()

End Function

Public Function SubRefresh()
  Dim strDatos, strCon, strAux As String

  If lvwDatos.SelectedItem Is Nothing Then        ' arma condicion para SubConcesion
    strDatos = "0"
  Else
    strDatos = lvwGetValue(lvwDatos, "IDentregaTra")
  End If

  If lvwCon.SelectedItem Is Nothing Then        ' arma condicion para SubConcesion
    strCon = "0"
  Else
    strCon = lvwGetValue(lvwCon, "concesionID")
  End If
  
  ' guardo nombre tabla actual
  strAux = strTableNameActual
  strTableNameActual = "entregasTraDetSub"
  
  strSQL = "SELECT * FROM ViewEntregasTraDetSub WHERE EntregaTraID = " & strDatos & " AND ConcesionID = " & strCon
  intRes = ListViewAppearanceChange(lvwSub)
  intRes = ListViewRefresh(lvwSub, strSQL)
  intRes = lvwHideColumn(lvwSub, "entregaTraID")
  intRes = lvwHideColumn(lvwSub, "concesionID")
  
  ' recupero tabla actual
  strTableNameActual = strAux
 
End Function

Private Sub Form_Load()

  Dim intRes As Integer

  intRes = entRefresh()

End Sub

Private Sub lvwCon_Click()
  
  intRes = SubRefresh()

  ' cargo frmInfo
  Load frmInfo
  
  ' paso datos a mostrar
  frmInfo.txtInfo = vbCrLf & _
  " Total Hidratado 15: " & Format(objSumColumn(lvwCon, "hidratado15"), "##,###,##0.000") & vbCrLf & _
  " Total Agua y Sedim: " & Format(objSumColumn(lvwCon, "aguaysedim"), "##,###,##0.000") & vbCrLf & _
  " Total Seco Seco 15: " & Format(objSumColumn(lvwCon, "secoseco15"), "##,###,##0.000") & vbCrLf & _
  " Total Mermas Tpte : " & Format(objSumColumn(lvwCon, "mermastpte"), "##,###,##0.000") & vbCrLf & _
  " Total Mermas Otras: " & Format(objSumColumn(lvwCon, "mermasotras"), "##,###,##0.000") & vbCrLf & _
  " Total Neto 15     : " & Format(objSumColumn(lvwCon, "neto15"), "##,###,##0.000")
  
  ' ajusto ancho largo
  frmInfo.Width = 3500
  frmInfo.Height = 1800
  frmInfo.txtInfo.Width = 3500
  frmInfo.txtInfo.Height = 1800
  
  ' muestro form
  frmInfo.Show vbModal

End Sub

Private Sub lvwDatos_Click()

  ' refresh data
  intRes = ConRefresh()

End Sub


Private Sub lvwSub_Click()

  ' cargo frmInfo
  Load frmInfo
  
  ' paso datos a mostrar
  frmInfo.txtInfo = vbCrLf & _
  " Pje Sub Ent       : " & Format(objSumColumn(lvwSub, "pjesubent"), "##,###,##0.000") & vbCrLf & _
  " Total Hidratado 15: " & Format(objSumColumn(lvwSub, "hidratado15"), "##,###,##0.000") & vbCrLf & _
  " Total Agua y Sedim: " & Format(objSumColumn(lvwSub, "aguaysedim"), "##,###,##0.000") & vbCrLf & _
  " Total Seco Seco 15: " & Format(objSumColumn(lvwSub, "secoseco15"), "##,###,##0.000") & vbCrLf & _
  " Total Mermas Tpte : " & Format(objSumColumn(lvwSub, "mermastpte"), "##,###,##0.000") & vbCrLf & _
  " Total Mermas Otras: " & Format(objSumColumn(lvwSub, "mermasotras"), "##,###,##0.000") & vbCrLf & _
  " Total Neto 15     : " & Format(objSumColumn(lvwSub, "neto15"), "##,###,##0.000")
  
  ' ajusto ancho largo
  frmInfo.Width = 3500
  frmInfo.Height = 1900
  frmInfo.txtInfo.Width = 3500
  frmInfo.txtInfo.Height = 1900
  
  ' muestro form
  frmInfo.Show vbModal

End Sub


Private Sub tlbOperaciones_ButtonClick(ByVal Button As MSComctlLib.Button)
  Dim rs As ADODB.Recordset
  
  blnRefresh = False
  
  Select Case Button.Key
    
  Case Is = "agregar"   ' ------------------------------------------------------
  
      ' cargo formulario
    
      Load frmEntregasTraInfo
    
      ' muestro formulario
            
      frmEntregasTraInfo.Show vbModal
      
      ' si hizo click en Aceptar, en el formulario frmContratosInfo
      ' se pone en true la variable global blnAceptar y
      ' armo string y ejecuto funcion de INSERT
      
      If blnAceptar Then
      
        With frmEntregasTraInfo
        
        strSQL = "EXEC spEntregasTraInsert " & _
        .cboEmpresa.ItemData(.cboEmpresa.ListIndex) & "," & _
        .cboTransportista.ItemData(.cboTransportista.ListIndex) & "," & _
        "'" & .txtActa & "'," & _
        "'" & dateToIso(.txtFecha) & "'," & _
        .cboPuntoCarga.ItemData(.cboPuntoCarga.ListIndex) & "," & _
        .cboTerminal.ItemData(.cboTerminal.ListIndex) & "," & _
        Val(.txtAgua) & "," & _
        Val(.txtSedimento) & "," & _
        Val(.txtmermasOtras2) & "," & _
        Val(.txtDensiSecoSeco15) & "," & _
        Val(.txtAPI) & "," & _
        Val(.txtSalinidad) & "," & _
        Val(.txtVolHidratado15) & "," & _
        Val(.txtMermasTpte) & "," & _
        Val(.txtMermasOtras1) & "," & _
        Val(.txtSecoSeco15) & "," & _
        Val(.txtNeto15) & "," & _
        Val(.txtVolAguaSedim) & "," & _
        Val(.txtVolMermasOtras2) & "," & _
        Val(.txtVolMermasTpte) & "," & _
        Val(.txtVolMermasOtras1)

        End With
  
        intRes = adoExecSQL(strSQL)
        
        'chequeo errores
        If Not lngAdoErrNum = -1 Then
          adoError
          Exit Sub
        End If

        blnRefresh = True
      
      End If
      
      ' descargo formulario
      
      Unload frmEntregasTraInfo
  
  Case Is = "modificar"   '-----------------------------------------------------
  
    If lvwDatos Is Nothing Then
      intRes = MsgBox("No hay ningún item seleccionado.", vbApplicationModal + vbOKOnly + vbInformation, "Informacion")
      Exit Sub
    End If
    
    ' chequeo que no se haya procesado stock por terminales
    If lvwGetValue(lvwDatos, "status") <> 0 Then
      intRes = MsgBox("No es posible modificarlo, primero debe borrar el stock por subconcesiones.", vbApplicationModal + vbOKOnly + vbInformation, "Informacion")
      Exit Sub
    End If
    
    ' cargo formulario
    
    Load frmEntregasTraInfo
    
    ' paso los valores del list view al formulario
      
    With frmEntregasTraInfo
    .cboEmpresa.ListIndex = ComboBoxFindItem(.cboEmpresa, lvwGetValue(lvwDatos, "empre"))
    .cboTransportista.ListIndex = ComboBoxFindItem(.cboTransportista, lvwGetValue(lvwDatos, "trans"))
    .txtActa = lvwGetValue(lvwDatos, "acta")
    .txtFecha = lvwGetValue(lvwDatos, "fecha")
    .cboPuntoCarga.ListIndex = ComboBoxFindItem(.cboPuntoCarga, lvwGetValue(lvwDatos, "carga"))
    .cboTerminal.ListIndex = ComboBoxFindItem(.cboTerminal, lvwGetValue(lvwDatos, "terminal"))
    .txtAgua = lvwGetValue(lvwDatos, "agua")
    .txtSedimento = lvwGetValue(lvwDatos, "sedimento")
    .txtmermasOtras2 = lvwGetValue(lvwDatos, "mermasotras2")
    .txtDensiSecoSeco15 = lvwGetValue(lvwDatos, "densisecoseco15")
    .txtAPI = lvwGetValue(lvwDatos, "api")
    .txtSalinidad = lvwGetValue(lvwDatos, "sali")
    .txtVolHidratado15 = lvwGetValue(lvwDatos, "volhi")
    .txtMermasTpte = lvwGetValue(lvwDatos, "mermastp")
    .txtMermasOtras1 = lvwGetValue(lvwDatos, "mermasotras1")
    .txtSecoSeco15 = lvwGetValue(lvwDatos, "volsecoseco15")
    .txtNeto15 = lvwGetValue(lvwDatos, "volneto15")
    .txtVolAguaSedim = lvwGetValue(lvwDatos, "VolAguaSedim")
    .txtVolMermasOtras2 = lvwGetValue(lvwDatos, "VolMermasOtras2")
    .txtVolMermasTpte = lvwGetValue(lvwDatos, "VolMermasTpte")
    .txtVolMermasOtras1 = lvwGetValue(lvwDatos, "VolMermasOtras1")
    
    .Show vbModal
    End With
      
    If blnAceptar Then
      
      ' si hizo click en Aceptar, genero string y ejecuto
      ' funcion de UPDATE, el primer argumento enviado es
      ' el campo clave por el cual aplica el WHERE
        
      With frmEntregasTraInfo
      strSQL = "EXEC spEntregasTraUpdate " & _
      Me.lvwDatos.SelectedItem & "," & _
      .cboEmpresa.ItemData(.cboEmpresa.ListIndex) & "," & _
      .cboTransportista.ItemData(.cboTransportista.ListIndex) & "," & _
      "'" & .txtActa & "'," & _
      "'" & dateToIso(.txtFecha) & "'," & _
      .cboPuntoCarga.ItemData(.cboPuntoCarga.ListIndex) & "," & _
      .cboTerminal.ItemData(.cboTerminal.ListIndex) & "," & _
      Val(.txtAgua) & "," & _
      Val(.txtSedimento) & "," & _
      Val(.txtmermasOtras2) & "," & _
      Val(.txtDensiSecoSeco15) & "," & _
      Val(.txtAPI) & "," & _
      Val(.txtSalinidad) & "," & _
      Val(.txtVolHidratado15) & "," & _
      Val(.txtMermasTpte) & "," & _
      Val(.txtMermasOtras1) & "," & _
      Val(.txtSecoSeco15) & "," & _
      Val(.txtNeto15) & "," & _
      Val(.txtVolAguaSedim) & "," & _
      Val(.txtVolMermasOtras2) & "," & _
      Val(.txtVolMermasTpte) & "," & _
      Val(.txtVolMermasOtras1)
      
      End With

      intRes = adoExecSQL(strSQL)
      
      'chequeo errores
      If Not lngAdoErrNum = -1 Then
        adoError
        Exit Sub
      End If
      
      blnRefresh = True
      
    End If
      
    ' descargo formulario
    Unload frmEntregasTraInfo
  
  Case Is = "eliminar"  '-------------------------------------------------------
  
    If lvwDatos Is Nothing Then
      intRes = MsgBox("No hay ningún item seleccionado.", vbApplicationModal + vbOKOnly + vbInformation, "Informacion")
      Exit Sub
    End If
    
    ' chequeo que no se haya procesado stock por terminales
    If lvwGetValue(lvwDatos, "status") <> 0 Then
      intRes = MsgBox("No es posible eliminarlo, primero debe borrar el stock por subconcesiones.", vbApplicationModal + vbOKOnly + vbInformation, "Informacion")
      Exit Sub
    End If
      
    intRes = MsgBox("Esta seguro.", vbQuestion + vbYesNo, "Confirmacón")
      
    If intRes = vbNo Then Exit Sub
        
    strSQL = "EXEC spEntregasTraDelete " & lvwGetValue(lvwDatos, "IDentrega")
    intRes = adoExecSQL(strSQL)
    
    'chequeo errores
    If Not lngAdoErrNum = -1 Then
      adoError
      Exit Sub
    End If
    
    blnRefresh = True
     
  Case Is = "buscar"
    intRes = FindData(lvwDatos)
  
  Case Is = "ordenar"
    intRes = lvwSortColumn(lvwDatos)
  
  Case Is = "filtrar"
    strWhere = FilterData(lvwDatos)
    If blnAceptar Then blnRefresh = True

  Case Is = "exportar"
    intRes = ExportData(lvwDatos)
  
  Case Is = "ajustar"
    ' ajusta y envia a INI
    intRes = lvwAdjustColumn(lvwDatos, True)
    intRes = lvwWidthToKeyIni(lvwDatos, strTableNameActual)
  
  End Select
      
  'si hubo alguna actualizacion de datos hago refresh
  If blnRefresh Then
          
    If strWhere = "" Then
      
      'muestro solo ultimo periodo ingresado
      Dim strRangoIso As isoLastPeriod
      strRangoIso = adoLastPeriod("viewEntregasTra", "fecha")
      strWhere = "fecha between " & strRangoIso.strDesde & " and " & strRangoIso.strHasta
    
    End If
    
    strSQL = "SELECT * FROM ViewEntregasTra" & _
             IIf(Not strWhere = "", " WHERE " & strWhere, "") & " " & _
             "ORDER BY IDentregatra desc"
    intRes = ListViewRefresh(lvwDatos, strSQL)
    intRes = lvwHideColumn(lvwDatos, "IDEnt")
    intRes = lvwHideColumn(lvwDatos, "empresaID")
    intRes = lvwHideColumn(lvwDatos, "status")
    intRes = ConRefresh()
  
  End If

End Sub

Private Sub tlbOperaciones_ButtonMenuClick(ByVal ButtonMenu As MSComctlLib.ButtonMenu)
  
  Dim VHidratado15100, VHidratado15, PAgua, PSedimentos, VAguasedimentos, VMermasOtras2, VSecoseco15 As Currency
  Dim PMermastpte, PMermasotras, VMermastpte, VMermasOtras1, VNeto15 As Currency
  Dim blnRefresh As Boolean
  Dim rs As ADODB.Recordset

  blnRefresh = False

  Select Case ButtonMenu.Key
      
  Case Is = "concesionAgregar"   '-------------------------------------------------
      
    If lvwDatos.SelectedItem Is Nothing Then
      intRes = MsgBox("No hay ningun item seleccionado.", vbApplicationModal + vbOKOnly + vbInformation, "Informacion")
      Exit Sub
    End If
      
    ' cargo formulario
    
    Load frmEntregasTraDet
    
    ' paso valor de hidratado y el ID de empresa para filtro
    frmEntregasTraDet.txtVolumenHidratado15 = lvwGetValue(Me.lvwDatos, "volhidratado")
    frmEntregasTraDet.txtEmpresaID = lvwGetValue(Me.lvwDatos, "Empresaid")
    
    ' muestro formulario
            
    frmEntregasTraDet.Show vbModal
      
    ' si hizo click en Aceptar, en el formulario frmContratosInfo
    ' se pone en true la variable global blnAceptar y
    ' armo string y ejecuto funcion de INSERT
      
    If blnAceptar Then
      
      With frmEntregasTraDet

      ' agrego encabezado de concesion en EntregasTraCon
      strSQL = "EXEC spEntregasTraDetConInsert " & _
                Val(lvwGetValue(lvwDatos, "IDentrega")) & "," & _
                .cboConcesion.ItemData(.cboConcesion.ListIndex) & "," & _
                Val(.txtVolumenHidratado15) & "," & _
                Val(.txtVolumenAguaSedimento) & "," & _
                Val(.txtmermasOtras2) & "," & _
                Val(.txtVolumenSecoSeco15) & "," & _
                Val(.txtMermasTpte) & "," & _
                Val(.txtMermasOtras1) & "," & _
                Val(.txtVolumenNeto15)
          intRes = adoExecSQL(strSQL)
        
      ' abro recordset tomo todas las subconcesiones de la
      ' concesion seleccionada para repartir los volumnes y
      ' agrego detalle en Subconcesiones en EntregasTraSub
      strSQL = "SELECT * FROM SubConcesionesParam_View WHERE empresaID = " & Val(lvwGetValue(lvwDatos, "empresaID")) & " AND " & _
               "ConcesionID = " & .cboConcesion.ItemData(.cboConcesion.ListIndex)
      Set rs = adoGetRS(strSQL)
        
      If Not rs.EOF() Then
        
        'declaro
        Dim blnGuardoPrimera As Boolean
        Dim intIDEntregaTraAUX, intConcesionAUX, intIDSubconcesionAUX As Integer
        Dim sngVH, sngVAS, sngVM1, sngVM2, sngVS, sngVT, sngVN
                
        'inicializo
        sngVH = 0
        sngVAS = 0
        sngVM1 = 0
        sngVM2 = 0
        sngVS = 0
        sngVT = 0
        sngVN = 0
        
        While Not rs.EOF()
          
          VHidratado15 = Round(CCur(.txtVolumenHidratado15) * rs!PjeSubEnt / 100, 3)
          VAguasedimentos = Round(CCur(.txtVolumenAguaSedimento) * rs!PjeSubEnt / 100, 3)
          VMermasOtras2 = Round(CCur(.txtmermasOtras2) * rs!PjeSubEnt / 100, 3)
          VSecoseco15 = Round(CCur(.txtVolumenSecoSeco15) * rs!PjeSubEnt / 100, 3)
          VMermastpte = Round(CCur(.txtMermasTpte) * rs!PjeSubEnt / 100, 3)
          VMermasOtras1 = Round(CCur(.txtMermasOtras1) * rs!PjeSubEnt / 100, 3)
          VNeto15 = Round(CCur(.txtVolumenNeto15) * rs!PjeSubEnt / 100, 3)
          
          strSQL = "EXEC spEntregasTraDetSubInsert " & _
                    Val(lvwGetValue(lvwDatos, "IDentregaTra")) & "," & _
                    rs!IDsubconcesion & "," & _
                    .cboConcesion.ItemData(.cboConcesion.ListIndex) & "," & _
                    Val(.txtVolumenHidratado15) & "," & _
                    rs!PjeSubEnt & "," & _
                    VHidratado15 & "," & _
                    VAguasedimentos & "," & _
                    VMermasOtras2 & "," & _
                    VSecoseco15 & "," & _
                    VMermastpte & "," & _
                    VMermasOtras1 & "," & _
                    VNeto15
          intRes = adoExecSQL(strSQL)
          
          'acumulo
          sngVH = sngVH + VHidratado15
          sngVAS = sngVAS + VAguasedimentos
          sngVM1 = sngVM1 + VMermasOtras1
          sngVM2 = sngVM2 + VMermasOtras2
          sngVS = sngVS + VSecoseco15
          sngVT = sngVT + VMermastpte
          sngVN = sngVN + VNeto15
            
          'utilizo una flag para saber guardar valores ID de la primer subconcesion
          'si hay diferencia cuando calculo los porcentajes se los aplico a la primera
          'subconcesion
          If Not blnGuardoPrimera Then
            intIDEntregaTraAUX = Val(lvwGetValue(lvwDatos, "IDentregaTra"))
            intIDConcesionAUX = .cboConcesion.ItemData(.cboConcesion.ListIndex)
            intIDSubconcesionAUX = rs!IDsubconcesion
            blnGuardoPrimera = True
          End If
           
          'puntero al siguiente
          rs.MoveNext
          
        Wend
        
        'tomo el procenjaje de participacion de cada concesion
        Dim rsCon As ADODB.Recordset
        
        strSQL = "SELECT SUM(PjeSubEnt) AS PjeSubEnt From SubConcesiones " & _
                 "where concesionID = " & .cboConcesion.ItemData(.cboConcesion.ListIndex)
        
        Set rsCon = adoGetRS(strSQL)
                
        'si hay diferencia
        If sngVH <> Round(CCur(.txtVolumenHidratado15 * rsCon!PjeSubEnt / 100), 3) Or sngVAS <> Round(CCur(.txtVolumenAguaSedimento * rsCon!PjeSubEnt / 100), 3) Or _
           sngVM1 <> Round(CCur(.txtMermasOtras1 * rsCon!PjeSubEnt / 100), 3) Or sngVM2 <> Round(CCur(.txtmermasOtras2 * rsCon!PjeSubEnt / 100), 3) Or _
           sngVS <> Round(CCur(.txtVolumenSecoSeco15 * rsCon!PjeSubEnt / 100), 3) Or sngVT <> Round(CCur(.txtMermasTpte * rsCon!PjeSubEnt / 100), 3) Or _
           sngVN <> Round(CCur(.txtVolumenNeto15 * rsCon!PjeSubEnt / 100), 3) Then
          
          'se la suma o resta a la primera subconcesion que encontro
          strSQL = "EXEC spEntregasTraDetSubUpdate " & _
                    intIDEntregaTraAUX & "," & _
                    intIDConcesionAUX & "," & _
                    intIDSubconcesionAUX & "," & _
                    Round(CCur(.txtVolumenHidratado15 * rsCon!PjeSubEnt / 100), 3) - sngVH & "," & _
                    Round(CCur(.txtVolumenAguaSedimento * rsCon!PjeSubEnt / 100), 3) - sngVAS & "," & _
                    Round(CCur(.txtmermasOtras2 * rsCon!PjeSubEnt / 100), 3) - sngVM2 & "," & _
                    Round(CCur(.txtVolumenSecoSeco15 * rsCon!PjeSubEnt / 100), 3) - sngVS & "," & _
                    Round(CCur(.txtMermasTpte * rsCon!PjeSubEnt / 100), 3) - sngVT & "," & _
                    Round(CCur(.txtMermasOtras1 * rsCon!PjeSubEnt / 100), 3) - sngVM1 & "," & _
                    Round(CCur(.txtVolumenNeto15 * rsCon!PjeSubEnt / 100), 3) - sngVN
          intRes = adoExecSQL(strSQL)
          
        End If
                
      End If
      
      rs.Close
      
      End With
  
      blnRefresh = True
      
    End If
      
    ' descargo formulario
      
    Unload frmEntregasTraDet
      
  Case Is = "concesionEliminar"    '-----------------------------------------------
       
    If lvwCon.SelectedItem Is Nothing Then
      intRes = MsgBox("No hay ningun Concesión seleccionado.", vbApplicationModal + vbOKOnly + vbInformation, "Informacion")
      Exit Sub
    End If
      
    intRes = MsgBox("Esta seguro que desea eliminar la Concesión: " & lvwGetValue(lvwCon, "concesion") & ".", vbQuestion + vbYesNo, "Confirmacón")
      
    If intRes = vbYes Then
        
      strSQL = "EXEC spEntregasTraDetConDelete " & Val(lvwGetValue(lvwCon, "entregaTraID")) & "," & Val(lvwGetValue(lvwCon, "ConcesionID"))
      a = adoExecSQL(strSQL)
      
      blnRefresh = True
      
    End If
      
  End Select

  ' si hubo alguna actualizacion de datos hago refresh
    
  If blnRefresh Then
          
    intRes = ConRefresh()
  
  End If

End Sub

