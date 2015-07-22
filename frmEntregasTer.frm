VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmEntregasTer 
   BorderStyle     =   0  'None
   Caption         =   "Entregas Terminales"
   ClientHeight    =   11520
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   12600
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   11520
   ScaleWidth      =   12600
   ShowInTaskbar   =   0   'False
   Begin MSComctlLib.Toolbar tlbOperaciones 
      Height          =   570
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   11730
      _ExtentX        =   20690
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
                  Key             =   "entregasTraAgregar"
                  Text            =   "Entregas Transportistas"
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
                  Key             =   "entregasTraEliminar"
                  Text            =   "Entregas Transportistas"
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
   Begin MSComctlLib.ListView lvwCon 
      Height          =   5160
      Left            =   0
      TabIndex        =   3
      Top             =   5670
      Width           =   9675
      _ExtentX        =   17066
      _ExtentY        =   9102
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
      Height          =   4845
      Left            =   0
      TabIndex        =   2
      Top             =   870
      Width           =   9675
      _ExtentX        =   17066
      _ExtentY        =   8546
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
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   45
      Top             =   10890
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
            Picture         =   "frmEntregasTer.frx":0000
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmEntregasTer.frx":08DA
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmEntregasTer.frx":11B4
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmEntregasTer.frx":1A8E
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmEntregasTer.frx":3420
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmEntregasTer.frx":3CFA
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmEntregasTer.frx":45D4
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmEntregasTer.frx":5F66
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Label lblTitulo 
      Alignment       =   2  'Center
      BackColor       =   &H00808000&
      Caption         =   "Entregas Terminales"
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
      TabIndex        =   1
      Top             =   585
      Width           =   9675
   End
End
Attribute VB_Name = "frmEntregasTer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public Function entRefresh()
  Dim strRangoIso As isoLastPeriod
  Dim strAux As String
  
  'muestro solo ultimo periodo ingresado
  strRangoIso = adoLastPeriod("ViewEntregasTer", "fecha")
  
  strSQL = "SELECT * FROM ViewEntregasTer where fecha between " & strRangoIso.strDesde & " and " & strRangoIso.strHasta & " " & _
           "ORDER BY IDEntregaTer desc"

  intRes = ListViewAppearanceChange(lvwDatos)
  intRes = ListViewRefresh(lvwDatos, strSQL, strStruc)
  intRes = lvwHideColumn(lvwDatos, "IDEntregaTer")
  intRes = lvwHideColumn(lvwDatos, "terminalID")
  intRes = lvwHideColumn(lvwDatos, "empresaID")
  intRes = lvwHideColumn(lvwDatos, "status")
    
  ' refresh data
  intRes = ConRefresh()

End Function

Public Function ConRefresh()

  Dim strSQL, strCon  As String
  Dim intRes As Integer

  If lvwDatos.SelectedItem Is Nothing Then        ' arma condicion para Concesion
    strCon = "0"
  Else
    strCon = lvwGetValue(lvwDatos, "IDentregaTer")
  End If

  ' guardo nombre tabla actual
  strAux = strTableNameActual
  strTableNameActual = "entregasTerDet"

  strSQL = "SELECT * FROM ViewEntregasTerDet WHERE EntregaTerID = " & strCon
  intRes = ListViewAppearanceChange(lvwCon)
  intRes = ListViewRefresh(lvwCon, strSQL)
  intRes = lvwHideColumn(lvwCon, "concesionID")
  intRes = lvwHideColumn(lvwCon, "entregaTerID")
  intRes = lvwHideColumn(lvwCon, "entregaTraID")

  ' recupero tabla actual
  strTableNameActual = strAux

End Function

Private Sub Form_Load()

  Dim intRes As Integer

  intRes = entRefresh()

End Sub

Private Sub lvwDatos_Click()
  
  ' muestra formulario
  intRes = ConRefresh()

End Sub

Private Sub tlbOperaciones_ButtonClick(ByVal Button As MSComctlLib.Button)
  Dim rs As ADODB.Recordset
  
  blnRefresh = False
  
  Select Case Button.Key
    
  Case Is = "agregar"   ' ------------------------------------------------------
  
      ' cargo formulario
    
      Load frmEntregasTerInfo
    
      ' muestro formulario
            
      frmEntregasTerInfo.Show vbModal
      
      ' si hizo click en Aceptar, en el formulario frmContratosInfo
      ' se pone en true la variable global blnAceptar y
      ' armo string y ejecuto funcion de INSERT
      
      If blnAceptar Then
      
        With frmEntregasTerInfo
        
        strSQL = "EXEC spEntregasTerInsert " & _
        .cboEmpresa.ItemData(.cboEmpresa.ListIndex) & "," & _
        .cboTransportista.ItemData(.cboTransportista.ListIndex) & "," & _
        .cboTerminal.ItemData(.cboTerminal.ListIndex) & "," & _
        "'" & .txtCertificado & "'," & _
        "'" & dateToIso(.txtFecha) & "'," & _
        Val(.txtAPI) & "," & _
        Val(.txtPjeMermas) & "," & _
        Val(.txtAPICoefAju)
        
        intRes = adoExecSQL(strSQL)
        
        'chequeo errores
        If Not lngAdoErrNum = -1 Then
          adoError
          Exit Sub
        End If

        blnRefresh = True

        End With
       
      
      End If
      
      ' descargo formulario
      
      Unload frmEntregasTerInfo
  
  Case Is = "modificar"   '-----------------------------------------------------
  
    If lvwDatos Is Nothing Then
      intRes = MsgBox("No hay ningún item seleccionado.", vbApplicationModal + vbOKOnly + vbInformation, "Informacion")
      Exit Sub
    End If
    
    ' chequeo que no se haya procesado stock por terminales
    If lvwGetValue(lvwDatos, "status") <> 0 Then
      intRes = MsgBox("No es posible modificarlo, primero debe borrar el stock por terminales.", vbApplicationModal + vbOKOnly + vbInformation, "Informacion")
      Exit Sub
    End If
    
    ' cargo formulario
    Load frmEntregasTerInfo
    
    ' paso los valores del list view al formulario
    With frmEntregasTerInfo
    .cboEmpresa.ListIndex = ComboBoxFindItem(.cboEmpresa, lvwGetValue(lvwDatos, "empre"))
    .cboTransportista.ListIndex = ComboBoxFindItem(.cboTransportista, lvwGetValue(lvwDatos, "trans"))
    .cboTerminal.ListIndex = ComboBoxFindItem(.cboTerminal, lvwGetValue(lvwDatos, "terminal"))
    .txtCertificado = lvwGetValue(lvwDatos, "certi")
    .txtFecha = lvwGetValue(lvwDatos, "fecha")
    .txtAPI = lvwGetValue(lvwDatos, "api")
    .txtPjeMermas = lvwGetValue(lvwDatos, "(%) Mermas")
    .txtAPICoefAju = lvwGetValue(lvwDatos, "API Coef Ajuste")
    .Show vbModal
    End With
      
    If blnAceptar Then
    
      ' si hizo click en Aceptar, genero string y ejecuto
      ' funcion de UPDATE, el primer argumento enviado es
      ' el campo clave por el cual aplica el WHERE
       
      With frmEntregasTerInfo
        
      strSQL = "EXEC spEntregasTerUpdate " & _
      lvwGetValue(Me.lvwDatos, "IDentregaTer") & "," & _
      .cboEmpresa.ItemData(.cboEmpresa.ListIndex) & "," & _
      .cboTransportista.ItemData(.cboTransportista.ListIndex) & "," & _
      .cboTerminal.ItemData(.cboTerminal.ListIndex) & "," & _
      "'" & .txtCertificado & "'," & _
      "'" & dateToIso(.txtFecha) & "'," & _
      Val(.txtAPI) & "," & _
      Val(.txtPjeMermas) & "," & _
      Val(.txtAPICoefAju)
        
      intRes = adoExecSQL(strSQL)
      blnRefresh = True
 
      End With
      
    End If
      
    ' descargo formulario
    Unload frmEntregasTerInfo
  
  Case Is = "eliminar"  '-------------------------------------------------------
  
    If lvwDatos Is Nothing Then
      intRes = MsgBox("No hay ningún item seleccionado.", vbApplicationModal + vbOKOnly + vbInformation, "Informacion")
      Exit Sub
    End If
    
    ' chequeo que no se haya procesado stock por terminales
    If lvwGetValue(lvwDatos, "status") <> 0 Then
      intRes = MsgBox("No es posible eliminarlo, primero debe borrar el stock por terminales.", vbApplicationModal + vbOKOnly + vbInformation, "Informacion")
      Exit Sub
    End If
      
    intRes = MsgBox("Esta seguro que desea eliminar el elemento seleccionado.", vbQuestion + vbYesNo, "Confirmacón")
      
    If intRes = vbYes Then
        
      strSQL = "EXEC spEntregasterDelete " & lvwGetValue(Me.lvwDatos, "IDentregaTer")
      intRes = adoExecSQL(strSQL)
      blnRefresh = True
      
    End If

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

  ' si hubo alguna actualizacion de datos hago refresh
    
  If blnRefresh Then
          
    If strWhere = "" Then
      
      'muestro solo ultimo periodo ingresado
      Dim strRangoIso As isoLastPeriod
      strRangoIso = adoLastPeriod("ViewEntregasTer", "fecha")
      strWhere = "fecha between " & strRangoIso.strDesde & " and " & strRangoIso.strHasta
    
    End If
          
    strSQL = "SELECT * FROM ViewEntregasTer" & _
             IIf(Not strWhere = "", " WHERE " & strWhere, "") & " " & _
             "ORDER BY IDEntregaTer desc"
    intRes = ListViewRefresh(lvwDatos, strSQL)
    intRes = lvwHideColumn(lvwDatos, "IDEnt")
    intRes = lvwHideColumn(lvwDatos, "terminalID")
    intRes = lvwHideColumn(lvwDatos, "empresaID")
    intRes = lvwHideColumn(lvwDatos, "status")
    
    intRes = ConRefresh()
  
  End If
  
End Sub


Private Sub tlbOperaciones_ButtonMenuClick(ByVal ButtonMenu As MSComctlLib.ButtonMenu)
  Dim intInd As Integer

  blnRefresh = False

  Select Case ButtonMenu.Key
      
  Case Is = "entregasTraAgregar"   '-------------------------------------------------
      
    If lvwDatos.SelectedItem Is Nothing Then
      intRes = MsgBox("No hay algún item seleccionado.", vbApplicationModal + vbOKOnly + vbInformation, "Informacion")
      Exit Sub
    End If
      
    ' cargo formulario
    Load frmEntregasTerDet
    
    ' le paso el codigo de empresa para poder utilizarlo para filtrar
    frmEntregasTerDet.txtEmpresaID = lvwGetValue(Me.lvwDatos, "empresaID")
    
    ' muestro formulario
    frmEntregasTerDet.Show vbModal
      
    ' si hizo click en Aceptar, en el formulario frmContratosInfo
    ' se pone en true la variable global blnAceptar y
    ' armo string y ejecuto funcion de INSERT
      
    If blnAceptar Then
      
      ' Inserta en EntregasTerDet y Actualiza EntregasTraDetCon
      For intInd = 1 To frmEntregasTerDet.lvwCon.ListItems.Count
      
        ' recorre cada item seleccionado
        If frmEntregasTerDet.lvwCon.ListItems(intInd).Checked Then
          
          ' pongo puntero en la seleccion
          frmEntregasTerDet.lvwCon.ListItems(intInd).Selected = True
          
          ' actualizo datos
          strSQL = "EXEC spEntregasTerDetInsert " & _
                  lvwGetValue(Me.lvwDatos, "IDentregaTer") & "," & _
                  lvwGetValue(frmEntregasTerDet.lvwCon, "entregaTraID") & "," & _
                  lvwGetValue(frmEntregasTerDet.lvwCon, "concesionID")
          intRes = adoExecSQL(strSQL)
        End If
  
      Next
  
      blnRefresh = True
      
    End If
      
    ' descargo formulario
      
    Unload frmEntregasTerDet
      
  Case Is = "entregasTraEliminar"    '-----------------------------------------------
       
    If lvwCon.SelectedItem Is Nothing Then
      intRes = MsgBox("No hay ningun Concesión seleccionado.", vbApplicationModal + vbOKOnly + vbInformation, "Informacion")
      Exit Sub
    End If
      
    intRes = MsgBox("Esta seguro que desea eliminar la Concesión: " & lvwGetValue(lvwCon, "conce") & ".", vbQuestion + vbYesNo, "Confirmacón")
      
    If intRes = vbYes Then
        
      strSQL = "EXEC spEntregasTerDetDelete " & Val(lvwGetValue(lvwDatos, "IDentregaTer")) & "," & Val(lvwGetValue(lvwCon, "entregaTraID")) & "," & Val(lvwGetValue(lvwCon, "ConcesionID"))
      intRes = adoExecSQL(strSQL)
      
      blnRefresh = True
      
    End If
      
  End Select

  ' si hubo alguna actualizacion de datos hago refresh
    
  If blnRefresh Then
          
    intRes = ConRefresh()
  
  End If


End Sub
