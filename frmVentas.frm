VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmVentas 
   BorderStyle     =   0  'None
   Caption         =   "Ventas"
   ClientHeight    =   7320
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   11175
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7320
   ScaleWidth      =   11175
   ShowInTaskbar   =   0   'False
   Begin MSComctlLib.Toolbar tblOperaciones 
      Align           =   1  'Align Top
      Height          =   570
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   11175
      _ExtentX        =   19711
      _ExtentY        =   1005
      ButtonWidth     =   2170
      ButtonHeight    =   1005
      Style           =   1
      TextAlignment   =   1
      ImageList       =   "ImageList1"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   8
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Agrega"
            Key             =   "agregar"
            ImageIndex      =   1
            Style           =   5
            BeginProperty ButtonMenus {66833FEC-8583-11D1-B16A-00C0F0283628} 
               NumButtonMenus  =   3
               BeginProperty ButtonMenu1 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "FacOil"
                  Text            =   "Facturas Oil"
               EndProperty
               BeginProperty ButtonMenu2 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "OilDebCre"
                  Text            =   "Créditos - Débitos (Oil)"
               EndProperty
               BeginProperty ButtonMenu3 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "VarDebCre"
                  Text            =   "Facturas - Débitos - Créditos (Varios)"
               EndProperty
            EndProperty
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Genera"
            Key             =   "Genera"
            ImageIndex      =   7
            Style           =   5
            BeginProperty ButtonMenus {66833FEC-8583-11D1-B16A-00C0F0283628} 
               NumButtonMenus  =   4
               BeginProperty ButtonMenu1 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "comprobante"
                  Text            =   "Impresión de Comprobantre"
               EndProperty
               BeginProperty ButtonMenu2 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "comprobanteBorrador"
                  Text            =   "Impresión de Comprobantre Borrador"
               EndProperty
               BeginProperty ButtonMenu3 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Text            =   "--"
               EndProperty
               BeginProperty ButtonMenu4 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "comprobante_PDF"
                  Text            =   "Generación de Comprobante en PDF"
               EndProperty
            EndProperty
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Aprobar"
            Key             =   "Aprobar"
            ImageIndex      =   3
            Style           =   5
            BeginProperty ButtonMenus {66833FEC-8583-11D1-B16A-00C0F0283628} 
               NumButtonMenus  =   2
               BeginProperty ButtonMenu1 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "autorizar"
                  Text            =   "Enviar a AFIP para autorizar"
               EndProperty
               BeginProperty ButtonMenu2 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "eliminar"
                  Text            =   "Eliminar Comprobante"
               EndProperty
            EndProperty
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Filtra"
            Key             =   "filtrar"
            ImageIndex      =   4
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Busca"
            Key             =   "buscar"
            ImageIndex      =   5
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Ordena"
            Key             =   "ordenar"
            ImageIndex      =   6
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Exporta"
            Key             =   "exportar"
            ImageIndex      =   8
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Ajusta"
            Key             =   "ajustar"
            ImageIndex      =   9
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ListView lvwDatos 
      Height          =   5730
      Left            =   0
      TabIndex        =   2
      Top             =   855
      Width           =   9240
      _ExtentX        =   16298
      _ExtentY        =   10107
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
      Left            =   90
      Top             =   6660
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   9
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmVentas.frx":0000
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmVentas.frx":08DA
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmVentas.frx":11B4
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmVentas.frx":1A8E
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmVentas.frx":3420
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmVentas.frx":3CFA
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmVentas.frx":45D4
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmVentas.frx":4EAE
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmVentas.frx":6840
            Key             =   ""
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
      TabIndex        =   1
      Top             =   585
      Width           =   9225
   End
End
Attribute VB_Name = "frmVentas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Sub Sleep Lib "kernel32.dll" (ByVal dwMilliseconds As Long)




Private Sub CRV1_CloseButtonClicked(UseDefault As Boolean)
CRV1.re
End Sub

Private Sub Form_Load()
  Dim strRangoIso As isoLastPeriod
  
  'muestro solo ultimo periodo ingresado
  strRangoIso = adoLastPeriod("ViewVentas", "fecha")
  
  strSQL = "SELECT * FROM ViewVentas where fecha between " & strRangoIso.strDesde & " and " & strRangoIso.strHasta & " " & _
           "ORDER BY fecha DESC"

  intRes = ListViewAppearanceChange(lvwDatos)
  intRes = ListViewRefresh(lvwDatos, strSQL, strStruc)
  intRes = lvwHideColumn(lvwDatos, "concepto")
  intRes = lvwHideColumn(lvwDatos, "condicion")
  intRes = lvwHideColumn(lvwDatos, "status")
  intRes = lvwHideColumn(lvwDatos, "IDempresa")

'   Mazzu
'   04/05/2015
'  intRes = lvwHideColumn(lvwDatos, "emitida")

End Sub



Private Sub lvwDatos_DblClick()
  Dim rs As ADODB.Recordset
  
  'tomo rs con consulta
  strSQL = "select * from ventasDetalleConsulta where factura = '" & lvwGetValue(lvwDatos, "factura") & "'"
  Set rs = adoGetRS(strSQL)
    
  Load ventasDetalleConsultaFrm               'cargo frm
  ventasDetalleConsultaFrm.txtConsulta = ""   'vacio consulta
  
  'recorro detalle
  While Not rs.EOF
  
    ventasDetalleConsultaFrm.txtConsulta = ventasDetalleConsultaFrm.txtConsulta & "Item          : " & str(rs!idItem) & vbCrLf
    ventasDetalleConsultaFrm.txtConsulta = ventasDetalleConsultaFrm.txtConsulta & "Concepto      : " & Format(rs!concepto, "###,###,##0.00") & vbCrLf
    ventasDetalleConsultaFrm.txtConsulta = ventasDetalleConsultaFrm.txtConsulta & "Cantidad      : " & Format(rs!Cantidad, "###,###,##0.00") & vbCrLf
    ventasDetalleConsultaFrm.txtConsulta = ventasDetalleConsultaFrm.txtConsulta & "Cantidad1     : " & Format(rs!cantidadInfo, "###,###,##0.00") & vbCrLf
    ventasDetalleConsultaFrm.txtConsulta = ventasDetalleConsultaFrm.txtConsulta & "Cantidad2     : " & Format(rs!cantidadInfo1, "###,###,##0.00") & vbCrLf
    ventasDetalleConsultaFrm.txtConsulta = ventasDetalleConsultaFrm.txtConsulta & "Precio        : " & Format(rs!precio, "###,###,##0.000") & vbCrLf
    ventasDetalleConsultaFrm.txtConsulta = ventasDetalleConsultaFrm.txtConsulta & "Importe       : " & Format(rs!Importe, "###,###,##0.00") & vbCrLf
    ventasDetalleConsultaFrm.txtConsulta = ventasDetalleConsultaFrm.txtConsulta & vbCrLf
    ventasDetalleConsultaFrm.txtConsulta = ventasDetalleConsultaFrm.txtConsulta & "Contrato" & vbCrLf
    ventasDetalleConsultaFrm.txtConsulta = ventasDetalleConsultaFrm.txtConsulta & "Cliente       : " & rs!cliente & vbCrLf
    ventasDetalleConsultaFrm.txtConsulta = ventasDetalleConsultaFrm.txtConsulta & "Fecha Desde   : " & rs!fechaDesde & vbCrLf
    ventasDetalleConsultaFrm.txtConsulta = ventasDetalleConsultaFrm.txtConsulta & "Fecha Hasta   : " & rs!fechaHasta & vbCrLf
    ventasDetalleConsultaFrm.txtConsulta = ventasDetalleConsultaFrm.txtConsulta & "Fecha Entrega : " & rs!fechaentregaCli & vbCrLf & vbCrLf
  
    'avanzo puntero
    rs.MoveNext
    
  Wend
  
  'muestro frm
  ventasDetalleConsultaFrm.Show vbModal
  
  'cierro rs
  rs.Close

End Sub

Private Sub tblOperaciones_ButtonClick(ByVal Button As MSComctlLib.Button)

Dim rs As ADODB.Recordset

  blnRefresh = False
  
  Select Case Button.Key
  
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

  ' hago un refresh de los datos para ver los cambios
  If blnRefresh Then
    
    If strWhere = "" Then
      
      'muestro solo ultimo periodo ingresado
      Dim strRangoIso As isoLastPeriod
      strRangoIso = adoLastPeriod("ViewVentas", "fecha")
      strWhere = "fecha between " & strRangoIso.strDesde & " and " & strRangoIso.strHasta
    
    End If
     
    strSQL = "SELECT * FROM ViewVentas" & _
              IIf(Not strWhere = "", " WHERE " & strWhere, "") & " order by fecha desc"
    intRes = ListViewRefresh(lvwDatos, strSQL)
    intRes = lvwHideColumn(lvwDatos, "concepto")
    intRes = lvwHideColumn(lvwDatos, "condicion")
  
  End If

End Sub


Private Sub tblOperaciones_ButtonMenuClick(ByVal ButtonMenu As MSComctlLib.ButtonMenu)
    
    
    Dim rptFRM As ReportesFRM
    Dim strT As String
    Dim strBorrador As String


        'DECLARE variables para generaciòn - impresiòn de comprobante
        Dim dblTotal As Double
        Dim strLetras As String
        Dim strFormula As String



  '-------------------------------------------------------------------------------------------------------------------------
  '27/04/2007, nuevo, vamos a ver como lo imprementamos.
  '
  'get rs con permisos con grupo y opcion
  Set rsPER = adoGetRS("select * from menuPermis where IDgrupo = '" & SQLparam.Usuario & "' and IDopcion = '" & ButtonMenu.Key & "'")
  
  'check si rs EOF, exit con mensaje sin permisos
  If rsPER.EOF Then
    intRes = MsgBox("No tiene permisos para ingresar a la opción seleccionada.", vbCritical + vbApplicationModal, "Atención...")
    Exit Sub
  End If
    
  'close rs
  rsPER.Close
  '
  '-------------------------------------------------------------------------------------------------------------------------


  Select Case ButtonMenu.Key
    
    
  '-------------------------------------------------------------------------------------------------------------------------
  ' 29/04/2015 - Mazzu
  '
  ' Generaciòn de Comprobante en PDF, x medio de Reporting Services
  '
  'get rs con permisos con grupo y opcion
  Case Is = "comprobante_PDF"
    
    'CHECK si tengo un item seleccionado
    If lvwDatos.SelectedItem.Selected Then
            


          'GET total y convierto a letras
          dblTotal = lvwGetValue(lvwDatos, "total")
          
          '------------------------------------------------------------------------------
          '20/07/2015 - Mazzu - Deshabilitado

          strLetras = NumToLeyend(dblTotal)

          'convierto primera letra a mayuscula
          strLetras = UCase(Left(strLetras, 1)) & Right(strLetras, Len(strLetras) - 1)

          'Dim strMoneda As String
          
          If lvwGetValue(lvwDatos, "moneda") = "$" Then
            strMoneda = "Son pesos: "
            ElseIf lvwGetValue(lvwDatos, "moneda") = "u$s" Then
                  strMoneda = "Son dolares estadounidenses: "
          End If
          '------------------------------------------------------------------------------
          
          'creo nuevo form, toma propiedades, muestro
          'Set rptFRM = New ReportesFRM
          
          'abre comprobante segun Exportacion o cualquier otro
          If lvwGetValue(lvwDatos, "tipo") = "Exp" Then
            
            'rptFRM.DataReportName = App.Path & "\reportes\comprobantesOilExportacion_PDF.rpt"
          
                    'SEND mensaje aviso, que este comprobante se hace mediante AFIP - Comprobantes en Linea
                    strT = "Las exportacions, se deben realizar mediante la web de la AFIP."
                    intRes = MsgBox(strT, vbQuestion + vbYesNo, "Confirmacón")
          
          '
          '
          '
          ElseIf lvwGetValue(lvwDatos, "tipo") = "Fac" And lvwGetValue(lvwDatos, "oper") = "Oil" Then
            
                    '   29/04/2015 - Mazzu
                    '   Abre explorador, apuntando a SSRS, con URL del comprobante a imprimir
                    '
                    ShellExecute 0, vbNullString, _
                    "http://applications.sinopecarg.com.ar/APP/_vti_bin/reportserver?http://applications.sinopecarg.com.ar/APP/RSReports/Ventas/Comprobante_OIL.rdl" + _
                    "&Comprobante=" + lvwGetValue(lvwDatos, "Factura") + "&importeLetras=" + strMoneda + strLetras + _
                    "&Copia=ORIGINAL" + _
                    "&rs:Format=PDF&rs:Command=Render", _
                    vbNullString, vbNullString, vbNormalFocus
          
                    Sleep 2500
                    
                    ShellExecute 0, vbNullString, _
                    "http://applications.sinopecarg.com.ar/APP/_vti_bin/reportserver?http://applications.sinopecarg.com.ar/APP/RSReports/Ventas/Comprobante_OIL.rdl" + _
                     "&Comprobante=" + lvwGetValue(lvwDatos, "Factura") + "&importeLetras=" + strMoneda + strLetras + _
                    "&Copia=DUPLICADO" + _
                    "&rs:Format=PDF&rs:Command=Render", _
                    vbNullString, vbNullString, vbNormalFocus
          
                    Sleep 2500
          
                    ShellExecute 0, vbNullString, _
                    "http://applications.sinopecarg.com.ar/APP/_vti_bin/reportserver?http://applications.sinopecarg.com.ar/APP/RSReports/Ventas/Comprobante_OIL.rdl" + _
                     "&Comprobante=" + lvwGetValue(lvwDatos, "Factura") + "&importeLetras=" + strMoneda + strLetras + _
                    "&Copia=TRIPLICADO" + _
                    "&rs:Format=PDF&rs:Command=Render", _
                    vbNullString, vbNullString, vbNormalFocus
         
        
          '
          '
          '
          ElseIf (lvwGetValue(lvwDatos, "tipo") = "Cre" Or lvwGetValue(lvwDatos, "tipo") = "Deb") And lvwGetValue(lvwDatos, "oper") = "Oil" Then
          
                    '   21/07/2015 - Mazzu
                    '   Abre explorador, apuntando a SSRS, con URL del comprobante a imprimir
                    '
                    ShellExecute 0, vbNullString, _
                    "http://applications.sinopecarg.com.ar/APP/_vti_bin/reportserver?http://applications.sinopecarg.com.ar/APP/RSReports/Ventas/Comprobante_DEB_CRE.rdl" + _
                    "&Comprobante=" + lvwGetValue(lvwDatos, "Factura") + "&importeLetras=" + strMoneda + strLetras + _
                    "&Copia=ORIGINAL" + _
                    "&rs:Format=PDF&rs:Command=Render", _
                    vbNullString, vbNullString, vbNormalFocus
          
                    Sleep 2500
                    
                    ShellExecute 0, vbNullString, _
                    "http://applications.sinopecarg.com.ar/APP/_vti_bin/reportserver?http://applications.sinopecarg.com.ar/APP/RSReports/Ventas/Comprobante_DEB_CRE.rdl" + _
                     "&Comprobante=" + lvwGetValue(lvwDatos, "Factura") + "&importeLetras=" + strMoneda + strLetras + _
                    "&Copia=DUPLICADO" + _
                    "&rs:Format=PDF&rs:Command=Render", _
                    vbNullString, vbNullString, vbNormalFocus
          
                    Sleep 2500
          
                    ShellExecute 0, vbNullString, _
                    "http://applications.sinopecarg.com.ar/APP/_vti_bin/reportserver?http://applications.sinopecarg.com.ar/APP/RSReports/Ventas/Comprobante_DEB_CRE.rdl" + _
                     "&Comprobante=" + lvwGetValue(lvwDatos, "Factura") + "&importeLetras=" + strMoneda + strLetras + _
                    "&Copia=TRIPLICADO" + _
                    "&rs:Format=PDF&rs:Command=Render", _
                    vbNullString, vbNullString, vbNormalFocus
          
          
          '
          '
          '
          ElseIf lvwGetValue(lvwDatos, "tipo") = "Fac" And lvwGetValue(lvwDatos, "oper") = "Gas" And lvwGetValue(lvwDatos, "moneda") = "$" Then
          
                
                    '   29/04/2015 - Mazzu
                    '   Abre explorador, apuntando a SSRS, con URL del comprobante a imprimir
                    '
                    ShellExecute 0, vbNullString, _
                    "http://applicationsqa.sinopecarg.com.ar/APP/_vti_bin/reportserver?http://applicationsqa.sinopecarg.com.ar/APP/RSReports/Ventas/Comprobante_GAS_ARS.rdl" + _
                     "&Comprobante=" + lvwGetValue(lvwDatos, "Factura") + "&importeLetras=" + strMoneda + strLetras + _
                    "&Copia=ORIGINAL" + _
                    "&rs:Format=PDF&rs:Command=Render", _
                    vbNullString, vbNullString, vbNormalFocus
          
                    Sleep 2500
                    
                    ShellExecute 0, vbNullString, _
                    "http://applicationsqa.sinopecarg.com.ar/APP/_vti_bin/reportserver?http://applicationsqa.sinopecarg.com.ar/APP/RSReports/Ventas/Comprobante_GAS_ARS.rdl" + _
                     "&Comprobante=" + lvwGetValue(lvwDatos, "Factura") + "&importeLetras=" + strMoneda + strLetras + _
                    "&Copia=DUPLICADO" + _
                    "&rs:Format=PDF&rs:Command=Render", _
                    vbNullString, vbNullString, vbNormalFocus
          
                    Sleep 2500
          
                    ShellExecute 0, vbNullString, _
                    "http://applicationsqa.sinopecarg.com.ar/APP/_vti_bin/reportserver?http://applicationsqa.sinopecarg.com.ar/APP/RSReports/Ventas/Comprobante_GAS_ARS.rdl" + _
                     "&Comprobante=" + lvwGetValue(lvwDatos, "Factura") + "&importeLetras=" + strMoneda + strLetras + _
                    "&Copia=TRIPLICADO" + _
                    "&rs:Format=PDF&rs:Command=Render", _
                    vbNullString, vbNullString, vbNormalFocus
                
                
                
          ElseIf lvwGetValue(lvwDatos, "tipo") = "Fac" And lvwGetValue(lvwDatos, "oper") = "Gas" And lvwGetValue(lvwDatos, "moneda") = "u$s" Then
                
                    '   29/04/2015 - Mazzu
                    '   Abre explorador, apuntando a SSRS, con URL del comprobante a imprimir
                    '
                    ShellExecute 0, vbNullString, _
                    "http://applicationsqa.sinopecarg.com.ar/APP/_vti_bin/reportserver?http://applicationsqa.sinopecarg.com.ar/APP/RSReports/Ventas/Comprobante_GAS_USD.rdl" + _
                     "&Comprobante=" + lvwGetValue(lvwDatos, "Factura") + "&importeLetras=" + strMoneda + strLetras + _
                    "&Copia=ORIGINAL" + _
                    "&rs:Format=PDF&rs:Command=Render", _
                    vbNullString, vbNullString, vbNormalFocus
          
                    Sleep 2500
                    
                    ShellExecute 0, vbNullString, _
                    "http://applicationsqa.sinopecarg.com.ar/APP/_vti_bin/reportserver?http://applicationsqa.sinopecarg.com.ar/APP/RSReports/Ventas/Comprobante_GAS_USD.rdl" + _
                     "&Comprobante=" + lvwGetValue(lvwDatos, "Factura") + "&importeLetras=" + strMoneda + strLetras + _
                    "&Copia=DUPLICADO" + _
                    "&rs:Format=PDF&rs:Command=Render", _
                    vbNullString, vbNullString, vbNormalFocus
          
                    Sleep 2500
          
                    ShellExecute 0, vbNullString, _
                    "http://applicationsqa.sinopecarg.com.ar/APP/_vti_bin/reportserver?http://applicationsqa.sinopecarg.com.ar/APP/RSReports/Ventas/Comprobante_GAS_USD.rdl" + _
                     "&Comprobante=" + lvwGetValue(lvwDatos, "Factura") + "&importeLetras=" + strMoneda + strLetras + _
                    "&Copia=TRIPLICADO" + _
                    "&rs:Format=PDF&rs:Command=Render", _
                    vbNullString, vbNullString, vbNormalFocus
              
          End If
            
          'armo string con formulas
          'strFormula = "conceptoGeneral;" & lvwGetValue(lvwDatos, "conceptoventa") & ";" & _
                       "strTituloImporte;" & "'" & strMoneda & "'" & ";" & _
                       "strImporteLetras;" & "'" & strLetras & "'" & ";" & _
                       "strBorrador;" & "'" & strBorrador & "'"
          
          'cambio propiedades
          'rptFRM.DataTitulo = "Impresion de Comprobante Numero " & lvwGetValue(lvwDatos, "factura")
          'rptFRM.DataSource = "SELECT * FROM viewVentasImprime"
          'rptFRM.DataWhere = "factura = '" & lvwGetValue(lvwDatos, "factura") & "'"
          'rptFRM.DataFormula = strFormula
          
          'rptFRM.Show vbModal
                    
                    
        '08/05/2015
        'DESHABILITADO - Al ser un PDF, no necesita confirmaciòn de que se imprimio correctamente
        
            'comprobante se imprimio correctamente?
            'intRes = MsgBox("Haga clic en Si, una vez generado correctamente el PDF, para que el comprobante quede marcado como emitido.", vbQuestion + vbYesNo, "atención...")
            
            'marco comprobante como Emitido
            'If intRes = vbYes Then
              
              
'               21/07/2015
'               Emazzu      - Deshabilito porque no quiero que la marque como emitida.
'
'              'marco en la tabla
'              strSQL = "exec spVentasUpdate " & _
'              "'" & lvwGetValue(lvwDatos, "empresa") & "'," & _
'              "'" & lvwGetValue(lvwDatos, "factura") & "'"
'
'              intRes = adoExecSQL(strSQL)
'
'              'chequeo errores
'              If Not lngAdoErrNum = -1 Then
'                adoError
'                Exit Sub
'              End If
'
'              'marco en listView en pantalla
'              intRes = lvwSetValue(lvwDatos, "Emitida", True)
              
            'End If
                    
                    
    Else
      intRes = MsgBox("No hay ningun item seleccionado.", vbApplicationModal + vbOKOnly + vbInformation, "Informaciòn")
    End If
    
    
    
    
  Case Is = "comprobante", "comprobanteBorrador"
    
    'armo la leyenda en pantalla
    If lvwDatos.SelectedItem.Selected Then
            
      'chequeo que no haya sido emitida en forma definitiva
      If ButtonMenu.Key = "comprobante" And lvwGetValue(lvwDatos, "Emitida") = "True" Then
        intRes = MsgBox("El comprobante ya fue emitido en forma definitiva.", vbCritical + vbOKOnly, "atención...")
        Exit Sub
      End If
            
      If ButtonMenu.Key = "comprobante" Then
        strT = "Esta seguro que desea imprimir el comprobante: " & lvwGetValue(lvwDatos, "factura") & " Definitivamente."
      Else
        strT = "Esta seguro que desea imprimir el comprobante: " & lvwGetValue(lvwDatos, "factura") & " como Borrador."
        strBorrador = "Si"
      End If
      
      'pregunto
      intRes = MsgBox(strT, vbQuestion + vbYesNo, "Confirmacón")
      
      If intRes = vbYes Then
    
    
          'GET total y convierto a letras
          dblTotal = lvwGetValue(lvwDatos, "total")
          
          '------------------------------------------------------------------------------
          '29/04/2015 - Mazzu - Deshabilitado

          strLetras = NumToLeyend(dblTotal)

          'convierto primera letra a mayuscula
          strLetras = UCase(Left(strLetras, 1)) & Right(strLetras, Len(strLetras) - 1)

          'Dim strMoneda As String
          
          If lvwGetValue(lvwDatos, "moneda") = "$" Then
            strMoneda = "Son pesos: "
            ElseIf lvwGetValue(lvwDatos, "moneda") = "u$s" Then
                  strMoneda = "Son dólares estadounidenses: "
          End If
          '------------------------------------------------------------------------------
          
          
        'creo nuevo form, toma propiedades, muestro
        Set rptFRM = New ReportesFRM
        
        'abre comprobante segun Exportacion o cualquier otro
        If lvwGetValue(lvwDatos, "tipo") = "Exp" Then
          
          rptFRM.DataReportName = App.Path & "\reportes\comprobantesOilExportacion.rpt"
        
        ElseIf lvwGetValue(lvwDatos, "tipo") = "Fac" And lvwGetValue(lvwDatos, "oper") = "Oil" Then
          
          rptFRM.DataReportName = App.Path & "\reportes\comprobantesOil.rpt"
            
            Else
              
              rptFRM.DataReportName = App.Path & "\reportes\comprobantesVarios.rpt"
            
            End If
          
        'armo string con formulas
        strFormula = "conceptoGeneral;" & lvwGetValue(lvwDatos, "conceptoventa") & ";" & _
                     "strTituloImporte;" & "'" & strMoneda & "'" & ";" & _
                     "strImporteLetras;" & "'" & strLetras & "'" & ";" & _
                     "strBorrador;" & "'" & strBorrador & "'"
        
        'cambio propiedades
        rptFRM.DataTitulo = "Impresion de Comprobante Numero " & lvwGetValue(lvwDatos, "factura")
        rptFRM.DataSource = "SELECT * FROM viewVentasImprime"
        rptFRM.DataWhere = "factura = '" & lvwGetValue(lvwDatos, "factura") & "'"
        rptFRM.DataFormula = strFormula
        
        rptFRM.Show vbModal
                  
        'si es comprobante definitivo
        If strBorrador <> "Si" Then
          
          'comprobante se imprimio correctamente?
          intRes = MsgBox("El comprobante se imprimio correctamente.", vbQuestion + vbYesNo, "atención...")
          
          'marco comprobante como Emitido
          If intRes = vbYes Then
            
            'marco en la tabla
            strSQL = "exec spVentasUpdate " & _
            "'" & lvwGetValue(lvwDatos, "empresa") & "'," & _
            "'" & lvwGetValue(lvwDatos, "factura") & "'"
            
            intRes = adoExecSQL(strSQL)
              
            'chequeo errores
            If Not lngAdoErrNum = -1 Then
              adoError
              Exit Sub
            End If
            
            'marco en listView en pantalla
            intRes = lvwSetValue(lvwDatos, "Emitida", True)
            
          End If
          
        End If
                  
      End If
      
    Else
      intRes = MsgBox("No hay ningun item seleccionado.", vbApplicationModal + vbOKOnly + vbInformation, "Informacion")
    End If
    
  Case Is = "soporte"
    
    If lvwDatos.SelectedItem.Selected Then
      
      intRes = MsgBox("Esta seguro que desea imprimir el soporte seleccionado: " & lvwGetValue(lvwDatos, "factura"), vbQuestion + vbYesNo, "Confirmacón")
      
      If intRes = vbYes Then
          
        'creo nuevo form, toma propiedades, muestro
        Set rptFRM = New ReportesFRM
        rptFRM.DataReportName = App.Path & "\reportes\comprobantesOilSoporte.rpt"
         
        'cambio propiedades
        rptFRM.DataTitulo = "Impresion de Soporte Numero " & lvwGetValue(lvwDatos, "factura")
        rptFRM.DataSource = "SELECT * FROM viewVentasImprime"
        rptFRM.DataWhere = "factura = '" & lvwGetValue(lvwDatos, "factura") & "'"
        rptFRM.Show
          
      End If
      
    Else
      intRes = MsgBox("No hay ningun item seleccionado.", vbApplicationModal + vbOKOnly + vbInformation, "Informacion")
    End If
    
    
  Case Is = "FacOil"
    
    ' cargo formulario
    Load frmVentasInfo
    
    ' muestro formulario
    frmVentasInfo.Show vbModal
      
    ' si hizo click en Aceptar, en el formulario frmxxxxxInfo
    ' se pone en true la variable global blnAceptar
    If blnAceptar Then
      blnRefresh = True
    End If
      
    ' descargo formulario
    Unload frmVentasInfo
    
  Case Is = "OilDebCre"
    
    ' cargo formulario
    Load frmDebitoCreditoOIL
    
    ' muestro formulario
    frmDebitoCreditoOIL.Show vbModal
      
    ' si hizo click en Aceptar, en el formulario frmxxxxxInfo
    ' se pone en true la variable global blnAceptar
    If blnAceptar Then
      blnRefresh = True
    End If
      
    ' descargo formulario
    Unload frmDebitoCreditoOIL
    
  Case Is = "VarDebCre"
    
    ' cargo formulario
    Load frmDebitoCredito
    
    ' muestro formulario
    frmDebitoCredito.Show vbModal
      
    ' si hizo click en Aceptar, en el formulario frmxxxxxInfo
    ' se pone en true la variable global blnAceptar
    If blnAceptar Then
      blnRefresh = True
    End If
      
    ' descargo formulario
    Unload frmDebitoCredito
      
      
  Case Is = "autorizar"
          
    'CHECK si se selecciono un comprobante
    If lvwDatos Is Nothing Then
      intRes = MsgBox("No hay ningun comprobante seleccionado.", vbApplicationModal + vbOKOnly + vbInformation, "Informacion")
      Exit Sub
    End If
    
    'CHECK si emitida esta en True
    If lvwGetValue(lvwDatos, "emitida") = "True" Then
      intRes = MsgBox("El comprobante esta marcado como True, ya fue enviado a autorizar con anterioridad.", vbCritical + vbOKOnly, "atención...")
      Exit Sub
    End If
    
    intRes = MsgBox("Esta seguro, que desea enviar al AFIP para autorizar el comprobante?", vbApplicationModal + vbQuestion + vbYesNo, "Autorizar comprobante...")

    'CHECK si confirmo, enviar a autorizar
    If intRes = vbYes Then
        
        'marco en la tabla
        strSQL = "exec spVentasUpdate '" & lvwGetValue(lvwDatos, "empresa") & "','" & lvwGetValue(lvwDatos, "factura") & "'"
        
        intRes = adoExecSQL(strSQL)
        blnRefresh = True
        
        'chequeo errores
        If Not lngAdoErrNum = -1 Then
        adoError
        Exit Sub
        End If
        
        'marco en listView en pantalla
        intRes = lvwSetValue(lvwDatos, "Emitida", True)
        
    End If
      
         
'   08/05/2015
'   Edu Mazzu   -   Deshabilitado, un comprobante ya no se puede anular


'  Case Is = "anular"
'
'    If lvwDatos Is Nothing Then
'      intRes = MsgBox("No hay ningun comprobante seleccionado.", vbApplicationModal + vbOKOnly + vbInformation, "Informacion")
'      Exit Sub
'    End If
'
'    If Val(lvwGetValue(lvwDatos, "status")) = 2 Then
'      intRes = MsgBox("No puede anular un comprobante correspondiente a un período cerrado.", vbApplicationModal + vbOKOnly + vbInformation, "Informacion")
'      Exit Sub
'    End If
'
'    'chequeo que se haya emitido en forma definitiva
'    If lvwGetValue(lvwDatos, "emitida") = "" Then
'      intRes = MsgBox("No puede anular un comprobante que todavia no fue emitido en forma definitivo.", vbCritical + vbOKOnly, "atención...")
'      Exit Sub
'    End If
'
'    If Val(lvwGetValue(lvwDatos, "Total")) = 0 Then
'      intRes = MsgBox("El comprobante ya fue anulado.", vbApplicationModal + vbOKOnly + vbInformation, "Informacion")
'      Exit Sub
'    End If
'
'    intRes = MsgBox("Esta seguro?", vbApplicationModal + vbQuestion + vbYesNo, "anular comprobante...")
'
'    If intRes = vbYes Then
'
'      strSQL = "exec ventasAnulaComprobanteSP '" & lvwGetValue(lvwDatos, "factura") & "'"
'      intRes = adoExecSQL(strSQL)
'      blnRefresh = True
'
'      'chequeo errores
'      If Not lngAdoErrNum = -1 Then
'        adoError
'        Exit Sub
'      End If
'
'    End If

  Case Is = "eliminar"
      
      
    If lvwDatos Is Nothing Then
      intRes = MsgBox("No hay ningun comprobante seleccionado.", vbApplicationModal + vbOKOnly + vbInformation, "Informacion")
      Exit Sub
    End If
      
    If Val(lvwGetValue(lvwDatos, "status")) = 2 Then
      intRes = MsgBox("No puede eliminar un comprobante correspondiente a un período cerrado.", vbApplicationModal + vbOKOnly + vbInformation, "Informacion")
      Exit Sub
    End If
    
    
    'chequeo que no haya sido emitida en forma definitiva
    If lvwGetValue(lvwDatos, "CAE") <> "" Then
      intRes = MsgBox("No puede eliminar un comprobante que fue aprobado por AFIP.", vbCritical + vbOKOnly, "atención...")
      Exit Sub
    End If
    
    'chequeo que no haya sido emitida en forma definitiva
    If lvwGetValue(lvwDatos, "emitida") = "True" Then
      intRes = MsgBox("No puede eliminar un comprobante que fue enviado para la atorizacion de AFIP.", vbCritical + vbOKOnly, "atención...")
      Exit Sub
    End If
    
    '   GET - ultimo comprobante emitido
    '
    strSQL = "SELECT * FROM empresas_Puntos_Venta " & _
             "WHERE IDempresa = " & lvwGetValue(lvwDatos, "IDempresa") & _
             " AND " & _
             "Operacion = '" & Mid(lvwGetValue(lvwDatos, "Factura"), 16, 3) & "'" & _
             " AND " & _
             "Comprobante = '" & Mid(lvwGetValue(lvwDatos, "Factura"), 20, 3) & "'"
             
    Set rs = adoGetRS(strSQL)
  
    'CHECK errores
    If Not lngAdoErrNum = -1 Then
      adoError
      Exit Sub
    End If
    
    'CHECK, si encontro algo.
    If Not rs.EOF Then
        
        '   CHECK   - Solo permite eliminar el ùltimo comprobante emitido
        If Mid(lvwGetValue(lvwDatos, "Factura"), 7, 8) <> Format(rs!Numero, "00000000") Then
            
            intRes = MsgBox("Solo se permite eliminar el ultimo comprobante - ultimo numero.", vbCritical + vbOKOnly, "atención...")
            Exit Sub
            
        End If

    End If
    
    
    intRes = MsgBox("Esta seguro, que desea eliminar el comprobante?", vbApplicationModal + vbQuestion + vbYesNo, "eliminar comprobante...")
    If intRes = vbYes Then
        
      strSQL = "exec ventasEliminarComprobanteSP '" & lvwGetValue(lvwDatos, "factura") & "'"
      intRes = adoExecSQL(strSQL)
      blnRefresh = True
        
      'chequeo errores
      If Not lngAdoErrNum = -1 Then
        adoError
        Exit Sub
      End If
       
    End If

  End Select

  ' hago un refresh de los datos para ver los cambios
  If blnRefresh Then
    
    If strWhere = "" Then
      
      'muestro solo ultimo periodo ingresado
      Dim strRangoIso As isoLastPeriod
      strRangoIso = adoLastPeriod("ViewVentas", "fecha")
      strWhere = "fecha between " & strRangoIso.strDesde & " and " & strRangoIso.strHasta
    
    End If
     
    strSQL = "SELECT * FROM ViewVentas" & _
              IIf(Not strWhere = "", " WHERE " & strWhere, "") & _
              " ORDER BY fecha DESC"
    intRes = ListViewRefresh(lvwDatos, strSQL)
    
    intRes = lvwHideColumn(lvwDatos, "concepto")
    intRes = lvwHideColumn(lvwDatos, "condicion")
    intRes = lvwHideColumn(lvwDatos, "status")
'    intRes = lvwHideColumn(lvwDatos, "emitida")
  
  End If

End Sub

