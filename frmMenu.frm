VERSION 5.00
Object = "{BDDD132C-614B-11D3-B85E-85ADB7D07209}#1.0#0"; "dXSBar.dll"
Begin VB.Form frmMenu 
   BorderStyle     =   0  'None
   Caption         =   "Menu"
   ClientHeight    =   6375
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   1665
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6375
   ScaleWidth      =   1665
   ShowInTaskbar   =   0   'False
   Begin DXSIDEBARLibCtl.dxSideBar dxSideBar1 
      Height          =   6300
      Left            =   30
      OleObjectBlob   =   "frmMenu.frx":0000
      TabIndex        =   0
      Top             =   45
      Width           =   1575
   End
End
Attribute VB_Name = "frmMenu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub dxSideBar1_OnClickItemLink(ByVal pGroup As DXSIDEBARLibCtl.IdxGroup, ByVal pLink As DXSIDEBARLibCtl.IdxItemLink, ByVal GroupIndex As Integer, ByVal ItemLinkIndex As Integer)
  
  Dim intRes As Integer
  Dim rsPER As ADODB.Recordset
  
  '-------------------------------------------------------------------------------------------------------------------------
  '27/04/2007, nuevo, vamos a ver como lo imprementamos.
  '
  'get rs con permisos con grupo y opcion
  Set rsPER = adoGetRS("select * from menuPermis where IDgrupo = '" & SQLparam.Usuario & "' and IDopcion = '" & pLink.item.Caption & "'")
  
  'check si rs EOF, exit con mensaje sin permisos
  If rsPER.EOF Then
    intRes = MsgBox("No tiene permisos para ingresar a la opción seleccionada.", vbCritical + vbApplicationModal, "Atención...")
    Exit Sub
  End If
    
  'close rs
  rsPER.Close
  '
  '-------------------------------------------------------------------------------------------------------------------------

  'check opcion
  Select Case Format(pLink.item.Caption, "<")

  Case Is = "contratos"
    
    ' muestra formulario
    strTableNameActual = "contratos"
    intRes = frmToShow(frmMenu, frmActivo, frmContratos, True)
    
  Case Is = "entregas clientes"
    
    ' muestra formulario
    strTableNameActual = "entregascli"
    intRes = frmToShow(frmMenu, frmActivo, frmEntregasCli, True)
  
  Case Is = "precios"
    
    ' muestra formulario
    strTableNameActual = "precios"
    intRes = frmToShow(frmMenu, frmActivo, frmPrecios, True)
  
  Case Is = "cotizacion dolar"
    
    ' muestra formulario
    strTableNameActual = "cotizacionDolar"
    intRes = frmToShow(frmMenu, frmActivo, frmCotizacionDolar, True)
  
  Case Is = "entregas transportistas"
    
    ' muestra formulario
    strTableNameActual = "entregastra"
    intRes = frmToShow(frmMenu, frmActivo, frmEntregasTra)
  
  Case Is = "entregas terminales"
    
    ' muestra formulario
    strTableNameActual = "entregaster"
    intRes = frmToShow(frmMenu, frmActivo, frmEntregasTer)
  
  Case Is = "subconcesiones"
    
    ' muestra formulario
    strTableNameActual = "subconcesiones"
    intRes = frmToShow(frmMenu, frmActivo, frmSubConcesiones, True)
  
  '-----------------------------------------------------------------
  ' parametros
  
  Case Is = "generales"
    
    ' muestra formulario
    strTableNameActual = "parametros"
    intRes = frmToShow(frmMenu, frmActivo, frmParametros, True)
  
  Case Is = "porsubconcesion"
    
    ' muestra formulario
    strTableNameActual = "porSubconcesiones"
    intRes = frmToShow(frmMenu, frmActivo, frmSubconcesionParam, True)
  
  Case Is = "porterminal"
    
    ' muestra formulario
    strTableNameActual = "porTerminalales"
    intRes = frmToShow(frmMenu, frmActivo, frmTerminalParam, True)
  
  '-----------------------------------------------------------------
  
  Case Is = "comprobantes oil"
    
    ' muestra formulario
    strTableNameActual = "comprobantesOil"
    intRes = frmToShow(frmMenu, frmActivo, frmVentas, True)
 
  Case Is = "produccion mensual"
    
    ' muestra formulario
    strTableNameActual = "produccionmensual"
    intRes = frmToShow(frmMenu, frmActivo, frmProduccionMensual, True)
 
  Case Is = "stock subconcesiones"
    
    ' muestra formulario
    strTableNameActual = "stocksubconcesiones"
    intRes = frmToShow(frmMenu, frmActivo, frmStockSubconcesiones, True)
  
  Case Is = "stock terminales"
    
    ' muestra formulario
    strTableNameActual = "stockterminales"
    intRes = frmToShow(frmMenu, frmActivo, frmStockTerminales, True)
 
  Case Is = "valorizacion stock"
    
    ' muestra formulario
    strTableNameActual = "valorizacionstock"
    intRes = frmToShow(frmMenu, frmActivo, frmValorizacionStock, True)
 
  Case Is = "cierre stock"
    
    ' muestra formulario
    strTableNameActual = "cierrestock"
    intRes = frmToShow(frmMenu, frmActivo, frmCierreStock, True)
 
  Case Is = "rg 1361"
        
    'show form
    rg1361.Show vbModal
 
  '**********************************************************************************
  ' REPORTES
  '**********************************************************************************
 
  Case Is = "entregas a clientes"
    
    'creo nuevo form, toma propiedades, muestro
    Set rptFRM = New ReportesFRM
    rptFRM.DataIDReporte = "rpt02"
    rptFRM.DataTitulo = "Entregas a Clientes"
    rptFRM.DataReportName = App.Path & "\reportes\entregasClientes.rpt"
    rptFRM.DataSource = "SELECT * FROM entregasClientes_vw_rpt"
    rptFRM.Show
 
  Case Is = "ventas totales"
    
    'creo nuevo form, toma propiedades, muestro
    Set rptFRM = New ReportesFRM
    rptFRM.DataIDReporte = "rpt05"
    rptFRM.DataTitulo = "Ventas Totales"
    rptFRM.DataReportName = App.Path & "\reportes\ventasTotales.rpt"
    rptFRM.DataSource = "SELECT * FROM ventas_totales1_vw"
    rptFRM.Show
 
  Case Is = "ventas por yacimiento"
    
    'creo nuevo form, toma propiedades, muestro
    Set rptFRM = New ReportesFRM
    rptFRM.DataIDReporte = "rpt06"
    rptFRM.DataTitulo = "Ventas por Yacimiento"
    rptFRM.DataReportName = App.Path & "\reportes\ventasXyacimiento.rpt"
    rptFRM.DataSource = "SELECT * FROM ventasXyacimiento_vw_rpt"
    rptFRM.Show
 
  Case Is = "precios por día"
    
    'creo nuevo form, toma propiedades, muestro
    Set rptFRM = New ReportesFRM
    rptFRM.DataIDReporte = "rpt07"
    rptFRM.DataTitulo = "Precios por Día"
    rptFRM.DataReportName = App.Path & "\reportes\precios_lista.rpt"
    rptFRM.DataSource = "SELECT * FROM ViewPrecios"
    rptFRM.Show
 
  Case Is = "precios promedio mensual"
    
    'creo nuevo form, toma propiedades, muestro
    Set rptFRM = New ReportesFRM
    rptFRM.DataIDReporte = "rpt08"
    rptFRM.DataTitulo = "Precios Promedio Mensual"
    rptFRM.DataReportName = App.Path & "\reportes\precios_avg.rpt"
    rptFRM.DataSource = "SELECT * FROM Precios_avg_vw"
    rptFRM.Show
 
  Case Is = "produccion capitulo iv"
    
    'creo nuevo form, toma propiedades, muestro
    Set rptFRM = New ReportesFRM
    rptFRM.DataIDReporte = "rpt09"
    rptFRM.DataTitulo = "Producción Capítulo IV"
    rptFRM.DataReportName = App.Path & "\reportes\prodXsubconcesion.rpt"
    rptFRM.DataSource = "SELECT * FROM produccionMensual_vw_rpt"
    rptFRM.Show
 
  Case Is = "stock por yacimientos"
    
    'creo nuevo form, toma propiedades, muestro
    Set rptFRM = New ReportesFRM
    rptFRM.DataIDReporte = "rpt10"
    rptFRM.DataTitulo = "Stock por Yacimientos"
    rptFRM.DataReportName = App.Path & "\reportes\stockSubXyacimiento.rpt"
    rptFRM.DataSource = "SELECT * FROM stockSubXyacimiento_vw_rpt"
    rptFRM.Show
    
  Case Is = "stock por terminales"
    
    'creo nuevo form, toma propiedades, muestro
    Set rptFRM = New ReportesFRM
    rptFRM.DataIDReporte = "rpt11"
    rptFRM.DataTitulo = "Stock por Terminales"
    rptFRM.DataReportName = App.Path & "\reportes\stockTerXyacimiento.rpt"
    rptFRM.DataSource = "SELECT * FROM stockTerXyacimiento_vw_rpt"
    rptFRM.Show
    
  Case Is = "valorizacion de stock"
    
    'creo nuevo form, toma propiedades, muestro
    Set rptFRM = New ReportesFRM
    rptFRM.DataIDReporte = "rpt12"
    rptFRM.DataTitulo = "Valorización de Stock"
    rptFRM.DataReportName = App.Path & "\reportes\valorizacionStockXyacimiento.rpt"
    rptFRM.DataSource = "SELECT * FROM valorizacionStock_xSub_xArea_vw_rpt"
    rptFRM.Show
    
  Case Is = "entrega a transportistas"
    
    'creo nuevo form, toma propiedades, muestro
    Set rptFRM = New ReportesFRM
    rptFRM.DataIDReporte = "rpt13"
    rptFRM.DataTitulo = "Entrega a Transportistas por Concesión"
    rptFRM.DataReportName = App.Path & "\reportes\entregasTransporte.rpt"
    rptFRM.DataSource = "SELECT * FROM entregasTra_vw_rpt"
    rptFRM.Show
    
  Case Is = "entrega a transportistas por concesion"
    
    'creo nuevo form, toma propiedades, muestro
    Set rptFRM = New ReportesFRM
    rptFRM.DataIDReporte = "rpt14"
    rptFRM.DataTitulo = "Entrega a Transportistas por Concesión"
    rptFRM.DataReportName = App.Path & "\reportes\entregasTransporte_Con.rpt"
    rptFRM.DataSource = "SELECT * FROM entregasTra_Con_vw_rpt"
    rptFRM.Show
    
  Case Is = "entrega a transportistas por subconcesion"
    
    'creo nuevo form, toma propiedades, muestro
    Set rptFRM = New ReportesFRM
    rptFRM.DataIDReporte = "rpt15"
    rptFRM.DataTitulo = "Entrega a Transportistas por Subconcesión"
    rptFRM.DataReportName = App.Path & "\reportes\entregasTransporte_Sub.rpt"
    rptFRM.DataSource = "SELECT * FROM entregasTra_Sub_vw_rpt"
    rptFRM.Show
    
  Case Is = "entregas a terminales por concesion"
    
    'creo nuevo form, toma propiedades, muestro
    Set rptFRM = New ReportesFRM
    rptFRM.DataIDReporte = "rpt16"
    rptFRM.DataTitulo = "Entregas a Terminales por Concesión"
    rptFRM.DataReportName = App.Path & "\reportes\entregas_terminales_con.rpt"
    rptFRM.DataSource = "SELECT * FROM entregas_terminales_con_vw"
    rptFRM.Show
    
  Case Is = "entregas a terminales por subconcesion"
    
    'creo nuevo form, toma propiedades, muestro
    Set rptFRM = New ReportesFRM
    rptFRM.DataIDReporte = "rpt17"
    rptFRM.DataTitulo = "Entregas a Terminales por SubConcesión"
    rptFRM.DataReportName = App.Path & "\reportes\entregas_terminales_sub.rpt"
    rptFRM.DataSource = "SELECT * FROM entregas_terminales_sub_vw"
    rptFRM.Show
    
  Case Is = "ventas de petroleo"
    
    'creo nuevo form, toma propiedades, muestro
    Set rptFRM = New ReportesFRM
    rptFRM.DataIDReporte = "rpt18"
    rptFRM.DataTitulo = "Ventas de Petroleo"
    rptFRM.DataReportName = App.Path & "\reportes\ventasPetroleo.rpt"
    rptFRM.DataSource = "SELECT * FROM ventas_petroleo_vw"
    rptFRM.Show
    
  Case Is = "ventas de gas"
    
    'creo nuevo form, toma propiedades, muestro
    Set rptFRM = New ReportesFRM
    rptFRM.DataIDReporte = "rpt19"
    rptFRM.DataTitulo = "Ventas de Gas"
    rptFRM.DataReportName = App.Path & "\reportes\ventas_Gas.rpt"
    rptFRM.DataSource = "SELECT * FROM ventas_gas_vw"
    rptFRM.Show
    
  Case Is = "ventas de varios"
    
    'creo nuevo form, toma propiedades, muestro
    Set rptFRM = New ReportesFRM
    rptFRM.DataIDReporte = "rpt20"
    rptFRM.DataTitulo = "Ventas de Varios"
    rptFRM.DataReportName = App.Path & "\reportes\ventas_varios.rpt"
    rptFRM.DataSource = "SELECT * FROM ventas_varios_vw"
    rptFRM.Show
    
  Case Is = "seguimiento de contratos"
    
    'creo nuevo form, toma propiedades, muestro
    Set rptFRM = New ReportesFRM
    rptFRM.DataIDReporte = "rpt21"
    rptFRM.DataTitulo = "Seguimiento de Contratos"
    rptFRM.DataReportName = App.Path & "\reportes\contratos_seguimiento.rpt"
    rptFRM.DataSource = "SELECT * FROM contratos_seguimiento_vw"
    rptFRM.Show
    
  Case Is = "transporte de crudo por consecion"
    
    'creo nuevo form, toma propiedades, muestro
    Set rptFRM = New ReportesFRM
    rptFRM.DataIDReporte = "rpt22"
    rptFRM.DataTitulo = "transporte de crudo por consecion"
    rptFRM.DataReportName = App.Path & "\reportes\transporte_Crudo_con.rpt"
    rptFRM.DataSource = "SELECT * FROM transporteCrudo_vw_rpt"
    rptFRM.Show
    
  Case Is = "transporte de crudo por subconcesion"
    
    'creo nuevo form, toma propiedades, muestro
    Set rptFRM = New ReportesFRM
    rptFRM.DataIDReporte = "rpt23"
    rptFRM.DataTitulo = "trnasporte de crudo por subconcesion"
    rptFRM.DataReportName = App.Path & "\reportes\transporte_Crudo_sub.rpt"
    rptFRM.DataSource = "SELECT * FROM transporteCrudo_vw_rpt"
    rptFRM.Show
    
  Case Is = "clientes"
    
    ' muestra formulario
    strTableNameActual = "Clientes"
    intRes = frmToShow(frmMenu, frmActivo, frmClientes, True)
    
  Case Is = "comprobantes"
    
    ' muestra formulario
    strTableNameActual = "Comprobantes"
    intRes = frmToShow(frmMenu, frmActivo, frmEmpresasComprobantes, True)
    
  End Select
  
End Sub

Private Sub Form_Load()
    
  Set frmActivo = frmEntregasCli    ' por primera vez fuerzo frm activo
    
  strTableNameActual = "entregascli"
  intRes = frmToShow(frmMenu, frmActivo, frmEntregasCli, True)
  
End Sub
