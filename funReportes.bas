Attribute VB_Name = "funReportes"
'
'SIRVE PARA ESPECIFICAR CARACTERISTICAS ADICIONALES A LOS REPORTES
'
'strNombre:   nombre del reporte que aparece en el mnu de opciones
'rpt:         referencia al reporte principal
'rptSub:      referencia a subreportes
'strWhere:    string con where
'strWhereArr: array con where separado para poder acceder a cada condicion por separado
'
Public Function reportesFRMOtros(strIDReporte As String, rpt As CRAXDRT.Report, strWhere As String, strWhereArr As Variant) As Boolean
  Dim rptSub As CRAXDRT.Report
  Dim rs As ADODB.Recordset
  Dim strWhereActual As String
  
  Select Case strIDReporte

  'totales de ventas por yacimiento
  'pasando parametros a los totales por
  'subconcesion y tipo de comprobante
  Case "rpt06"
  
    'armo where
    strWhereActual = ""
    If strWhere <> "" Then
      strWhereActual = " where " & strWhere
    End If
      
    'le tiro rs a subreport
    Set rptSub = rpt.OpenSubreport("totalesXventasLocales")
    strSQL = "SELECT Subconcesion, TipoComprobante, SUM(Mts15) as Mts15, SUM(Mts1556) as Mts1556, SUM(Bbls) as Bbls, SUM(importe) as Importe, sum(pjeSubTerm*Bbls) as PjExBbls, sum(nApiGravity*Bbls) as APIxBbls " & _
             "FROM ventasXyacimiento_vw_rpt " & strWhereActual & " GROUP BY subconcesion,TipoComprobante"
    Set rs = adoGetRS(strSQL)
    rptSub.Database.SetDataSource rs
    
    Set rptSub = rpt.OpenSubreport("totalesXventasExportacion")
    strSQL = "SELECT Subconcesion, TipoComprobante, SUM(Mts15) as Mts15, SUM(Mts1556) as Mts1556, SUM(Bbls) as Bbls, SUM(importe) as Importe, sum(pjeSubTerm*Bbls) as PjExBbls, sum(nApiGravity*Bbls) as APIxBbls " & _
             "FROM ventasXyacimiento_vw_rpt " & strWhereActual & " GROUP BY subconcesion,TipoComprobante"
    Set rs = adoGetRS(strSQL)
    rptSub.Database.SetDataSource rs
    
    Set rptSub = rpt.OpenSubreport("totalesXventasTodas")
    strSQL = "SELECT Subconcesion, SUM(Mts15) as Mts15, SUM(Mts1556) as Mts1556, SUM(Bbls) as Bbls, SUM(importe) as Importe, sum(pjeSubTerm*Bbls) as PjExBbls, sum(nApiGravity*Bbls) as APIxBbls " & _
             "FROM ventasXyacimiento_vw_rpt " & strWhereActual & " GROUP BY Subconcesion"
    Set rs = adoGetRS(strSQL)
    rptSub.Database.SetDataSource rs
  
  'totales produccion capitulo IV x yacimiento y x area
  Case "rpt09"
  
    'armo where
    strWhereActual = ""
    If strWhere <> "" Then
      strWhereActual = " where " & strWhere
    End If
      
    'le tiro rs a subreport
    Set rptSub = rpt.OpenSubreport("prodXarea")
    strSQL = "SELECT * FROM produccionMensual_vw_rpt " & strWhereActual
    Set rs = adoGetRS(strSQL)
    rptSub.Database.SetDataSource rs
  
  'stock Subconcesiones x yacimiento
  Case "rpt10"
  
    'armo where
    strWhereActual = ""
    If strWhere <> "" Then
      strWhereActual = " where " & strWhere
    End If
      
    'le tiro rs a subreport
    Set rptSub = rpt.OpenSubreport("StockXarea")
    strSQL = "SELECT * FROM stockSubXyacimiento_vw_rpt " & strWhereActual
    Set rs = adoGetRS(strSQL)
    rptSub.Database.SetDataSource rs
  
  'stock terminales x yacimiento
  Case "rpt11"
  
    'armo where
    strWhereActual = ""
    If strWhere <> "" Then
      strWhereActual = " where " & strWhere
    End If
      
    'le tiro rs a subreport
    'Set rptSub = rpt.OpenSubreport("stockXarea")
    'strSQL = "SELECT * FROM stockTerXyacimiento_vw_rpt " & strWhereActual
    'Set rs = adoGetRS(strSQL)
    'rptSub.Database.SetDataSource rs
  
  'valorizacion de stock
  Case "rpt12"
  
    'armo where
    strWhereActual = ""
    If strWhere <> "" Then
      strWhereActual = " where " & strWhere
    End If
      
    'le tiro rs a subreport
    Set rptSub = rpt.OpenSubreport("total_subconcesion")
    strSQL = "SELECT * FROM valorizacionStock_xSub_xArea_vw_rpt" & strWhereActual
    Set rs = adoGetRS(strSQL)
    rptSub.Database.SetDataSource rs
  
    Set rptSub = rpt.OpenSubreport("total_terminal")
    strSQL = "SELECT * FROM valorizacionStock_xSub_xArea_vw_rpt" & strWhereActual
    Set rs = adoGetRS(strSQL)
    rptSub.Database.SetDataSource rs
  
  End Select
  
End Function
