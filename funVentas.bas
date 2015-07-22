Attribute VB_Name = "funVentas"
' anula factura

Public Function getParamCierre(ByVal strParam As String) As Single
  Dim rsParam As ADODB.Recordset
  
  getParamCierre = -999
  
  strSQL = "SELECT * FROM " & conDBParam & " WHERE Referencia = '" & Format(strParam, "<") & "'"
  Set rsParam = SQLexec(strSQL)
  
  If Not rsParam.EOF Then
    getParamCierre = rsParam!Valor
  End If
  rsParam.Close
  
End Function

' importacion de produccion PAP

Public Function ImportaProduccionMensual(ByVal dtmFecha As Date)
  Dim rsSub, rsPAP As ADODB.Recordset
  Dim curOil, curGas As Currency
  
  ' cambio puntero mouse
  Screen.MousePointer = vbHourglass
  
  ' abro recordset con todas las subconcesiones
  strSQL = "SELECT * FROM SubConcesionesParam_View"
  Set rsSub = SQLexec(strSQL)
  
  ' recorro recordset
  While Not rsSub.EOF
      
    ' chequeo existencia de archivo PAP
    If rsSub!PAPPath <> "" Then
      
      ' chequea que exista el archivo
      If Dir(rsSub!PAPPath) <> "" Then
        
        ' abro recordset con produccion de Oil y Gas
        strSQL = " SELECT sum(PR.pet_neto_dia * PR.dee * HC.coef_pet) As Oil, " & _
                 " sum(PR.gas_dia * PR.dee * HC.coef_gas) As Gas " & _
                 " FROM (Produccion PR INNER JOIN HistoriaCoef HC ON " & _
                 " Month(PR.Fecha) = Month(HC.Fecha) AND " & _
                 " Year(PR.Fecha) = Year(HC.Fecha)) " & _
                 " WHERE Month(PR.Fecha) = " & Month(dtmFecha) & " AND " & _
                 " year(PR.Fecha) = " & Year(dtmFecha) & " AND " & _
                 " UC = " & rsSub!PAPControl
        Set rsPAP = adoGetRSAccess(rsSub!PAPPath, strSQL)
      
        ' chequeo existencia de produccion
        If Not rsPAP.EOF And Not IsNull(rsPAP!Oil) Then
      
          ' agrego produccion
          strSQL = "EXEC spProduccionMensualInsert " & _
                   "'" & dateToIso(dtmFecha) & "'," & _
                   rsSub!empresaID & "," & _
                   rsSub!IDsubconcesion & "," & _
                   Round(rsPAP!Oil, 3) & "," & _
                   Round(rsPAP!Gas, 3)
          SQLexec (strSQL)
      
        End If
      
        ' cierro recordset
        rsPAP.Close
      
      End If
    
    Else
    
      ' agrego produccion vacia, cuando no tiene asociado un mdb
      strSQL = "EXEC spProduccionMensualInsert " & _
               "'" & dateToIso(dtmFecha) & "'," & _
               rsSub!empresaID & "," & _
               rsSub!IDsubconcesion & "," & _
               0 & "," & _
               0
      SQLexec (strSQL)
    
    End If
    
    ' muevo siguiente subconcesion
    rsSub.MoveNext
    
  Wend

  'cierro Recordset
  rsSub.Close
  
  ' agrega control de proceso
  strSQL = "EXEC spStockCierreInsert " & _
           "'" & dateToIso(dtmFecha) & "'," & _
           "'PMEN'" & "," & _
           1
  SQLexec (strSQL)
  
  'cierro conexion
  SQLclose
  
  ' vuelvo puntero mouse
  Screen.MousePointer = vbDefault
  
  intRes = MsgBox("El proceso finalizó con éxito.", vbInformation + vbOKOnly, "Información")

End Function

'
' calculando stock en subconcesiones con Entregas y Produccion Mensual
'
Public Function CalcularStockSubconcesiones(dtmPeriodoAct As Date)
  Dim rsSub, rsIni, rsPRO, rsTRA As ADODB.Recordset
  Dim dtmPeriodoAnt As Date
  Dim curSdoInicial, curProMensual, curEntregasTra As Currency
  
  ' cambio puntero mouse
  Screen.MousePointer = vbHourglass
  
  ' tomo el ultimo dia del periodo anterior
  dtmPeriodoAnt = dtmPeriodoAct - Day(dtmPeriodoAct)

  ' abro recordset con las subconcesiones por empresa
  strSQL = "SELECT * FROM ViewSubconcesionesStockSub"
  Set rsSub = SQLexec(strSQL)

  ' recorro subconcesiones para calcular stock
  While Not rsSub.EOF
  
    ' busco saldo inicial para la subconcesion que
    ' es el saldo final del periodo anterior
    strSQL = "SELECT subconcesionID, empresaID, fecha, stock FROM SubconcesionesStock WHERE " & _
          "subconcesionID = " & rsSub!IDsubconcesion & " AND " & _
          "empresaID = " & rsSub!empresaID & " AND " & _
          "fecha = '" & dateToIso(dtmPeriodoAnt) & "'"
    Set rsIni = SQLexec(strSQL)
    
    ' tomo el saldo inicial
    curSdoInicial = 0
    If Not rsIni.EOF Then
      curSdoInicial = rsIni!stock
    End If
    rsIni.Close
    
    ' busco produccion mensual para la subconcesion y periodo actual
    strSQL = "SELECT fecha, empresaID, subconcesionID, Oil, Gas FROM ProduccionMensual WHERE " & _
          "fecha = '" & dateToIso(dtmPeriodoAct) & "' AND " & _
          "empresaID = " & rsSub!empresaID & " AND " & _
          "subconcesionID = " & rsSub!IDsubconcesion
    Set rsPRO = SQLexec(strSQL)
    
    ' tomo la produccion mensual
    curProMensual = 0
    If Not rsPRO.EOF Then
      curProMensual = rsPRO!Oil
    End If
    rsPRO.Close
    
    ' busco todas las entregas transportistas del periodo actual
    strSQL = "SELECT * FROM ViewSubconcesionesStockTra WHERE " & _
          "subconcesionID = " & rsSub!IDsubconcesion & " AND " & _
          "empresaID = " & rsSub!empresaID & " AND " & _
          "fecha BETWEEN '" & dateToIso(dateToFirstDay(dtmPeriodoAct)) & "' AND " & _
          "'" & dateToIso(dtmPeriodoAct) & "'"
    Set rsTRA = SQLexec(strSQL)
    
    ' recorro entregas a transportistas
    
    curEntregasTra = 0
    While Not rsTRA.EOF
      
      ' sumo volumen SecoSeco15
      curEntregasTra = curEntregasTra + rsTRA!VolSecoSeco15
      
      ' avanzo puntero entregas
      rsTRA.MoveNext
    
    Wend
    rsTRA.Close
    
    ' agrego stock periodo actual
  
    strSQL = "EXEC spSubconcesionesStockInsert " & _
          rsSub!IDsubconcesion & "," & _
          rsSub!empresaID & "," & _
          "'" & dateToIso(dtmPeriodoAct) & "'," & _
          Round(curSdoInicial, 3) & "," & _
          Round(curProMensual, 3) & "," & _
          Round(curEntregasTra, 3) & "," & _
          Round(curSdoInicial + curProMensual - curEntregasTra, 3)

    SQLexec (strSQL)
  
    ' avanzo puntero subconcesiones
    rsSub.MoveNext
  
  Wend
  rsSub.Close

  ' agrega control de proceso
  strSQL = "EXEC spStockCierreInsert " & _
           "'" & dateToIso(dtmPeriodoAct) & "'," & _
           "'SSUB'" & "," & _
           1
  SQLexec (strSQL)
  
  'cierro conexion
  SQLclose
  
  ' recupero mouse standard
  Screen.MousePointer = vbDefault

  intRes = MsgBox("El proceso finalizó con éxito.", vbInformation + vbOKOnly, "Información")

End Function

'**********************************************************************************************
'
'                       CALCULO DE STOCK EN TERMINALES
'
'**********************************************************************************************
Public Function CalcularStockTerminales(dtmPeriodoAct As Date, strRecalcularVentas As String)
  Dim rsTotTer, rsTotSub, rsTerStock, rsSubStock, rsTraTer, rsEntCli, rsSub, rsIni, rsTer As ADODB.Recordset
  Dim intDiaActual, intIDEntregaCli As Integer
  Dim curTotalEnt, curSdoInicialOil, curSdoFinalOil, curAPICliOld, curAPICliNew, curVolSeco15Cli As Currency
  Dim dtmPeriodoAnt, dtmDiaActual As Date
  Dim curImpurezas, curVolNeto15Tra, curAPItra, curAPITer, curVolseco15Ter, curVolSeco15Per, curPjeMermas As Currency
  Dim strActa As String
  Dim blnPrimeraVez As Boolean
  Dim intIDEmpresa, intIDTerminal, intIDCliente, intIDEntrega, intIDSubconcesion As Integer
  Dim curStockTerAnt, curStockSubAnt As Currency
  Dim intEmpresaAUX As Integer
  Dim blnB As Boolean
  
  'cambio puntero mouse
  Screen.MousePointer = vbHourglass
  
  '-------------------------------------------------------------------------------------------
  'PASO CERO:   BLANQUEAMOS LOS PORCENTAJES DE LA TABLA SUBCONCESIONES ANTES DE COMENZAR
  '             ES PARA QUE LAS SUBCONCESIONES QUE NO HAY INFO ESTEN EN 0 Y NO ARRASTRE
  '             EL VALOR DEL MES ANTERIOR
  '-------------------------------------------------------------------------------------------
    
  'abro rs con todas las subconcesiones
  strSQL = "exec spSubconcesionesLimpiaPje"
  SQLexec (strSQL)
      
  'chequeo errores
  If Not lngAdoErrNum = -1 Then
    adoError
    Exit Function
  End If
      
  '-------------------------------------------------------------------------------------------
  'PRIMER PASO:           CALCULA PORCENTAJES DE VENTAS POR SUBCONCESION ESTO SE
  '                       GUARDA EN LA COLUMNA PJESUBTERM DE LA TABLA SUBCONCESIONES
  '-------------------------------------------------------------------------------------------
    
  'chequeo que opcion se selecciono, para recalcular porcentajes de distribucion de ventas o no
  
  'abro recordset con total de las entregas x terminales
  strSQL = "SELECT * FROM ViewTerminalesStockTotalxTer WHERE " & _
          "anio = " & Year(dtmPeriodoAct) & " AND " & _
          "mes = " & Month(dtmPeriodoAct)
  Set rsTotTer = SQLexec(strSQL)
  
  'chequeo errores
  If Not lngAdoErrNum = -1 Then
    adoError
    Exit Function
  End If
  
  'abro recordset con total de las entregas x Subconcesion
  strSQL = "SELECT * FROM ViewTerminalesStockTotalxTerySub WHERE " & _
          "anio = " & Year(dtmPeriodoAct) & " AND " & _
          "mes = " & Month(dtmPeriodoAct)
  Set rsTotSub = SQLexec(strSQL)
  
  'abro recordset con stock x subconcesiones al ultimo dia del periodo anterior
  dtmPeriodoAnt = dtmPeriodoAct - Day(dtmPeriodoAct)
  strSQL = "SELECT * FROM ViewTerminalesStock WHERE " & _
          "fecha = '" & dateToIso(dtmPeriodoAnt) & "'"
  Set rsSubStock = SQLexec(strSQL)
  
  'chequeo errores
  If Not lngAdoErrNum = -1 Then
    adoError
    Exit Function
  End If
  
  'recorro total por terminales
  While Not rsTotTer.EOF
    
    'abro recordset con stock x empresa y terminal al ultimo dia del periodo anterior
    'dtmPeriodoAnt = dtmPeriodoAct - Day(dtmPeriodoAct)
    strSQL = "SELECT SUM(finalOil) as FinalOil,empresaID, terminalID FROM ViewTerminalesStock WHERE " & _
             "fecha = '" & dateToIso(dtmPeriodoAnt) & "' and " & _
             "empresaID = " & rsTotTer!empresaID & " and " & _
             "terminalID = " & rsTotTer!terminalid & " group by empresaID, terminalID"
    Set rsTerStock = SQLexec(strSQL)
    
  'chequeo errores
  If Not lngAdoErrNum = -1 Then
    adoError
    Exit Function
  End If
    
    'si encontro stock lo guardo
    curStockTerAnt = 0
    If Not rsTerStock.EOF Then
      curStockTerAnt = rsTerStock!FinalOil
    End If
    
    'cierro
    rsTerStock.Close
    
    'filtro total por subconcesiones por la empresa y terminal
    'dada por cada registro de total por terminales
    rsTotSub.Filter = "empresaID = " & rsTotTer!empresaID & " AND " & _
                      "TerminalID = " & rsTotTer!terminalid
    
    'puntero al principio
    rsTotSub.MoveFirst
    
    Dim curPJETotal As Currency
    curPJETotal = 0
    
    'recorro total por subconcesiones
    While Not rsTotSub.EOF
      
      'la primera pasada guardo datos de clave primaria por si hay diferencia
      'entre los totales de entregas por terminal y los totales de entregas
      'por subconcesion, la diferencia a la primera subconcesion de cada terminal
      If Not blnPrimeraVez Then
        intIDEmpresa = rsTotSub!empresaID
        intIDTerminal = rsTotSub!terminalid
        intIDSubconcesion = rsTotSub!subconcesionID
        blnPrimeraVez = True
      End If
      
      'tomo stock inicial que es al ultimo dia del mes anterior y se los sumo al total de lo entregado
      rsSubStock.Filter = "empresaID = " & rsTotSub!empresaID & " and subconcesionID = " & rsTotSub!subconcesionID & _
                          " and fecha = " & dtmPeriodoAnt
      
      'si encontro stock anterior lo guardo
      curStockSubAnt = 0
      If Not rsSubStock.EOF Then
        curStockSubAnt = rsSubStock!FinalOil
      End If
      
      'actualizo porcentaje de distribucion en subconcesiones para colum PjeSubTerm
      strSQL = "EXEC spTerminalesStockPjeSubTermUpdate " & _
      rsTotSub!empresaID & "," & _
      rsTotSub!subconcesionID & "," & _
      IIf(rsTotTer!volNeto15Aju <> 0, IIf(Round(((curStockSubAnt + rsTotSub!volNeto15Aju) / (rsTotTer!volNeto15Aju + curStockTerAnt)) * 100, 2) < 0, 0.05, Round(((curStockSubAnt + rsTotSub!volNeto15Aju) / (rsTotTer!volNeto15Aju + curStockTerAnt)) * 100, 2)), 0)
      SQLexec (strSQL)
      
      'chequeo errores
      If Not lngAdoErrNum = -1 Then
        adoError
        Exit Function
      End If
      
      'acumulador de porcentajes
      curPJETotal = curPJETotal + IIf(rsTotTer!volNeto15Aju <> 0, IIf(Round(((curStockSubAnt + rsTotSub!volNeto15Aju) / (rsTotTer!volNeto15Aju + curStockTerAnt)) * 100, 2) < 0, 0.05, Round(((curStockSubAnt + rsTotSub!volNeto15Aju) / (rsTotTer!volNeto15Aju + curStockTerAnt)) * 100, 2)), 0)
        
      'avanzo
      rsTotSub.MoveNext
    
    Wend
    
    'chequeo totales de terminales con acumuladores x subconcesion para calcular diferencia
    If (100 - curPJETotal) <> 0 Then
      
      'actualizo porcentaje de distribucion en subconcesiones por la diferencia
      strSQL = "EXEC spTerminalesStockPjeSubTermUpdateDif " & _
      intIDEmpresa & "," & _
      intIDSubconcesion & "," & _
      (100 - curPJETotal)
      SQLexec (strSQL)
    
      'chequeo errores
      If Not lngAdoErrNum = -1 Then
        adoError
        Exit Function
      End If
    
    End If
    
    'avanzo
    rsTotTer.MoveNext
  
  Wend
  
  'cierro
  rsTotSub.Close
  rsTotTer.Close
    
  '-------------------------------------------------------------------------------------------
  'SEGUNDO PASO:           APERTURA DE VENTAS POR SUBCONCESION TOMA EL PJESUBTERM
  '                        DE LA TABLA SUBCONCESIONES GENERADA EN EL PASO ANTERIOR
  '-------------------------------------------------------------------------------------------
  
  Dim rsVentas, rsPorcen As ADODB.Recordset
  Dim curTotalMts15, curTotalMts1556, curTotalBbls, curTotalImporte, curTotalPorcen As Currency
  
  'abro porcentajes de ventas por terminal y subconcesion
  strSQL = "SELECT * FROM pjeVentasXterYsub_vw"
  Set rsPorcen = SQLexec(strSQL)
  
  'chequeo errores
  If Not lngAdoErrNum = -1 Then
    adoError
    Exit Function
  End If
  
  'abro entregas a clientes relacionadas con las ventas x terminal
  strSQL = "SELECT * FROM distribucionVentas_vw WHERE " & _
           "anio = " & Year(dtmPeriodoAct) & " AND " & _
           "mes = " & Month(dtmPeriodoAct)
  Set rsVentas = SQLexec(strSQL)
  
  'chequeo errores
  If Not lngAdoErrNum = -1 Then
    adoError
    Exit Function
  End If
  
  'recorro total por terminales
  While Not rsVentas.EOF
      
    'filtro subconcesiones para la empresa y terminal de la venta actual
    rsPorcen.Filter = "empresaID = " & rsVentas!empresaID & " and terminalID = " & rsVentas!terminalid
    
    'puntero al primer lugar
    rsPorcen.MoveFirst
    
    'totalizadores a cero
    curTotalMts15 = 0
    curTotalMts1556 = 0
    curTotalBbls = 0
    curTotalImporte = 0
    curTotalPorcen = 0
    blnPrimeraVez = False
    
    'recorro subconcesiones
    While Not rsPorcen.EOF
      
      'la primera pasada guardo datos de clave primaria por si hay diferencia
      'entre los totales y los acumuladores de ventas, puedo agregarle o sacarle
      'la diferencia a la primera subconcesion de cada terminal
      If Not blnPrimeraVez Then
        intIDEmpresa = rsVentas!empresaID
        intIDCliente = rsVentas!clienteID
        intIDEntrega = rsVentas!identregaCli
        intIDSubconcesion = rsPorcen!IDsubconcesion
        strFactura = rsVentas!factura
        blnPrimeraVez = True
      End If
      
      'guardo apertura de ventas
      strSQL = "EXEC distribucionVentasXsubInsert " & _
              rsVentas!empresaID & "," & _
              rsVentas!clienteID & "," & _
              rsVentas!identregaCli & "," & _
              rsVentas!terminalid & "," & _
              rsPorcen!IDsubconcesion & "," & _
              rsVentas!nAPIGravity & "," & _
              rsVentas!nDensity & "," & _
              rsPorcen!pjeSubTerm & "," & _
              "'" & dateToIso(rsVentas!fechaEntrega) & "'," & _
              rsVentas!barcoID & "," & _
              Round((rsVentas!Mts15 * rsPorcen!pjeSubTerm / 100), 3) & "," & _
              Round((rsVentas!mts1556 * rsPorcen!pjeSubTerm / 100), 3) & "," & _
              Round((rsVentas!Bbls * rsPorcen!pjeSubTerm / 100), 3) & "," & _
              "'" & rsVentas!Moneda & "'," & _
              Round((rsVentas!Importe * rsPorcen!pjeSubTerm / 100), 2) & "," & _
              "'" & rsVentas!factura & "'," & _
              rsVentas!precio & ",'" & _
              rsVentas!tipoComprobante & "'"
     SQLexec (strSQL)
      
      'chequeo errores
      If Not lngAdoErrNum = -1 Then
        adoError
        Exit Function
      End If
      
      'acumuladores
      curTotalPorcen = curTotalPorcen + rsPorcen!pjeSubTerm
      curTotalMts15 = curTotalMts15 + Round((rsVentas!Mts15 * rsPorcen!pjeSubTerm / 100), 3)
      curTotalMts1556 = curTotalMts1556 + Round((rsVentas!mts1556 * rsPorcen!pjeSubTerm / 100), 3)
      curTotalBbls = curTotalBbls + Round((rsVentas!Bbls * rsPorcen!pjeSubTerm / 100), 3)
      curTotalImporte = curTotalImporte + Round((rsVentas!Importe * rsPorcen!pjeSubTerm / 100), 2)
    
      'avanzo proxima venta
      rsPorcen.MoveNext
  
    Wend
    
    'chequeo totales de ventas con acumuladores x subconcesion
    If (rsVentas!Mts15 - curTotalMts15) <> 0 Or (rsVentas!mts1556 - curTotalMts1556) <> 0 Or (rsVentas!Bbls - curTotalBbls) <> 0 Or (rsVentas!Importe - curTotalImporte) <> 0 Then
      
      'actualizo diferencia
      strSQL = "EXEC distribucionVentasXsubUpdate " & _
              intIDEmpresa & "," & _
              intIDCliente & "," & _
              intIDEntrega & "," & _
              intIDSubconcesion & "," & _
              "'" & strFactura & "'," & _
              rsVentas!Mts15 - curTotalMts15 & "," & _
              rsVentas!mts1556 - curTotalMts1556 & "," & _
              rsVentas!Bbls - curTotalBbls & "," & _
              rsVentas!Importe - curTotalImporte
      SQLexec (strSQL)
    
      'chequeo errores
      If Not lngAdoErrNum = -1 Then
        adoError
        Exit Function
      End If
    
    End If
    
    'avanzo proxima venta
    rsVentas.MoveNext
  
  Wend
  
  'cierro rs
  rsVentas.Close
  rsPorcen.Close
    
  '-------------------------------------------------------------------------------------------
  'SEGUNDO PASO:          CALCULO STOCK DE TERMINALES
  '-------------------------------------------------------------------------------------------
    
  'abro recordset con las entregas transportistas y entregas terminales
  strSQL = "SELECT * FROM ViewTerminalesStockEntregas WHERE " & _
          "Fecha BETWEEN '" & dateToIso(dateToFirstDay(dtmPeriodoAct)) & "' AND '" & _
          dateToIso(dateToLastDay(dtmPeriodoAct)) & "'"
  Set rsTraTer = SQLexec(strSQL)
  
  'chequeo errores
  If Not lngAdoErrNum = -1 Then
    adoError
    Exit Function
  End If
  
  'abro recordset con las entregas clientes - embarques
  strSQL = "SELECT * FROM distribucionVentasXsub WHERE " & _
          "FechaEntrega BETWEEN '" & dateToIso(dateToFirstDay(dtmPeriodoAct)) & "' AND '" & _
          dateToIso(dateToLastDay(dtmPeriodoAct)) & "'"
  Set rsEntCli = SQLexec(strSQL)
    
  'chequeo errores
  If Not lngAdoErrNum = -1 Then
    adoError
    Exit Function
  End If
    
  'abro recordset con las subconcesiones por empresa y recorro
  strSQL = "SELECT * FROM ViewSubconcesionesStockSub"
  Set rsSub = SQLexec(strSQL)
  While Not rsSub.EOF
  
    'busco saldo inicial para la subconcesion que
    'es el saldo final del periodo anterior
    'tomo el ultimo dia del periodo anterior
    dtmPeriodoAnt = dtmPeriodoAct - Day(dtmPeriodoAct)
    strSQL = "SELECT subconcesionID, empresaID, fecha, APICli, FinalOil FROM TerminalesStock WHERE " & _
            "subconcesionID = " & rsSub!IDsubconcesion & " AND " & _
            "empresaID = " & rsSub!empresaID & " AND " & _
            "fecha = '" & dateToIso(dtmPeriodoAnt) & "'"
    Set rsIni = SQLexec(strSQL)
    
  'chequeo errores
  If Not lngAdoErrNum = -1 Then
    adoError
    Exit Function
  End If
    
    curSdoInicialOil = 0
    curAPICliOld = 0
    curAPICliNew = 0
    If Not rsIni.EOF Then
      curSdoInicialOil = rsIni!FinalOil          ' Stock Inicial
      curAPICliOld = rsIni!APICli                ' API Inicial
      curAPICliNew = rsIni!APICli
    End If
    rsIni.Close
    
    ' recorro los 28, 30 o 31 dias del mes actual la para poder generar el dia a dia
    For intDiaActual = Day(dateToFirstDay(dtmPeriodoAct)) To Day(dateToLastDay(dtmPeriodoAct))
          
      ' genero fecha del dia actual para poder filtrar entregas
      dtmDiaActual = CVDate(Format(intDiaActual, "00") & Right(str(dtmPeriodoAct), 8))
          
      ' filtro entregas transportistas y entregas terminales
      ' por empresa, subconcesion y dia actual
      rsTraTer.Filter = "empresaID = " & rsSub!empresaID & " AND " & _
                      "SubconcesionID = " & rsSub!IDsubconcesion & " AND " & _
                      "Fecha = " & dtmDiaActual
    
      ' inicializo variables
      strActa = ""
      curImpurezas = 0
      curVolNeto15Tra = 0
      curAPItra = 0
      curAPITer = 0
      curVolseco15Ter = 0
      curVolSeco15Per = 0
      curPjeMermas = 0

      ' hay algun acta este dia ?
      If Not rsTraTer.EOF Then
        
        ' busco parametro de PjeMermas en terminal
        strSQL = "SELECT PjeMermas FROM Terminales WHERE idTerminal = " & rsTraTer!terminalid
        Set rsTer = SQLexec(strSQL)
        If Not rsTer.EOF Then
          curPjeMermas = rsTer!PjeMermas
        End If
        rsTer.Close
        
        ' chequeo si la subconcesion actual si de debe ajustar la merma de la terminal
        If rsSub!ajuMermaTerm = "No" Then
          curPjeMermas = 0
        End If
        
        strActa = rsTraTer!Acta
        curImpurezas = rsTraTer!Impurezas
        curVolNeto15Tra = rsTraTer!VolNeto15
        curAPItra = rsTraTer!APITRa
        curAPITer = rsTraTer!ApiTer
        curVolseco15Ter = Round(curVolNeto15Tra * (1 + getParamCierre("apiCoeficienteAju") * (curAPItra - curAPITer)), 3)
        curVolSeco15Per = Round(curVolseco15Ter * (1 - curPjeMermas / 100), 3)
        End If
         
      ' filtro entregas clientes (embarques)
      ' por empresa, subconcesiones y dia actual
      rsEntCli.Filter = "IDempresa = " & rsSub!empresaID & " AND " & _
                        "IDSubconcesion = " & rsSub!IDsubconcesion & " AND " & _
                        "FechaEntrega = " & dtmDiaActual
      
      ' inicializo variables
      curVolSeco15Cli = 0
      intIDEntregaCli = 0
      
      Dim curPjeSubTerm As Currency
      curPjeSubTerm = 0
        
      ' hay alguna venta este dia ?
      If Not rsEntCli.EOF Then
        
        'guardo los datos de cabezera
        intIDEntregaCli = rsEntCli!identregaCli
        curAPICliNew = rsEntCli!nAPIGravity
        curPjeSubTerm = rsEntCli!pjeSubTerm
        curVolSeco15Cli = 0
        
        'recorro por puede ser el caso que hay
        '2 entregas para el mismo dia entonces las acumulo
        While Not rsEntCli.EOF
        
          curVolSeco15Cli = curVolSeco15Cli + rsEntCli!Mts15
          rsEntCli.MoveNext
          
        Wend
      
      Else
      
        ' verifico si debe ajustar el API para esta subconcesion
        If rsSub!AjustaApiStk = "Si" Then
      
          ' si no hay venta este dia, debe buscar la proxima venta para tomar el API
          ' si no hay proximas ventas hasta fin de mes, tomo API anterior
          Dim dtmDiaActualAUX As Date
          Dim intDiaActualAUX As Integer
        
          For intDiaActualAUX = intDiaActual + 1 To Day(dateToLastDay(dtmPeriodoAct))
          
            ' genero fecha del dia actual para buscar nuevo API
            dtmDiaActualAUX = CVDate(Format(intDiaActualAUX, "00") & Right(str(dtmPeriodoAct), 8))
          
            ' filtro entregas clientes (embarques)
            ' por empresa, subconcesiones y proximo dia al dia actual hasta el final
            rsEntCli.Filter = "IDempresa = " & rsSub!empresaID & " AND " & _
                              "IDSubconcesion = " & rsSub!IDsubconcesion & " AND " & _
                              "FechaEntrega = " & dtmDiaActualAUX
            ' hay alguna venta este dia ?
            If Not rsEntCli.EOF Then
              curAPICliNew = rsEntCli!nAPIGravity
              Exit For
            End If
        
          Next
              
        Else
          curAPICliNew = 0
        End If
              
      End If
      
     'verifico si debe ajustar el API para esta subconcesion
     If rsSub!AjustaApiStk = "Si" Then
        
       'llevo VolSeco15 con perdidas de dia actual a API cliente saldo final dia actual
       curVolSeco15Per = Round(curVolSeco15Per * (1 + getParamCierre("apiCoeficienteAju") * (curAPITer - curAPICliNew)), 3)
        
       'llevo saldo inicial de api cliente dia anterior a api cliente dia actual
       curSdoInicialOil = Round(curSdoInicialOil * (1 + getParamCierre("apiCoeficienteAju") * (curAPICliOld - curAPICliNew)), 3)
        
    End If
    
    'calculo saldo final del dia Inicial + Entregas - Ventas
    curSdoFinalOil = Round(curSdoInicialOil + curVolSeco15Per - curVolSeco15Cli, 3)
    
    'agrego stock periodo actual
    strSQL = "EXEC spTerminalesStockInsert " & _
            rsSub!empresaID & "," & _
            rsSub!IDsubconcesion & "," & _
            "'" & dateToIso(dtmDiaActual) & "','" & _
            strActa & "'," & _
            curImpurezas & "," & _
            curVolNeto15Tra & "," & _
            curAPItra & "," & _
            curAPITer & "," & _
            curVolseco15Ter & "," & _
            curVolSeco15Per & "," & _
            curVolSeco15Cli & "," & _
            curPjeSubTerm & "," & _
            curAPICliNew & "," & _
            curSdoInicialOil & "," & _
            curSdoFinalOil & "," & _
            rsSub!terminalid & "," & _
            0 & "," & _
            intIDEntregaCli
      SQLexec (strSQL)
    
      'chequeo errores
      If Not lngAdoErrNum = -1 Then
        adoError
        Exit Function
      End If
    
      ' le paso el SaldoFinalOIL al SaldoInicialOIL
      ' para que el saldo inicial pase para el proximo dia
      ' le paso el APINew al APIOld para el proximo dia
      curSdoInicialOil = curSdoFinalOil
      curAPICliOld = curAPICliNew
    
    Next
    
    ' avanzo
    rsSub.MoveNext
  
  Wend
  
  'cierro rs
  rsSub.Close
  
  ' agrega control de proceso
  strSQL = "EXEC spStockCierreInsert " & _
           "'" & dateToIso(dateToLastDay(dtmDiaActual)) & "'," & _
           "'STER'" & "," & _
           1
  SQLexec (strSQL)

  'chequeo errores
  If Not lngAdoErrNum = -1 Then
    adoError
    Exit Function
  End If

  'cierro cn
  SQLclose

  ' recupero mouse standard
  Screen.MousePointer = vbDefault

  intRes = MsgBox("El proceso finalizó con éxito.", vbInformation + vbOKOnly, "Información")

End Function





