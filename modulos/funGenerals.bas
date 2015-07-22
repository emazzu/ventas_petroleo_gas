Attribute VB_Name = "funGenerals"
'
'FUNCIONES GENERALES
'convierte un tipo date con formato dd/mm/yyyy a ISO yyyymmdd
'
Public Function dateToIso(ByVal dtm As Variant) As String

  dateToIso = ""
  If IsDate(dtm) Then
    dateToIso = Format(dtm, "yyyymmdd")
  End If

End Function

'
'convierte un tipo date con formato dd/mm/yyyy a Periodo yyyy/mm
'
Public Function dateToPeriodo(ByVal dtm As Variant) As String

  dateToPeriodo = ""
  If IsDate(dtm) Then
    dateToPeriodo = Format(dtm, "yyyy") & "/" & Format(dtm, "mm")
  End If

End Function

'
' Separa informacion en texto separado
'  por comas en un array unidimencional
'
Public Function separateText(ByVal str As String, Optional ByVal strSimbolo As String) As Variant
  Dim arrAUX(), strSimboloAux As String
  Dim intInd, intFind, intCantidad As Integer

  If str = "" Then
    separateText = Array()
    Exit Function
  End If
  
  'si no existe simbolo ponemos ; por default
  If strSimbolo = "" Then
    strSimboloAux = ";"
  Else
    strSimboloAux = strSimbolo
  End If

  ' recorro el string hasta que se acabe
  intInd = 1
  intCantidad = 0
  Do While intInd <= Len(str)

    ' buscamos la primer coma y vamos corriendo
    ' el inicio desde donde busca para la proxima coma
    intFind = InStr(intInd, str, strSimboloAux)
    
    ' si encuentra
    If intFind <> 0 Then
      intCantidad = intCantidad + 1
      ReDim Preserve arrAUX(intCantidad)
      arrAUX(intCantidad) = Mid(str, intInd, intFind - intInd)
      intInd = intFind + 1
    Else
      ' cuando ya no encuentra es el ultimo dato
      intCantidad = intCantidad + 1
      ReDim Preserve arrAUX(intCantidad)
      arrAUX(intCantidad) = Mid(str, intInd, Len(str))
      Exit Do
    End If

  Loop

  separateText = arrAUX

End Function

'
'toma un valor de un array con nombre,valor,nombre,valor, etc.
'
Public Function arrayGetValue(ByVal arrName As Variant, ByVal strColumnName As String) As String
  Dim intInd As Integer
  
  ' valor default
  arrayGetValue = ""
  
  ' valido si array tiene datos
  If UBound(arrName) > 0 Then
    For intInd = 1 To UBound(arrName) - 1 Step 2
      If Format(strColumnName, "<") = Format(arrName(intInd), "<") Then
        arrayGetValue = arrName(intInd + 1)
      End If
    Next
  End If
  
End Function

