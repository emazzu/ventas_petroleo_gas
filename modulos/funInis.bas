Attribute VB_Name = "funInis"
'
' API para leer y grabar INIS
'
Private Declare Function GetPrivateProfileString Lib "Kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
Private Declare Function WritePrivateProfileString Lib "Kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpString As Any, ByVal lpFileName As String) As Long

'
' leo formato de un ini y devuelvo array
' 2 filas y cantidad de columnas como info halla
'
Public Function keyIniToArray(ByVal strHead As String, ByVal strKey As String, Optional ByVal strIniName As String) As Variant
  Dim strDato As String
  Dim intInd, intColumn As Integer
    
  ' nombre de ini default
  If strIniName = "" Then
    strIniName = SQLparam.iniName
  End If
    
  ' leo informacion de INI
  strDato = ReadIni(strHead, strKey, strIniName)

  ' si no encontro ini o clave
  ' devuelvo array vacio y exit
  If strDato = "" Then
    keyIniToArray = Array()
    Exit Function
  End If

  ' devuelvo array con formato
  keyIniToArray = separateText(strDato)

End Function

' strSección , se refiere a lo que va entre corchetes en el <.ini>
' strClave , lo que quieres leer
' Por ejemplo:  de uno llamado <Configuracion.ini>
' [Seccion1]  --> strSección
' MiNombre=JJ --> strClave
'
Public Function ReadIni(strHead As String, strKey As String, Optional strIniName As String) As String
    
    ' nombre de ini default
  If strIniName = "" Then
    strIniName = SQLparam.iniName
  End If
  
  'Los parámetros son:
  'vDefault:      Valor opcional que devolverá
  '               si no se encuentra la clave.
  Dim lpString As String
  Dim LTmp As Long
  Dim sRetVal As String
    
  'Si no se especifica el valor por defecto,
  'asignar incialmente una cadena vacía
  If IsMissing(vDefault) Then
    lpString = ""
  Else
    lpString = vDefault
  End If
    
  sRetVal = String$(2000, 0)
  LTmp = GetPrivateProfileString(strHead, strKey, _
            "", sRetVal, Len(sRetVal), strIniName)
    
  If LTmp = 0 Then
    ReadIni = ""
  Else
    ReadIni = Left(sRetVal, LTmp)
  End If

End Function


'
' guarda una clave en un INI
'
Public Function WriteIni(ByVal strHead As String, ByVal strKey As String, ByVal strValue As String, Optional ByVal strIniName As String)
  Dim LTmp As Long

  ' nombre de ini default
  If strIniName = "" Then
    strIniName = SQLparam.iniName
  End If
    
  LTmp = WritePrivateProfileString(strHead, strKey, strValue, strIniName)

End Function

