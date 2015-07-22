Attribute VB_Name = "varGenerals"
'utilizadas en formularios en donde se confirma o cancela
Global blnAceptar As Boolean
Global blnCancelar As Boolean

'uso general
Global intRes As Integer
Global strSQL As String

'guardo form activo, sirve para el filtro avanzado
Global activeFRM As gridFRM

'guardo errores de conexion
Global lngAdoErrNum As Long
Global strAdoErrDesc As String
