Attribute VB_Name = "funConectSQL"
'
'leo parametros de conexion de un INI
'
'ejemplo de INI
'[conexion]
'IDmenu = 1
'Provider = SQLOLEDB
'serverSecurity = vintage_arg
'ServerName = voa
'DatabaseName = vintageData
'IntegratedSecurity = true
'TimeOut = 30
'UserID=
'Password=
'
Function adoGetParam()
  
  SQLparam.iniName = App.Path & "\" & App.EXEName & ".ini"
  SQLparam.datName = App.EXEName & ".dat"

  SQLparam.IDmenu = ReadIni("conexion", "idMenu", SQLparam.iniName)
  SQLparam.Provider = ReadIni("conexion", "Provider", SQLparam.iniName)
  SQLparam.ServerSecurity = ReadIni("conexion", "serverSecurity", SQLparam.iniName)
  SQLparam.ServerName = ReadIni("conexion", "serverName", SQLparam.iniName)
  SQLparam.DatabaseName = ReadIni("conexion", "databaseName", SQLparam.iniName)
  SQLparam.IntegratedSecurity = ReadIni("conexion", "integratedSecurity", SQLparam.iniName)
  SQLparam.TimeOut = ReadIni("conexion", "timeOut", SQLparam.iniName)
  SQLparam.UserID = ReadIni("conexion", "userID", SQLparam.iniName)
  SQLparam.Password = ReadIni("conexion", "password", SQLparam.iniName)
    
End Function

'leo usuario conectado
Public Function SQLgetUsuario() As String
  
  Dim strT As String
  Dim rs As New ADODB.Recordset
    
  'dafault devuelve ningun usuario
  SQLgetUsuario = ""
  
  'busco usuario
  strT = "select system_user"
  Set rs = adoGetRS(strT)
    
  'chequeo error
  If Not lngAdoErrNum = -1 Then
    adoError
    Exit Function
  End If
          
  'si encuentro devuelvo nombre
  If Not rs.EOF Then
    SQLgetUsuario = rs(0)
  End If
    
End Function

'
'devuelve nombre de usuario de conexion
'sirve para usuarios SQL, usuarios NT, grupos NT
'
Public Function adoGetUsuario() As String
  Dim rs, rs1 As ADODB.Recordset
    
  'valor por dafault
  adoGetUsuario = ""
  
  'seguridad SQL
  If SQLparam.IntegratedSecurity = False Then
      
    'busco usuario
    strSQL = "select name From sysusers Where isSqlUser = 1 and name = '" & SQLparam.UserID & "'"
    Set rs = adoGetRS(strSQL)
    
    'chequeo error
    If Not lngAdoErrNum = -1 Then
      adoError
      End
    End If
          
    'si encuentro devuelvo nombre
    If Not rs.EOF Then
      adoGetUsuario = rs!Name
    End If
            
  'seguridad NT
  Else
    
    'abro rs con grupos
    strSQL = "select name From sysusers Where isntgroup = 1"
    Set rs = adoGetRS(strSQL)
    
    'chequeo error
    If Not lngAdoErrNum = -1 Then
      adoError
      End
    End If
    
    'recorro grupos
    Do While Not rs.EOF
      
      'busco si usuario es miembro de algun grupo
      strSQL = "select is_member('" & SQLparam.ServerSecurity & "\" & rs!Name & "')"
      Set rs1 = adoGetRS(strSQL)
      
      'chequeo error
      If Not lngAdoErrNum = -1 Then
        adoError
        End
      End If
      
      'si encontro, devuelvo nombre
      If rs1(0) = 1 Then
        adoGetUsuario = rs!Name
        Exit Do
      End If
      
      'puntero al siguiente
      rs.MoveNext
      
    Loop
    
  End If
    
End Function

'
'RETURN nombre del role asociado al usuario de conexion
'22/08/2008 ya no se utiliza seguridad por roles
'
'Public Function adoGetRole() As String
'  Dim rs As ADODB.Recordset
'
'  'defa
'  adoGetRole = ""
'
'  Set rs = adoGetRS("select role from menuRoles where usuario = '" & SQLparam.Usuario & "'")
'
'  'chequeo errores
'  If Not lngAdoErrNum = -1 Then
'    Exit Function
'  End If
'
'  If Not rs.EOF Then
'    adoGetRole = rs!Role
'  End If
'
'End Function

'
'devuelvo recordset
'si conexion esta cerrada, la abro, sino l a utilizo
'
Public Function adoGetRS(ByVal strSQL As String, Optional ByVal strProvider As String, Optional ByVal strServerName As String, Optional ByVal strDatabaseName As String, Optional ByVal blnIntegratedSecurity As Boolean, Optional ByVal intTimeOut As Integer, Optional ByVal strUserID As String, Optional ByVal strPassword As String) As ADODB.Recordset
  Dim Cn As New ADODB.Connection
  Dim rs As New ADODB.Recordset
  
  'control de errores
  On Error GoTo controlError
    
  'default sin error
  lngAdoErrNum = True
  strAdoErrDesc = ""
    
  'si conexion cerrada, abro
  If Cn.State = adStateClosed Then
    
    'si default vacio, asigno false
    If SQLparam.IntegratedSecurity = "" Then SQLparam.IntegratedSecurity = "False"
    
    'si parametros vacios, asigno default
    If strProvider = "" Then strProvider = SQLparam.Provider
    If strServerName = "" Then strServerName = SQLparam.ServerName
    If strDatabaseName = "" Then strDatabaseName = SQLparam.DatabaseName
    If blnIntegratedSecurity = False Then blnIntegratedSecurity = CBool(SQLparam.IntegratedSecurity)
    If intTimeOut = 0 Then intTimeOut = Val(SQLparam.TimeOut)
    If strUserID = "" Then strUserID = SQLparam.UserID
    If strPassword = "" Then strPassword = SQLparam.Password
    
    'proveedor y server
    Cn.Provider = strProvider
    Cn.Properties("Data Source") = strServerName
    
    'seguridad integrada
    If blnIntegratedSecurity = True Then
      Cn.Properties("Integrated Security") = "SSPI"
    Else
      Cn.Properties("User Id") = strUserID
      Cn.Properties("Password") = strPassword
    End If
    
    'time out
    If intTimeOut <> 0 Then
      Cn.ConnectionTimeout = intTimeOut
      Cn.CommandTimeout = intTimeOut
    End If
    
    'abro conexion
    Cn.Open                              'abro coneccion
    Cn.DefaultDatabase = strDatabaseName 'base de datos default
    
  End If
      
'  'ejecuto role
'  22/08/2008 ya no se utiliza mas seguridad por roles
'
'  If SQLparam.Role <> "" Then
'    Cn.Execute "exec sp_setapprole '" & SQLparam.Role & "', 'petroleo!15092004'"
'  End If
    
  'abro recordset
  Set rs.ActiveConnection = Cn
  rs.CursorLocation = adUseClient   ' cursor cliente
  rs.CursorType = adOpenStatic      ' tipo estatico
  rs.Source = strSQL                ' origen igual a query
  rs.Open , , , , adCmdText         ' abro le indico que le estoy pasando un string con el Query
  Set adoGetRS = rs                 ' devuelvo recordset
  rs.ActiveConnection = Nothing     ' desconecto recordset con coneccion
  Cn.Close                          'cierra conexion
    
  Exit Function                     'exit funcion
  
  'control errores
controlError:
  lngAdoErrNum = Err.Number
  strAdoErrDesc = Err.Description

End Function

'
' ABRO, EXECUTO SQL, CIERRO CONECCION
'
Public Function adoExecSQL(ByVal strSQL As String, Optional ByVal strProvider As String, Optional ByVal strServerName As String, Optional ByVal strDatabaseName As String, Optional ByVal blnIntegratedSecurity As Boolean, Optional ByVal intTimeOut As Integer, Optional ByVal strUserID As String, Optional ByVal strPassword As String) As Boolean
  Dim Cn As New ADODB.Connection
  
  'control de errores
  On Error GoTo controlError
    
  'default sin error
  lngAdoErrNum = True
  strAdoErrDesc = ""
    
  'si conexion cerrada, abro
  If Cn.State = adStateClosed Then
    
    'si default vacio, asigno false
    If SQLparam.IntegratedSecurity = "" Then SQLparam.IntegratedSecurity = "False"
    
    'si parametros vacios, asigno default
    If strProvider = "" Then strProvider = SQLparam.Provider
    If strServerName = "" Then strServerName = SQLparam.ServerName
    If strDatabaseName = "" Then strDatabaseName = SQLparam.DatabaseName
    If blnIntegratedSecurity = False Then blnIntegratedSecurity = CBool(SQLparam.IntegratedSecurity)
    If intTimeOut = 0 Then intTimeOut = Val(SQLparam.TimeOut)
    If strUserID = "" Then strUserID = SQLparam.UserID
    If strPassword = "" Then strPassword = SQLparam.Password
    
    'proveedor y server
    Cn.Provider = strProvider
    Cn.Properties("Data Source") = strServerName
    
    'seguridad integrada
    If blnIntegratedSecurity = True Then
      Cn.Properties("Integrated Security") = "SSPI"
    Else
      Cn.Properties("User Id") = strUserID
      Cn.Properties("Password") = strPassword
    End If
    
    'time out
    If intTimeOut <> 0 Then
      Cn.ConnectionTimeout = intTimeOut
      Cn.CommandTimeout = intTimeOut
    End If
    
    'abro conexion
    Cn.Open                              'abro coneccion
    Cn.DefaultDatabase = strDatabaseName 'base de datos default
  
  End If
  
  'ejecuto role
  If SQLparam.Role <> "" Then
    Cn.Execute "exec sp_setapprole '" & SQLparam.Role & "', 'petroleo!15092004'"
  End If
  
  Cn.Execute strSQL                       'ejecuta SQL
  Cn.Close                                'cierra conexion
  
  Exit Function                           'fin funcion

'control errores
controlError:
  lngAdoErrNum = Err.Number
  strAdoErrDesc = Err.Description

End Function

'--------------------------------------------------------------------------------------------------------
' ABRO, CONECTO, DEVUELVO CONECCION Y NO CIERRO
'
'Public Function adoOpenCn(Optional ByVal strProvider As String, Optional ByVal strServerName As String, Optional ByVal strDatabaseName As String, Optional ByVal blnIntegratedSecurity As Boolean, Optional ByVal intTimeOut As Integer, Optional ByVal strUserID As String, Optional ByVal strPassword As String) As ADODB.Connection
'  Dim cn As New ADODB.Connection
  
'  'control de errores
'  On Error GoTo controlError
    
'  'default sin error
'  lngAdoErrNum = True
'  strAdoErrDesc = ""
    
'  'si conexion cerrada, abro
'  If cn.State = adStateClosed Then
    
'    'si default vacio, asigno false
'    If SQLparam.IntegratedSecurity = "" Then SQLparam.IntegratedSecurity = "False"
    
'    'si parametros vacios, asigno default
'    If strProvider = "" Then strProvider = SQLparam.Provider
'    If strServerName = "" Then strServerName = SQLparam.ServerName
'    If strDatabaseName = "" Then strDatabaseName = SQLparam.DatabaseName
'    If blnIntegratedSecurity = False Then blnIntegratedSecurity = CBool(SQLparam.IntegratedSecurity)
'    If intTimeOut = 0 Then intTimeOut = Val(SQLparam.TimeOut)
'    If strUserID = "" Then strUserID = SQLparam.UserID
'    If strPassword = "" Then strPassword = SQLparam.Password
    
'    'proveedor y server
'    cn.Provider = strProvider
'    cn.Properties("Data Source") = strServerName
    
'    'seguridad integrada
'    If blnIntegratedSecurity = True Then
'      cn.Properties("Integrated Security") = "SSPI"
'    Else
'      cn.Properties("User Id") = strUserID
'      cn.Properties("Password") = strPassword
'    End If
    
'    'time out
'    If intTimeOut <> 0 Then
'      cn.ConnectionTimeout = intTimeOut
'      cn.CommandTimeout = intTimeOut
'    End If
    
'    'abro conexion
'    cn.Open                              'abro coneccion
'    cn.DefaultDatabase = strDatabaseName 'base de datos default
  
'  End If
  
'  'ejecuto role
'  If SQLparam.Role <> "" Then
'    cn.Execute "exec sp_setapprole '" & SQLparam.Role & "', 'petroleo!15092004'"
'  End If
  
'  Set adoOpenCn = cn                      'devuelvo conection
  
'  Exit Function                           'fin funcion
  
'  'control errores
'controlError:
'  lngAdoErrNum = Err.Number
'  strAdoErrDesc = Err.Description

'End Function
'--------------------------------------------------------------------------------------------------------


' abro una conexion en Access
' abro un recordset
' devuelvo un recordset en el cliente y estatico
' cierro conexion

Public Function adoGetRSAccess(ByVal strPath As String, ByVal strQuery As String) As ADODB.Recordset
  Dim Cn As New ADODB.Connection
  Dim rs As New ADODB.Recordset
  Dim strCadena As String
  
  strCadena = "Provider=" & conProviderAccess & _
              ";Data Source=" & strPath & _
              ";User ID=" & conUserIDAccess & _
              ";Password=" & conPasswordAccess
  
  Cn.Open strCadena
  
  Set rs.ActiveConnection = Cn
  rs.CursorLocation = adUseClient   ' cursor cliente
  rs.CursorType = adOpenStatic      ' tipo estatico
  rs.Source = strQuery              ' origen igual a query
  rs.Open , , , , adCmdText         ' abro le indico que le estoy pasando un string con el Query
  Set adoGetRSAccess = rs         ' devuelvo recordset
  rs.ActiveConnection = Nothing
  
  Cn.Close

End Function

' abro una conexion en Access
' abro un recordset
' devuelvo un recordset en el cliente y estatico
' cierro conexion

Public Function adoOpenCnAccess(ByVal strPath As String) As ADODB.Connection
  Dim Cn As New ADODB.Connection
  Dim strCadena As String
  
  strCadena = "Provider=" & conProviderAccess & _
              ";Data Source=" & strPath & _
              ";User ID=" & conUserIDAccess & _
              ";Password=" & conPasswordAccess
  
  Cn.Open strCadena
  Set adoOpenCnAccess = Cn     'devuelvo conection
  
End Function

'ejecuta un comando SQL
Public Function SQLexec(ByVal strSQL As String, Optional ByVal strProvider As String, Optional ByVal strServerName As String, Optional ByVal strDatabaseName As String, Optional ByVal blnIntegratedSecurity As Boolean, Optional ByVal intTimeOut As Integer, Optional ByVal strUserID As String, Optional ByVal strPassword As String) As ADODB.Recordset
  
  Dim rs As New ADODB.Recordset
  
  'control de errores
  On Error GoTo controlError
    
  'default sin error
  lngAdoErrNum = True
  strAdoErrDesc = ""
  
  'si conexion cerrada, abro
  If SQLparam.Cn.State = adStateClosed Then
    
    'si default vacio, asigno false
    If SQLparam.IntegratedSecurity = "" Then SQLparam.IntegratedSecurity = "False"
    
    'si parametros vacios, asigno default
    If strProvider = "" Then strProvider = SQLparam.Provider
    If strServerName = "" Then strServerName = SQLparam.ServerName
    If strDatabaseName = "" Then strDatabaseName = SQLparam.DatabaseName
    If blnIntegratedSecurity = False Then blnIntegratedSecurity = CBool(SQLparam.IntegratedSecurity)
    If intTimeOut = 0 Then intTimeOut = Val(SQLparam.TimeOut)
    If strUserID = "" Then strUserID = SQLparam.UserID
    If strPassword = "" Then strPassword = SQLparam.Password
    
    'proveedor y server
    SQLparam.Cn.Provider = strProvider
    SQLparam.Cn.Properties("Data Source") = strServerName
    
    'seguridad integrada
    If blnIntegratedSecurity = True Then
      SQLparam.Cn.Properties("Integrated Security") = "SSPI"
    Else
      SQLparam.Cn.Properties("User Id") = strUserID
      SQLparam.Cn.Properties("Password") = strPassword
    End If
    
    'time out
    If intTimeOut <> 0 Then
      SQLparam.Cn.ConnectionTimeout = intTimeOut
      SQLparam.Cn.CommandTimeout = intTimeOut
    End If
    
    SQLparam.Cn.Open                              'abro coneccion
    SQLparam.Cn.DefaultDatabase = strDatabaseName 'base de datos default
    
    'ejecuto role
    If SQLparam.Role <> "" Then
      SQLparam.Cn.Execute "exec sp_setapprole '" & SQLparam.Role & "', 'petroleo!15092004'"
    End If
    
  End If
    
  'determino si ejecuto comando o abro recordset
  strSQL = LCase(strSQL)
  If InStr(strSQL, "exec") Or InStr(strSQL, "insert") Or InStr(strSQL, "update") Or InStr(strSQL, "delete") Then
    SQLparam.Cn.Execute strSQL
    Set SQLexec = Nothing
  Else
    Set rs.ActiveConnection = SQLparam.Cn
    rs.CursorLocation = adUseClient       'cursor cliente
    rs.CursorType = adOpenStatic          'tipo estatico
    rs.Source = strSQL                    'origen igual a query
    rs.Open , , , , adCmdText             'abro le indico que le estoy pasando un string con el Query
    Set SQLexec = rs                      'devuelvo recordset
    rs.ActiveConnection = Nothing         'desconecto recordset con coneccion
  End If
  
  Exit Function                     'exit funcion
  
  'control errores
controlError:
  lngAdoErrNum = Err.Number
  strAdoErrDesc = Err.Description

End Function

'cierra conexion
Public Function SQLclose()
  
  'si conexion abierta, cierro
  If SQLparam.Cn.State = adStateOpen Then
  
    SQLparam.Cn.Close
    Set SQLparam.Cn = Nothing
    
  End If
  
End Function

'
'control de errores de ADO
'
Function adoError() As Boolean
  
  Select Case lngAdoErrNum
      
  Case 20              'no hay errores
    adoError = True
      
  Case Else
    MsgBox ("Error: " & " " & lngAdoErrNum & " " & strAdoErrDesc)
    adoError = False
      
  End Select

End Function


