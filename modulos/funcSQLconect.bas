Attribute VB_Name = "funcSQLconect"

'parametros para conexion SQL
Type SQLparam
  nombreINI As String
  nombreDAT As String
  IDmenu As String
  Proveedor As String
  ServerSeguridad As String
  ServerDatos As String
  BaseDEdatos As String
  SeguridadIntegrada As String
  TiempoEspera As String
  Usuario As String
  UsuarioClave As String
  UsuarioConectado As String
  GrupoConectado As String
  CantidadFilas As String
  Role As String
  RoleClave As String
  cn As New ADODB.Connection
  CnErrNumero As Long
  CnErrTexto As String
End Type

Global SQLparam As SQLparam

'constantes para tipos de dato SQL
Global Const conChar = 129
Global Const conNchar = 130
Global Const conVarchar = 200
Global Const conText = 201
Global Const conNVarchar = 202
Global Const conNtext = 203
Global Const conDateTime = 135
Global Const conSmallDateTime = 135
Global Const conSmallInt = 2
Global Const conInt = 3
Global Const conTinyInt = 17
Global Const conReal = 4
Global Const conFloat = 5
Global Const conMoney = 6
Global Const conSmallMoney = 6
Global Const conNumeric = 131
Global Const conDecimal = 131
Global Const conBit = 11

'leo parametros de conexion de un INI
Function SQLgetParam() As Boolean
    
  'default devuelve true
  SQLgetParam = True
  
  SQLparam.nombreINI = App.Path & "\" & App.EXEName & ".ini"
  SQLparam.nombreDAT = App.Path & "\" & App.EXEName & ".dat"
  
  SQLparam.IDmenu = ReadIni("conexion", "idMenu", SQLparam.nombreINI)
  SQLparam.Proveedor = ReadIni("conexion", "Proveedor", SQLparam.nombreINI)
  SQLparam.ServerSeguridad = ReadIni("conexion", "ServerSeguridad", SQLparam.nombreINI)
  SQLparam.ServerDatos = ReadIni("conexion", "ServerDatos", SQLparam.nombreINI)
  SQLparam.BaseDEdatos = ReadIni("conexion", "BaseDEdatos", SQLparam.nombreINI)
  SQLparam.SeguridadIntegrada = ReadIni("conexion", "SeguridadIntegrada", SQLparam.nombreINI)
  SQLparam.TiempoEspera = ReadIni("conexion", "TiempoEspera", SQLparam.nombreINI)
  SQLparam.Usuario = ReadIni("conexion", "Usuario", SQLparam.nombreINI)
  SQLparam.UsuarioClave = ReadIni("conexion", "UsuarioClave", SQLparam.nombreINI)
  SQLparam.CantidadFilas = ReadIni("conexion", "CantidadFilas", SQLparam.nombreINI)
    
  'si no puede leer los datos basicos para la conexion devuelvo false y salgo
  If SQLparam.Proveedor = "" Or SQLparam.ServerSeguridad = "" Or SQLparam.ServerDatos = "" Or SQLparam.BaseDEdatos = "" Then
      SQLgetParam = False
      Exit Function
  End If
    
  'leo usuario conectado
  SQLparam.UsuarioConectado = SQLgetUsuario()
    
  'leo grupo al que pertenece el usuario conectado
  SQLparam.GrupoConectado = SQLgetGrupo()
  
  'busco role asociado al usuario o grupo conectado
  SQLparam.Role = SQLgetRole()
  
  'cierro cn
  SQLclose
  
End Function

'leo usuario conectado
Public Function SQLgetUsuario() As String
  
  Dim strT As String
  Dim rs As New ADODB.Recordset
    
  'dafault devuelve ningun usuario
  SQLgetUsuario = ""
  
  'busco usuario
  strT = "select system_user"
  Set rs = SQLexec(strT)
    
  'chequeo error
  If Not SQLparam.CnErrNumero = -1 Then
    SQLError
    Exit Function
  End If
          
  'si encuentro devuelvo nombre
  If Not rs.EOF Then
    SQLgetUsuario = rs(0)
  End If
    
End Function


'leo nombre de usuario SQL, usuario de dominio, grupo de dominio que se conecto
Public Function SQLgetGrupo() As String
  
  Dim strT As String
  Dim rs, rs1 As New ADODB.Recordset
    
  'dafault devuelve ningun usuario
  SQLgetGrupo = ""
  
  'seguridad SQL
  If SQLparam.SeguridadIntegrada = False Then
      
    'busco usuario
    strT = "select name From sysusers Where isSqlUser = 1 and name = '" & SQLparam.Usuario & "'"
    Set rs = SQLexec(strT)
    
    'chequeo error
    If Not SQLparam.CnErrNumero = -1 Then
      SQLError
      Exit Function
    End If
          
    'si encuentro devuelvo nombre
    If Not rs.EOF Then
      SQLgetGrupo = rs!Name
    End If
            
  'seguridad NT
  Else
    
    'leo grupos
    strT = "select name From sysusers Where isntgroup = 1"
    Set rs = SQLexec(strT)
    
    'chequeo error
    If Not SQLparam.CnErrNumero = -1 Then
      SQLError
      Exit Function
    End If
    
    'recorro grupos
    Do While Not rs.EOF
      
      'busco si usuario es miembro de algun grupo del dominio encargado de la seguridad
      strT = "select is_member('" & SQLparam.ServerSeguridad & "\" & rs!Name & "')"
      Set rs1 = SQLexec(strT)
      
      'chequeo error
      If Not SQLparam.CnErrNumero = -1 Then
        SQLError
        Exit Function
      End If
      
      'si encontro, devuelvo nombre
      If rs1(0) = 1 Then
        SQLgetGrupo = rs!Name
        Exit Do
      End If
      
      'cierro
      rs1.Close
      
      'puntero al siguiente
      rs.MoveNext
      
    Loop
    
  End If
    
End Function

'leo nombre del role asociado al usuario de conexion
Public Function SQLgetRole() As String
  
  Dim strT As String
  Dim rs As New ADODB.Recordset
  
  'default devuelve ningun role
  SQLgetRole = ""
  
  'chequeo si existe la tabla menuRoles para no generar un error
  strT = "select name from sysObjects where name = 'menuRoles'"
  Set rs = SQLexec(strT)
  
  'chequeo errores
  If Not SQLparam.CnErrNumero = -1 Then
    Exit Function
  End If
    
  'si no existe fin
  If rs.EOF Then
    Exit Function
  End If
    
  'busco role correspondiente a usuario conectado
  strT = "select role from menuRoles where usuario = '" & SQLparam.UsuarioConectado & "'"
  Set rs = SQLexec(strT)
  
  'chequeo errores
  If Not SQLparam.CnErrNumero = -1 Then
    Exit Function
  End If
          
  'si no encontro usuario fin
  If rs.EOF Then
    Exit Function
  End If
    
  'devuelvo role
  SQLgetRole = rs!Role
      
End Function

'
'devuelve true o false para otorgar o denegar acceso a determinada opcion del menu
'recibe ID grupo, ID opcion, operacion: VER, INSERTAR, EDITAR, ELIMINAR
'1ro. busca en tabla menuPermisos
'2do. si el paso 1ro. da false, busca por roles de SQL server
'
Public Function adoOPCPermisos(opcion As Integer, Optional menu As String, Optional Usuario As String) As Variant
  Dim rs, rs1 As ADODB.Recordset
    
  'valor por default
  adoOPCPermisos = 0
    
  'si parametros vacios, asigno default
  If menu = "" Then menu = IDmenu
  If Usuario = "" Then Usuario = strUsuario
      
  'chequeo permisos
  strSQL = "select ins, edi, eli From dsiOPCpermisos where IDmenu = '" & menu & "' and usuario = '" & Usuario & "' and IDopc = " & opcion
  Set rs = adoGetRS(strSQL)
  
  'chequeo error
  If Not lngAdoErrNum = -1 Then
    adoError
    End
  End If
  
  'si encontro algo
  If Not rs.EOF Then
    
    If strOperacion = "ver" Then
      adoGetPermisos = IIf(rs!ver = True, -1, 0)
    End If
    
    If strOperacion = "insertar" Then
      adoGetPermisos = IIf(rs!ins = True, -1, 0)
    End If
    
    If strOperacion = "editar" Then
      adoGetPermisos = IIf(rs!edi = True, -1, 0)
    End If
    
    If strOperacion = "eliminar" Then
      adoGetPermisos = IIf(rs!eli = True, -1, 0)
    End If
    
  End If
  
  'si hasta este momento no tiene acceso a la opcion busco acceso
  'por roles de base, si es miembro de db_securityAdmin lo dejo pasar
  If Not adoGetPermisos Then
    
    '16386 corresponde a db_securityAdmin
    strSQL = "select memberUid from sysMembers where memberUid = " & intGrupo & " and groupUid = 16386"
    
    'averiguo si grupo tiene algun role standard de SQL asociado
    'strSQL = "select IS_SRVROLEMEMBER('sysadmin') as sysA, IS_SRVROLEMEMBER('securityadmin') as secA, is_member('db_datareader') as [Read], is_member('db_dataWriter') as [Write]"
    
    Set rs1 = adoGetRS(strSQL)
    
    'chequeo error
    If Not lngAdoErrNum = -1 Then
      adoError
      End
    End If
    
    'si encontro es securityAdmin
    If Not rs1.EOF Then
    
      'doy acceso segun roles al cual pertenece
      If strOperacion = "ver" Then
        adoGetPermisos = True
      End If
      
      If strOperacion = "insertar" Then
        adoGetPermisos = False
      End If
      
      If strOperacion = "editar" Then
        adoGetPermisos = False
      End If
      
      If strOperacion = "eliminar" Then
        adoGetPermisos = False
      End If
      
    End If
    
    'cierro
    rs1.Close
    
  End If
    
  'doy acceso segun roles al cual pertenece
  '  If strOperacion = "ver" Then
  '    adoGetPermisos = IIf(rs1!sysA = 1 Or rs1!secA = 1 Or rs1!Read = 1 Or rs1!Write = 1, -1, 0)
  '  End If
    
  '  If strOperacion = "insertar" Then
  '    adoGetPermisos = IIf(rs1!sysA = 1 Or rs1!Write = 1, -1, 0)
  '  End If
    
  '  If strOperacion = "editar" Then
  '    adoGetPermisos = IIf(rs1!sysA = 1 Or rs1!Write = 1, -1, 0)
  '  End If
    
  '  If strOperacion = "eliminar" Then
  '    adoGetPermisos = IIf(rs1!sysA = 1 Or rs1!Write = 1, -1, 0)
  '  End If
    
    'cierro
  '  rs1.Close
    
  'End If
  
End Function

'ejecuta un comando SQL
Public Function SQLexec(ByVal strSQL As String, Optional ByVal strProvider As String, Optional ByVal strServerName As String, Optional ByVal strDatabaseName As String, Optional ByVal blnIntegratedSecurity As Boolean, Optional ByVal intTimeOut As Integer, Optional ByVal strUserID As String, Optional ByVal strPassword As String) As ADODB.Recordset
  
  Dim strT As String
  Dim rs As New ADODB.Recordset
  
  'control de errores
  On Error GoTo controlError
    
  'default sin error
  SQLparam.CnErrNumero = True
  SQLparam.CnErrTexto = ""
    
  'si conexion cerrada, abro
  If SQLparam.cn.State = adStateClosed Then
    
    'si default vacio, asigno false
    If SQLparam.SeguridadIntegrada = "" Then SQLparam.SeguridadIntegrada = "False"
    
    'si parametros vacios, asigno default
    If strProvider = "" Then strProvider = SQLparam.Proveedor
    If strServerName = "" Then strServerName = SQLparam.ServerDatos
    If strDatabaseName = "" Then strDatabaseName = SQLparam.BaseDEdatos
    If blnIntegratedSecurity = False Then blnIntegratedSecurity = CBool(SQLparam.SeguridadIntegrada)
    If intTimeOut = 0 Then intTimeOut = Val(SQLparam.TiempoEspera)
    If strUserID = "" Then strUserID = SQLparam.Usuario
    If strPassword = "" Then strPassword = SQLparam.UsuarioClave
    
    'proveedor y server
    SQLparam.cn.Provider = strProvider
    SQLparam.cn.Properties("Data Source") = strServerName
    
    'seguridad integrada
    If blnIntegratedSecurity = True Then
      SQLparam.cn.Properties("Integrated Security") = "SSPI"
    Else
      SQLparam.cn.Properties("User Id") = strUserID
      SQLparam.cn.Properties("Password") = strPassword
    End If
    
    'time out
    If intTimeOut <> 0 Then
      SQLparam.cn.ConnectionTimeout = intTimeOut
      SQLparam.cn.CommandTimeout = intTimeOut
    End If
    
    'abro conexion
    SQLparam.cn.Open
    
    'base de datos default
    SQLparam.cn.DefaultDatabase = strDatabaseName
    
    'si usamos seguridad por roles de aplicacion, la activo
    If SQLparam.Role <> "" Then
      strT = "exec sp_setappRole '" & SQLparam.Role & "', '" & SQLparam.RoleClave & "'"
      SQLparam.cn.Execute strT
    End If
    
  End If
      
  'determino si ejecuto comando o abro recordset
  If InStr(strSQL, "exec") Or InStr(strSQL, "insert") Or InStr(strSQL, "update") Or InStr(strSQL, "delete") Then
    SQLparam.cn.Execute strSQL
    Set SQLexec = Nothing
  Else
    Set rs.ActiveConnection = SQLparam.cn
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
  SQLparam.CnErrNumero = Err.Number
  SQLparam.CnErrTexto = Err.Description

End Function

'cierra conexion
Public Function SQLclose()
  
  'si conexion abierta, cierro
  If SQLparam.cn.State = adStateOpen Then
  
    SQLparam.cn.Close
    Set SQLparam.cn = Nothing
    
  End If
  
End Function

'
'control de errores de ADO
Function SQLError() As Boolean
  
  Dim intN As Integer
  
  Select Case SQLparam.CnErrNumero
      
  Case 20               'no hay errores
    SQLError = True
      
  Case -2147217873      'clave primaria repetida
    intN = MsgBox("Esta intentando agregar información que ya existe.", vbCritical + vbOKOnly, "atención...")
      
  Case -2147217911      'no tiene permisos
    intN = MsgBox("Esta intentando realizar una operación, para la cual no tiene permisos necesarios.", vbCritical + vbOKOnly, "atención...")
      
  Case Else             'otros errores
    MsgBox ("Error: " & " " & SQLparam.CnErrNumero & " " & SQLparam.CnErrTexto)
      
  End Select

End Function

