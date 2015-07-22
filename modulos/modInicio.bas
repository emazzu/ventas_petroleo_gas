Attribute VB_Name = "modInicio"

'
'modulo de inicio del sistema
'
Sub Main()
  
  adoGetParam                             'leo parametros
  
  SQLparam.Usuario = adoGetUsuario()      'leo grupo
  
  '22/08/2008 no se utiliza mas seguridad por roles
  '
  'SQLparam.Role = adoGetRole()            'leo role
  
  SQLparam.UsuarioLogon = SQLgetUsuario() 'leo usuario
        
  'abro frm de inicio
  MDIMenu.Show
  
End Sub

