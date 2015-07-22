Attribute VB_Name = "modConstants"
'
' DEFINO UN TIPO DE DATO NUEVO UTILIZADO PARA CUANDO ABRO UNA VISTA O TABLA
' PODER ABRIR SOLO EL ULTIMO MES EXISTENTE Y ME DEVUELVE EL RANGO DESDE HASTA
'
Public Type isoLastPeriod
  strDesde As String
  strHasta As String
End Type

'
'Constantes de Conexion Access
Global Const conProviderAccess = "Microsoft.Jet.OLEDB.4.0"
Global conUserIDAccess As String
Global conPasswordAccess As String

'
'parametros de conexion
Type SQLconectParam
  iniName As String
  datName As String
  IDmenu As String
  Provider As String
  ServerSecurity As String
  ServerName As String
  DatabaseName As String
  IntegratedSecurity As String
  TimeOut As String
  UserID As String
  Password As String
  Usuario As String
  Role As String
  UsuarioLogon As String
  Cn As New ADODB.Connection
End Type

'utilizados para la conexion
Global SQLparam As SQLconectParam
Global lngAdoErrNum As Long
Global strAdoErrDesc As String

'
'variables globales de Conexion a SQL
'
'Global conProvider As String
'Global conServerName As String
'Global conDatabaseName As String
'Global conIntegratedSecurity As String
'Global conUserID As String
'Global conPassword As String
'Global conRole As String

' utilizadas para todos los formularios en donde se confirma o cancela
Global blnAceptar As Boolean
Global blnCancelar As Boolean

' utilizada en todos los formularios para hacer o no un refresh del ListView
Global blnRefresh As Boolean

' de uso general para llamar a funciones que devuelven un valor
Global intRes As Integer

' para guardar un string con un query, utilizado para los recordset
Global strSQL As String

' para guardar condicion de filtro, devuelto por la funcion FilterData
Global strWhere As String

' array dinamico utilizado por la funcion de filtrar
' informacion, el array se llena con la estructura de
' una tabla o vista perteneciente a un listview
Global strStruc() As String

' utilizada para conocer frm activo en el menu
Global frmActivo As Form

' array para formatos y anchos de columnas para INIS
Global arrFormat As Variant
Global arrWidth As Variant

' nombre de tabla actual
Global strTableNameActual As String

' tabla default de parametros
Global Const conDBParam = "ViewParametros"

' Constantes de Apariencia de ListView

Global Const conListView_BackColor = &HFFFFFF
Global Const conListView_ForeColor = &H808080

' Constantes de ubicacion para Form

Global Const conFrmHeight = 6270
Global Const conFrmWidth = 9650
Global Const conFrmBorderStyle = 0        ' none

' Constantes de ubicacion para ToolBar

Global Const conTlbAlign = vbAlignNone
Global Const conTlbHeight = 600
Global Const conTlbWidth = conFrmWidth
Global Const conTlbTop = 0
Global Const conTlbLeft = 0
Global Const conTlbButtomHeight = 0
Global Const conTlbButtomWidth = 0
Global Const conTlbAppearance = 0         ' ccFlat

' Constantes de ubicacion para Label

Global Const conLblHeight = 290
Global Const conLblWidth = conFrmWidth
Global Const conLblTop = 600
Global Const conLblLeft = 0
Global Const conLblAlignment = vbCenter
Global Const conLblBackColor = &H808080
Global Const conLblForeColor = &HFFFFFF
Global Const conLblBackStyle = 1          ' Opaque
Global Const conLblBorderStyle = 0        ' None
Global Const conLblFont = "Arial"
Global Const conLblFontBold = True
Global Const conLblFontSize = 11

' Constantes de ubicacion para ListView

Global Const conLvwHeight = 5720
Global Const conLvwWidth = 9680
Global Const conLvwTop = conTlbHeight + conLblHeight
Global Const conLvwLeft = 0
