VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form frmCierreStock 
   BorderStyle     =   0  'None
   Caption         =   "CierreStock"
   ClientHeight    =   6750
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   12765
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   6750
   ScaleWidth      =   12765
   ShowInTaskbar   =   0   'False
   Begin MSComctlLib.ListView lvwDatos 
      Height          =   5055
      Left            =   0
      TabIndex        =   0
      Top             =   990
      Width           =   9240
      _ExtentX        =   16298
      _ExtentY        =   8916
      View            =   3
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      _Version        =   393217
      ForeColor       =   8421504
      BackColor       =   16777215
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   0
   End
   Begin MSComctlLib.Toolbar tlbOperaciones 
      Height          =   570
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   12420
      _ExtentX        =   21908
      _ExtentY        =   1005
      ButtonWidth     =   2461
      ButtonHeight    =   1005
      Style           =   1
      TextAlignment   =   1
      ImageList       =   "ImageList1"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   8
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "&Agregar"
            Key             =   "agregar"
            ImageIndex      =   1
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "&Modificar"
            Key             =   "modificar"
            ImageIndex      =   2
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "&Eliminar"
            Key             =   "eliminar"
            ImageIndex      =   3
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "&Buscar"
            Key             =   "buscar"
            ImageIndex      =   5
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "&Filtrar"
            Key             =   "filtrar"
            ImageIndex      =   4
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "&Ordenar"
            Key             =   "ordenar"
            ImageIndex      =   6
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "&Exportar"
            Key             =   "exportar"
            ImageIndex      =   7
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "&Ajustar"
            Key             =   "ajustar"
            ImageIndex      =   8
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   90
      Top             =   6075
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   8
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCierreStock.frx":0000
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCierreStock.frx":08DA
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCierreStock.frx":11B4
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCierreStock.frx":1A8E
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCierreStock.frx":3420
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCierreStock.frx":3CFA
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCierreStock.frx":45D4
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCierreStock.frx":5F66
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Label lblTitulo 
      Alignment       =   2  'Center
      BackColor       =   &H00808000&
      Caption         =   "Cierre Stock"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   0
      TabIndex        =   2
      Top             =   585
      Width           =   9225
   End
End
Attribute VB_Name = "frmCierreStock"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
  Dim strRangoIso As isoLastPeriod
  
  'muestro solo ultimo periodo ingresado
  strRangoIso = adoLastPeriod("viewStockCierre", "fecha")
  
  strSQL = "SELECT * FROM viewStockCierre where fecha between " & strRangoIso.strDesde & " and " & strRangoIso.strHasta & " " & _
           "ORDER BY fecha DESC"
  
  intRes = ListViewAppearanceChange(lvwDatos)
  intRes = ListViewRefresh(lvwDatos, strSQL, strStruc)
  intRes = lvwHideColumn(lvwDatos, "proce")

End Sub

Private Sub tlbOperaciones_ButtonClick(ByVal Button As MSComctlLib.Button)

  blnRefresh = False
  
  Select Case Button.Key
    
  Case Is = "agregar"
  
    ' para cerrar mes, selecciona maximo de fecha cuando el status esta en 1 y cantidad en 3
    ' procesos fueron calculados, produccion mensual, stock subconcesiones, stock terminales
    Dim rs As New ADODB.Recordset
    
    ' abro vista con ultimo periodo a cerrar
    strSQL = "select * from stockCierreProcesados_view"
    Set rs = adoGetRS(strSQL)
    
    ' valido que exista
    If rs.EOF Or rs!Cantidad = 0 Then
      intRes = MsgBox("No se encuentra ningun periodo pendiente.", vbApplicationModal + vbOKOnly + vbInformation, "Informacion")
      Exit Sub
    End If
    
    ' confirmo cierre
    intRes = MsgBox("Esta seguro que desea cerrar el periodo " & rs!fecha & ".", vbApplicationModal + vbQuestion + vbYesNo, "Cerrando periodo...")
    If intRes = vbYes Then
      
      ' agrega cierre
      strSQL = "EXEC spStockCierreInsert " & _
               "'" & dateToIso(rs!fecha) & "'," & _
               "'CERR'" & "," & _
               2
      intRes = adoExecSQL(strSQL)
    
      blnRefresh = True
    
    End If
    
    ' cierro
    rs.Close
  
  Case Is = "modificar"
  
    intRes = MsgBox("Operación deshabilitada.", vbApplicationModal + vbOKOnly + vbInformation, "Informacion")
  
  Case Is = "eliminar"
  
    ' chequeo seleccion de algo
    If lvwDatos Is Nothing Then
      intRes = MsgBox("No hay ningún item seleccionado.", vbApplicationModal + vbOKOnly + vbInformation, "Informacion")
      Exit Sub
    End If
    
    ' chequeo que proceso no este cerrado
    If lvwGetValue(lvwDatos, "status") = "Cerrado" Then
      intRes = MsgBox("No es posible eliminarlo, el período esta cerrado.", vbApplicationModal + vbOKOnly + vbInformation, "Informacion")
      Exit Sub
    End If
  
    ' pregunto si deseo eliminar
    intRes = MsgBox("Esta por eliminar " & lvwGetValue(lvwDatos, "descripcion") & ", correspondiente al perirodo " & lvwGetValue(lvwDatos, "fecha") & ", esta seguro.", vbQuestion + vbYesNo, "Eliminando proceso...")
  
    If intRes = vbYes Then
    
      ' agrega cierre
      strSQL = "EXEC spStockCierreDelete " & _
               "'" & dateToIso(lvwGetValue(lvwDatos, "fecha")) & "'," & _
               "'" & lvwGetValue(lvwDatos, "proce") & "'"
      intRes = adoExecSQL(strSQL)
    
      blnRefresh = True
    
    End If
  
    Case Is = "buscar"
      intRes = FindData(lvwDatos)
  
    Case Is = "ordenar"
      intRes = lvwSortColumn(lvwDatos)
  
    Case Is = "filtrar"
    
      strWhere = FilterData(lvwDatos)
      If blnAceptar Then blnRefresh = True

    Case Is = "exportar"
     
      intRes = ExportData(lvwDatos)
  
    Case Is = "ajustar"
     
      ' ajusta y envia a INI
      intRes = lvwAdjustColumn(lvwDatos, True)
      intRes = lvwWidthToKeyIni(lvwDatos, strTableNameActual)
  
  End Select

  If blnRefresh Then
    
    strSQL = "SELECT * FROM ViewStockCierre" & _
             IIf(Not strWhere = "", " WHERE " & strWhere, "") & " " & _
             "ORDER BY fecha DESC"
    intRes = ListViewAppearanceChange(lvwDatos)
    intRes = ListViewRefresh(lvwDatos, strSQL)
    intRes = lvwHideColumn(lvwDatos, "proce")

  End If

End Sub
