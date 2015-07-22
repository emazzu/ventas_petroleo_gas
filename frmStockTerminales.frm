VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form frmStockTerminales 
   BorderStyle     =   0  'None
   Caption         =   "Stock Terminales"
   ClientHeight    =   6690
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   9690
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6690
   ScaleWidth      =   9690
   ShowInTaskbar   =   0   'False
   Begin MSComctlLib.ListView lvwDatos 
      Height          =   5055
      Left            =   0
      TabIndex        =   0
      Top             =   900
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
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   90
      Top             =   6030
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
            Picture         =   "frmStockTerminales.frx":0000
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmStockTerminales.frx":08DA
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmStockTerminales.frx":11B4
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmStockTerminales.frx":1A8E
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmStockTerminales.frx":3420
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmStockTerminales.frx":3CFA
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmStockTerminales.frx":45D4
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmStockTerminales.frx":5F66
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar tlbOperaciones 
      Height          =   570
      Left            =   0
      TabIndex        =   2
      Top             =   0
      Width           =   12195
      _ExtentX        =   21511
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
   Begin VB.Label lblTitulo 
      Alignment       =   2  'Center
      BackColor       =   &H00808000&
      Caption         =   "Stock Terminales"
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
      TabIndex        =   1
      Top             =   585
      Width           =   9225
   End
End
Attribute VB_Name = "frmStockTerminales"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
  Dim strRangoIso As isoLastPeriod
  
  'muestro solo ultimo periodo ingresado
  strRangoIso = adoLastPeriod("ViewTerminalesStock", "fecha")
  
  strSQL = "SELECT * FROM ViewTerminalesStock where fecha between " & strRangoIso.strDesde & " and " & strRangoIso.strHasta & " " & _
           "ORDER BY subconcesion, fecha, empresa"
  intRes = ListViewAppearanceChange(lvwDatos)
  intRes = ListViewRefresh(lvwDatos, strSQL, strStruc)
  intRes = lvwHideColumn(lvwDatos, "concesionID")
  intRes = lvwHideColumn(lvwDatos, "terminalID")
  intRes = lvwHideColumn(lvwDatos, "subconcesionID")
  intRes = lvwHideColumn(lvwDatos, "empresaID")

End Sub

Private Sub tlbOperaciones_ButtonClick(ByVal Button As MSComctlLib.Button)
  
  blnRefresh = False
  
  Select Case Button.Key
    
  Case Is = "agregar"
  
    Dim rsStock As ADODB.Recordset
    Dim strUltimo  As String
    Dim strNew As Date
  
    strSQL = "select top 1 fecha, status FROM StockCierre where proceso = 'STER' order by fecha desc"
    Set rsStock = adoGetRS(strSQL)
    
    ' cheque que existan movimientos
    If rsStock.EOF Then
      intRes = MsgBox("No existen movimientos, no se puede determinar el mes de cierre.", vbApplicationModal + vbInformation + vbOKOnly, "información...")
      Exit Sub
    End If
    
    ' chequea que ultimo movimiento este cerrado
    If rsStock!Status <> 2 Then
      intRes = MsgBox("El ultimo periodo encontrado es " & rsStock!fecha & ", pero todavia no fue cerrado.", vbApplicationModal + vbInformation + vbOKOnly, "información...")
      Exit Sub
    End If
  
    strNew = CDate(dateToLastDay(rsStock!fecha + 1))
    strUltimo = rsStock!fecha
    
    'cierro
    rsStock.Close
    
    'confirmo proceso
    intRes = MsgBox("El último período cerrado es " & strUltimo & ", el período a procesar será: " & strNew & ".", vbApplicationModal + vbQuestion + vbYesNo, "procesando stock...")
    If intRes = vbYes Then
      
      intRes = CalcularStockTerminales(strNew, "Si")
      blnRefresh = True
    
    End If
  
  Case Is = "modificar"
  
    intRes = MsgBox("Operación deshabilitada.", vbApplicationModal + vbOKOnly + vbInformation, "Informacion")
  
  Case Is = "eliminar"
  
    intRes = MsgBox("Operación deshabilitada.", vbApplicationModal + vbOKOnly + vbInformation, "Informacion")
  
  Case Is = "buscar"
  
    intRes = FindData(lvwDatos)
  
  Case Is = "filtrar"
    
    strWHERE = FilterData(lvwDatos)
    If blnAceptar Then blnRefresh = True
  
    Case Is = "ordenar"
      intRes = lvwSortColumn(lvwDatos)
  
  Case Is = "exportar"
     
    intRes = ExportData(lvwDatos)
  
  Case Is = "ajustar"
     
    ' ajusta y envia a INI
    intRes = lvwAdjustColumn(lvwDatos, True)
    intRes = lvwWidthToKeyIni(lvwDatos, strTableNameActual)
  
  End Select

  If blnRefresh Then
      
    If strWHERE = "" Then
      
      'muestro solo ultimo periodo ingresado
      Dim strRangoIso As isoLastPeriod
      strRangoIso = adoLastPeriod("ViewTerminalesStock", "fecha")
      strWHERE = "fecha between " & strRangoIso.strDesde & " and " & strRangoIso.strHasta
    
    End If
      
    strSQL = "SELECT * FROM ViewTerminalesStock" & _
             IIf(Not strWHERE = "", " WHERE " & strWHERE, "") & " " & _
             "ORDER BY subconcesion, fecha, empresa"
    intRes = ListViewAppearanceChange(lvwDatos)
    intRes = ListViewRefresh(lvwDatos, strSQL)
    intRes = lvwHideColumn(lvwDatos, "concesionID")
    intRes = lvwHideColumn(lvwDatos, "terminalID")
    intRes = lvwHideColumn(lvwDatos, "subconcesionID")
    intRes = lvwHideColumn(lvwDatos, "empresaID")

  End If

End Sub



