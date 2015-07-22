VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmParametros 
   BorderStyle     =   0  'None
   Caption         =   "Parametros"
   ClientHeight    =   6855
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   13215
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   6855
   ScaleWidth      =   13215
   ShowInTaskbar   =   0   'False
   Begin MSComctlLib.ListView lvwDatos 
      Height          =   5055
      Left            =   0
      TabIndex        =   0
      Top             =   855
      Width           =   9240
      _ExtentX        =   16298
      _ExtentY        =   8916
      View            =   3
      Sorted          =   -1  'True
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
      Left            =   45
      Top             =   6210
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
            Picture         =   "frmParametros.frx":0000
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmParametros.frx":08DA
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmParametros.frx":11B4
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmParametros.frx":1A8E
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmParametros.frx":3420
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmParametros.frx":3CFA
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmParametros.frx":45D4
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmParametros.frx":5F66
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar tlbOperaciones 
      Height          =   570
      Left            =   0
      TabIndex        =   2
      Top             =   0
      Width           =   11655
      _ExtentX        =   20558
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
      Caption         =   "Parametros"
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
Attribute VB_Name = "frmParametros"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Form_Load()
  Dim strQuery As String

  strQuery = "SELECT * FROM ViewParametros"
  intRes = ListViewAppearanceChange(lvwDatos)
  intRes = ListViewRefresh(lvwDatos, strQuery, strStruc)
  intRes = lvwHideColumn(lvwDatos, "idparametro")
  intRes = lvwHideColumn(lvwDatos, "editable")

End Sub

Private Sub tlbOperaciones_ButtonClick(ByVal Button As MSComctlLib.Button)

  Dim strSQL As String
  Dim intRes As Integer
  
  blnRefresh = False
  
  Select Case Button.Key
    
  Case Is = "agregar"
  
  Case Is = "modificar"
  
    If lvwDatos Is Nothing Then
      a = MsgBox("No hay ningún item seleccionado.", vbApplicationModal + vbOKOnly + vbInformation, "Informacion")
      Exit Sub
    End If
    
    'cargo formulario
    Load frmParametrosUpdate
          
    'paso valores para poder editar
    frmParametrosUpdate.txtDato1 = lvwGetValue(lvwDatos, "parametro")
    frmParametrosUpdate.txtDato2 = lvwGetValue(lvwDatos, "valor")
    frmParametrosUpdate.txtDato2.SelLength = Len(frmParametrosUpdate.txtDato2)
    
    If Val(lvwGetValue(lvwDatos, "editable")) = 1 Then
      frmParametrosUpdate.txtDato2.Enabled = False
    Else
      frmParametrosUpdate.txtDato1.Enabled = False
    End If
    
    'muestro formulario
    frmParametrosUpdate.Show vbModal
      
    If blnAceptar Then
      
      'si hizo click en Aceptar, genero string y ejecuto
      'funcion de UPDATE, el primer argumento enviado es
      'el campo clave por el cual aplica el WHERE
        
      With frmParametrosUpdate
      strSQL = "EXEC spParametrosUpdate " & _
      Val(lvwGetValue(lvwDatos, "idparametro")) & "," & _
      "'" & dateToIso(Date) & "'," & _
      "'" & .txtDato1 & "'," & _
      Val(.txtDato2)
      End With
        
      a = adoExecSQL(strSQL)
      blnRefresh = True
 
    End If
      
      ' descargo formulario
      Unload frmParametrosUpdate
  
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
      
    strSQL = "SELECT * FROM Viewparametros" & _
    IIf(Not strWhere = "", " WHERE " & strWhere, "")
    intRes = ListViewRefresh(lvwDatos, strSQL)
    intRes = lvwHideColumn(lvwDatos, "idparametro")
    intRes = lvwHideColumn(lvwDatos, "editable")

  End If

End Sub
