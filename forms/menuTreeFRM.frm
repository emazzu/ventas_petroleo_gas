VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form menuTreeFRM 
   Caption         =   "Menu"
   ClientHeight    =   5370
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   2895
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   5370
   ScaleWidth      =   2895
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   630
      Top             =   3960
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      MaskColor       =   12632256
      _Version        =   393216
   End
   Begin MSComctlLib.TreeView tvwMenu 
      Height          =   5460
      Left            =   -45
      TabIndex        =   0
      Top             =   -45
      Width           =   2985
      _ExtentX        =   5265
      _ExtentY        =   9631
      _Version        =   393217
      HideSelection   =   0   'False
      LabelEdit       =   1
      LineStyle       =   1
      Style           =   7
      HotTracking     =   -1  'True
      SingleSel       =   -1  'True
      ImageList       =   "ImageList1"
      Appearance      =   1
   End
End
Attribute VB_Name = "menuTreeFRM"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Form_Activate()

  'set botones menu invisibles
  MainMDI.tlbMenu.Buttons("insertar").Visible = False
  MainMDI.tlbMenu.Buttons("editar").Visible = False
  MainMDI.tlbMenu.Buttons("eliminar").Visible = False
  MainMDI.tlbMenu.Buttons("buscar").Visible = False
  MainMDI.tlbMenu.Buttons("rapido").Visible = False
  MainMDI.tlbMenu.Buttons("avanzado").Visible = False
  MainMDI.tlbMenu.Buttons("borrar").Visible = False
  MainMDI.tlbMenu.Buttons("excel").Visible = False
'  MainMDI.tlbMenu.Buttons("reportes").Visible = False

End Sub

Private Sub Form_Load()
  Dim rs As ADODB.Recordset
  Dim imagen As ListImage
  Dim nodo As Node

  ' llena treeView
  strSQL = "select * from menuOpciones where opcionTipo = 'TVW' and idAplicacion = " & conIDAplicacion & " order by father"
  Set rs = adoGetRS(strSQL, , "comun")

  While Not rs.EOF
      
    ' agrego el icono relacionado al control image
    Set imagen = ImageList1.ListImages.Add(, , LoadPicture(App.Path & "\iconos\" & rs!Icon))
    
    'image.Picture = LoadPicture(pps.Path & "\iconos\" & rs!Icon)
    
    ' si no tiene father es nodo padre
    If rs!father = 0 Then
      Set nodo = tvwMenu.Nodes.Add(, , "'" & rs!idmenuoption & "'", rs!menuOption, imagen.Index)
      nodo.Expanded = True
    Else
      ' si tiene padre es un hijo
      Set nodo = tvwMenu.Nodes.Add("'" & rs!father & "'", tvwChild, "'" & rs!idmenuoption & "'", rs!menuOption, imagen.Index)
      nodo.Expanded = True
    End If
    rs.MoveNext
  
  Wend

  ' definicion de variables de ubicacion
  Dim strUbico As Variant
  Dim sngLeft, sngTop, sngWidth, sngHeight As Single
 
  ' busco en ini si hay propiedades de ubicacion top, left, width, height
  strUbico = keyIniToArray(Me.Caption, "ubicacion")
  sngLeft = Val(arrayGetValue(strUbico, "left"))
  sngTop = Val(arrayGetValue(strUbico, "top"))
  sngWidth = Val(arrayGetValue(strUbico, "width"))
  sngHeight = Val(arrayGetValue(strUbico, "height"))
    
  tvwMenu.Indentation = 20      'separacion entre lineas verticales
  tvwMenu.HideSelection = False 'no se ve lo que se selecciono cuando pierde el foco
  tvwMenu.SingleSel = False     'cuando selecciono un item no comprime los demas
  tvwMenu.HotTracking = True    'estilo wew, manito con dedo en cada opcion
    
  ' cambio propiedades de ubicacion segun ini
  Me.Move sngLeft, sngTop, sngWidth, sngHeight

End Sub

Private Sub Form_Resize()

  ' ajusto tamaño de treeView segun FRM
  tvwMenu.Height = Me.Height
  tvwMenu.Width = Me.Width

End Sub

Private Sub Form_Unload(Cancel As Integer)

  strValor = "left;" & Me.Left & ";top;" & Me.Top & ";width;" & Me.Width & ";height;" & Me.Height
  intRes = WriteIni(Me.Caption, "ubicacion", strValor)

End Sub

Private Sub tvwMenu_NodeClick(ByVal Node As MSComctlLib.Node)
  Dim frm As gridFRM
  Dim a As Form
  Dim rs As ADODB.Recordset
  
  ' abro FRM segun item seleccionado en el treeView
  strSQL = "select * from menuOpciones where idaplicacion = " & conIDAplicacion & " and father is not null and menuOption = '" & tvwMenu.SelectedItem.Text & "'"
  Set rs = adoGetRS(strSQL, , "comun")

  ' si no encuentra algo o la columna DataSource esta vacia exit
  If rs.EOF Or rs!DataSource = "" Then
    Exit Sub
  End If
    
  ' chequeo que no haya sido abierto antes
  Dim frmOpen As Form
  Dim blnOpen As Boolean
    
  blnOpen = True
  For Each frmOpen In Forms
    If frmOpen.Caption = rs!CaptionFrm Then
      'si esta abierto lo activa
      frmOpen.SetFocus
      blnOpen = False
    End If
  Next frmOpen
    
  ' si ha sido abierto antes exit
  If Not blnOpen Then
    Exit Sub
  End If
  
  If rs!Namefrm = "" Then
    'form standard
    Set frm = New gridFRM
  Else
    'form armado
    Set frm = Forms.Add(rs!Namefrm)
  End If
      
  ' definicion de variables de ubicacion
  Dim strGet  As Variant
  Dim strDataWhere, strDataOrder, strDataTopMax As String
  Dim sngLeft, sngTop, sngWidth, sngHeight As Single
  
  ' busco en ini si hay propiedades de ubicacion
  strGet = keyIniToArray(rs!CaptionFrm, "ubicacion")
  
  sngLeft = Val(arrayGetValue(strGet, "left"))
  sngTop = Val(arrayGetValue(strGet, "top"))
  sngWidth = Val(arrayGetValue(strGet, "width"))
  sngHeight = Val(arrayGetValue(strGet, "height"))
  
  'muevo ubicacion de frm segun ini
  frm.Move sngLeft, sngTop, sngWidth, sngHeight
      
  ' busco en ini si hay valores para propiedades de datos
  strGet = keyIniToArray(rs!CaptionFrm, "data")
      
  strDataWhere = arrayGetValue(strGet, "datawhere")
  strDataOrder = arrayGetValue(strGet, "dataorder")
  strDataTopMax = Val(arrayGetValue(strGet, "datatopmax"))
      
  'asigno propiedades al frm
  frm.Caption = rs!CaptionFrm
  frm.DataSource = rs!DataSource
  frm.DataWhere = strDataWhere
  frm.DataOrder = strDataOrder
  frm.DataMaximo = rs!DataMaximo
  frm.DataTopMax = strDataTopMax
  frm.DataStoreProcedure = rs!StoreProcedure
  frm.DataNoMuestraEnGrilla = rs!NoMuestraEnGrilla
  frm.DataNoMuestraEnEdit = rs!NoMuestraEnEdit
  frm.DataSoloLecturaEnEdit = rs!soloLecturaEnEdit
  frm.DataObligatorioEnEdit = rs!ObligatorioEnEdit
  frm.DataComboBox = rs!ComboBox
  frm.DataImportRelaciones = rs!DataImportRelaciones
  
  ' hago un refresh de datos
  frm.DataRefresh = True
  
End Sub
