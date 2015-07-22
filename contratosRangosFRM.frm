VERSION 5.00
Object = "{F856EC8B-F03C-4515-BDC6-64CBD617566A}#6.0#0"; "fpspr60.ocx"
Begin VB.Form contratosRangosFRM 
   BorderStyle     =   5  'Sizable ToolWindow
   Caption         =   "Rangos"
   ClientHeight    =   2865
   ClientLeft      =   60
   ClientTop       =   330
   ClientWidth     =   7950
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2865
   ScaleWidth      =   7950
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin FPSpreadADO.fpSpread spd 
      Height          =   2055
      Left            =   90
      TabIndex        =   0
      Top             =   120
      Width           =   7755
      _Version        =   393216
      _ExtentX        =   13679
      _ExtentY        =   3625
      _StockProps     =   64
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      SpreadDesigner  =   "contratosRangosFRM.frx":0000
   End
   Begin VB.Menu mnu 
      Caption         =   "mnu"
      Visible         =   0   'False
      Begin VB.Menu mnuINS 
         Caption         =   "Insertar Linea"
      End
      Begin VB.Menu mnuELI 
         Caption         =   "Eliminar Linea"
      End
   End
End
Attribute VB_Name = "contratosRangosFRM"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Form_Activate()

  spd.ReDraw = False

  'alineo operadores al centro
  spd.Col = 4
  spd.TypeHAlign = TypeHAlignCenter
  spd.Col = 6
  spd.TypeHAlign = TypeHAlignCenter

  'ajusto decimales en valores
  spd.Col = 5
  spd.TypeNumberDecPlaces = 3
  spd.Col = 7
  spd.TypeNumberDecPlaces = 3
    
  spd.ReDraw = True

End Sub

Private Sub Form_Load()
  
  'cambio apariencia de grilla
  spd.UnitType = UnitTypeTwips
  spd.OperationMode = OperationModeRow
  spd.EditModeReplace = True
  spd.EditEnterAction = EditEnterActionRight
  spd.SelBackColor = RGB(210, 230, 210)
  spd.GrayAreaBackColor = spd.BackColor
      
  'escondo col 1
  spd.Col = 1
  spd.ColHidden = True
    
  'altura titulo
  spd.RowHeight(0) = 300
    
  Dim rs As New ADODB.Recordset
  Set rs = adoGetRS("select * from contratosRangos order by nombre")
  
  'chequeo errores
  If Not lngAdoErrNum = -1 Then
    adoError
    Exit Sub
  End If
   
  'asigno rs a grilla
  Set spd.DataSource = rs
    
  'set maximo filas igual a rs
  spd.MaxRows = rs.RecordCount
    
  'bloqueo operadores
  spd.Col = 4
  spd.Lock = True
  spd.Col = 6
  spd.Lock = True
    
End Sub

Private Sub Form_Resize()
  
  spd.ReDraw = False
    
  'ajuste spd
  spd.Top = Me.ScaleTop
  spd.Left = Me.ScaleLeft
  spd.Width = Me.ScaleWidth
  spd.Height = Me.ScaleHeight
  
  Dim lngAncho As Long
  lngAncho = (spd.Width - spd.ColWidth(0) - (spd.ColWidth(0) / 1.2) - 600) / 6
  
  'ajusto cols
  spd.ColWidth(2) = lngAncho
  spd.ColWidth(3) = lngAncho
  spd.ColWidth(5) = lngAncho
  spd.ColWidth(7) = lngAncho
  spd.ColWidth(8) = lngAncho
  spd.ColWidth(9) = lngAncho
  
  spd.ColWidth(4) = 300
  spd.ColWidth(6) = 300
  
  spd.ReDraw = True
  
End Sub

Private Sub mnuELI_Click()
  
  If spd.ActiveRow > 0 Then
  
    Dim varID As Variant
    spd.GetText 1, spd.ActiveRow, varID
        
    'elimino
    adoExecSQL ("exec contratosRangos_ELI_sp " & varID)
    
    'chequeo errores
    If Not lngAdoErrNum = -1 Then
      adoError
      Exit Sub
    End If
    
    'elimino linea de grilla, hago todo esto porque sino no borra
    spd.MaxRows = spd.MaxRows + 1
    spd.InsertRows spd.ActiveRow, 1
    spd.DeleteRows spd.ActiveRow, 2
    spd.MaxRows = spd.MaxRows - 2
    
  End If
  
End Sub

Private Sub mnuINS_Click()
  
  If spd.ActiveRow > -1 Then
    
    'genero id nuevo
    Dim rs As ADODB.Recordset
    Dim intN As Integer
    Dim strT As String
        
    'leo id maximo
    Set rs = adoGetRS("select max(id) from contratosRangos")
    
    'chequeo errores
    If Not lngAdoErrNum = -1 Then
      adoError
      Exit Sub
    End If
    
    'si es primer caso id = 1
    If IsNull(rs(0)) Then
      intN = 1
    Else
      intN = rs(0) + 1
    End If
      
    'agrego en tabla
    strT = "exec contratosRangos_INS_sp " & intN
    adoExecSQL (strT)
    
    'chequeo errores
    If Not lngAdoErrNum = -1 Then
      adoError
      Exit Sub
    End If

    'inserto linea en grilla
    spd.MaxRows = spd.MaxRows + 1
    spd.InsertRows spd.ActiveRow, 1
    
    'asigno id
    spd.SetText 1, spd.ActiveRow, intN
        
    'fuerzo operadores
    spd.Row = spd.ActiveRow
    spd.Col = 4
    spd.Text = ">="
    spd.Col = 6
    spd.Text = "<="
        
  End If

End Sub

Private Sub spd_Change(ByVal Col As Long, ByVal Row As Long)
    
  Dim strT As String
  
  Dim varID, varNombre, varFormula, varCond1, varCond2, varCoef, varCiefTest, varCoefIIBB
      
  'chequeo que haya seleccionado alguna fila
  If spd.ActiveRow > 0 Then
    
    spd.GetText 1, spd.ActiveRow, varID
    spd.GetText 2, spd.ActiveRow, varNombre
    spd.GetText 3, spd.ActiveRow, varFormula
    spd.GetText 5, spd.ActiveRow, varCond1
    spd.GetText 7, spd.ActiveRow, varCond2
    spd.GetText 8, spd.ActiveRow, varCoef
    spd.GetText 9, spd.ActiveRow, varCoefIIBB
            
    'chequeo que esten todos las columnas necesarias con info y despues update
    If varNombre <> "" And varFormula <> "" And Val(varCond1) <> 0 And Val(varCond2) <> 0 And varCoef <> "" And varCoefIIBB <> "" Then
      
      'actualizo en tabla
      strT = "exec contratosRangos_EDI_sp " & varID & ",'" & varNombre & "','" & varFormula & "'," & varCond1 & "," & varCond2 & ",'" & varCoef & "','" & varCoefIIBB & "'"
      adoExecSQL (strT)
      
      'chequeo errores
      If Not lngAdoErrNum = -1 Then
        adoError
        Exit Sub
      End If
        
    End If
      
  End If
     
End Sub

Private Sub spd_DataColConfig(ByVal Col As Long, ByVal DataField As String, ByVal DataType As Integer)
  
  'set ancho de columna a 100 para formulas
  If Col = 8 Or Col = 9 Then
    
    'set puntero fila, col
    spd.Row = -1
    spd.Col = Col
      
    'set maximo 100
    spd.TypeMaxEditLen = 100
  
  End If
  
End Sub

Private Sub spd_DblClick(ByVal Col As Long, ByVal Row As Long)
  
  'nombre de tabla
  If Col = 2 Then
    
    'nombre de tabla como comboBox
    spd.Row = Row
    spd.Col = Col
    spd.CellType = CellTypeComboBox
    spd.TypeComboBoxEditable = True
    
    Dim rs As ADODB.Recordset
    Dim str As String
    
    Set rs = adoGetRS("select nombre from contratosRangos group by nombre")
    
    'chequeo errores
    If Not lngAdoErrNum = -1 Then
      adoError
      Exit Sub
    End If
    
    str = ""
    While Not rs.EOF
      
      If str <> "" Then
        str = str & Chr(9)
      End If
      
      str = str & rs!Nombre
      
      rs.MoveNext
    
    Wend
    
    spd.TypeComboBoxList = str
    
  End If
    
  'formula
  If Col = 3 Then
    
    spd.Row = Row
    spd.Col = Col
    spd.CellType = CellTypeComboBox
    spd.TypeComboBoxEditable = False
    spd.TypeComboBoxList = "AVG-DISC" & Chr(9) & "AVG" & Chr(9) & "DISC"
  
  End If
  
End Sub

Private Sub spd_RightClick(ByVal ClickType As Integer, ByVal Col As Long, ByVal Row As Long, ByVal MouseX As Long, ByVal MouseY As Long)
  
  PopupMenu mnu
  
End Sub

