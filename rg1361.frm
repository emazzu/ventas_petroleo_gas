VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{F856EC8B-F03C-4515-BDC6-64CBD617566A}#6.0#0"; "fpSpr60.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form rg1361 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "RG 1361"
   ClientHeight    =   7890
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   12195
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7890
   ScaleWidth      =   12195
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Height          =   915
      Left            =   90
      TabIndex        =   6
      Top             =   45
      Width           =   12030
      Begin MSComDlg.CommonDialog comDestino 
         Left            =   6930
         Top             =   315
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
      Begin VB.CommandButton cmdExportar 
         Height          =   330
         Left            =   4275
         Picture         =   "rg1361.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   3
         ToolTipText     =   "Exportar a Excel"
         Top             =   450
         Width           =   375
      End
      Begin VB.CommandButton cmdSalir 
         Caption         =   "&Salir"
         Height          =   330
         Left            =   10440
         TabIndex        =   4
         Top             =   450
         Width           =   1455
      End
      Begin VB.CommandButton cmdAplicar 
         Height          =   330
         Left            =   3600
         Picture         =   "rg1361.frx":038A
         Style           =   1  'Graphical
         TabIndex        =   2
         ToolTipText     =   "Actualizar"
         Top             =   450
         Width           =   375
      End
      Begin MSComCtl2.DTPicker dtpD 
         Height          =   330
         Left            =   135
         TabIndex        =   0
         Top             =   450
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   582
         _Version        =   393216
         Format          =   49152001
         CurrentDate     =   39258
      End
      Begin MSComCtl2.DTPicker dtpH 
         Height          =   330
         Left            =   1845
         TabIndex        =   1
         Top             =   450
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   582
         _Version        =   393216
         Format          =   49152001
         CurrentDate     =   39258
      End
      Begin VB.Label Label2 
         Caption         =   "Hasta"
         Height          =   195
         Left            =   1890
         TabIndex        =   8
         Top             =   225
         Width           =   1050
      End
      Begin VB.Label Label1 
         Caption         =   "Desde"
         Height          =   195
         Left            =   180
         TabIndex        =   7
         Top             =   225
         Width           =   1050
      End
   End
   Begin FPSpreadADO.fpSpread spd 
      Height          =   6765
      Left            =   90
      TabIndex        =   5
      Top             =   1035
      Width           =   12030
      _Version        =   393216
      _ExtentX        =   21220
      _ExtentY        =   11933
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
      SpreadDesigner  =   "rg1361.frx":0714
   End
End
Attribute VB_Name = "rg1361"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmdAplicar_Click()
  
  Dim intRes As Integer
  Dim rs As ADODB.Recordset
  Dim strT As String
  
  'adjust columnas
  Me.spd.AutoSize = True
  Me.spd.DAutoSizeCols = DAutoSizeColsMax
  
  'set tamaño letra
  Me.spd.FontSize = 9
    
  'modo seleccion multiple
  Me.spd.OperationMode = OperationModeRead
    
  'build datos
  strT = "exec rg1361lineaFinal_sp '" & dateToIso(Me.dtpD.Value) & "','" & dateToIso(Me.dtpH.Value) & "'"
  adoExecSQL (strT)
  
  'check error
  If Not lngAdoErrNum = -1 Then
    adoError
    Exit Sub
  End If
  
  'get rs
  strT = "select * from rg1361final order by factura"
  Set rs = adoGetRS(strT)
  
  'check error
  If Not lngAdoErrNum = -1 Then
    adoError
    Exit Sub
  End If
  
  'fill spread con rs
  Me.spd.DataSource = rs.DataSource
      
  'set limite de filas
  Me.spd.MaxRows = Me.spd.DataRowCnt
      
  'delete ultima columna
'  Me.spd.DeleteCols Me.spd.DataColCnt, 1
'  Me.spd.MaxCols = Me.spd.MaxCols - 1
      
  'close rs
  rs.Close
    
  'show 3 decimales, no show coma de liles y simbolo $
  Me.spd.BlockMode = True
  Me.spd.Col = 2
  Me.spd.Col2 = 5
  Me.spd.Row = 1
  Me.spd.Row2 = Me.spd.DataRowCnt
  Me.spd.TypeCurrencyDecPlaces = 3
  Me.spd.TypeCurrencyShowSep = False
  Me.spd.TypeCurrencyShowSymbol = False
  
'  Me.spd.Col = 7
'  Me.spd.Col2 = 7
'  Me.spd.TypeCurrencyDecPlaces = 3
'  Me.spd.TypeCurrencyShowSep = False
'  Me.spd.TypeCurrencyShowSymbol = False
  
  Me.spd.BlockMode = False
    
End Sub

Private Sub cmdExportar_Click()

  Dim blnB As Boolean
  Dim strT As String
    
  'filter xls, sino txt
  Me.comDestino.Filter = "Archivos de Excel (*.xls)|*.xls"
    
  'titulo de ventana
  Me.comDestino.DialogTitle = "Exportando..."
  
  Me.comDestino.FileName = ""
  
  'abro cuadro de dialogo
  Me.comDestino.ShowSave
      
  'si cancelar salgo
  If Me.comDestino.FileName = "" Then
    Exit Sub
  End If
    
  'mouse reloj
  Screen.MousePointer = vbHourglass
  
  'esto es para no generar una planilla bloqueada y exporto
  Me.spd.Protect = False
  blnB = Me.spd.ExportToExcel(Me.comDestino.FileName, "", "")
  
  'mouse defa
  Screen.MousePointer = vbDefault
    
  'status
  If blnB Then
    blnB = MsgBox("La exportación se realizó con éxito.", vbInformation + vbOKOnly, "atención...")
  Else
    blnB = MsgBox("La exportación fallo.", vbCritical + vbOKOnly, "atención...")
  End If
    
End Sub

Private Sub cmdSalir_Click()
  
  'exit
  Unload Me
  
End Sub

