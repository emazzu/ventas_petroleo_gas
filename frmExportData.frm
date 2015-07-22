VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmExportData 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Export..."
   ClientHeight    =   4620
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6060
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4620
   ScaleWidth      =   6060
   StartUpPosition =   2  'CenterScreen
   Begin MSComDlg.CommonDialog comNombre 
      Left            =   2745
      Top             =   3510
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton cmdCancelar 
      Caption         =   "&Cancel"
      Height          =   315
      Left            =   4920
      TabIndex        =   5
      Top             =   4140
      Width           =   960
   End
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "&Ok"
      Height          =   315
      Left            =   3800
      TabIndex        =   4
      Top             =   4140
      Width           =   960
   End
   Begin VB.CommandButton cmdRemove 
      Caption         =   "De&lete Column"
      Height          =   315
      Left            =   2340
      TabIndex        =   3
      Top             =   2790
      Width           =   1320
   End
   Begin VB.CommandButton cmdAdd 
      Caption         =   "&Add Column"
      Height          =   315
      Left            =   2340
      TabIndex        =   2
      Top             =   1215
      Width           =   1320
   End
   Begin VB.ListBox lstSelecting 
      ForeColor       =   &H00808080&
      Height          =   3570
      Left            =   3780
      MultiSelect     =   2  'Extended
      TabIndex        =   1
      Top             =   405
      Width           =   2085
   End
   Begin VB.ListBox lstSelect 
      ForeColor       =   &H00808080&
      Height          =   3570
      Left            =   135
      MultiSelect     =   2  'Extended
      TabIndex        =   0
      Top             =   405
      Width           =   2085
   End
   Begin VB.Label Label2 
      Caption         =   "Columns Selecting"
      Height          =   240
      Left            =   3825
      TabIndex        =   7
      Top             =   180
      Width           =   1995
   End
   Begin VB.Label Label1 
      Caption         =   "Select Columns"
      Height          =   240
      Left            =   180
      TabIndex        =   6
      Top             =   180
      Width           =   1995
   End
End
Attribute VB_Name = "frmExportData"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmdAceptar_Click()
  Dim intInd As Integer
  Dim blnSelected As Boolean

  ' chequeo seleccion de por lo menos 1 columna
  If lstSelecting.ListCount = 0 Then Exit Sub
  
  ' abro cuadro de dialogo para seleccionar nombre y tipo de formato a exportar
  comNombre.DialogTitle = "Exporting To..."
  comNombre.Filter = "Excel (*.xls)|*.xls"
  comNombre.FileName = ""
  comNombre.ShowSave

  ' si aceptar en cuadro de dialogo guardar como
  If comNombre.FileName <> "" Then
    blnAceptar = True
    blnCancelar = False
  Else
    blnAceptar = False
    blnCancelar = True
  End If
  
  Me.Hide

End Sub


Private Sub cmdAdd_Click()
  Dim intRow As Integer

  ' recorro elementos de la lista}
  For intRow = 0 To lstSelect.ListCount - 1

    ' si item seleccionado
    If lstSelect.Selected(intRow) Then
      
      ' agrego item a lstselecting
      lstSelecting.AddItem lstSelect.List(intRow)
    
    End If

  Next

  ' los elementos que pase a lstselecting los borro de lstselect
  ' para que el usuario no pueda seleccionar 2 veces el mismo
  intRow = lstSelect.ListCount
  While intRow > 0
    
    ' avanzo puntero
    intRow = intRow - 1
    
    ' si item seleccionado lo elimino de
    If lstSelect.Selected(intRow) = True Then
      
      ' elimino item a lstselect
      lstSelect.RemoveItem intRow
    
    End If
  
  Wend

End Sub

Private Sub cmdCancelar_Click()

  blnAceptar = False
  blnCancelar = True
  Me.Hide

End Sub

Private Sub cmdRemove_Click()

  Dim intRow As Integer

  ' recorro elementos de la lista}
  For intRow = 0 To lstSelecting.ListCount - 1

    ' si item seleccionado
    If lstSelecting.Selected(intRow) = True Then
      
      ' agrego item a lstselecting
      lstSelect.AddItem lstSelecting.List(intRow)
    
    End If

  Next

  ' los elementos que pase a lstselecting los borro de lstselect
  ' para que el usuario no pueda seleccionar 2 veces el mismo
  intRow = lstSelecting.ListCount
  While intRow > 0
    
    ' avanzo puntero
    intRow = intRow - 1
    
    ' si item seleccionado lo elimino de
    If lstSelecting.Selected(intRow) = True Then
      
      ' elimino item a lstselect
      lstSelecting.RemoveItem intRow
    
    End If
  
  Wend

End Sub

