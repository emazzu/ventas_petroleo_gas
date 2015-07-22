VERSION 5.00
Begin VB.Form frmFilterData 
   BackColor       =   &H80000013&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Filtrar..."
   ClientHeight    =   2670
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7455
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2670
   ScaleWidth      =   7455
   StartUpPosition =   2  'CenterScreen
   Begin VB.ComboBox cboLogica 
      Enabled         =   0   'False
      ForeColor       =   &H00808080&
      Height          =   315
      ItemData        =   "frmFilterData.frx":0000
      Left            =   90
      List            =   "frmFilterData.frx":000A
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   315
      Width           =   870
   End
   Begin VB.TextBox txtDato 
      Height          =   315
      Left            =   4275
      TabIndex        =   11
      Top             =   315
      Width           =   2085
   End
   Begin VB.CommandButton cmdAnular 
      Caption         =   "An&ula Item"
      Height          =   315
      Left            =   90
      TabIndex        =   4
      Top             =   2250
      Width           =   1185
   End
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "&Aceptar"
      Height          =   300
      Left            =   6480
      TabIndex        =   3
      Top             =   315
      Width           =   870
   End
   Begin VB.CommandButton cmdCancelar 
      Caption         =   "&Cancelar"
      Height          =   315
      Left            =   4860
      TabIndex        =   7
      Top             =   2250
      Width           =   1185
   End
   Begin VB.CommandButton cmdAplicar 
      Caption         =   "A&plicar"
      Height          =   315
      Left            =   6165
      TabIndex        =   6
      Top             =   2250
      Width           =   1185
   End
   Begin VB.ListBox lstCondicion 
      ForeColor       =   &H00808080&
      Height          =   1425
      Left            =   90
      TabIndex        =   5
      Top             =   720
      Width           =   7260
   End
   Begin VB.ComboBox cboOperacion 
      ForeColor       =   &H00808080&
      Height          =   315
      ItemData        =   "frmFilterData.frx":0017
      Left            =   2880
      List            =   "frmFilterData.frx":0021
      Style           =   2  'Dropdown List
      TabIndex        =   2
      Top             =   315
      Width           =   1365
   End
   Begin VB.ComboBox cboColumna 
      ForeColor       =   &H00808080&
      Height          =   315
      Left            =   990
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   315
      Width           =   1860
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "Logica"
      Height          =   195
      Left            =   135
      TabIndex        =   12
      Top             =   90
      Width           =   780
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Operación"
      Height          =   240
      Left            =   2925
      TabIndex        =   10
      Top             =   90
      Width           =   1275
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Columna"
      Height          =   240
      Left            =   4320
      TabIndex        =   9
      Top             =   90
      Width           =   2040
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Columna"
      Height          =   195
      Left            =   990
      TabIndex        =   8
      Top             =   90
      Width           =   1815
   End
End
Attribute VB_Name = "frmFilterData"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim strFmat As String



Private Sub cboColumna_Click()

  ' se selecciono algo ?
  If Me.cboColumna.ListIndex = -1 Then Exit Sub
  
    ' determino que tipo de datos tiene la columna seleccionada
    Select Case strStruc(2, Me.cboColumna.ListIndex)

    Case "numeric"
      
      ' agrego las posibles operaciones en el combo
      Me.cboOperacion.Clear
      Me.cboOperacion.AddItem "="
      Me.cboOperacion.AddItem ">"
      Me.cboOperacion.AddItem "<"
      Me.cboOperacion.AddItem ">="
      Me.cboOperacion.AddItem "<="
      Me.cboOperacion.AddItem "<>"
      
      ' tomo el formato de la columna seleccionada para validacion
      strFmat = strStruc(3, Me.cboColumna.ListIndex)
      
    Case "string"
      
      Me.cboOperacion.Clear
      Me.cboOperacion.AddItem "like"
      Me.cboOperacion.AddItem "not like"
      
      ' tomo el formato de la columna seleccionada para validacion
      strFmat = strStruc(3, Me.cboColumna.ListIndex)
    
    Case "date"
    
      Me.cboOperacion.Clear
      Me.cboOperacion.AddItem "="
      Me.cboOperacion.AddItem ">="
      Me.cboOperacion.AddItem "<="
      Me.cboOperacion.AddItem "<>"
    
      ' tomo el formato de la columna seleccionada para validacion
      strFmat = strStruc(3, Me.cboColumna.ListIndex)
    
    Case "boolean"
    
      Me.cboOperacion.Clear
      Me.cboOperacion.AddItem "True"
      Me.cboOperacion.AddItem "False"
    
      ' tomo el formato de la columna seleccionada para validacion
      strFmat = strStruc(3, Me.cboColumna.ListIndex)
    
    End Select

End Sub

Private Sub cmdAceptar_Click()
  Dim strParcial As String

  ' valido seleccion de algun operador logico
  If Me.cboLogica.Enabled = True And Me.cboLogica.ListIndex = -1 Then
    Exit Sub
  End If

  ' valido seleccion de alguna columna
  If Me.cboColumna.ListIndex = -1 Then
    Exit Sub
  End If
  
  ' valido seleccion de alguna operacion
  If Me.cboOperacion.ListIndex = -1 Then
    Exit Sub
  End If
  
  ' valido ingreso de algun dato
  If Not DataValidate(Me.txtDato, strFmat, True) Then Exit Sub

  ' agrego nombre de columna a la condicion parcial
  strParcial = strStruc(1, Me.cboColumna.ListIndex)

  ' agrego operacion a la condicion parcial
  strParcial = strParcial & " " & Me.cboOperacion.List(Me.cboOperacion.ListIndex)
  
  ' agrego dato a la condicion parcial
  Select Case strStruc(2, Me.cboColumna.ListIndex)

  Case "numeric"
      
      ' si es numeric lo agrego con la funcion val
      strParcial = strParcial & " " & Me.txtDato
      
    Case "string"
      
      ' si es string lo agrego con comillas
      strParcial = strParcial & " '" & Me.txtDato & "%'"
    
    Case "date"
    
      ' si es string lo agrego con la funcion datetoiso y comillas
      strParcial = strParcial & " '" & dateToIso(Me.txtDato) & "'"
    
    Case "boolean"
    
      ' si es string lo agrego con la funcion datetoiso y comillas
      'strParcial = strParcial & " '" & dateToIso(Me.txtDato) & "'"
    
    End Select
  
  ' si la lista de condiciones esta vacia agrego la primera condicion
  If Me.lstCondicion.ListCount = 0 Then
  
    ' agrego condicion parcial a lista de condiciones
    Me.lstCondicion.AddItem strParcial
    ' habilito la condicion logica AND y OR
    Me.cboLogica.Enabled = True
  
  Else
  
    ' si la operacion logica es OR inserto condicion parcial
    ' en el mismo item de la lista a continuacion del existente
    ' si la operacion logica es AND inserto condicion parcial
    ' en el item siguiente de la lista
    If Me.cboLogica = "and" Then
      
      ' agrego al ultimo item la logica AND y lo cierro entre parentesis
      Me.lstCondicion.List(Me.lstCondicion.ListCount - 1) = "(" & Me.lstCondicion.List(Me.lstCondicion.ListCount - 1) & ") " & _
      Me.cboLogica.List(Me.cboLogica.ListIndex)
      ' agrego la condicion nueva en el item siguiente
      Me.lstCondicion.AddItem strParcial
    
    Else
    
      ' agrego al ultimo item la logica OR y la nueva condicion a continuacion
      Me.lstCondicion.List(Me.lstCondicion.ListCount - 1) = Me.lstCondicion.List(Me.lstCondicion.ListCount - 1) & " " & _
      Me.cboLogica.List(Me.cboLogica.ListIndex) & " " & _
      strParcial
    
    End If
  
  End If

End Sub

Private Sub cmdAnular_Click()

  ' valido seleccion de algun item
  If Me.lstCondicion.ListIndex = -1 Then Exit Sub
  
  ' si item seleccionado es el ultimo de la lista, pero no es el unico
  ' borro condicion logica AND y los parentesis () en el item anterior
  If Me.lstCondicion.ListIndex = Me.lstCondicion.ListCount - 1 And Me.lstCondicion.ListCount <> 1 Then
    Me.lstCondicion.List(Me.lstCondicion.ListIndex - 1) = Replace(Me.lstCondicion.List(Me.lstCondicion.ListIndex - 1), "(", "")
    Me.lstCondicion.List(Me.lstCondicion.ListIndex - 1) = Replace(Me.lstCondicion.List(Me.lstCondicion.ListIndex - 1), ") and", "")
  End If
  
  ' borro item seleccionado
  Me.lstCondicion.RemoveItem Me.lstCondicion.ListIndex
  
  ' si lista vacia ? deshabilito condicion logica
  If Me.lstCondicion.ListCount = 0 Then
    Me.cboLogica.Enabled = False
  End If

End Sub

Private Sub cmdAplicar_Click()

  blnAceptar = True
  blnCancelar = False
  Me.Hide

End Sub

Private Sub cmdCancelar_Click()

  blnAceptar = False
  blnCancelar = True
  Me.Hide

End Sub

