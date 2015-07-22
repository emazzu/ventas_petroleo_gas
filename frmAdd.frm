VERSION 5.00
Begin VB.Form frmAdd 
   BorderStyle     =   0  'None
   ClientHeight    =   330
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4695
   LinkTopic       =   "Form1"
   ScaleHeight     =   330
   ScaleWidth      =   4695
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox txtDato 
      BackColor       =   &H00C0FFFF&
      ForeColor       =   &H00FF0000&
      Height          =   330
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   4695
   End
End
Attribute VB_Name = "frmAdd"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub txtDato_KeyPress(KeyAscii As Integer)
  
  Dim intRes As Integer

  If KeyAscii = 13 Then       ' sale cuando apreta enter
    
    If Me.txtDato <> "" Then  ' valido si ingreso algo
    
      If MsgBox("Esta seguro que desea agregar la información ingresada.", vbYesNo + vbQuestion, "Confirmación") = vbYes Then
        blnAceptar = True
        blnCancelar = False
      Else
        blnAceptar = False
        blnCancelar = True
      End If
    
    Else                      ' si no ingreso fuerzo cancelar
      blnAceptar = False
      blnCancelar = True
    End If
    
    Me.Hide
  
  End If

  If KeyAscii = 27 Then       ' escapar
    blnAceptar = False
    blnCancelar = True
    Me.Hide
  End If

End Sub

