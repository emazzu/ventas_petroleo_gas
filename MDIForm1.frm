VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.MDIForm MDIMenu 
   BackColor       =   &H8000000C&
   Caption         =   "Ventas v.2015.07.06"
   ClientHeight    =   6615
   ClientLeft      =   1800
   ClientTop       =   1680
   ClientWidth     =   11325
   Icon            =   "MDIForm1.frx":0000
   LinkTopic       =   "MDIForm1"
   ScrollBars      =   0   'False
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin MSComctlLib.StatusBar staGeneral 
      Align           =   2  'Align Bottom
      Height          =   285
      Left            =   0
      TabIndex        =   0
      Top             =   6330
      Width           =   11325
      _ExtentX        =   19976
      _ExtentY        =   503
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   4
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Text            =   "Filtro: ninguno"
            TextSave        =   "Filtro: ninguno"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Text            =   "Orden: ninguno"
            TextSave        =   "Orden: ninguno"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Text            =   "Info: "
            TextSave        =   "Info: "
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            AutoSize        =   2
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "MDIMenu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False



Private Sub MDIForm_Load()

  'check configuracion regional
  If Not checkConfigRegional() Then
    
    'show mensaje
    blnB = MsgBox("El sistema detecto que la configuración regional no es correcta." & vbCrLf & vbCrLf & _
           "Configurar el formato para números de esta forma: 123,456,789.00." & vbCrLf & vbCrLf & _
           "Configurar el formato para fechas  de esta forma: dd/MM/yyyy." & vbCrLf & vbCrLf & _
           "El sistema modificará la configuración regional automáticamente.", vbCritical + vbOKOnly, "Atención...")
        
    'change configuracion regional
    Call changeConfigRegional(".", ",", "dd/MM/yyyy")
      
  End If
    
  'SHOW en barra de estado usuario y grupo
  '22/08/2008 ya no se utiliza seguridad por roles, solo por grupos
  
  'Me.staGeneral.Panels(4).Text = " " & SQLparam.UsuarioLogon & " - " & SQLparam.Usuario & " - " & SQLparam.Role & " "
  
  Me.staGeneral.Panels(4).Text = " " & SQLparam.UsuarioLogon & " - " & SQLparam.Usuario
  
End Sub

Private Sub MDIForm_Resize()
  Dim a, intIndice, intLVW, intPositionLVW As Integer
  Dim Frm As Form

  ' chequeo que no se pase del minimo posible para ajustar
  If MDIMenu.Height < (MDIMenu.staGeneral.Height + 450) Then
    Exit Sub
  End If
  
  ' ajusto tamaño form menu 450 incluye borde de MDI y Alto barra de estado
  frmMenu.Height = MDIMenu.Height - (MDIMenu.staGeneral.Height + 500)
  frmMenu.dxSideBar1.Height = MDIMenu.Height - (MDIMenu.staGeneral.Height + 550)
  
  Set Frm = frmActivo
  
  ' ajusto tamaño form activo
  Frm.Width = MDIMenu.Width - frmMenu.Width
  Frm.Height = MDIMenu.Height
  
  ' cuanto cuentos listview tiene el form activo
  intLVW = 0
  For intIndice = 0 To Frm.Count - 1
    If TypeName(Frm.Controls(intIndice)) = "ListView" Then
      intLVW = intLVW + 1
      intPositionLVW = intIndice
    End If
  Next
  
  ' recorro controles
  For intIndice = 0 To Frm.Controls.Count - 1
    
    Select Case TypeName(Frm.Controls(intIndice))
    
    Case "Toolbar"
    
    Case "Label"
      Frm.Controls(intIndice).Width = MDIMenu.Width - frmMenu.Width - 150
    
    Case "ListView"
    
      ' si el form tiene solo un LVW lo acomoda a la izquiqerda arriba
      If intLVW = 1 Then
        'Frm.Controls(intIndice).Top = conLvwTop
        'Frm.Controls(intIndice).Left = conLvwLeft
        Frm.Controls(intIndice).Height = MDIMenu.Height - (conTlbHeight + conLblHeight + 700)
      Else
        If intIndice = intPositionLVW Then
          ' largo si es ultimo ListView del Form lo estira hasta el final de mdi
          ' los 700 incluye borde del formulario MDI y altura de barra de estado
          Frm.Controls(intIndice).Height = MDIMenu.Height - (conTlbHeight + conLblHeight + 700)
        End If
      End If
        
      ' ancho para todos ajuste al maximo de formulario
      Frm.Controls(intIndice).Width = MDIMenu.Width - frmMenu.Width - 150
    
    End Select
  
  Next

End Sub

Private Sub MDIForm_Unload(Cancel As Integer)
  
  Do While Not (Me.ActiveForm Is Nothing)
    Unload Me.ActiveForm
  Loop
  
End Sub
