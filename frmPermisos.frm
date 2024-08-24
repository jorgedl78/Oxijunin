VERSION 5.00
Begin VB.Form frmPermisos 
   Caption         =   "Permisos de Usuario"
   ClientHeight    =   6420
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   9780
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   ScaleHeight     =   6420
   ScaleWidth      =   9780
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdCancelar 
      Caption         =   "Cancelar"
      Height          =   855
      Left            =   5880
      Picture         =   "frmPermisos.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   5400
      Width           =   975
   End
   Begin VB.CommandButton cmdGuardar 
      Caption         =   "Guardar"
      Height          =   855
      Left            =   2880
      Picture         =   "frmPermisos.frx":08CA
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   5400
      Width           =   975
   End
   Begin VB.Frame Frame1 
      Caption         =   "Permisos"
      Height          =   3015
      Left            =   240
      TabIndex        =   14
      Top             =   2160
      Width           =   9015
      Begin VB.CheckBox chVerInformes 
         Caption         =   "Ver Informes"
         Height          =   375
         Left            =   360
         TabIndex        =   15
         Top             =   2400
         Width           =   1695
      End
      Begin VB.CheckBox chFacturar 
         Caption         =   "Facturar"
         Height          =   375
         Left            =   360
         TabIndex        =   7
         Top             =   1920
         Width           =   2535
      End
      Begin VB.CheckBox chModificarArticulos 
         Caption         =   "Modificar Artículos"
         Height          =   375
         Left            =   360
         TabIndex        =   6
         Top             =   1440
         Width           =   2535
      End
      Begin VB.CheckBox chModificarClientes 
         Caption         =   "ModificarClientes"
         Height          =   375
         Left            =   360
         TabIndex        =   5
         Top             =   960
         Width           =   2535
      End
      Begin VB.CheckBox chEstablecerUsuariosyPermisos 
         Caption         =   "Establecer Usuarios y Permisos"
         Height          =   375
         Left            =   360
         TabIndex        =   4
         Top             =   480
         Width           =   2535
      End
   End
   Begin VB.TextBox txtEmail 
      Height          =   375
      Left            =   5760
      TabIndex        =   3
      Top             =   1320
      Width           =   3135
   End
   Begin VB.TextBox txttelefono 
      Height          =   375
      Left            =   1080
      TabIndex        =   2
      Top             =   1320
      Width           =   3135
   End
   Begin VB.TextBox txtDomicilio 
      Height          =   375
      Left            =   5760
      TabIndex        =   1
      Top             =   600
      Width           =   3135
   End
   Begin VB.TextBox txtUsuario 
      Height          =   375
      Left            =   1080
      TabIndex        =   0
      Top             =   600
      Width           =   3135
   End
   Begin VB.Label Label4 
      Caption         =   "E_mail:"
      Height          =   255
      Left            =   5040
      TabIndex        =   13
      Top             =   1440
      Width           =   615
   End
   Begin VB.Label Label3 
      Caption         =   "Teléfonos:"
      Height          =   255
      Left            =   240
      TabIndex        =   12
      Top             =   1440
      Width           =   855
   End
   Begin VB.Label Label2 
      Caption         =   "Domicilio:"
      Height          =   255
      Left            =   4920
      TabIndex        =   11
      Top             =   720
      Width           =   735
   End
   Begin VB.Label Label1 
      Caption         =   "Nombre:"
      Height          =   255
      Left            =   360
      TabIndex        =   10
      Top             =   720
      Width           =   615
   End
End
Attribute VB_Name = "frmPermisos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdCancelar_Click()
    Unload Me
End Sub

Private Sub cmdGuardar_Click()
    cn.Open
    If Estado = "Agregando" Then
        cn.Execute ("AgregarUsuario '" & txtUsuario & "','" & txtDomicilio & "','" & txttelefono & "','" & txtEmail & "'," & chEstablecerUsuariosyPermisos & "," & chModificarClientes & "," & chModificarArticulos & "," & chFacturar & "," & chVerInformes)
    Else
        cn.Execute ("GuardarUsuario '" & txtUsuario & "','" & txtDomicilio & "','" & txttelefono & "','" & txtEmail & "'," & chEstablecerUsuariosyPermisos & "," & chModificarClientes & "," & chModificarArticulos & "," & chFacturar & "," & chVerInformes & "," & idUsuarioPermiso)
    End If
    cn.Close
    Unload Me
End Sub

Private Sub Form_Load()
    If Estado = "Modificando" Then
        Dim rs As New ADODB.Recordset
        cn.Open
        Set rs = cn.Execute("VerPermsisosDeUsuario " & idUsuarioPermiso)
        txtUsuario = rs!Nombre
        chEstablecerUsuariosyPermisos.Value = rs!EstablecerUsuariosyPermisos
        chModificarClientes = rs!ModificarClientes
        chModificarArticulos = rs!ModificarArticulos
        chFacturar = rs!Facturar
        chVerInformes = rs!VerInformes
        rs.Close
        Set rs = Nothing
        cn.Close
    End If
End Sub

