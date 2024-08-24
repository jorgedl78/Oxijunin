VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form frmUsuarios 
   Caption         =   "Usuarios"
   ClientHeight    =   5385
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   11430
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   ScaleHeight     =   5385
   ScaleWidth      =   11430
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdCancelar 
      Caption         =   "Cancelar"
      Height          =   855
      Left            =   7320
      Picture         =   "frmUsuarios.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   4320
      Width           =   975
   End
   Begin VB.CommandButton cmdAgregar 
      Caption         =   "Agregar"
      Height          =   855
      Left            =   4560
      Picture         =   "frmUsuarios.frx":08CA
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   4320
      Width           =   975
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid grUsuarios 
      Height          =   3375
      Left            =   240
      TabIndex        =   0
      Top             =   600
      Width           =   10935
      _ExtentX        =   19288
      _ExtentY        =   5953
      _Version        =   393216
      Rows            =   1
      FixedRows       =   0
      FixedCols       =   0
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
   End
End
Attribute VB_Name = "frmUsuarios"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdAgregar_Click()
    Estado = "Agregando"
    frmPermisos.Show 1
    CargarUsuarios
End Sub

Private Sub cmdCancelar_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    grUsuarios.Cols = 5
    grUsuarios.ColWidth(0) = 700
    grUsuarios.ColWidth(1) = 2500
    grUsuarios.ColWidth(2) = 2500
    grUsuarios.ColWidth(3) = 2500
    grUsuarios.ColWidth(4) = 2500
    CargarUsuarios
End Sub
Sub CargarUsuarios()
    grUsuarios.Rows = 1
    grUsuarios.TextArray(0) = "Codigo"
    grUsuarios.TextArray(1) = "Usuario"
    grUsuarios.TextArray(2) = "Domicilio"
    grUsuarios.TextArray(3) = "Teléfono"
    grUsuarios.TextArray(4) = "E_mail"
    Dim rs As New ADODB.Recordset
    cn.Open
    Set rs = cn.Execute("PedirUsuarios")
    With grUsuarios
    Do While rs.EOF = False
        .Rows = .Rows + 1
        .TextMatrix(.Rows - 1, 0) = rs!idUsuario
        .TextMatrix(.Rows - 1, 1) = rs!Nombre
        .TextMatrix(.Rows - 1, 2) = rs!Domicilio
        .TextMatrix(.Rows - 1, 3) = rs!Telefono
        .TextMatrix(.Rows - 1, 4) = rs!email
        rs.MoveNext
        .FixedRows = 1
    Loop
    End With
    cn.Close
End Sub

Private Sub grUsuarios_DblClick()
    idUsuarioPermiso = grUsuarios.TextMatrix(grUsuarios.Row, 0)
    Estado = "Modificando"
    frmPermisos.Show 1
    CargarUsuarios
End Sub
