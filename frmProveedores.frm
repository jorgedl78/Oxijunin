VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form frmProveedores 
   Caption         =   "Proveedores"
   ClientHeight    =   8880
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   11970
   LinkTopic       =   "Form1"
   ScaleHeight     =   8880
   ScaleWidth      =   11970
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame2 
      Caption         =   "Frame2"
      Height          =   1695
      Left            =   240
      TabIndex        =   9
      Top             =   7080
      Width           =   11415
      Begin VB.CommandButton cmdPagos 
         Caption         =   "Pagos"
         Height          =   1095
         Left            =   4800
         Picture         =   "frmProveedores.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   15
         Top             =   360
         Width           =   1335
      End
      Begin VB.CommandButton cmdRecibos 
         Caption         =   "Recibos"
         Height          =   1095
         Left            =   3360
         Picture         =   "frmProveedores.frx":0646
         Style           =   1  'Graphical
         TabIndex        =   14
         Top             =   360
         Width           =   1335
      End
      Begin VB.CommandButton cmdCuentaCorriente 
         Caption         =   "Cta Cte"
         Height          =   1095
         Left            =   1920
         Picture         =   "frmProveedores.frx":6E98
         Style           =   1  'Graphical
         TabIndex        =   13
         Top             =   360
         Width           =   1335
      End
      Begin VB.CommandButton cmdAgregar 
         Caption         =   "Agregar"
         Height          =   1095
         Left            =   6240
         Picture         =   "frmProveedores.frx":7762
         Style           =   1  'Graphical
         TabIndex        =   12
         Top             =   360
         Width           =   1215
      End
      Begin VB.CommandButton cmdSalir 
         Caption         =   "Salir"
         Height          =   1095
         Left            =   9840
         Picture         =   "frmProveedores.frx":802C
         Style           =   1  'Graphical
         TabIndex        =   11
         Top             =   360
         Width           =   1215
      End
      Begin VB.CommandButton cmCompras 
         Caption         =   "Compras"
         Height          =   1095
         Left            =   480
         Picture         =   "frmProveedores.frx":88F6
         Style           =   1  'Graphical
         TabIndex        =   10
         Top             =   360
         Width           =   1335
      End
   End
   Begin VB.TextBox txtBusca 
      Height          =   375
      Left            =   960
      TabIndex        =   3
      Top             =   240
      Width           =   3735
   End
   Begin VB.Frame Frame1 
      Caption         =   "Ordenado Por"
      Height          =   615
      Left            =   5400
      TabIndex        =   0
      Top             =   120
      Width           =   2415
      Begin VB.OptionButton optNombre 
         Caption         =   "Nombre"
         Height          =   255
         Left            =   240
         TabIndex        =   2
         Top             =   240
         Value           =   -1  'True
         Width           =   975
      End
      Begin VB.OptionButton optCodigo 
         Caption         =   "Código"
         Height          =   255
         Left            =   1320
         TabIndex        =   1
         Top             =   240
         Width           =   975
      End
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid grclientes 
      Height          =   5535
      Left            =   240
      TabIndex        =   4
      Top             =   840
      Width           =   11415
      _ExtentX        =   20135
      _ExtentY        =   9763
      _Version        =   393216
      Rows            =   1
      FixedRows       =   0
      FixedCols       =   0
      GridColorFixed  =   255
      TextStyleFixed  =   3
      HighLight       =   2
      SelectionMode   =   1
      GridLineWidthFixed=   1
      FontWidthFixed  =   0
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
   End
   Begin VB.Label Label1 
      Caption         =   "Buscar:"
      Height          =   255
      Left            =   240
      TabIndex        =   8
      Top             =   360
      Width           =   615
   End
   Begin VB.Label lblEncontrados 
      Caption         =   "lblEncontrados"
      Height          =   255
      Left            =   10080
      TabIndex        =   7
      Top             =   360
      Width           =   1095
   End
   Begin VB.Label Label2 
      Caption         =   "Total de Proveedores:"
      Height          =   255
      Left            =   8280
      TabIndex        =   6
      Top             =   360
      Width           =   1815
   End
   Begin VB.Label lblDescripcion 
      BackStyle       =   0  'Transparent
      Caption         =   "Descripcion"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   240
      TabIndex        =   5
      Top             =   6480
      Width           =   5175
   End
End
Attribute VB_Name = "frmProveedores"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmCompras_Click()
    If grclientes.TextMatrix(grclientes.Row, 0) = "" Then MsgBox ("Debe seleccionar un Proveedor"): txtBusca.SetFocus: Exit Sub
    
    frmCompras.lblidProveedor = grclientes.TextMatrix(grclientes.Row, 0)
    idProveedor = grclientes.TextMatrix(grclientes.Row, 0)
    frmCompras.txtNombre = grclientes.TextMatrix(grclientes.Row, 1)
    frmCompras.lblCategoria = grclientes.TextMatrix(grclientes.Row, 5)
    frmCompras.lblTipo = grclientes.TextMatrix(grclientes.Row, 6)
    frmCompras.lblNumero = grclientes.TextMatrix(grclientes.Row, 7)
    frmCompras.lblSaldo = "Saldo: $ " & Format(grclientes.TextMatrix(grclientes.Row, 4), "0.00")
    frmCompras.Show 1
End Sub

Private Sub cmdAgregar_Click()
    cn.Open
    Set rs = cn.Execute("VerPermsisosDeUsuario " & idUsuario)
    If rs!ModificarClientes = 0 Then MsgBox ("Función no permitida"): cn.Close: Exit Sub
    cn.Close
    Estado = "Agregando"
    frmFichaProveedor.Show 1
End Sub

Private Sub cmdCobrar_Click()
    idCliente = grclientes.TextMatrix(grclientes.Row, 0)
    frmRecibo.txtNombre = grclientes.TextMatrix(grclientes.Row, 1)
    frmRecibo.Show 1
End Sub

Private Sub cmdCuentaCorriente_Click()
    If grclientes.TextMatrix(grclientes.Row, 0) = "" Then MsgBox ("Debe seleccionar un Proveedor"): txtBusca.SetFocus: Exit Sub
    idProveedor = grclientes.TextMatrix(grclientes.Row, 0)
    frmCuentaCorrienteProveedor.Show 1
End Sub

Private Sub cmdFacturas_Click()

End Sub



Private Sub cmdPagos_Click()
    If grclientes.TextMatrix(grclientes.Row, 0) = "" Then MsgBox ("Debe seleccionar un Proveedor"): txtBusca.SetFocus: Exit Sub
    idProveedor = grclientes.TextMatrix(grclientes.Row, 0)
    frmReciboProveedores.txtNombre = grclientes.TextMatrix(grclientes.Row, 1)
    frmReciboProveedores.Show 1
End Sub

Private Sub cmdRecibos_Click()
    If grclientes.TextMatrix(grclientes.Row, 0) = "" Then MsgBox ("Debe seleccionar un Proveedor"): txtBusca.SetFocus: Exit Sub
    idProveedor = grclientes.TextMatrix(grclientes.Row, 0)
    frmRecibosProveedores.Show 1
End Sub

Private Sub cmdSalir_Click()
    Unload Me
End Sub








Private Sub Form_Activate()
    txtBusca.SetFocus
End Sub

Private Sub Form_Load()
    grclientes.Cols = 8
    grclientes.ColWidth(0) = 700
    grclientes.ColWidth(1) = 3000
    grclientes.ColWidth(2) = 2500
    grclientes.ColWidth(3) = 2500
    grclientes.ColAlignment(3) = 1
    grclientes.ColWidth(4) = 1000
    grclientes.ColWidth(5) = 1200
    grclientes.ColWidth(6) = 500
    grclientes.ColWidth(7) = 1200

    
    lblEncontrados = 0
    lblDescripcion = ""

    'BuscarClientes
End Sub

Private Sub grClientes_DblClick()
    EditarCliente
End Sub

Private Sub grClientes_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then EditarCliente
End Sub

Private Sub grClientes_RowColChange()
    lblDescripcion = grclientes.TextMatrix(grclientes.Row, 1)
End Sub

Private Sub optCodigo_Click()
    BuscarProveedores
End Sub

Private Sub optNombre_Click()
    BuscarProveedores
End Sub



Private Sub txtBusca_Change()
    'BuscarClientes
End Sub

Private Sub txtBusca_KeyPress(KeyAscii As Integer)
    If KeyAscii = 27 Then cmdSalir_Click
    If KeyAscii <> 13 Then Exit Sub
    BuscarProveedores
End Sub
Sub EditarCliente()
    If grclientes.Rows > 1 Then
        If EligiendoCliente = 1 Then
            frmCompras.lblidProveedor = grclientes.TextMatrix(grclientes.Row, 0)
            idProveedor = grclientes.TextMatrix(grclientes.Row, 0)
            frmCompras.txtNombre = grclientes.TextMatrix(grclientes.Row, 1)
            frmCompras.lblCategoria = grclientes.TextMatrix(grclientes.Row, 5)
            frmCompras.lblTipo = grclientes.TextMatrix(grclientes.Row, 6)
            frmCompras.lblNumero = grclientes.TextMatrix(grclientes.Row, 7)
            frmCompras.lblSaldo = "Saldo: $ " & Format(grclientes.TextMatrix(grclientes.Row, 4), "0.00")
            
            Unload Me
        Else
            'If grclientes.TextMatrix(grclientes.Row, 0) = 1 Then Exit Sub 'no permito editar al cliente CONSUMIDOR FINAL
            idProveedor = grclientes.TextMatrix(grclientes.Row, 0)
            ProveedorSeleccionado = grclientes.Row
            Saltar = 1
            Estado = "Modificando"
            frmFichaProveedor.Show 1
            If Saltar = 0 Then
                BuscarProveedores
                grclientes.Row = ProveedorSeleccionado
                grClientes_RowColChange
            End If
        End If
    End If
End Sub

Sub BuscarProveedores()
    cn.Open
    Dim rs As ADODB.Recordset
    ordenaPor = "Nombre"
    If optCodigo.Value = True Then
        ordenaPor = "idProveedor"
    End If
    Set rs = cn.Execute("ABMProveedores '" & txtBusca & "'," & ordenaPor)
    lblEncontrados = rs.RecordCount
   ' Set grClientes.DataSource = rs
    With grclientes
    .Rows = 1
    .TextArray(0) = "Codigo"
    .TextArray(1) = "Nombre"
    .TextArray(2) = "Domicilio"
    .TextArray(3) = "Telefonos"
    .TextArray(4) = "Saldo"
    Do While rs.EOF = False
        .Rows = .Rows + 1
        .TextMatrix(.Rows - 1, 0) = rs!idProveedor
        .TextMatrix(.Rows - 1, 1) = rs!Nombre
        .TextMatrix(.Rows - 1, 2) = rs!Domicilio
        .TextMatrix(.Rows - 1, 3) = rs!Telefonos
        .TextMatrix(.Rows - 1, 4) = Format(rs!saldo, "0.00")
        .TextMatrix(.Rows - 1, 5) = rs!Categoria
        .TextMatrix(.Rows - 1, 6) = rs!TipoDocumento
        .TextMatrix(.Rows - 1, 7) = rs!NumeroDocumento
      
        rs.MoveNext
        .FixedRows = 1
    Loop
    End With
    If rs.RecordCount > 0 Then
        'grclientes.SetFocus
        grClientes_RowColChange
    Else
        txtBusca.SetFocus
    End If
    rs.Close
    Set rs = Nothing
    cn.Close
End Sub

