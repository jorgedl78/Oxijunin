VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form frmClientes 
   Caption         =   "Clientes"
   ClientHeight    =   8235
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   11805
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   ScaleHeight     =   8235
   ScaleWidth      =   11805
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdTubosPropios 
      Caption         =   "Ver Tubos Propios"
      Height          =   855
      Left            =   360
      Picture         =   "frmClientes.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   17
      Top             =   7320
      Width           =   1095
   End
   Begin VB.CommandButton cmdVerTubos 
      Caption         =   "Ver Tubos Vendidos"
      Height          =   855
      Left            =   1560
      Picture         =   "frmClientes.frx":0A02
      Style           =   1  'Graphical
      TabIndex        =   16
      Top             =   7320
      Width           =   1095
   End
   Begin VB.Frame Frame1 
      Caption         =   "Ordenado Por"
      Height          =   615
      Left            =   5280
      TabIndex        =   13
      Top             =   0
      Width           =   2415
      Begin VB.OptionButton optCodigo 
         Caption         =   "Código"
         Height          =   255
         Left            =   1320
         TabIndex        =   15
         Top             =   240
         Width           =   975
      End
      Begin VB.OptionButton optNombre 
         Caption         =   "Nombre"
         Height          =   255
         Left            =   240
         TabIndex        =   14
         Top             =   240
         Value           =   -1  'True
         Width           =   975
      End
   End
   Begin VB.CommandButton cmdUltimo 
      Caption         =   "Ultimo Ingresado"
      Height          =   855
      Left            =   6480
      Picture         =   "frmClientes.frx":1404
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   7320
      Width           =   1095
   End
   Begin VB.CommandButton cmdRecibos 
      Caption         =   "Recibos"
      Height          =   855
      Left            =   5160
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   7320
      Width           =   1095
   End
   Begin VB.CommandButton cmdFacturas 
      Caption         =   "Facturas"
      Height          =   855
      Left            =   7680
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   7320
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.CommandButton cmdCobrar 
      Caption         =   "Cobrar"
      Height          =   855
      Left            =   3960
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   7320
      Width           =   1095
   End
   Begin VB.CommandButton cmdCuentaCorriente 
      Caption         =   "Cta Cte"
      Height          =   855
      Left            =   2760
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   7320
      Width           =   1095
   End
   Begin VB.CommandButton cmdAgregar 
      Caption         =   "Agregar"
      Height          =   855
      Left            =   9240
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   7320
      Width           =   1095
   End
   Begin VB.TextBox txtBusca 
      Height          =   375
      Left            =   840
      TabIndex        =   0
      Top             =   120
      Width           =   3735
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid grclientes 
      Height          =   5775
      Left            =   120
      TabIndex        =   2
      Top             =   720
      Width           =   11415
      _ExtentX        =   20135
      _ExtentY        =   10186
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
   Begin VB.CommandButton cmdSalir 
      Caption         =   "Salir"
      Height          =   855
      Left            =   10440
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   7320
      Width           =   1095
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
      Left            =   360
      TabIndex        =   6
      Top             =   6600
      Width           =   5175
   End
   Begin VB.Label Label2 
      Caption         =   "Total de clientes:"
      Height          =   255
      Left            =   8160
      TabIndex        =   5
      Top             =   240
      Width           =   1335
   End
   Begin VB.Label lblEncontrados 
      Caption         =   "lblEncontrados"
      Height          =   255
      Left            =   9600
      TabIndex        =   4
      Top             =   240
      Width           =   1095
   End
   Begin VB.Label Label1 
      Caption         =   "Buscar:"
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   240
      Width           =   615
   End
End
Attribute VB_Name = "frmClientes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdAgregar_Click()
    cn.Open
    Set rs = cn.Execute("VerPermsisosDeUsuario " & idUsuario)
    If rs!ModificarClientes = 0 Then MsgBox ("Función no permitida"): cn.Close: Exit Sub
    cn.Close
    Estado = "Agregando"
    frmFichaCliente.Show 1
End Sub

Private Sub cmdCobrar_Click()
    idCliente = grclientes.TextMatrix(grclientes.Row, 0)
    frmRecibo.txtNombre = grclientes.TextMatrix(grclientes.Row, 1)
    frmRecibo.Show 1
End Sub

Private Sub cmdCuentaCorriente_Click()
    idCliente = grclientes.TextMatrix(grclientes.Row, 0)
    frmCuentaCorrienteCliente.Show 1
End Sub

Private Sub cmdRecibos_Click()
    idCliente = grclientes.TextMatrix(grclientes.Row, 0)
    frmRecibos.Show 1
End Sub

Private Sub cmdSalir_Click()
    Unload Me
End Sub


Private Sub cmdTubosPropios_Click()
    idCliente = grclientes.TextMatrix(grclientes.Row, 0)
    frmVerTubosPropios.Show 1
End Sub

Private Sub cmdUltimo_Click()
  cn.Open
  Set rs = cn.Execute("select idCliente, Nombre from Clientes where idCliente=(SELECT max(idcliente) from clientes)")
  MsgBox ("Ultimo código de Cliente: " & rs!idCliente & " - " & rs!Nombre)
  cn.Close
End Sub

Private Sub cmdVerTubos_Click()
    idCliente = grclientes.TextMatrix(grclientes.Row, 0)
    frmVerTubosEnCliente.Show 1
End Sub

Private Sub Form_Activate()
        txtBusca.SetFocus
End Sub

Private Sub Form_Load()
    grclientes.Cols = 9
    grclientes.ColWidth(0) = 700
    grclientes.ColWidth(1) = 3000
    grclientes.ColWidth(2) = 2500
    grclientes.ColWidth(3) = 2500
    grclientes.ColAlignment(3) = 1
    grclientes.ColWidth(4) = 1000
    grclientes.ColWidth(5) = 2500
    grclientes.ColWidth(6) = 500
    grclientes.ColWidth(7) = 1200
    grclientes.ColWidth(8) = 0
    
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
    BuscarClientes
End Sub

Private Sub optNombre_Click()
    BuscarClientes
End Sub



Private Sub txtBusca_Change()
    'BuscarClientes
End Sub

Private Sub txtBusca_KeyPress(KeyAscii As Integer)
    If KeyAscii = 27 Then cmdSalir_Click
    If KeyAscii <> 13 Then Exit Sub
    BuscarClientes
End Sub
Sub EditarCliente()
    If grclientes.Rows > 1 Then
        If EligiendoCliente = 1 Then
            Dim rs As ADODB.Recordset
            cn.Open
            frmFacturador.lblIdCliente = grclientes.TextMatrix(grclientes.Row, 0)
            idCliente = grclientes.TextMatrix(grclientes.Row, 0)
            frmFacturador.lblCliente = grclientes.TextMatrix(grclientes.Row, 1)
            frmFacturador.lblCategoria = grclientes.TextMatrix(grclientes.Row, 5)
            frmFacturador.lblTipoDocumento = grclientes.TextMatrix(grclientes.Row, 6)
            frmFacturador.lblNumeroDocumento = grclientes.TextMatrix(grclientes.Row, 7)
            frmFacturador.lblSaldo = "Saldo: $ " & Format(grclientes.TextMatrix(grclientes.Row, 4), "0.00")
            
            'buco el utimo uso mensual facturado para informar
            Set rs = cn.Execute("SELECT  top 1 DetalleVenta.Cantidad, Articulos.Descripcion, DetalleVenta.PrecioTotal, Ventas.Fecha, Ventas.Tipo, Ventas.Puesto, Ventas.Numero, Ventas.idCliente, Articulos.idArticulo FROM Articulos INNER JOIN DetalleVenta ON Articulos.idArticulo = DetalleVenta.idArticulo INNER JOIN Ventas ON DetalleVenta.idVenta = Ventas.idVenta CROSS JOIN Clientes WHERE (Ventas.idCliente = " & grclientes.TextMatrix(grclientes.Row, 0) & ") AND (Articulos.idArticulo between 101 and 112)ORDER BY DetalleVenta.idDetalleVenta DESC")
            If rs.EOF = True Then
                frmFacturador.lblUltimoUsoFacturado = "No registra último uso facturado"
            Else
                frmFacturador.lblUltimoUsoFacturado = rs!Cantidad & "-" & rs!Descripcion & " " & rs!Tipo & Format(rs!Puesto, "0000") & "-" & Format(rs!numero, "00000000") & " (" & rs!Fecha & ")"
            End If
            
            'cuento la cantidad de tubos en propiedad del cliente
            Set rs = cn.Execute("select COUNT(idTubo) as Cantidad from Tubos where ClienteActual=" & grclientes.TextMatrix(grclientes.Row, 0) & " and Propietario<>" & grclientes.TextMatrix(grclientes.Row, 0))
            If rs!Cantidad = 0 Then
                frmFacturador.lblCantidadTubosEnCliente = "No se registran tubos en su propiedad"
            Else
                 frmFacturador.lblCantidadTubosEnCliente = "Tubos en su propiedad: " & rs!Cantidad
            End If
            
            If grclientes.TextMatrix(grclientes.Row, 8) = 1 Then
                frmFacturador.lblTipoPrecio = "Revendedor"
            Else
                frmFacturador.lblTipoPrecio = "Público"
            End If
            

            
            If frmFacturador.lblCategoria = "Responsable Inscripto" Or frmFacturador.lblCategoria = "Monotributo" Then
                frmFacturador.lblLetra = "A"
                Set rs = cn.Execute("SELECT (NumeroA + 1) as Numero, Puesto from Parametros")
            Else
                frmFacturador.lblLetra = "B"
                Set rs = cn.Execute("SELECT (NumeroB + 1) as Numero, Puesto from Parametros")
            End If
            frmFacturador.lblPuesto = Format(rs!Puesto, "0000")
            frmFacturador.lblNumero = Format(rs!numero, "00000000")
            cn.Close
            Unload Me
        Else
            If grclientes.TextMatrix(grclientes.Row, 0) = 1 Then Exit Sub 'no permito editar al cliente CONSUMIDOR FINAL
            idCliente = grclientes.TextMatrix(grclientes.Row, 0)
            ClienteSeleccionado = grclientes.Row
            Saltar = 1
            Estado = "Modificando"
            frmFichaCliente.Show 1
            If Saltar = 0 Then
                BuscarClientes
                grclientes.Row = ClienteSeleccionado
                grClientes_RowColChange
            End If
        End If
    End If
End Sub

Sub BuscarClientes()
    cn.Open
    Dim rs As ADODB.Recordset
    ordenaPor = "Nombre"
    If optCodigo.Value = True Then
        ordenaPor = "idCliente"
    End If
    Set rs = cn.Execute("ABMClientes '" & txtBusca & "'," & ordenaPor)
    lblEncontrados = rs.RecordCount
   ' Set grClientes.DataSource = rs
    With grclientes
    .Rows = 1
    .TextArray(0) = "Codigo"
    .TextArray(1) = "Nombre"
    .TextArray(2) = "Domicilio"
    .TextArray(3) = "Telefonos"
    .TextArray(4) = "Saldo"
    .TextArray(5) = "Categoría"
    Do While rs.EOF = False
        .Rows = .Rows + 1
        .TextMatrix(.Rows - 1, 0) = rs!idCliente
        .TextMatrix(.Rows - 1, 1) = rs!Nombre
        .TextMatrix(.Rows - 1, 2) = rs!Domicilio
        .TextMatrix(.Rows - 1, 3) = rs!Telefonos
        .TextMatrix(.Rows - 1, 4) = Format(rs!saldo, "0.00")
        .TextMatrix(.Rows - 1, 5) = rs!Categoria
        .TextMatrix(.Rows - 1, 6) = rs!TipoDocumento
        .TextMatrix(.Rows - 1, 7) = rs!NumeroDocumento
        .TextMatrix(.Rows - 1, 8) = rs!PrecioRevendedor
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
