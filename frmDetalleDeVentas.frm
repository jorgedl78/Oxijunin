VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmDetalleDeVentas 
   Caption         =   "Detalle de Ventas"
   ClientHeight    =   4020
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8475
   LinkTopic       =   "Form1"
   ScaleHeight     =   4020
   ScaleWidth      =   8475
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Height          =   3735
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   8175
      Begin VB.CommandButton cmdImprimirReporteDeMovimientos 
         Caption         =   "Imprimir Detalle de Movimientos"
         Height          =   1095
         Left            =   4080
         Picture         =   "frmDetalleDeVentas.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   2040
         Width           =   1575
      End
      Begin MSComCtl2.DTPicker CalendarHasta 
         Height          =   375
         Left            =   5280
         TabIndex        =   6
         Top             =   960
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   661
         _Version        =   393216
         Format          =   85000193
         CurrentDate     =   42391
      End
      Begin MSComCtl2.DTPicker CalendarDesde 
         Height          =   375
         Left            =   1320
         TabIndex        =   5
         Top             =   960
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   661
         _Version        =   393216
         Format          =   85000193
         CurrentDate     =   42391
      End
      Begin VB.CommandButton cmdImprimirDetalleArticulosVendidos 
         Caption         =   "Imprimir Detalle de Artículos Vendidos"
         Height          =   1095
         Left            =   2160
         Picture         =   "frmDetalleDeVentas.frx":08CA
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   2040
         Width           =   1575
      End
      Begin VB.CommandButton cmdImprimirDetalle 
         Caption         =   "Imprimir Detalle de Totales"
         Enabled         =   0   'False
         Height          =   1095
         Left            =   480
         Picture         =   "frmDetalleDeVentas.frx":1194
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   1560
         Visible         =   0   'False
         Width           =   1575
      End
      Begin VB.Label Label1 
         Caption         =   "Desde"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1560
         TabIndex        =   3
         Top             =   480
         Width           =   975
      End
      Begin VB.Label Label2 
         Caption         =   "Hasta"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   5280
         TabIndex        =   2
         Top             =   480
         Width           =   975
      End
   End
End
Attribute VB_Name = "frmDetalleDeVentas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rst As ADODB.Recordset

Private Sub cmdImprimirDetalle_Click()
    cn.Open
    'Set rs = cn.Execute("SELECT Fecha, Tipo, Numero, Neto, IVA, Total, idCliente, Puesto, pedido FROM Ventas")
    With DetalleDeVenta.Sections("Sección4")
        .Controls("lblDesde").Caption = CalendarDesde.Value
        .Controls("lblHasta").Caption = CalendarHasta.Value
    End With
    
    With DetalleDeVenta.Sections("Sección5")
        Set rs = cn.Execute("SELECT Sum(Neto) as neto, sum(IVA) as iva, sum(Total) as total, sum(Impuestos) as Impuestos FROM Ventas WHERE Fecha BETWEEN '" & CalendarDesde.Value & "' and '" & CalendarHasta.Value & "'")
        .Controls("lblNeto").Caption = Format(rs!Neto, "#.00")
        .Controls("lblIva").Caption = Format(rs!Iva, "#.00")
        .Controls("lblTotal").Caption = Format(rs!Total, "#.00")
        .Controls("lblImpuestos").Caption = Format(rs!Impuestos, "#.00")
        Set rs = cn.Execute("SELECT Isnull(Sum(Total),0) as Contado FROM Ventas WHERE Condicion='CONTADO' AND Fecha BETWEEN '" & CalendarDesde.Value & "' and '" & CalendarHasta.Value & "'")
        .Controls("lblContado").Caption = Format(rs!contado, "#.00")
        Dim contado As Double
        contado = rs!contado
        Set rs = cn.Execute("SELECT Isnull(Sum(Total),0) as CuentaCorriente FROM Ventas WHERE Condicion='CUENTA CORRIENTE' AND Fecha BETWEEN '" & CalendarDesde.Value & "' and '" & CalendarHasta.Value & "'")
        .Controls("lblCuentaCorriente").Caption = Format(rs!CuentaCorriente, "#.00")
        Set rs = cn.Execute("SELECT IsNull(SUM(Importe),0) as Importe FROM Recibos WHERE Fecha BETWEEN '" & CalendarDesde.Value & "' and '" & CalendarHasta.Value & "'")
        .Controls("lblTotalCobrado").Caption = Format(rs!Importe, "#.00")
        Dim cobrado As Double
        cobrado = rs!Importe
        .Controls("lblResultadoObtenido").Caption = Format(contado + cobrado, "#.00")

        
    End With
        
        
    Set rs = cn.Execute("SELECT Ventas.Fecha, Ventas.Tipo, Ventas.Numero, Ventas.Neto, Ventas.IVA, Ventas.Total, Ventas.idCliente, Ventas.Puesto, Ventas.pedido, Clientes.Nombre, Ventas.Impuestos , Ventas.Condicion FROM Ventas INNER JOIN Clientes ON Ventas.idCliente = Clientes.idCliente WHERE Fecha BETWEEN '" & CalendarDesde.Value & "' and '" & CalendarHasta.Value & "'")
    Set DetalleDeVenta.DataSource = rs
    DetalleDeVenta.WindowState = 2
    
    DetalleDeVenta.Show 1
    
    cn.Close
End Sub


Private Sub cmdImprimirDetalleArticulosVendidos_Click()
    cn.Open
    With DetalleDeArticulosVendidos.Sections("Sección4")
        .Controls("lblDesde").Caption = CalendarDesde.Value
        .Controls("lblHasta").Caption = CalendarHasta.Value
    End With
    
    With DetalleDeArticulosVendidos.Sections("Sección5")
        Set rs = cn.Execute("SELECT  sum(DetalleVenta.PrecioTotal) as Importe FROM DetalleVenta INNER JOIN Articulos ON DetalleVenta.idArticulo = Articulos.idArticulo INNER JOIN Ventas ON DetalleVenta.idVenta = Ventas.idVenta WHERE ventas.Fecha BETWEEN '" & CalendarDesde.Value & "' and '" & CalendarHasta.Value & "'")
        .Controls("lblTotal").Caption = Format(rs!Importe, "#.00")
    End With
        
    
    Set rs = cn.Execute("SELECT '' as Fecha, Articulos.CodBar, Articulos.Descripcion, SUM(DetalleVenta.Cantidad) as Cantidad, sum(DetalleVenta.PrecioTotal) as Importe, '' as Motivo FROM DetalleVenta INNER JOIN Articulos ON DetalleVenta.idArticulo = Articulos.idArticulo INNER JOIN Ventas ON DetalleVenta.idVenta = Ventas.idVenta WHERE ventas.Fecha BETWEEN '" & CalendarDesde.Value & "' and '" & CalendarHasta.Value & "' group by Articulos.CodBar, Articulos.Descripcion " _
    & " union all " _
    & " SELECT convert(varchar(10),Ajustes_Stock.Fecha) as Fecha, articulos.CodBar, Articulos.Descripcion,  Ajustes_Stock.Cantidad, '0.00',  Ajustes_Stock.Movimiento + ': ' + Ajustes_Stock.Motivo FROM Ajustes_Stock INNER JOIN Articulos ON Ajustes_Stock.idArticulo = Articulos.idArticulo Where Ajustes_Stock.Fecha Between '" & CalendarDesde.Value & "' AND '" & CalendarHasta.Value & "'")
    Set DetalleDeArticulosVendidos.DataSource = rs
    DetalleDeArticulosVendidos.WindowState = 2
    
    DetalleDeArticulosVendidos.Show 1
    
    cn.Close
End Sub

Private Sub Command1_Click()

End Sub

Private Sub cmdImprimirReporteDeMovimientos_Click()
    cn.Open
    Set rs = cn.Execute("SELECT Fecha, Tipo, Numero, Neto, IVA, Total, idCliente, Puesto, pedido FROM Ventas")
    With ReporteDeMovimientos.Sections("Sección4")
        .Controls("lblDesde").Caption = CalendarDesde.Value
        .Controls("lblHasta").Caption = CalendarHasta.Value
    End With
    
    With ReporteDeMovimientos.Sections("Sección5")
        'Set rs = cn.Execute("SELECT Sum(Neto) as neto, sum(IVA) as iva, sum(Total) as total, sum(Impuestos) as Impuestos FROM Ventas WHERE Fecha BETWEEN '" & CalendarDesde.Value & "' and '" & CalendarHasta.Value & "'")
    
        '.Controls("lblNeto").Caption = Format(rs!Neto, "#.00")
        '.Controls("lblIva").Caption = Format(rs!Iva, "#.00")
        '.Controls("lblTotal").Caption = Format(rs!Total, "#.00")
        '.Controls("lblImpuestos").Caption = Format(rs!Impuestos, "#.00")
        Set rs = cn.Execute("SELECT (SELECT Isnull(Sum(Total),0) as Contado FROM Ventas WHERE Condicion='CONTADO' AND Fecha BETWEEN '" & CalendarDesde.Value & "' and '" & CalendarHasta.Value & "' AND Comprobante <> 'Nota de Crédito') - (SELECT Isnull(Sum(Total),0) as Contado FROM Ventas WHERE Condicion='CONTADO' AND Fecha BETWEEN '" & CalendarDesde.Value & "' and '" & CalendarHasta.Value & "' AND Comprobante = 'Nota de Crédito') as Contado")
        .Controls("lblContado").Caption = Format(rs!contado, "#0.00")
        Dim contado As Double
       contado = rs!contado
       Set rs = cn.Execute("SELECT (SELECT Isnull(Sum(Total),0) as Contado FROM Ventas WHERE Condicion='CUENTA CORRIENTE' AND Fecha BETWEEN '" & CalendarDesde.Value & "' and '" & CalendarHasta.Value & "' AND Comprobante <> 'Nota de Crédito') - (SELECT Isnull(Sum(Total),0) as Contado FROM Ventas WHERE Condicion='CUENTA CORRIENTE' AND Fecha BETWEEN '" & CalendarDesde.Value & "' and '" & CalendarHasta.Value & "' AND Comprobante = 'Nota de Crédito') as CuentaCorriente")
       .Controls("lblCuentaCorriente").Caption = Format(rs!CuentaCorriente, "#0.00")
       Set rs = cn.Execute("SELECT IsNull(SUM(Importe),0) as Importe FROM Recibos WHERE Anulado is Null and Fecha BETWEEN '" & CalendarDesde.Value & "' and '" & CalendarHasta.Value & "'")
       .Controls("lblTotalCobrado").Caption = Format(rs!Importe, "#0.00")
       Dim cobrado As Double
       cobrado = rs!Importe
       .Controls("lblResultadoObtenido").Caption = Format(contado + cobrado, "#0.00")

    End With
        
     
    Set rs = cn.Execute("select 'Comprobantes' as Grupo,Fecha, clientes.Nombre, Comprobante, Tipo + REPLICATE('0',(4 - LEN(convert(varchar(8),Puesto) ))) + convert(varchar(8),Puesto) + REPLICATE('0',(8 - LEN(convert(varchar(8),Numero) ))) + convert(varchar(8),Numero) as Detalle, Total, Condicion  from Ventas inner join Clientes on ventas.idCliente=Clientes.idCliente Where Ventas.Fecha Between '" & CalendarDesde.Value & "' AND '" & CalendarHasta.Value & "' Union All select 'Recibos' as grupo, Fecha, Nombre, 'Recibo ' + IsNull(Anulado,'')  as Comprobante, convert(varchar,(Numero)) as Numero, Importe, 'CUENTA CORRIENTE' as Condicion from Recibos inner join Clientes on Recibos.idCliente=clientes.idCliente Where Recibos.Fecha Between '" & CalendarDesde.Value & "' AND '" & CalendarHasta.Value & "'")
    Set ReporteDeMovimientos.DataSource = rs
    ReporteDeMovimientos.WindowState = 2
    
    ReporteDeMovimientos.Show 1
    
    cn.Close
    
End Sub

Private Sub Form_Load()
    CalendarDesde.Value = Date
    CalendarHasta.Value = Date
End Sub
