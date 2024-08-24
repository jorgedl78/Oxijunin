VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmConsultaComprobantes 
   Caption         =   "Consulta Comprobantes"
   ClientHeight    =   8640
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   14535
   ControlBox      =   0   'False
   DrawMode        =   14  'Copy Pen
   LinkTopic       =   "Form1"
   ScaleHeight     =   8640
   ScaleWidth      =   14535
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdExcel 
      Caption         =   "Exportar"
      Height          =   855
      Left            =   8400
      Picture         =   "frmConsultaComprobantes.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   16
      Top             =   7440
      Width           =   1095
   End
   Begin VB.CommandButton cmdLibroIVAVentas 
      Caption         =   "IVA Ventas"
      Height          =   855
      Left            =   7080
      Picture         =   "frmConsultaComprobantes.frx":05C3
      Style           =   1  'Graphical
      TabIndex        =   15
      Top             =   7440
      Width           =   1095
   End
   Begin VB.Frame Frame3 
      Caption         =   "Reimprimir Comprobante"
      Height          =   1215
      Left            =   240
      TabIndex        =   11
      Top             =   7200
      Width           =   3015
      Begin VB.OptionButton optDuplicado 
         Caption         =   "Duplicado"
         Height          =   255
         Left            =   240
         TabIndex        =   14
         Top             =   720
         Width           =   1215
      End
      Begin VB.OptionButton optOriginal 
         Caption         =   "Original"
         Height          =   255
         Left            =   240
         TabIndex        =   13
         Top             =   360
         Width           =   1095
      End
      Begin VB.CommandButton cmdReimprimir 
         Caption         =   "Reimprimir"
         Height          =   855
         Left            =   1680
         Picture         =   "frmConsultaComprobantes.frx":0924
         Style           =   1  'Graphical
         TabIndex        =   12
         Top             =   240
         Width           =   1095
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Resultados"
      Height          =   5895
      Left            =   120
      TabIndex        =   9
      Top             =   1320
      Width           =   14295
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid grComprobantes 
         Height          =   5535
         Left            =   120
         TabIndex        =   10
         Top             =   240
         Width           =   14055
         _ExtentX        =   24791
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
   End
   Begin VB.Frame Frame1 
      Height          =   1215
      Left            =   120
      TabIndex        =   3
      Top             =   120
      Width           =   14295
      Begin VB.CommandButton cmdConsultar 
         Caption         =   "Consultar"
         Height          =   855
         Left            =   9120
         Picture         =   "frmConsultaComprobantes.frx":11EE
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   240
         Width           =   1095
      End
      Begin MSComCtl2.DTPicker CalendarHasta 
         Height          =   375
         Left            =   5040
         TabIndex        =   4
         Top             =   480
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   661
         _Version        =   393216
         Format          =   91095041
         CurrentDate     =   42391
      End
      Begin MSComCtl2.DTPicker CalendarDesde 
         Height          =   375
         Left            =   1440
         TabIndex        =   5
         Top             =   480
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   661
         _Version        =   393216
         Format          =   82903041
         CurrentDate     =   42391
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
         Left            =   4080
         TabIndex        =   7
         Top             =   480
         Width           =   975
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
         Left            =   360
         TabIndex        =   6
         Top             =   480
         Width           =   975
      End
   End
   Begin VB.CommandButton cmdSalir 
      Caption         =   "Salir"
      Height          =   855
      Left            =   12120
      Picture         =   "frmConsultaComprobantes.frx":1AB8
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   7440
      Width           =   1095
   End
   Begin VB.CommandButton cmdAnular 
      Caption         =   "Anular"
      Enabled         =   0   'False
      Height          =   855
      Left            =   4200
      Picture         =   "frmConsultaComprobantes.frx":2382
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   7440
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.CommandButton cmdDetalle 
      Caption         =   "Detalle"
      Height          =   855
      Left            =   5400
      Picture         =   "frmConsultaComprobantes.frx":2497
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   7440
      Visible         =   0   'False
      Width           =   1095
   End
End
Attribute VB_Name = "frmConsultaComprobantes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim obj As PageSet.PrinterControl
Dim objExcel As Object
Dim objWorkbook As Object
Dim objWorksheet As Object


Private Sub cmdAnular_Click()
    If grComprobantes.TextMatrix(grComprobantes.Row, 11) = "Anulada" Then MsgBox ("El comprobante ya se encuentra anulado"): Exit Sub
   
    Respuesta = MsgBox("¿Confirma la anulación del comprobante " & grComprobantes.TextMatrix(grComprobantes.Row, 3) & "?", vbYesNo, "Atencion!")
    If Respuesta = vbNo Then Exit Sub
    cn.Open
    cn.Execute ("UPDATE Ventas SET Anulada='Anulada', Neto=0.00, Iva=0.00, Impuestos=0.00, Total=0.00 WHERE idVenta=" & grComprobantes.TextMatrix(grComprobantes.Row, 12))
    If grComprobantes.TextMatrix(grComprobantes.Row, 1) = "Nota de Crédito" Then
        debe = Replace(grComprobantes.TextMatrix(grComprobantes.Row, 9), ",", ".")
        haber = 0
    Else
        debe = 0
        haber = Replace(grComprobantes.TextMatrix(grComprobantes.Row, 10), ",", ".")
    End If
    'si es cta cte genero el movimiento de anulacion en el cliente
    If grComprobantes.TextMatrix(grComprobantes.Row, 14) = "CUENTA CORRIENTE" Then
        cn.Execute ("AgregarCuentaCorriente '" & Format(Date, "yyyy/mm/dd") & "','Anulación " & grComprobantes.TextMatrix(grComprobantes.Row, 1) & " " & grComprobantes.TextMatrix(grComprobantes.Row, 2) + grComprobantes.TextMatrix(grComprobantes.Row, 3) & "'," & debe & "," & haber & "," & grComprobantes.TextMatrix(grComprobantes.Row, 13) & ",'Anu',0")
    End If
    Dim detalleArticulos As New Recordset
    Set detalleArticulos = cn.Execute("select idArticulo,Cantidad  from DetalleVenta where idVenta = " & grComprobantes.TextMatrix(grComprobantes.Row, 12))
    If grComprobantes.TextMatrix(grComprobantes.Row, 1) = "Nota de Crédito" Then
        multiplicador = -1
    Else
        multiplicador = 1
    End If
    While detalleArticulos.EOF = False
        Cantidad = detalleArticulos!Cantidad
        cn.Execute ("UPDATE Articulos set Stock = Stock + " & (Cantidad * multiplicador) & " WHERE idArticulo = " & detalleArticulos!idArticulo)
        detalleArticulos.MoveNext
    Wend
    'Pongo en cero el detalle de articulos vendidos
    cn.Execute ("UPDATE DetalleVenta SET PrecioNeto=0, Cantidad=0, Costo=0, PrecioTotal=0 WHERE idVenta=" & grComprobantes.TextMatrix(grComprobantes.Row, 12))
    cn.Close
    Consultar
End Sub

Private Sub cmdConsultar_Click()
    Consultar
End Sub

Private Sub cmdExcel_Click()
    On Error GoTo ErrorHandler
    
    If Dir("C:\Libro_de_iva_ventas.xlsx") <> "" Then
        ' Eliminar el archivo
        Kill "C:\Libro_de_iva_ventas.xlsx"
        'MsgBox "El archivo ha sido eliminado exitosamente."
    Else
        'MsgBox "El archivo no existe."
    End If

    ' Crear una nueva instancia de Excel
    Set objExcel = CreateObject("Excel.Application")
    
    ' Hacer visible la aplicación de Excel (opcional)
    objExcel.Visible = True
    
    ' Agregar un nuevo libro
    Set objWorkbook = objExcel.Workbooks.Add
    
    ' Usar la primera hoja de trabajo
    Set objWorksheet = objWorkbook.Worksheets(1)
    
    ' Rellenar la hoja de trabajo con datos
    objWorksheet.Cells(1, 1).Value = "Fecha"
    objWorksheet.Cells(1, 2).Value = "Comprobante"
    objWorksheet.Cells(1, 3).Value = "Tipo"
    objWorksheet.Cells(1, 4).Value = "Puesto"
    objWorksheet.Cells(1, 5).Value = "Numero"
    objWorksheet.Cells(1, 6).Value = "Cliente"
    objWorksheet.Cells(1, 7).Value = "Categoría"
    objWorksheet.Cells(1, 8).Value = "CUIT"
    objWorksheet.Cells(1, 9).Value = "Neto 21%"
    objWorksheet.Cells(1, 10).Value = "Neto 10.5 %"
    objWorksheet.Cells(1, 11).Value = "Iva 21%"
    objWorksheet.Cells(1, 12).Value = "Iva 10.5 %"
    objWorksheet.Cells(1, 13).Value = "Impuesto"
    objWorksheet.Cells(1, 14).Value = "Total"
    objWorksheet.Cells(1, 15).Value = "CAE"
    objWorksheet.Cells(1, 15).Value = "Condición"

    cn.Open
    Dim rs As ADODB.Recordset
    Set rs = cn.Execute("ConsultaDeComprobantes '" & CalendarDesde & "','" & CalendarHasta & "', 0 , 999999")
    'lblEncontrados = rs.RecordCount
   ' Set grClientes.DataSource = rs

    
    fila = 2
    Do While rs.EOF = False
        objWorksheet.Cells(fila, 1).Value = CDate(rs!Fecha)
        objWorksheet.Cells(fila, 2).Value = rs!Comprobante
        objWorksheet.Cells(fila, 3).Value = rs!Tipo
        objWorksheet.Cells(fila, 4).Value = rs!Puesto
        objWorksheet.Cells(fila, 5).Value = rs!numero
        objWorksheet.Cells(fila, 6).Value = rs!Nombre
        objWorksheet.Cells(fila, 7).Value = rs!Categoria
        objWorksheet.Cells(fila, 8).Value = rs!NumeroDocumento
        objWorksheet.Cells(fila, 9).Value = rs!Neto
        objWorksheet.Cells(fila, 10).Value = rs!Neto2
        objWorksheet.Cells(fila, 11).Value = rs!Iva
        objWorksheet.Cells(fila, 12).Value = rs!Iva2
        objWorksheet.Cells(fila, 13).Value = rs!Impuestos
        objWorksheet.Cells(fila, 14).Value = rs!Total
        objWorksheet.Cells(fila, 15).Value = rs!CAE
        objWorksheet.Cells(fila, 16).Value = rs!Condicion
        
        fila = fila + 1
        
 
        rs.MoveNext
        
    Loop

    rs.Close
    Set rs = Nothing
    cn.Close

    
    ' Guardar el archivo en una ruta específica
    objWorkbook.SaveAs "C:\Libro_de_iva_ventas.xlsx"
    
    ' Cerrar el libro y la aplicación de Excel
    'objWorkbook.Close
    'objExcel.Quit
    
    ' Liberar los objetos
    'Set objWorksheet = Nothing
    'Set objWorkbook = Nothing
    'Set objExcel = Nothing
    
    'MsgBox "Exportación completada!"
    Exit Sub
    
ErrorHandler:
    MsgBox "Ocurrió un error al intentar eliminar el archivo: " & Err.Description
End Sub



Private Sub cmdLibroIVAVentas_Click()
    fechaDesde = CalendarDesde.Value
    fechaHasta = CalendarHasta
    'LibroIVAVentas.WindowState = 2
    'LibroIVAVentas.Show 1
      Set obj = New PrinterControl
      obj.ChngOrientationLandscape
      'DataReport1.Show
      'DataReport1.PrintReport False, rptRangeFromTo, 1, 1

    ReporteLibroIVAVentas.WindowState = 2
    ReporteLibroIVAVentas.Show 1
    obj.ChngOrientationPortrait


End Sub

Private Sub cmdReimprimir_Click()
    idComprobante = grComprobantes.TextMatrix(grComprobantes.Row, 12)
    If grComprobantes.TextMatrix(grComprobantes.Row, 15) = 1 Then
        If optOriginal.Value = True Then
            'frmImprimeFacturaElectronica.lblfacturaOriginal = "Original"
            condicionComprobante = "Original"
        Else
            'frmImprimeFacturaElectronica.lblfacturaOriginal = "Duplicado"
            condicionComprobante = "Duplicado"
        End If
        'frmImprimeFacturaElectronica.Show 1
        'frmImprimeFacturaElectronica.PrintForm
        'Unload frmImprimeFacturaElectronica
        cn.Open
        'Set rs = cn.Execute("SELECT max(idVenta) as idVenta FROM Ventas WHERE idVenta=" & idComprobante)
        'idRecibo = rs!UltimoRecibo
        'Set ImprimeFacturaElectronica.DataSource = rs
        
        ImprimeFacturaElectronica.WindowState = 2
        ImprimeFacturaElectronica.Show 1
        cn.Close

    Else
        'frmImprimeFactura.lblVencimiento = ""
        'frmImprimeFactura.lblRemito = ""
        'frmImprimeFactura.lblComentario = ""
        'frmImprimeFactura.PrintForm
        'Unload frmImprimeFactura
    End If
End Sub

Private Sub cmdSalir_Click()
    Unload Me
End Sub

Private Sub Command1_Click()
End Sub

Private Sub Form_Load()
    CalendarDesde.Value = Date
    CalendarHasta.Value = Date
    optDuplicado.Value = True
    
    With grComprobantes
        .Cols = 16
        .ColWidth(0) = 1000 'fecha
        .ColWidth(1) = 1400 'comprobante
        .ColWidth(2) = 300 'tipo
        .ColWidth(3) = 1200 'detalle
        .ColWidth(4) = 2500 'cliente
        .ColWidth(5) = 1800 'categoria
        .ColWidth(6) = 800 'neto
        .ColWidth(7) = 800 'iva 21
        .ColWidth(8) = 800 'iva 10,5
        .ColWidth(9) = 800 'impuesto
        .ColWidth(10) = 800 'total
        .ColWidth(11) = 0 'anulada
        .ColWidth(12) = 0 'idventa
        .ColWidth(13) = 0 'idcliente
        .ColWidth(14) = 2000 'condicion
        .ColWidth(15) = 0 'electronica
        
    End With
End Sub

Private Sub Consultar()
    cn.Open
    Dim rs As ADODB.Recordset
    Set rs = cn.Execute("ConsultaDeComprobantes '" & CalendarDesde & "','" & CalendarHasta & "', 0 , 999999")
    'lblEncontrados = rs.RecordCount
   ' Set grClientes.DataSource = rs
    With grComprobantes
    .Rows = 1
    .TextArray(0) = "Fecha"
    .TextArray(1) = "Comprobante"
    .TextArray(2) = "Tipo"
    .TextArray(3) = "Detalle"
    .TextArray(4) = "Cliente"
    .TextArray(5) = "Categoria"
    .TextArray(6) = "Neto"
    .TextArray(7) = "Iva 21%"
    .TextArray(8) = "Iva 10.5%"
    .TextArray(9) = "Impuesto"
    .TextArray(10) = "Total"
    .TextArray(11) = "Anulada"
    .TextArray(12) = "idVenta"
    .TextArray(13) = "idCliente"
    .TextArray(14) = "Condicion"
    .TextArray(15) = "Electronica"
    
    
    Do While rs.EOF = False
        .Rows = .Rows + 1
        .TextMatrix(.Rows - 1, 0) = rs!Fecha
        .TextMatrix(.Rows - 1, 1) = rs!Comprobante
        .TextMatrix(.Rows - 1, 2) = rs!Tipo
        .TextMatrix(.Rows - 1, 3) = rs!Detalle
        .TextMatrix(.Rows - 1, 4) = rs!Nombre
        .TextMatrix(.Rows - 1, 5) = rs!Categoria
        .TextMatrix(.Rows - 1, 6) = rs!Neto
        .TextMatrix(.Rows - 1, 7) = rs!Iva
        .TextMatrix(.Rows - 1, 8) = rs!Iva2
        .TextMatrix(.Rows - 1, 9) = rs!Impuestos
        .TextMatrix(.Rows - 1, 10) = rs!Total
        .TextMatrix(.Rows - 1, 11) = rs!Anulada
        .TextMatrix(.Rows - 1, 12) = rs!idVenta
        .TextMatrix(.Rows - 1, 13) = rs!idCliente
        .TextMatrix(.Rows - 1, 14) = rs!Condicion
        .TextMatrix(.Rows - 1, 15) = rs!Electronica
        rs.MoveNext
        
        .FixedRows = 1
    Loop
    End With
    rs.Close
    Set rs = Nothing
    cn.Close


End Sub

