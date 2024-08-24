VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmInformesDeCompras 
   Caption         =   "Informes de Compras"
   ClientHeight    =   4995
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   7710
   LinkTopic       =   "Form1"
   ScaleHeight     =   4995
   ScaleWidth      =   7710
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame3 
      Height          =   1695
      Left            =   240
      TabIndex        =   8
      Top             =   3240
      Width           =   7335
      Begin VB.CommandButton cmdSalir 
         Caption         =   "Salir"
         Height          =   855
         Left            =   3000
         Picture         =   "frmInformesDeCompras.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   9
         Top             =   480
         Width           =   1095
      End
   End
   Begin VB.Frame Frame2 
      Height          =   1695
      Left            =   240
      TabIndex        =   5
      Top             =   1440
      Width           =   7335
      Begin VB.CommandButton cmdLibroIVAVentas 
         Caption         =   "IVA Ventas"
         Height          =   855
         Left            =   1920
         Picture         =   "frmInformesDeCompras.frx":08CA
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   480
         Width           =   1095
      End
      Begin VB.CommandButton cmdExcel 
         Caption         =   "Exportar"
         Height          =   855
         Left            =   4320
         Picture         =   "frmInformesDeCompras.frx":0C2B
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   480
         Width           =   1095
      End
   End
   Begin VB.Frame Frame1 
      Height          =   1215
      Left            =   240
      TabIndex        =   0
      Top             =   120
      Width           =   7335
      Begin MSComCtl2.DTPicker CalendarHasta 
         Height          =   375
         Left            =   5040
         TabIndex        =   1
         Top             =   480
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   661
         _Version        =   393216
         Format          =   90636289
         CurrentDate     =   42391
      End
      Begin MSComCtl2.DTPicker CalendarDesde 
         Height          =   375
         Left            =   1440
         TabIndex        =   2
         Top             =   480
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   661
         _Version        =   393216
         Format          =   90701825
         CurrentDate     =   42391
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
         TabIndex        =   4
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
         Left            =   4080
         TabIndex        =   3
         Top             =   480
         Width           =   975
      End
   End
End
Attribute VB_Name = "frmInformesDeCompras"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdExcel_Click()
    On Error GoTo ErrorHandler
    
    If Dir("C:\Libro_de_iva_compras.xlsx") <> "" Then
        ' Eliminar el archivo
        Kill "C:\Libro_de_iva_compras.xlsx"
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
    
    objWorksheet.Cells(1, 1).Value = "Libro de IVA Compras: desde " & CalendarDesde & " hasta " & CalendarHasta
    
    ' Rellenar la hoja de trabajo con datos
    objWorksheet.Cells(3, 1).Value = "Fecha"
    objWorksheet.Cells(3, 2).Value = "Proveedor"
    objWorksheet.Cells(3, 3).Value = "Cuit"
    objWorksheet.Cells(3, 4).Value = "Tipo"
    objWorksheet.Cells(3, 5).Value = "Numero"
    objWorksheet.Cells(3, 6).Value = "Neto"
    objWorksheet.Cells(3, 7).Value = "IVA"
    objWorksheet.Cells(3, 8).Value = "Percepción IVA"
    objWorksheet.Cells(3, 9).Value = "Percepción IIBB"
    objWorksheet.Cells(3, 10).Value = "Impuestos"
    objWorksheet.Cells(3, 11).Value = "Total"

    cn.Open
    Dim rs As ADODB.Recordset
    Set rs = cn.Execute("SELECT Compras.Fecha, Compras.TipoComprobante, Compras.Tipo, Compras.Tipo + REPLICATE('0', 4 - LEN(CONVERT(varchar(8), Compras.Puesto))) + CONVERT(varchar(8), Compras.Puesto) + REPLICATE('0', 8 - LEN(CONVERT(varchar(8), Compras.Numero))) + CONVERT(varchar(8), Compras.Numero) AS Comprobante, Compras.Neto, Compras.IVA, Compras.PercepcionIva, Compras.PercepcionIIBB , Compras.Impuestos, Compras.Total, Proveedores.Nombre, Proveedores.NumeroDocumento FROM Compras INNER JOIN Proveedores ON Compras.idProveedor = Proveedores.idProveedor WHERE Compras.Fecha BETWEEN '" & CalendarDesde & "' AND '" & CalendarHasta & "' ORDER BY Compras.Fecha")
    'lblEncontrados = rs.RecordCount
   ' Set grClientes.DataSource = rs

    
    fila = 4
    Do While rs.EOF = False
        objWorksheet.Cells(fila, 1).Value = CDate(rs!Fecha)
        objWorksheet.Cells(fila, 2).Value = rs!Nombre
        objWorksheet.Cells(fila, 3).Value = rs!NumeroDocumento
        objWorksheet.Cells(fila, 4).Value = rs!TipoComprobante
        objWorksheet.Cells(fila, 5).Value = rs!Comprobante
        objWorksheet.Cells(fila, 6).Value = rs!Neto
        objWorksheet.Cells(fila, 7).Value = rs!Iva
        objWorksheet.Cells(fila, 8).Value = rs!PercepcionIva
        objWorksheet.Cells(fila, 9).Value = rs!PercepcionIIBB
        objWorksheet.Cells(fila, 10).Value = rs!Impuestos
        objWorksheet.Cells(fila, 11).Value = rs!Total
        
        fila = fila + 1
        
 
        rs.MoveNext
        
    Loop

    rs.Close
    Set rs = Nothing
    cn.Close

    
    ' Guardar el archivo en una ruta específica
    objWorkbook.SaveAs "C:\Libro_de_iva_compras.xlsx"
    
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

    ReporteLibroIVACompras.WindowState = 2
    ReporteLibroIVACompras.Show 1
    obj.ChngOrientationPortrait

End Sub

Private Sub cmdSalir_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    CalendarDesde.Value = Date
    CalendarHasta.Value = Date
End Sub
