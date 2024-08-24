VERSION 5.00
Begin VB.Form frmFichaTubo 
   Caption         =   "Form1"
   ClientHeight    =   5400
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10605
   LinkTopic       =   "Form1"
   ScaleHeight     =   5400
   ScaleWidth      =   10605
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame2 
      Caption         =   "Estado del Tubo"
      Height          =   3975
      Left            =   4800
      TabIndex        =   11
      Top             =   120
      Width           =   5655
      Begin VB.TextBox txtUltimoMovimiento 
         Alignment       =   1  'Right Justify
         Height          =   375
         Left            =   360
         Locked          =   -1  'True
         MaxLength       =   20
         TabIndex        =   18
         Top             =   1440
         Width           =   975
      End
      Begin VB.TextBox txtReside 
         Enabled         =   0   'False
         Height          =   375
         Left            =   360
         Locked          =   -1  'True
         MaxLength       =   20
         TabIndex        =   17
         Top             =   2400
         Width           =   5175
      End
      Begin VB.TextBox txtDetalleUltimo 
         Height          =   375
         Left            =   1440
         Locked          =   -1  'True
         MaxLength       =   50
         TabIndex        =   14
         Top             =   1440
         Width           =   4095
      End
      Begin VB.ComboBox cmMovimientos 
         Height          =   315
         Left            =   360
         Style           =   2  'Dropdown List
         TabIndex        =   12
         Top             =   600
         Width           =   5175
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Ultimo movimiento:"
         Height          =   195
         Left            =   360
         TabIndex        =   16
         Top             =   1200
         Width           =   1320
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Reside:"
         Height          =   195
         Left            =   360
         TabIndex        =   15
         Top             =   2160
         Width           =   540
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "Estado:"
         Height          =   195
         Left            =   360
         TabIndex        =   13
         Top             =   360
         Width           =   540
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Datos del Tubo"
      Height          =   3975
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   4575
      Begin VB.TextBox txtCodigoPropietario 
         Height          =   375
         Left            =   240
         MaxLength       =   20
         TabIndex        =   21
         Top             =   3360
         Width           =   735
      End
      Begin VB.TextBox txtPropietario 
         Enabled         =   0   'False
         Height          =   375
         Left            =   1080
         Locked          =   -1  'True
         MaxLength       =   20
         TabIndex        =   19
         Top             =   3360
         Width           =   3375
      End
      Begin VB.TextBox txtCapacidad 
         Alignment       =   1  'Right Justify
         Height          =   375
         Left            =   240
         TabIndex        =   7
         Top             =   2400
         Width           =   855
      End
      Begin VB.TextBox txtNumero 
         Height          =   375
         Left            =   240
         MaxLength       =   20
         TabIndex        =   6
         Top             =   600
         Width           =   3375
      End
      Begin VB.TextBox txtidTubo 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E0E0E0&
         Enabled         =   0   'False
         Height          =   375
         Left            =   1440
         TabIndex        =   5
         Top             =   120
         Visible         =   0   'False
         Width           =   735
      End
      Begin VB.ComboBox cmGas 
         Height          =   315
         Left            =   240
         Style           =   2  'Dropdown List
         TabIndex        =   4
         Top             =   1560
         Width           =   3735
      End
      Begin VB.ComboBox cmUnidad 
         Height          =   315
         Left            =   1200
         Style           =   2  'Dropdown List
         TabIndex        =   3
         Top             =   2400
         Width           =   975
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Propietario:"
         Height          =   195
         Left            =   240
         TabIndex        =   20
         Top             =   3120
         Width           =   795
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         Caption         =   "Numero:"
         Height          =   195
         Left            =   240
         TabIndex        =   10
         Top             =   360
         Width           =   600
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "Capacidad:"
         Height          =   195
         Left            =   240
         TabIndex        =   9
         Top             =   2160
         Width           =   810
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Gas:"
         Height          =   195
         Left            =   240
         TabIndex        =   8
         Top             =   1320
         Width           =   330
      End
   End
   Begin VB.CommandButton cmdGuardar 
      Caption         =   "Guardar"
      Height          =   855
      Left            =   3600
      Picture         =   "frmFichaTubo.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   4320
      Width           =   975
   End
   Begin VB.CommandButton cmdCancelar 
      Caption         =   "Cancelar"
      Height          =   855
      Left            =   5280
      Picture         =   "frmFichaTubo.frx":08CA
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   4320
      Width           =   975
   End
End
Attribute VB_Name = "frmFichaTubo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim idEstadoActual As Integer
Dim EnDestino As Integer
Dim rs As New ADODB.Recordset


Private Sub cmdCancelar_Click()
    Unload Me
End Sub

Private Sub cmdGuardar_Click()
    If txtNumero.Text = "" Then MsgBox ("Debe especificar un Número"): Exit Sub
    If cmMovimientos.ListIndex < 0 Then MsgBox ("Debe elejir el Estado el Tubo"): Exit Sub
    If cmGas.ListIndex < 0 Then MsgBox ("Debe especificar el gas"): Exit Sub
    If cmUnidad.ListIndex < 0 Then MsgBox ("Debe definir la Unidad de Medida"): Exit Sub
    If IsNumeric(txtCapacidad) = False Then MsgBox ("El capacidad no es un valor válido"): Exit Sub
    If IsNumeric(txtCodigoPropietario) = False Then MsgBox ("El código del propietario es inválido"): Exit Sub
    Respuesta = MsgBox("¿Esta seguro de guardar el Tubo?", vbYesNo, "Guardar")
    If Respuesta = vbNo Then Exit Sub
    cn.Open
    
    If cmMovimientos.ItemData(cmMovimientos.ListIndex) = 2 Then EnDestino = 1
    If cmMovimientos.ItemData(cmMovimientos.ListIndex) = 3 Then EnDestino = 2
    If cmMovimientos.ItemData(cmMovimientos.ListIndex) = 4 Then EnDestino = 1
    If cmMovimientos.ItemData(cmMovimientos.ListIndex) = 12 Then EnDestino = 916
    If cmMovimientos.ItemData(cmMovimientos.ListIndex) = 13 Then EnDestino = 1
    If cmMovimientos.ItemData(cmMovimientos.ListIndex) = 14 Then EnDestino = 909
    If cmMovimientos.ItemData(cmMovimientos.ListIndex) = 15 Then EnDestino = 1

    
    If Estado = "Modificando" Then
        cn.Execute ("UPDATE Tubos SET Numero='" & txtNumero & "', Capacidad=" & Replace(txtCapacidad, ",", ".") & ", idEstadoTubos=" & cmMovimientos.ItemData(cmMovimientos.ListIndex) & ", idArticulo=" & cmGas.ItemData(cmGas.ListIndex) & ", idUnidadMedida=" & cmUnidad.ItemData(cmUnidad.ListIndex) & ", Propietario=" & txtCodigoPropietario & " where idTubo=" & idTubo)
        If cmMovimientos.ItemData(cmMovimientos.ListIndex) <> idEstadoActual Then 'se cambio el estado del tubo. Graba movimiento 'se cambio el estado del tubo
            cn.Execute ("UPDATE Tubos SET UltimoMovimiento='" & Format(Date, "dd/mm/yyyy") & "', DetalleUltimo='" & cmMovimientos.Text & ": Movimiento manual', ClienteActual=" & EnDestino & "  where idTubo=" & idTubo)
            cn.Execute ("INSERT INTO MovimientosTubos (Fecha, Detalle, idTubo) VALUES ('" & Format(Date, "dd/mm/yyyy") & "', '" & cmMovimientos.Text & ": Movimiento manual'," & idTubo & ")")
        End If
    Else
        cn.Execute ("INSERT INTO Tubos (Numero, Capacidad, idEstadoTubos, idArticulo, idUnidadMedida, ClienteActual, Propietario, UltimoMovimiento, DetalleUltimo ) VALUES ('" & txtNumero & "', " & Replace(txtCapacidad, ",", ".") & " , " & cmMovimientos.ItemData(cmMovimientos.ListIndex) & ", " & cmGas.ItemData(cmGas.ListIndex) & ", " & cmUnidad.ItemData(cmUnidad.ListIndex) & ", " & EnDestino & "," & txtCodigoPropietario & ", '" & Format(Date, "dd/mm/yyyy") & "','Movimiento Manual')")
        Set rs = cn.Execute("SELECT max(idTubo) as UltimoId from Tubos")
        UltimoId = rs!UltimoId
        cn.Execute ("INSERT INTO MovimientosTubos (Fecha, Detalle, idTubo) VALUES ('" & Format(Date, "dd/mm/yyyy") & "', '" & cmMovimientos.Text & ": Movimiento manual'," & UltimoId & ")")
    End If
    cn.Close
    Saltar = 0
    Unload Me
End Sub

Private Sub Form_Load()

    cn.Open
    Set rs = cn.Execute("SELECT idEstadoTubos,Movimiento FROM EstadoTubos")
    Do While rs.EOF = False
        cmMovimientos.AddItem (rs!Movimiento)
        cmMovimientos.ItemData(cmMovimientos.NewIndex) = rs!idEstadoTubos
        rs.MoveNext
    Loop
    rs.Close
    Set rs = Nothing
    
    Set rs = cn.Execute("select  distinct tubos.idArticulo, Descripcion from Tubos inner join Articulos on tubos.idArticulo=articulos.idArticulo Where tubos.idArticulo <= 19 order by idArticulo")
    Do While rs.EOF = False
        cmGas.AddItem (rs!Descripcion)
        cmGas.ItemData(cmGas.NewIndex) = rs!idArticulo
        rs.MoveNext
    Loop
    rs.Close
    Set rs = Nothing
    
    Set rs = cn.Execute("select idUnidadMedida,Unidad from UnidadesMedidas")
    Do While rs.EOF = False
        cmUnidad.AddItem (rs!Unidad)
        cmUnidad.ItemData(cmUnidad.NewIndex) = rs!idUnidadMedida
        rs.MoveNext
    Loop
    rs.Close
    Set rs = Nothing
    
    If Estado = "Agregando" Then
        txtNumero = ""
        txtCapacidad = ""
        cmGas.ListIndex = 0
        For I = 0 To cmMovimientos.ListCount - 1
            If cmMovimientos.ItemData(I) = 16 Then cmMovimientos.ListIndex = I
        Next I
        cmMovimientos.Enabled = False
        cmUnidad.ListIndex = 0
        chkEtiquetar = 1
        EnDestino = 1
    Else
        txtNumero.Enabled = False
        Set rs = cn.Execute("SELECT idTubo,Numero,Capacidad,ClienteActual,UltimoMovimiento,isnull(DetalleUltimo, '') as DetalleUltimo,idUnidadMedida, idEstadoTubos, idArticulo, Clientes.Nombre, Clientes_1.Nombre as NombrePropietario, Propietario FROM Tubos INNER JOIN Clientes ON Tubos.ClienteActual = Clientes.idCliente INNER JOIN Clientes as Clientes_1 ON Tubos.Propietario = Clientes_1.idCliente where idTubo=" & idTubo)
        If rs.EOF = False Then
            txtidTubo = rs!idTubo
            txtNumero = rs!numero
            txtCapacidad = Format(rs!Capacidad, "0.00")
            For I = 0 To cmGas.ListCount - 1
                If cmGas.ItemData(I) = rs!idArticulo Then cmGas.ListIndex = I
            Next I
            For I = 0 To cmMovimientos.ListCount - 1
                If cmMovimientos.ItemData(I) = rs!idEstadoTubos Then cmMovimientos.ListIndex = I
            Next I
            For I = 0 To cmUnidad.ListCount - 1
                If cmUnidad.ItemData(I) = rs!idUnidadMedida Then cmUnidad.ListIndex = I
            Next I
            txtCodigoPropietario = rs!Propietario
            txtPropietario = rs!NombrePropietario
            idEstadoActual = rs!idEstadoTubos
            If rs!UltimoMovimiento <> Nulo Then txtUltimoMovimiento = rs!UltimoMovimiento
            txtDetalleUltimo = rs!DetalleUltimo
            txtReside = rs!Nombre
            EnDestino = rs!ClienteActual
        End If
    End If
    cn.Close
End Sub


Private Sub txtCapacidad_KeyPress(KeyAscii As Integer)
    If InStr(1, "0123456789." & Chr(13) & Chr(8), Chr(KeyAscii)) = 0 Then
        KeyAscii = 0
    End If
End Sub

Private Sub txtCodigoPropietario_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
    If IsNumeric(txtCodigoPropietario) = False Then MsgBox ("El código del propietario es inválido"): Exit Sub

        cn.Open
        Set rs = cn.Execute("select Nombre  from Clientes where idCliente=" & txtCodigoPropietario)
        If rs.EOF = True Then
            MsgBox ("No se encontro el propietario"): txtCodigoPropietario.SetFocus: cn.Close: Exit Sub
        Else
            txtPropietario = rs!Nombre
        End If
        
        cn.Close
    End If
End Sub
