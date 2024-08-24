VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmCompras 
   Caption         =   "Compras"
   ClientHeight    =   8070
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   13125
   LinkTopic       =   "Form1"
   ScaleHeight     =   8070
   ScaleWidth      =   13125
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame5 
      Height          =   1575
      Left            =   120
      TabIndex        =   14
      Top             =   6360
      Width           =   12855
      Begin VB.CommandButton cmdCancelar 
         Caption         =   "Cancelar"
         Height          =   855
         Left            =   7440
         Picture         =   "frmCompras.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   16
         Top             =   360
         Width           =   975
      End
      Begin VB.CommandButton cmdGuardar 
         Caption         =   "Guardar"
         Height          =   855
         Left            =   5040
         Picture         =   "frmCompras.frx":08CA
         Style           =   1  'Graphical
         TabIndex        =   15
         Top             =   360
         Width           =   975
      End
   End
   Begin VB.Frame Frame4 
      Caption         =   "Comentario"
      Height          =   1455
      Left            =   120
      TabIndex        =   12
      Top             =   4800
      Width           =   12855
      Begin VB.TextBox txtComentario 
         Height          =   975
         Left            =   120
         MaxLength       =   100
         TabIndex        =   13
         Top             =   240
         Width           =   12615
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "Importes"
      Height          =   1455
      Left            =   120
      TabIndex        =   7
      Top             =   3240
      Width           =   12855
      Begin VB.TextBox txtPercepcionIIBB 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   6000
         TabIndex        =   36
         Top             =   840
         Width           =   1815
      End
      Begin VB.TextBox txtPercepcionIVA 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   3960
         TabIndex        =   34
         Top             =   840
         Width           =   1815
      End
      Begin VB.TextBox txtTotal 
         Alignment       =   1  'Right Justify
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
         Left            =   10440
         TabIndex        =   28
         Top             =   840
         Width           =   2175
      End
      Begin VB.TextBox txtImpuestos 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   8040
         TabIndex        =   27
         Top             =   840
         Width           =   1815
      End
      Begin VB.TextBox txtIva 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1920
         TabIndex        =   26
         Top             =   840
         Width           =   1815
      End
      Begin VB.TextBox txtNeto 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   25
         Top             =   840
         Width           =   1575
      End
      Begin VB.Line Line1 
         X1              =   10080
         X2              =   10080
         Y1              =   120
         Y2              =   1440
      End
      Begin VB.Label Label11 
         Caption         =   "Percepción IIBB"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   6000
         TabIndex        =   37
         Top             =   600
         Width           =   1815
      End
      Begin VB.Label lclPercepIva 
         Caption         =   "Percepción IVA"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   3960
         TabIndex        =   35
         Top             =   600
         Width           =   1815
      End
      Begin VB.Label Label9 
         Caption         =   "Otros Impuestos:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   8040
         TabIndex        =   11
         Top             =   600
         Width           =   2055
      End
      Begin VB.Label Label7 
         Caption         =   "Total:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   10440
         TabIndex        =   10
         Top             =   360
         Width           =   1095
      End
      Begin VB.Label Label6 
         Caption         =   "Iva:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   1920
         TabIndex        =   9
         Top             =   600
         Width           =   1335
      End
      Begin VB.Label Label5 
         Caption         =   "Neto:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   8
         Top             =   600
         Width           =   1335
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Comprobante"
      Height          =   1215
      Left            =   120
      TabIndex        =   1
      Top             =   1920
      Width           =   12855
      Begin VB.ComboBox cmTipoComprobante 
         DataField       =   "Factura, Nota de Crédito, Nota de Débito"
         Height          =   315
         ItemData        =   "frmCompras.frx":1194
         Left            =   8520
         List            =   "frmCompras.frx":1196
         Style           =   2  'Dropdown List
         TabIndex        =   38
         Top             =   600
         Width           =   2055
      End
      Begin VB.ComboBox cmCondicion 
         DataField       =   "A,B,C"
         Height          =   315
         ItemData        =   "frmCompras.frx":1198
         Left            =   6120
         List            =   "frmCompras.frx":119A
         Style           =   2  'Dropdown List
         TabIndex        =   24
         Top             =   600
         Width           =   2055
      End
      Begin VB.ComboBox cmTipo 
         DataField       =   "A,B,C"
         Height          =   315
         ItemData        =   "frmCompras.frx":119C
         Left            =   2040
         List            =   "frmCompras.frx":119E
         Style           =   2  'Dropdown List
         TabIndex        =   22
         Top             =   600
         Width           =   615
      End
      Begin VB.TextBox txtPuesto 
         Alignment       =   1  'Right Justify
         Height          =   375
         Left            =   2760
         TabIndex        =   21
         Top             =   600
         Width           =   975
      End
      Begin VB.TextBox txtNumero 
         Alignment       =   1  'Right Justify
         Height          =   375
         Left            =   4080
         TabIndex        =   20
         Top             =   600
         Width           =   1575
      End
      Begin MSComCtl2.DTPicker dateFecha 
         Height          =   375
         Left            =   360
         TabIndex        =   19
         Top             =   600
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   661
         _Version        =   393216
         Format          =   119078913
         CurrentDate     =   42366
      End
      Begin VB.Label Label12 
         Caption         =   "Tipo:"
         Height          =   375
         Left            =   8520
         TabIndex        =   39
         Top             =   360
         Width           =   1215
      End
      Begin VB.Label Label1 
         Caption         =   "-"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   3840
         TabIndex        =   23
         Top             =   480
         Width           =   135
      End
      Begin VB.Label Label10 
         Caption         =   "Condición:"
         Height          =   375
         Left            =   6120
         TabIndex        =   6
         Top             =   360
         Width           =   1215
      End
      Begin VB.Label Label8 
         Caption         =   "Puesto:"
         Height          =   255
         Left            =   2760
         TabIndex        =   5
         Top             =   360
         Width           =   735
      End
      Begin VB.Label Label4 
         Caption         =   "Nº:"
         Height          =   375
         Left            =   4080
         TabIndex        =   4
         Top             =   360
         Width           =   375
      End
      Begin VB.Label Label3 
         Caption         =   "Tipo:"
         Height          =   255
         Left            =   2040
         TabIndex        =   3
         Top             =   360
         Width           =   495
      End
      Begin VB.Label Label2 
         Caption         =   "Fecha:"
         Height          =   255
         Left            =   360
         TabIndex        =   2
         Top             =   360
         Width           =   615
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Proveedor"
      Height          =   1695
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   12855
      Begin VB.CommandButton cmdBuscar 
         BackColor       =   &H00E0E0E0&
         Height          =   495
         Left            =   4680
         Picture         =   "frmCompras.frx":11A0
         Style           =   1  'Graphical
         TabIndex        =   18
         Top             =   240
         Width           =   495
      End
      Begin VB.TextBox txtNombre 
         Enabled         =   0   'False
         Height          =   375
         Left            =   720
         TabIndex        =   17
         Top             =   360
         Width           =   3855
      End
      Begin VB.Label lblNumero 
         Caption         =   "Numero"
         Height          =   255
         Left            =   4320
         TabIndex        =   33
         Top             =   1080
         Width           =   1575
      End
      Begin VB.Label lblTipo 
         Alignment       =   1  'Right Justify
         Caption         =   "Tipo"
         Height          =   255
         Left            =   3600
         TabIndex        =   32
         Top             =   1080
         Width           =   615
      End
      Begin VB.Label lblCategoria 
         Caption         =   "Categoria"
         Height          =   255
         Left            =   360
         TabIndex        =   31
         Top             =   1080
         Width           =   3255
      End
      Begin VB.Label lblSaldo 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0C0C0&
         Caption         =   "Saldo"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   6720
         TabIndex        =   30
         Top             =   240
         Width           =   5655
      End
      Begin VB.Label lblidProveedor 
         Caption         =   "X"
         Height          =   255
         Left            =   480
         TabIndex        =   29
         Top             =   480
         Width           =   255
      End
   End
End
Attribute VB_Name = "frmCompras"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdBuscar_Click()
    'EligiendoCliente = 1
    'frmProveedores.Show 1
    'EligiendoCliente = 0

End Sub

Private Sub cmdCancelar_Click()
    Unload Me
End Sub

Private Sub cmdGuardar_Click()
   
    If txtNombre = "" Then MsgBox ("Debe especificar un Proveedor"): Exit Sub
    'If cmCategorias = "Responsable Inscripto" And cmbTipo <> "CUIT" Then
    '    MsgBox ("Para esta categoría el tipo debe ser CUIT")
    '    Exit Sub
    'End If
    'If cmCategorias = "Responsable Inscripto" And Len(txtNumeroDocumento) <> 11 Then
    '    MsgBox ("Es obligatorio un nro de CUIT válido para esta categoría")
    '    Exit Sub
    'End If
    
    If txtPuesto = "" Then MsgBox ("Debe especificar el puesto de facturación"): txtPuesto.SetFocus: Exit Sub
    If txtNumero = "" Then MsgBox ("Debe especificar el número de comprobante"): txtNumero.SetFocus: Exit Sub
    If Val(txtTotal) = 0 Or txtTotal = "" Then: MsgBox ("Rebice los importes del comprobante"): Exit Sub
    
    Respuesta = MsgBox("¿Esta seguro de guardar el proveedor?", vbYesNo, "Guardar")
    If Respuesta = vbNo Then Exit Sub
    If txtNeto = "" Then txtNeto = "0"
    If txtIva = "" Then txtIva = "0"
    If txtPercepcionIVA = "" Then txtPercepcionIVA = "0"
    If txtPercepcionIIBB = "" Then txtPercepcionIIBB = "0"
    If txtImpuestos = "" Then txtImpuestos = "0"
    If txtTotal = "" Then txtTotal = "0"
    cn.Open
    

    cn.Execute ("AgregarCompra '" & Format(dateFecha, "yyyy/mm/dd") & "','" & cmTipo.Text & "'," & txtPuesto & "," & txtNumero & "," & Replace(txtNeto, ",", ".") & "," & Replace(txtIva, ",", ".") & "," & Replace(txtPercepcionIVA, ",", ".") & "," & Replace(txtPercepcionIIBB, ",", ".") & "," & Replace(txtImpuestos, ",", ".") & "," & Replace(txtTotal, ",", ".") & "," & idProveedor & ",'" & cmCondicion & "','" & cmTipoComprobante & "','" & txtComentario & "'")

    
    If cmCondicion = "Cuenta Corriente" Then
    
        Set rs = cn.Execute("SELECT MAX(idCompra) AS Nuevoid FROM Compras")
        NuevoID = rs!NuevoID
        
        If cmTipoComprobante = "Nota de Crédito" Then
            cn.Execute ("AgregarCuentaCorrienteProveedor '" & Format(dateFecha, "yyyy/mm/dd") & "','" & cmTipoComprobante & " " & cmTipo & Format(txtPuesto, "0000") & "-" & Format(txtNumero, "00000000") & "'," & "0," & Replace(txtTotal, ",", ".") & "," & Val(lblidProveedor) & ",'Com'," & NuevoID)
        Else
            cn.Execute ("AgregarCuentaCorrienteProveedor '" & Format(dateFecha, "yyyy/mm/dd") & "','" & cmTipoComprobante & " " & cmTipo & Format(txtPuesto, "0000") & "-" & Format(txtNumero, "00000000") & "'," & Replace(txtTotal, ",", ".") & "," & "0," & Val(lblidProveedor) & ",'Com'," & NuevoID)
        End If
    End If
    cn.Close
    Unload Me
End Sub

Private Sub dateFecha_Change()
    txtPuesto.SetFocus
End Sub

Private Sub Form_Load()
    cmTipo.AddItem ("A")
    cmTipo.ItemData(cmTipo.NewIndex) = 0
    cmTipo.AddItem ("B")
    cmTipo.ItemData(cmTipo.NewIndex) = 1
    cmTipo.AddItem ("C")
    cmTipo.ItemData(cmTipo.NewIndex) = 2
    cmTipo.ListIndex = 0
    
    cmCondicion.AddItem ("Cuenta Corriente")
    cmCondicion.ItemData(cmCondicion.NewIndex) = 0
    cmCondicion.AddItem ("Contado")
    cmCondicion.ItemData(cmCondicion.NewIndex) = 1
    cmCondicion.ListIndex = 0
    
    cmTipoComprobante.AddItem ("Factura")
    cmTipoComprobante.ItemData(cmTipoComprobante.NewIndex) = 0
    cmTipoComprobante.AddItem ("Nota de Crédito")
    cmTipoComprobante.ItemData(cmTipoComprobante.NewIndex) = 1
    cmTipoComprobante.AddItem ("Nota de Débito")
    cmTipoComprobante.ItemData(cmTipoComprobante.NewIndex) = 2
    cmTipoComprobante.ListIndex = 0
    
    
    dateFecha = Date
    
    'cmdBuscar_Click
End Sub

Private Sub txtImpuestos_Change()
    If InStr(1, "0123456789" & Chr(13) & Chr(8), Chr(KeyAscii)) = 0 Then
        KeyAscii = 0
    End If
End Sub

Private Sub txtImpuestos_KeyPress(KeyAscii As Integer)
    If InStr(1, "0123456789" & Chr(13) & Chr(8), Chr(KeyAscii)) = 0 Then
        KeyAscii = 0
    End If
End Sub

Private Sub txtIva_Change()
    CalcularTotal
End Sub

Private Sub txtIva_KeyPress(KeyAscii As Integer)
    If InStr(1, "0123456789." & Chr(13) & Chr(8), Chr(KeyAscii)) = 0 Then
        KeyAscii = 0
    End If
End Sub

Private Sub txtNeto_Change()
    txtIva = Format(txtNeto * 0.21, "0.00")
    txtPercepcionIVA = Format(txtNeto * 0.3, "0.00")
    CalcularTotal
End Sub

Private Sub txtNeto_KeyPress(KeyAscii As Integer)
    If InStr(1, "0123456789." & Chr(13) & Chr(8), Chr(KeyAscii)) = 0 Then
        KeyAscii = 0
    End If

End Sub

Private Sub txtNumero_KeyPress(KeyAscii As Integer)
    If InStr(1, "0123456789" & Chr(13) & Chr(8), Chr(KeyAscii)) = 0 Then
        KeyAscii = 0
    End If
End Sub

Private Sub txtNumero_LostFocus()
    cn.Open
    Set rs = cn.Execute("SELECT * FROM Compras WHERE Puesto = " & txtPuesto & " AND Numero = " & txtNumero & " AND idProveedor = " & lblidProveedor)
    If Not rs.EOF Then
        MsgBox "El comprobante ya fue registrado para este Proveedor" & vbCrLf & "Fecha: " & rs!Fecha & " - Importe: " & rs!Total, vbInformation, "Atención"
    End If
    cn.Close
End Sub

Private Sub txtPercepcionIIBB_Change()
    CalcularTotal
End Sub

Private Sub txtPercepcionIIBB_KeyPress(KeyAscii As Integer)
    If InStr(1, "0123456789." & Chr(13) & Chr(8), Chr(KeyAscii)) = 0 Then
        KeyAscii = 0
    End If
End Sub

Private Sub txtPercepcionIVA_Change()
    CalcularTotal
End Sub

Private Sub txtPercepcionIVA_KeyPress(KeyAscii As Integer)
    If InStr(1, "0123456789." & Chr(13) & Chr(8), Chr(KeyAscii)) = 0 Then
        KeyAscii = 0
    End If
End Sub

Private Sub txtPuesto_KeyPress(KeyAscii As Integer)
    If InStr(1, "0123456789." & Chr(13) & Chr(8), Chr(KeyAscii)) = 0 Then
        KeyAscii = 0
    End If
End Sub

Private Sub txtTotal_KeyPress(KeyAscii As Integer)
    If InStr(1, "0123456789." & Chr(13) & Chr(8), Chr(KeyAscii)) = 0 Then
        KeyAscii = 0
    End If
End Sub

Private Sub CalcularTotal()
    txtTotal = Format(Val(Replace(txtNeto, ",", ".")) + Val(Replace(txtIva, ",", ".")) + Val(Replace(txtPercepcionIVA, ",", ".")) + Val(Replace(txtPercepcionIIBB, ",", ".")) + Val(Replace(txtImpuestos, ",", ".")), "0.00")
End Sub
