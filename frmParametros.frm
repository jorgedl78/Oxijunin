VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmParametros 
   Caption         =   "Parámetros"
   ClientHeight    =   7755
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   8580
   ControlBox      =   0   'False
   DrawMode        =   14  'Copy Pen
   LinkTopic       =   "Form1"
   ScaleHeight     =   7755
   ScaleWidth      =   8580
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdCancelar 
      Caption         =   "Cancelar"
      Height          =   855
      Left            =   4680
      Picture         =   "frmParametros.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   6600
      Width           =   975
   End
   Begin VB.CommandButton cmdGuardar 
      Caption         =   "Guardar"
      Height          =   855
      Left            =   2880
      Picture         =   "frmParametros.frx":08CA
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   6600
      Width           =   975
   End
   Begin TabDlg.SSTab tabParametros 
      Height          =   5895
      Left            =   240
      TabIndex        =   0
      Top             =   360
      Width           =   8175
      _ExtentX        =   14420
      _ExtentY        =   10398
      _Version        =   393216
      Tab             =   1
      TabHeight       =   520
      TabCaption(0)   =   "Datos de la Empresa"
      TabPicture(0)   =   "frmParametros.frx":1194
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "Frame4"
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "Facturación"
      TabPicture(1)   =   "frmParametros.frx":11B0
      Tab(1).ControlEnabled=   -1  'True
      Tab(1).Control(0)=   "Frame3"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "Frame2"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).ControlCount=   2
      TabCaption(2)   =   "Parámetros Generales"
      TabPicture(2)   =   "frmParametros.frx":11CC
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Frame1"
      Tab(2).ControlCount=   1
      Begin VB.Frame Frame1 
         Caption         =   "Parámetros Generales"
         Height          =   1095
         Left            =   -74640
         TabIndex        =   38
         Top             =   1620
         Width           =   7215
         Begin VB.TextBox txtPorcentaje 
            Alignment       =   1  'Right Justify
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   420
            Left            =   5160
            TabIndex        =   40
            Top             =   315
            Width           =   735
         End
         Begin VB.TextBox txtRecibo 
            Alignment       =   1  'Right Justify
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   420
            Left            =   1920
            TabIndex        =   39
            Top             =   360
            Width           =   1095
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "Último Recibo:"
            Height          =   195
            Left            =   600
            TabIndex        =   43
            Top             =   480
            Width           =   1035
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            Caption         =   "Porcentaje Cta. Cte.:"
            Height          =   195
            Left            =   3600
            TabIndex        =   42
            Top             =   480
            Width           =   1470
         End
         Begin VB.Label Label5 
            Caption         =   "%"
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
            Left            =   6000
            TabIndex        =   41
            Top             =   360
            Width           =   375
         End
      End
      Begin VB.Frame Frame4 
         Caption         =   "Datos de la Empresa"
         Height          =   4815
         Left            =   -74760
         TabIndex        =   19
         Top             =   540
         Width           =   7575
         Begin VB.TextBox txtEmail 
            Height          =   420
            Left            =   1440
            MaxLength       =   100
            TabIndex        =   28
            Top             =   4320
            Width           =   5895
         End
         Begin VB.TextBox txtTelefonos 
            Height          =   420
            Left            =   1440
            MaxLength       =   50
            TabIndex        =   27
            Top             =   3840
            Width           =   5895
         End
         Begin VB.TextBox txtCP 
            Height          =   420
            Left            =   1440
            MaxLength       =   15
            TabIndex        =   26
            Top             =   3360
            Width           =   2175
         End
         Begin VB.TextBox txtLocalidad 
            Height          =   420
            Left            =   1440
            MaxLength       =   50
            TabIndex        =   25
            Top             =   2880
            Width           =   5175
         End
         Begin VB.TextBox txtDomicilio 
            Height          =   420
            Left            =   1440
            MaxLength       =   50
            TabIndex        =   24
            Top             =   2400
            Width           =   5175
         End
         Begin VB.TextBox txtInicioActividades 
            Height          =   420
            Left            =   1440
            MaxLength       =   10
            TabIndex        =   23
            Top             =   1920
            Width           =   1455
         End
         Begin VB.TextBox txtIngresosBrutos 
            Height          =   420
            Left            =   1440
            MaxLength       =   50
            TabIndex        =   22
            Top             =   1440
            Width           =   3255
         End
         Begin VB.TextBox txtCuit 
            Height          =   420
            Left            =   1440
            MaxLength       =   11
            TabIndex        =   21
            Top             =   960
            Width           =   2175
         End
         Begin VB.TextBox txtNombre 
            Height          =   420
            Left            =   1440
            MaxLength       =   100
            TabIndex        =   20
            Top             =   480
            Width           =   5895
         End
         Begin VB.Label Label18 
            AutoSize        =   -1  'True
            Caption         =   "E-mail:"
            Height          =   195
            Left            =   840
            TabIndex        =   37
            Top             =   4440
            Width           =   465
         End
         Begin VB.Label Label17 
            AutoSize        =   -1  'True
            Caption         =   "Telefonos:"
            Height          =   195
            Left            =   600
            TabIndex        =   36
            Top             =   3960
            Width           =   750
         End
         Begin VB.Label Label16 
            AutoSize        =   -1  'True
            Caption         =   "Cod. Postal:"
            Height          =   195
            Left            =   480
            TabIndex        =   35
            Top             =   3480
            Width           =   855
         End
         Begin VB.Label Label15 
            AutoSize        =   -1  'True
            Caption         =   "Localidad:"
            Height          =   195
            Left            =   600
            TabIndex        =   34
            Top             =   3000
            Width           =   735
         End
         Begin VB.Label Label14 
            AutoSize        =   -1  'True
            Caption         =   "Domicilio:"
            Height          =   195
            Left            =   720
            TabIndex        =   33
            Top             =   2520
            Width           =   675
         End
         Begin VB.Label Label13 
            AutoSize        =   -1  'True
            Caption         =   "Inicio Actividades:"
            Height          =   195
            Left            =   120
            TabIndex        =   32
            Top             =   2040
            Width           =   1290
         End
         Begin VB.Label Label12 
            AutoSize        =   -1  'True
            Caption         =   "Ingresos Brutos:"
            Height          =   195
            Left            =   240
            TabIndex        =   31
            Top             =   1560
            Width           =   1140
         End
         Begin VB.Label Label11 
            AutoSize        =   -1  'True
            Caption         =   "CUIT:"
            Height          =   195
            Left            =   960
            TabIndex        =   30
            Top             =   1080
            Width           =   420
         End
         Begin VB.Label Label8 
            AutoSize        =   -1  'True
            Caption         =   "Nombre:"
            Height          =   195
            Left            =   720
            TabIndex        =   29
            Top             =   600
            Width           =   600
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   "Paramettros de Facturación Manual"
         Height          =   1215
         Left            =   600
         TabIndex        =   12
         Top             =   840
         Width           =   7215
         Begin VB.TextBox txtComprobanteB 
            Alignment       =   1  'Right Justify
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   420
            Left            =   5880
            TabIndex        =   15
            Top             =   480
            Width           =   975
         End
         Begin VB.TextBox txtComprobanteA 
            Alignment       =   1  'Right Justify
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   420
            Left            =   3360
            TabIndex        =   14
            Top             =   480
            Width           =   975
         End
         Begin VB.TextBox txtPuesto 
            Alignment       =   1  'Right Justify
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   420
            Left            =   840
            TabIndex        =   13
            Top             =   480
            Width           =   975
         End
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
            Caption         =   "Puesto:"
            Height          =   195
            Left            =   240
            TabIndex        =   18
            Top             =   600
            Width           =   540
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Comprobante ""A"":"
            Height          =   195
            Left            =   1920
            TabIndex        =   17
            Top             =   600
            Width           =   1290
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "Comprobante ""B"":"
            Height          =   195
            Left            =   4440
            TabIndex        =   16
            Top             =   600
            Width           =   1290
         End
      End
      Begin VB.Frame Frame3 
         Caption         =   "Parámetros de Facturación Electrónica"
         Height          =   3495
         Left            =   600
         TabIndex        =   3
         Top             =   2220
         Width           =   7215
         Begin VB.CommandButton cmdConsultar 
            Caption         =   "Consultar"
            Height          =   375
            Left            =   5760
            TabIndex        =   54
            Top             =   2640
            Width           =   855
         End
         Begin VB.Frame frTipo 
            Caption         =   "Tipo"
            Height          =   1095
            Left            =   4920
            TabIndex        =   51
            Top             =   2280
            Width           =   735
            Begin VB.OptionButton optA 
               Caption         =   "A"
               Height          =   255
               Left            =   120
               TabIndex        =   53
               Top             =   360
               Width           =   495
            End
            Begin VB.OptionButton optB 
               Caption         =   "B"
               Height          =   255
               Left            =   120
               TabIndex        =   52
               Top             =   720
               Width           =   495
            End
         End
         Begin VB.Frame frComprobante 
            Caption         =   "Comprobante"
            Height          =   1095
            Left            =   2760
            TabIndex        =   47
            Top             =   2280
            Width           =   2055
            Begin VB.OptionButton optFactura 
               Caption         =   "Factura"
               Height          =   195
               Left            =   120
               TabIndex        =   50
               Top             =   240
               Width           =   975
            End
            Begin VB.OptionButton optNC 
               Caption         =   "Nota de Crédito"
               Height          =   255
               Left            =   120
               TabIndex        =   49
               Top             =   480
               Width           =   1455
            End
            Begin VB.OptionButton optND 
               Caption         =   "Nota de Débito"
               Height          =   255
               Left            =   120
               TabIndex        =   48
               Top             =   750
               Width           =   1455
            End
         End
         Begin VB.CheckBox cheModoTesting 
            Height          =   255
            Left            =   3600
            TabIndex        =   45
            Top             =   480
            Width           =   255
         End
         Begin VB.TextBox txtPuestoElectrónico 
            Alignment       =   1  'Right Justify
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   420
            Left            =   960
            TabIndex        =   8
            Top             =   480
            Width           =   1215
         End
         Begin VB.TextBox txtCertificado 
            Height          =   420
            Left            =   960
            TabIndex        =   7
            Top             =   1080
            Width           =   5415
         End
         Begin VB.TextBox txtClave 
            Height          =   420
            Left            =   960
            TabIndex        =   6
            Top             =   1680
            Width           =   5415
         End
         Begin VB.CommandButton cmdCertificado 
            Caption         =   "..."
            Height          =   495
            Left            =   6480
            TabIndex        =   5
            Top             =   1080
            Width           =   495
         End
         Begin VB.CommandButton cmdClave 
            Caption         =   "..."
            Height          =   495
            Left            =   6480
            TabIndex        =   4
            Top             =   1680
            Width           =   495
         End
         Begin MSComDlg.CommonDialog dialogCerificado 
            Left            =   6600
            Top             =   720
            _ExtentX        =   847
            _ExtentY        =   847
            _Version        =   393216
            DialogTitle     =   "Certificado"
            FileName        =   "*.crt"
         End
         Begin MSComDlg.CommonDialog dialogClave 
            Left            =   6480
            Top             =   1320
            _ExtentX        =   847
            _ExtentY        =   847
            _Version        =   393216
            DialogTitle     =   "Certificado"
            FileName        =   "*.crt"
         End
         Begin VB.Label Label20 
            AutoSize        =   -1  'True
            Caption         =   "Consultar Ultimo Comprobante:"
            Height          =   195
            Left            =   240
            TabIndex        =   46
            Top             =   2880
            Width           =   2175
         End
         Begin VB.Label Label19 
            AutoSize        =   -1  'True
            Caption         =   "Modo Testing:"
            Height          =   195
            Left            =   3840
            TabIndex        =   44
            Top             =   480
            Width           =   1020
         End
         Begin VB.Label Label7 
            AutoSize        =   -1  'True
            Caption         =   "Puesto:"
            Height          =   195
            Left            =   240
            TabIndex        =   11
            Top             =   600
            Width           =   540
         End
         Begin VB.Label Label9 
            AutoSize        =   -1  'True
            Caption         =   "Certificado:"
            Height          =   195
            Left            =   120
            TabIndex        =   10
            Top             =   1200
            Width           =   795
         End
         Begin VB.Label Label10 
            AutoSize        =   -1  'True
            Caption         =   "Clave:"
            Height          =   195
            Left            =   360
            TabIndex        =   9
            Top             =   1800
            Width           =   450
         End
      End
   End
End
Attribute VB_Name = "frmParametros"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmdCancelar_Click()
    Unload Me
End Sub

Private Sub cmdCertificado_Click()
    dialogCerificado.FileName = "*.crt"
    dialogCerificado.ShowOpen
    txtCertificado = dialogCerificado.FileName
End Sub

Private Sub cmdClave_Click()
    dialogClave.FileName = "*.key"
    dialogClave.ShowOpen
    txtClave = dialogClave.FileName
End Sub

Private Sub cmdConsultar_Click()
    Dim Comprobante As String
    Dim Tipo As String
    If optFactura.Value = True Then Comprobante = "Factura"
    If optNC.Value = True Then Comprobante = "Nota de Crédito"
    If optND.Value = True Then Comprobante = "Nota de Débito"
    If optA.Value = True Then Tipo = "A"
    If optB.Value = True Then Tipo = "B"
    a = FacturaElectrónica.ConsultaUltimoComprobante(Comprobante, Tipo)
End Sub

Private Sub cmdGuardar_Click()
    If IsNumeric(txtPuesto) = False Then MsgBox ("El puesto no es válido"): Exit Sub
    If IsNumeric(txtComprobanteA) = False Then MsgBox ("El comprobante A no es válido"): Exit Sub
    If IsNumeric(txtComprobanteB) = False Then MsgBox ("El comprobante B no es válido"): Exit Sub
    If IsNumeric(txtRecibo) = False Then MsgBox ("El Recibo no es válido"): Exit Sub
    If IsNumeric(txtPorcentaje) = False Then MsgBox ("El Porcentaje no es válido"): Exit Sub
    If IsNumeric(txtPuestoElectronico) = False Then MsgBox ("El puesto Electrónico no es válido"): Exit Sub
    cn.Open
    cn.Execute ("UPDATE Parametros SET Puesto=" & txtPuesto & ",NumeroA=" & txtComprobanteA & " , NumeroB=" & txtComprobanteB & ",Recibo=" & txtRecibo & ",PorcentajeCtaCte=" & Replace(txtPorcentaje, ",", ".") & ", PuestoElectronico=" & txtPuestoElectrónico & ",Certificado='" & txtCertificado & "',Clave='" & txtClave & "',Nombre='" & txtNombre & "',Cuit='" & txtCuit & "',IngresosBrutos='" & txtIngresosBrutos & "',InicioActividades='" & txtInicioActividades & "',Domicilio='" & txtDomicilio & "',Localidad='" & txtLocalidad & "',CP='" & txtCP & "',Telefonos='" & txtTelefonos & "',Email='" & txtEmail & "',ModoTesting=" & cheModoTesting)
    cn.Close
    Unload Me
End Sub



Private Sub Form_Load()
    cn.Open
    Dim rs As ADODB.Recordset
    Set rs = cn.Execute("SELECT * from Parametros")
    txtPuesto = rs!Puesto
    txtComprobanteA = rs!NumeroA
    txtComprobanteB = rs!NumeroB
    txtRecibo = rs!Recibo
    txtPorcentaje = rs!PorcentajeCtaCte
    txtPuestoElectrónico = rs!PuestoElectronico
    txtCertificado = rs!Certificado
    txtClave = rs!Clave
    cheModoTesting = rs!ModoTesting
    txtNombre = rs!Nombre
    txtCuit = rs!Cuit
    txtIngresosBrutos = rs!IngresosBrutos
    txtInicioActividades = rs!InicioActividades
    txtDomicilio = rs!Domicilio
    txtLocalidad = rs!Localidad
    txtCP = rs!CP
    txtTelefonos = rs!Telefonos
    txtEmail = rs!Email
    cn.Close
    
    optFactura.Value = True
    optA.Value = True
   
End Sub

Private Sub Text1_Change()

End Sub

Private Sub txtComprobanteA_KeyPress(KeyAscii As Integer)
    If InStr(1, "0123456789." & Chr(13) & Chr(8), Chr(KeyAscii)) = 0 Then
        KeyAscii = 0
    End If
End Sub

Private Sub txtComprobanteB_KeyPress(KeyAscii As Integer)
    If InStr(1, "0123456789." & Chr(13) & Chr(8), Chr(KeyAscii)) = 0 Then
        KeyAscii = 0
    End If
End Sub

Private Sub txtPorcentaje_KeyPress(KeyAscii As Integer)
    If InStr(1, "0123456789." & Chr(13) & Chr(8), Chr(KeyAscii)) = 0 Then
        KeyAscii = 0
    End If
End Sub

Private Sub txtPuesto_KeyPress(KeyAscii As Integer)
    If InStr(1, "0123456789." & Chr(13) & Chr(8), Chr(KeyAscii)) = 0 Then
        KeyAscii = 0
    End If
End Sub

Private Sub txtRecibo_KeyPress(KeyAscii As Integer)
    If InStr(1, "0123456789." & Chr(13) & Chr(8), Chr(KeyAscii)) = 0 Then
        KeyAscii = 0
    End If
End Sub
