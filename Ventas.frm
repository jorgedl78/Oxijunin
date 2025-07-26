VERSION 5.00
Object = "{9C5C9460-5789-11DA-8CFB-0000E856BC17}#1.0#0"; "Fiscal051122.Ocx"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{648A5603-2C6E-101B-82B6-000000000014}#1.1#0"; "MSCOMM32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmFacturador 
   BackColor       =   &H00404040&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   11490
   ClientLeft      =   4860
   ClientTop       =   15
   ClientWidth     =   17445
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   FillStyle       =   0  'Solid
   BeginProperty Font 
      Name            =   "Courier New"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   FontTransparent =   0   'False
   ForeColor       =   &H00000000&
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   11490
   ScaleWidth      =   17445
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdQuitarDevuelto 
      Caption         =   "Quitar"
      Height          =   855
      Left            =   10200
      Picture         =   "Ventas.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   69
      Top             =   10200
      Width           =   975
   End
   Begin VB.CommandButton cmdAgregarDevuelto 
      Caption         =   "Agregar"
      Height          =   855
      Left            =   10200
      Picture         =   "Ventas.frx":0115
      Style           =   1  'Graphical
      TabIndex        =   68
      Top             =   9240
      Width           =   975
   End
   Begin VB.TextBox texIvaMitad 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00404040&
      BorderStyle     =   0  'None
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "0,00"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   11274
         SubFormatType   =   1
      EndProperty
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   360
      Left            =   14880
      Locked          =   -1  'True
      TabIndex        =   63
      Top             =   7920
      Width           =   1935
   End
   Begin VB.CommandButton cmdCaja 
      BackColor       =   &H00E0E0E0&
      Caption         =   "F3"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   21840
      MaskColor       =   &H00000000&
      Picture         =   "Ventas.frx":09DF
      Style           =   1  'Graphical
      TabIndex        =   62
      ToolTipText     =   "Caja"
      Top             =   18600
      UseMaskColor    =   -1  'True
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.CommandButton cmdAceptar 
      BackColor       =   &H00E0E0E0&
      Caption         =   "F2"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   15600
      MaskColor       =   &H00000000&
      Picture         =   "Ventas.frx":0FFC
      Style           =   1  'Graphical
      TabIndex        =   61
      ToolTipText     =   "Aceptar Ticket"
      Top             =   10200
      UseMaskColor    =   -1  'True
      Width           =   855
   End
   Begin VB.CommandButton cmdSalir 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Esc"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   13440
      Picture         =   "Ventas.frx":18C6
      Style           =   1  'Graphical
      TabIndex        =   60
      Top             =   10200
      Width           =   855
   End
   Begin VB.CommandButton cmdAgregar 
      Caption         =   "Agregar"
      Height          =   855
      Left            =   10200
      Picture         =   "Ventas.frx":2190
      Style           =   1  'Graphical
      TabIndex        =   56
      Top             =   6000
      Width           =   975
   End
   Begin VB.CommandButton cmdQuitar 
      Caption         =   "Quitar"
      Height          =   855
      Left            =   10200
      Picture         =   "Ventas.frx":2A5A
      Style           =   1  'Graphical
      TabIndex        =   55
      Top             =   6960
      Width           =   975
   End
   Begin VB.TextBox txtSubtotal 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00404040&
      BorderStyle     =   0  'None
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "0,00"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   11274
         SubFormatType   =   1
      EndProperty
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   360
      Left            =   11160
      Locked          =   -1  'True
      TabIndex        =   39
      Top             =   11280
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.TextBox txtRemito 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   6720
      TabIndex        =   54
      Top             =   11040
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.CheckBox chkConIva 
      BackColor       =   &H00404040&
      Caption         =   "Comprobante Con IVA"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   12120
      TabIndex        =   52
      Top             =   1920
      Value           =   1  'Checked
      Width           =   3375
   End
   Begin MSFlexGridLib.MSFlexGrid grDetalle 
      Height          =   2295
      Left            =   480
      TabIndex        =   51
      Top             =   2760
      Width           =   10455
      _ExtentX        =   18441
      _ExtentY        =   4048
      _Version        =   393216
      Rows            =   0
      Cols            =   0
      FixedRows       =   0
      FixedCols       =   0
      BackColor       =   14737632
      ForeColor       =   0
      BackColorFixed  =   14737632
      ForeColorFixed  =   0
      ForeColorSel    =   0
      BackColorBkg    =   14737632
      GridColor       =   14737632
      GridColorFixed  =   14737632
      GridLines       =   0
      GridLinesFixed  =   0
      MergeCells      =   1
      BorderStyle     =   0
      Appearance      =   0
   End
   Begin VB.CheckBox chkElectronica 
      BackColor       =   &H00404040&
      Caption         =   "Comprobante Electrónico"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   12120
      TabIndex        =   50
      Top             =   1560
      Value           =   1  'Checked
      Width           =   3855
   End
   Begin VB.TextBox txtComentario 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1410
      Left            =   12000
      MaxLength       =   100
      MultiLine       =   -1  'True
      TabIndex        =   48
      Top             =   3480
      Width           =   5175
   End
   Begin VB.TextBox txtNeto 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00404040&
      BorderStyle     =   0  'None
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "0,00"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   11274
         SubFormatType   =   1
      EndProperty
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   360
      Left            =   14400
      Locked          =   -1  'True
      TabIndex        =   41
      Top             =   6240
      Width           =   2535
   End
   Begin VB.TextBox txtImpuestos 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00404040&
      BorderStyle     =   0  'None
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "0,00"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   11274
         SubFormatType   =   1
      EndProperty
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   360
      Left            =   3240
      Locked          =   -1  'True
      TabIndex        =   40
      Top             =   11160
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.TextBox txtIva 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00404040&
      BorderStyle     =   0  'None
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "0,00"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   11274
         SubFormatType   =   1
      EndProperty
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   360
      Left            =   14760
      Locked          =   -1  'True
      TabIndex        =   38
      Top             =   7080
      Width           =   2055
   End
   Begin VB.CheckBox chkVencimiento 
      BackColor       =   &H00404040&
      Caption         =   "Check1"
      Height          =   255
      Left            =   9000
      TabIndex        =   37
      Top             =   11280
      Visible         =   0   'False
      Width           =   255
   End
   Begin MSComCtl2.DTPicker dateFechaFactura 
      Height          =   375
      Left            =   12960
      TabIndex        =   33
      Top             =   0
      Visible         =   0   'False
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   661
      _Version        =   393216
      Enabled         =   0   'False
      Format          =   40173569
      CurrentDate     =   42366
   End
   Begin VB.Frame frComprobante 
      BackColor       =   &H00404040&
      Caption         =   "Comprobante"
      ForeColor       =   &H00FFFFFF&
      Height          =   1095
      Left            =   13080
      TabIndex        =   30
      Top             =   120
      Width           =   4215
      Begin VB.OptionButton optNotaDebito 
         BackColor       =   &H00404040&
         Caption         =   "N. Débito"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   2760
         TabIndex        =   49
         Top             =   240
         Width           =   1335
      End
      Begin VB.OptionButton optNotaCredito 
         BackColor       =   &H00404040&
         Caption         =   "N. Crédito"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   1320
         TabIndex        =   32
         Top             =   240
         Width           =   1455
      End
      Begin VB.OptionButton optFactura 
         BackColor       =   &H00404040&
         Caption         =   "Factura"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   120
         TabIndex        =   31
         Top             =   240
         Width           =   1095
      End
      Begin VB.Label lblCondicion 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "FORMA DE PAGO"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   375
         Left            =   240
         TabIndex        =   65
         Top             =   600
         Width           =   3735
      End
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Enabled         =   0   'False
      Height          =   495
      Left            =   9720
      TabIndex        =   25
      Top             =   6480
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.CommandButton cmdBuscarCliente 
      BackColor       =   &H00E0E0E0&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   240
      Picture         =   "Ventas.frx":2B6F
      Style           =   1  'Graphical
      TabIndex        =   20
      Top             =   360
      Width           =   375
   End
   Begin MSCommLib.MSComm MSComm1 
      Left            =   16080
      Top             =   0
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      DTREnable       =   -1  'True
      InputLen        =   1
      RThreshold      =   1
   End
   Begin VB.TextBox txtCantidad 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   360
      TabIndex        =   0
      Top             =   1815
      Width           =   735
   End
   Begin VB.TextBox txtBarras 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   1440
      TabIndex        =   1
      Top             =   1800
      Width           =   8775
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid grDetalleDeprecated 
      Height          =   735
      Left            =   10680
      TabIndex        =   13
      Top             =   2400
      Visible         =   0   'False
      Width           =   855
      _ExtentX        =   1508
      _ExtentY        =   1296
      _Version        =   393216
      BackColor       =   14737632
      Rows            =   0
      FixedRows       =   0
      FixedCols       =   0
      BackColorFixed  =   14737632
      BackColorBkg    =   14737632
      BackColorUnpopulated=   14737632
      GridColor       =   14737632
      GridColorFixed  =   16776960
      Enabled         =   0   'False
      GridLinesUnpopulated=   1
      MergeCells      =   1
      BorderStyle     =   0
      Appearance      =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
      _Band(0).GridLinesBand=   1
      _Band(0).TextStyleBand=   0
      _Band(0).TextStyleHeader=   0
   End
   Begin VB.CommandButton cmdBuscar 
      BackColor       =   &H00E0E0E0&
      Caption         =   "F1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   10680
      Picture         =   "Ventas.frx":3571
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   1680
      Width           =   495
   End
   Begin VB.TextBox txtVuelto 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00404040&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   10560
      Locked          =   -1  'True
      TabIndex        =   11
      Top             =   6480
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.TextBox txtCredito 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00404040&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   10680
      TabIndex        =   10
      Top             =   6480
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.TextBox txtDebito 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00404040&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   10200
      TabIndex        =   9
      Top             =   6480
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.TextBox txtEfectivo 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00404040&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   9840
      TabIndex        =   8
      Top             =   6480
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.TextBox txtTotal 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00404040&
      BorderStyle     =   0  'None
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "0,00"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   11274
         SubFormatType   =   1
      EndProperty
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   360
      Left            =   14400
      Locked          =   -1  'True
      TabIndex        =   7
      Top             =   8760
      Width           =   2535
   End
   Begin MSComCtl2.DTPicker dateFechaVencimiento 
      Height          =   375
      Left            =   9240
      TabIndex        =   35
      Top             =   11160
      Visible         =   0   'False
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   661
      _Version        =   393216
      Enabled         =   0   'False
      Format          =   40173569
      CurrentDate     =   42366
   End
   Begin MSFlexGridLib.MSFlexGrid grDetalleTubos 
      Height          =   2175
      Left            =   480
      TabIndex        =   57
      Top             =   6000
      Width           =   9015
      _ExtentX        =   15901
      _ExtentY        =   3836
      _Version        =   393216
      Rows            =   0
      Cols            =   0
      FixedRows       =   0
      FixedCols       =   0
      BackColor       =   14737632
      ForeColor       =   0
      BackColorFixed  =   14737632
      ForeColorFixed  =   0
      ForeColorSel    =   0
      BackColorBkg    =   14737632
      GridColor       =   14737632
      GridColorFixed  =   14737632
      GridLines       =   0
      GridLinesFixed  =   0
      MergeCells      =   1
      BorderStyle     =   0
      Appearance      =   0
   End
   Begin MSFlexGridLib.MSFlexGrid grTubosDevueltos 
      Height          =   1695
      Left            =   480
      TabIndex        =   70
      Top             =   9240
      Width           =   9015
      _ExtentX        =   15901
      _ExtentY        =   2990
      _Version        =   393216
      Rows            =   0
      Cols            =   0
      FixedRows       =   0
      FixedCols       =   0
      BackColor       =   14737632
      ForeColor       =   0
      BackColorFixed  =   14737632
      ForeColorFixed  =   0
      ForeColorSel    =   0
      BackColorBkg    =   14737632
      GridColor       =   14737632
      GridColorFixed  =   14737632
      GridLines       =   0
      GridLinesFixed  =   0
      MergeCells      =   1
      BorderStyle     =   0
      Appearance      =   0
   End
   Begin VB.Label lblLetra 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "lblLetra"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   270
      Left            =   8400
      TabIndex        =   26
      Top             =   480
      Visible         =   0   'False
      Width           =   1200
   End
   Begin VB.Shape Shape26 
      BorderColor     =   &H00FFFFFF&
      Height          =   1455
      Left            =   11760
      Shape           =   4  'Rounded Rectangle
      Top             =   9840
      Width           =   5535
   End
   Begin VB.Label Label22 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Tubos a devolver:"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   270
      Left            =   600
      TabIndex        =   73
      Top             =   8760
      Width           =   2550
   End
   Begin VB.Label Label21 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Tubos a entregar:"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   270
      Left            =   480
      TabIndex        =   72
      Top             =   5520
      Width           =   2550
   End
   Begin VB.Label Label20 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Comentario"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   270
      Left            =   3600
      TabIndex        =   71
      Top             =   -1320
      Width           =   1500
   End
   Begin VB.Shape Shape25 
      BorderColor     =   &H00FFFFFF&
      Height          =   2655
      Left            =   240
      Shape           =   4  'Rounded Rectangle
      Top             =   8640
      Width           =   11175
   End
   Begin VB.Shape Shape12 
      BackColor       =   &H00E0E0E0&
      BackStyle       =   1  'Opaque
      Height          =   1935
      Left            =   360
      Shape           =   4  'Rounded Rectangle
      Top             =   9120
      Width           =   9615
   End
   Begin VB.Label Label13 
      BackStyle       =   0  'Transparent
      Caption         =   "Precio:"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   270
      Left            =   5040
      TabIndex        =   67
      Top             =   720
      Width           =   1170
   End
   Begin VB.Label lblTipoPrecio 
      BackStyle       =   0  'Transparent
      Caption         =   "Público"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   270
      Left            =   6120
      TabIndex        =   66
      Top             =   720
      Width           =   2130
   End
   Begin VB.Shape Shape24 
      BorderColor     =   &H00FFFFFF&
      Height          =   495
      Left            =   14280
      Shape           =   4  'Rounded Rectangle
      Top             =   7800
      Width           =   2775
   End
   Begin VB.Label Label19 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "IVA 10.5%:"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   405
      Left            =   12120
      TabIndex        =   64
      Top             =   7920
      Width           =   2100
   End
   Begin VB.Shape Shape23 
      BorderColor     =   &H00FFFFFF&
      Height          =   495
      Left            =   -6840
      Shape           =   4  'Rounded Rectangle
      Top             =   0
      Width           =   2775
   End
   Begin VB.Label lblCantidadTubosEnCliente 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "lblCantidadDeTubosEnCliente"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   270
      Left            =   7680
      TabIndex        =   59
      Top             =   840
      Width           =   5010
   End
   Begin VB.Label lblUltimoUsoFacturado 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "lblUltimoUsoFacturado"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   270
      Left            =   7680
      TabIndex        =   58
      Top             =   600
      Width           =   5010
   End
   Begin VB.Shape Shape22 
      BackColor       =   &H00E0E0E0&
      BackStyle       =   1  'Opaque
      Height          =   2535
      Left            =   360
      Shape           =   4  'Rounded Rectangle
      Top             =   5880
      Width           =   9615
   End
   Begin VB.Label lblSaldo 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "lblSaldo"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   270
      Left            =   10440
      TabIndex        =   53
      Top             =   360
      Width           =   2130
   End
   Begin VB.Shape Shape21 
      BorderColor     =   &H00FFFFFF&
      Height          =   975
      Left            =   120
      Shape           =   4  'Rounded Rectangle
      Top             =   1320
      Width           =   11295
   End
   Begin VB.Shape Shape20 
      BackColor       =   &H00E0E0E0&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      Height          =   1695
      Left            =   11880
      Shape           =   4  'Rounded Rectangle
      Top             =   3360
      Width           =   5415
   End
   Begin VB.Label Label18 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Comentario"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   270
      Left            =   12000
      TabIndex        =   47
      Top             =   2880
      Width           =   1500
   End
   Begin VB.Label Label17 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "IVA 21%:"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   405
      Left            =   12480
      TabIndex        =   46
      Top             =   7080
      Width           =   1680
   End
   Begin VB.Label Label15 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Impuestos:"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   405
      Left            =   2880
      TabIndex        =   44
      Top             =   11040
      Visible         =   0   'False
      Width           =   2100
   End
   Begin VB.Label Label14 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Neto:"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   405
      Left            =   13080
      TabIndex        =   43
      Top             =   6240
      Width           =   1050
   End
   Begin VB.Shape Shape19 
      BorderColor     =   &H00FFFFFF&
      Height          =   975
      Left            =   11760
      Shape           =   4  'Rounded Rectangle
      Top             =   1320
      Width           =   5535
   End
   Begin VB.Shape Shape18 
      BackColor       =   &H00E0E0E0&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      Height          =   255
      Left            =   6600
      Shape           =   4  'Rounded Rectangle
      Top             =   11160
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.Label Label12 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Remito Nro:"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   270
      Left            =   10800
      TabIndex        =   42
      Top             =   11160
      Visible         =   0   'False
      Width           =   1650
   End
   Begin VB.Shape Shape17 
      BorderColor     =   &H00FFFFFF&
      Height          =   2580
      Left            =   11640
      Shape           =   4  'Rounded Rectangle
      Top             =   2640
      Width           =   5775
   End
   Begin VB.Shape Shape16 
      BorderColor     =   &H00FFFFFF&
      Height          =   495
      Left            =   14280
      Shape           =   4  'Rounded Rectangle
      Top             =   6120
      Width           =   2775
   End
   Begin VB.Shape Shape15 
      BorderColor     =   &H00FFFFFF&
      Height          =   495
      Left            =   3120
      Shape           =   4  'Rounded Rectangle
      Top             =   11040
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.Shape Shape14 
      BorderColor     =   &H00FFFFFF&
      Height          =   495
      Left            =   5160
      Shape           =   4  'Rounded Rectangle
      Top             =   11040
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Shape Shape13 
      BorderColor     =   &H00FFFFFF&
      Height          =   495
      Left            =   14280
      Shape           =   4  'Rounded Rectangle
      Top             =   6960
      Width           =   2775
   End
   Begin VB.Label Label11 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Vencimiento"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   270
      Left            =   7200
      TabIndex        =   36
      Top             =   11280
      Visible         =   0   'False
      Width           =   1650
   End
   Begin VB.Label Label10 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Fecha:"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   270
      Left            =   12840
      TabIndex        =   34
      Top             =   0
      Visible         =   0   'False
      Width           =   900
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "-"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   270
      Left            =   15000
      TabIndex        =   29
      Top             =   840
      Visible         =   0   'False
      Width           =   150
   End
   Begin VB.Label lblNumero 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   270
      Left            =   15240
      TabIndex        =   28
      Top             =   840
      Visible         =   0   'False
      Width           =   150
   End
   Begin VB.Label lblPuesto 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   270
      Left            =   14280
      TabIndex        =   27
      Top             =   840
      Visible         =   0   'False
      Width           =   150
   End
   Begin VB.Label lblNumeroDocumento 
      BackStyle       =   0  'Transparent
      Caption         =   "lblNumeroDocumento"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   270
      Left            =   1560
      TabIndex        =   24
      Top             =   720
      Width           =   2130
   End
   Begin VB.Label lblTipoDocumento 
      BackStyle       =   0  'Transparent
      Caption         =   "lblTipoDocumento"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   270
      Left            =   240
      TabIndex        =   23
      Top             =   720
      Width           =   1170
   End
   Begin VB.Label lblCategoria 
      BackStyle       =   0  'Transparent
      Caption         =   "lblCategoria"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   270
      Left            =   5040
      TabIndex        =   22
      Top             =   360
      Width           =   3810
   End
   Begin VB.Label lblIdCliente 
      Caption         =   "Label10"
      Height          =   255
      Left            =   240
      TabIndex        =   21
      Top             =   240
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Label lblCliente 
      BackStyle       =   0  'Transparent
      Caption         =   "lblCliente"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   270
      Left            =   720
      TabIndex        =   19
      Top             =   360
      Width           =   3690
   End
   Begin FiscalPrinterLibCtl.HASAR HASAR1 
      Left            =   15120
      OleObjectBlob   =   "Ventas.frx":3F73
      Top             =   0
   End
   Begin VB.Label lblCaja 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Cerrada"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   270
      Left            =   1200
      TabIndex        =   18
      Top             =   0
      Width           =   1050
   End
   Begin VB.Label Label9 
      BackStyle       =   0  'Transparent
      Caption         =   "Caja Nº:"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   270
      Left            =   120
      TabIndex        =   17
      Top             =   0
      Width           =   1200
   End
   Begin VB.Label lblCajero 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Cajero:"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   270
      Left            =   3480
      TabIndex        =   16
      Top             =   0
      Width           =   1050
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Código de Barras"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   270
      Left            =   1800
      TabIndex        =   15
      Top             =   1365
      Width           =   2400
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Cantidad"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   270
      Left            =   195
      TabIndex        =   14
      Top             =   1365
      Width           =   1200
   End
   Begin VB.Shape Shape11 
      BackColor       =   &H00E0E0E0&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      Height          =   495
      Left            =   240
      Shape           =   4  'Rounded Rectangle
      Top             =   1680
      Width           =   975
   End
   Begin VB.Shape Shape10 
      BorderColor     =   &H00FFFFFF&
      Height          =   3135
      Left            =   240
      Shape           =   4  'Rounded Rectangle
      Top             =   5400
      Width           =   11175
   End
   Begin VB.Shape Shape9 
      BackColor       =   &H00E0E0E0&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      Height          =   495
      Left            =   1320
      Shape           =   4  'Rounded Rectangle
      Top             =   1680
      Width           =   9135
   End
   Begin VB.Shape Shape8 
      BorderColor     =   &H00FFFFFF&
      Height          =   735
      Left            =   8640
      Shape           =   4  'Rounded Rectangle
      Top             =   6480
      Visible         =   0   'False
      Width           =   2775
   End
   Begin VB.Shape Shape7 
      BorderColor     =   &H00FFFFFF&
      Height          =   735
      Left            =   8520
      Shape           =   4  'Rounded Rectangle
      Top             =   6480
      Visible         =   0   'False
      Width           =   2775
   End
   Begin VB.Shape Shape6 
      BorderColor     =   &H00FFFFFF&
      Height          =   735
      Left            =   8640
      Shape           =   4  'Rounded Rectangle
      Top             =   6600
      Visible         =   0   'False
      Width           =   2775
   End
   Begin VB.Shape Shape5 
      BorderColor     =   &H00FFFFFF&
      Height          =   735
      Left            =   8640
      Shape           =   4  'Rounded Rectangle
      Top             =   6480
      Visible         =   0   'False
      Width           =   2775
   End
   Begin VB.Shape Shape4 
      BorderColor     =   &H00FFFFFF&
      Height          =   495
      Left            =   14280
      Shape           =   4  'Rounded Rectangle
      Top             =   8640
      Width           =   2775
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Vuelto:"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   405
      Left            =   9720
      TabIndex        =   6
      Top             =   6840
      Visible         =   0   'False
      Width           =   1470
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Crédito:"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   405
      Left            =   9480
      TabIndex        =   5
      Top             =   6840
      Visible         =   0   'False
      Width           =   1680
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Débito:"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   405
      Left            =   9840
      TabIndex        =   4
      Top             =   6960
      Visible         =   0   'False
      Width           =   1470
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Efectivo:"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   405
      Left            =   9360
      TabIndex        =   3
      Top             =   6960
      Visible         =   0   'False
      Width           =   1890
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Total:"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   405
      Left            =   12960
      TabIndex        =   2
      Top             =   8760
      Width           =   1260
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H00FFFFFF&
      Height          =   975
      Left            =   120
      Shape           =   4  'Rounded Rectangle
      Top             =   240
      Width           =   12735
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00E0E0E0&
      BackStyle       =   1  'Opaque
      Height          =   2655
      Left            =   240
      Shape           =   4  'Rounded Rectangle
      Top             =   2640
      Width           =   11175
   End
   Begin VB.Label Label16 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Subtotal:"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   405
      Left            =   3720
      TabIndex        =   45
      Top             =   11040
      Visible         =   0   'False
      Width           =   1890
   End
   Begin VB.Shape Shape3 
      BackColor       =   &H00FFFFFF&
      BorderColor     =   &H00FFFFFF&
      Height          =   4215
      Left            =   11640
      Shape           =   4  'Rounded Rectangle
      Top             =   5400
      Width           =   5655
   End
End
Attribute VB_Name = "frmFacturador"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Total As Double
Dim Efectivo As Double
Dim Debito As Double
Dim Credito As Double
Dim Vuelto As Double
Dim TotalAcumulado As Double
Dim NetoAcumulado As Double
Dim NetoAcumuladoMitad As Double
Dim ImpuestoAcumulado As Double
Dim IvaAcumulado As Double
Dim IvaAcumuladoMitad As Double
Dim PorcentajeCtaCte As Double

Dim Nombre As String
Dim NumeroDoc As String
Dim TipoDoc As TiposDeDocumento
Dim Comprobante As DocumentosFiscales
Dim Responsable As TiposDeResponsabilidades
Dim HayScanner As String
Dim PuestoFiscal As Integer
Dim ComprobanteFiscal As String
Dim Buffer As String

Private Sub chkVencimiento_Click()
    If chkVencimiento.Value = 1 Then
        dateFechaVencimiento.Enabled = True
    Else
        dateFechaVencimiento.Enabled = False
    End If
        
End Sub

Private Sub cmdAceptar_Click()
    
'MsgBox (HASAR1.IndicadorFiscal(2048))
'HASAR1.CapacidadRestante
'Capacidad = Round(100 * (1 - PFiscal.Respuesta(4) / PFiscal.Respuesta(3)), 2)

    If TotalAcumulado = 0 Then MsgBox ("Comprobante sin movimientos"): Exit Sub
    'MsgBox (Efectivo - Vuelto + Debito + Credito)
    'If (Efectivo - Vuelto + Debito + Credito) <> Total Then
    '    MsgBox ("El detalle de pago no coincide con el total")
    '    Exit Sub
    'End If
    If lblNumeroDocumento = "0" And lblCategoria <> "Consumidor Final" Then MsgBox ("El número de CUIT no puede estar vacio para esta categoria de IVA"): Exit Sub
    'If lblNumeroDocumento = "0" And lblCategoria = "Consumidor Final" And txtTotal >= 1000 Then MsgBox ("Si el importe es igual o mayor a $1000 debe especificar el dni del cliente"): Exit Sub
    If lblCondicion.Caption = "FORMA DE PAGO" Then MsgBox ("Debe especificar la forma de Pago"): Exit Sub
    
    
    Respuesta = MsgBox("¿Confirma el comprobante?", vbYesNo, "")
    If Respuesta = vbNo Then Exit Sub
    'Confirmo comprobante
    Dim rs As ADODB.Recordset
    
    Dim TipoComprobante As String
    If optFactura.Value = True Then
        TipoComprobante = "Factura"
    End If
    If optNotaCredito.Value = True Then
        TipoComprobante = "Nota de Crédito"
    End If
    If optNotaDebito.Value = True Then
        TipoComprobante = "Nota de Débito"
    End If
    
    CAE = "0"
    CaeVencimiento = Format(dateFechaFactura, "yyyy/mm/dd")
    If ComprobanteFiscal <> "NO" Then
        Pedido = 0
        'ImpresionFiscal
        Letra = lblLetra
        
        If chkElectronica.Value = 1 Then
            a = FacturaElectronica(TipoComprobante, lblLetra, lblTipoDocumento, lblNumeroDocumento, NetoAcumulado, NetoAcumuladoMitad, txtIva, texIvaMitad, txtTotal, txtImpuestos, lblCategoria)
            If CAE = "" Then Exit Sub
            NumeroComprobante = cbte_nro
            PuestoFiscal = punto_vta
            CaeVencimiento = Mid(VencimientoCAE, 1, 4) & "/" & Mid(VencimientoCAE, 5, 2) & "/" & Mid(VencimientoCAE, 7, 2)
        Else
            NumeroComprobante = lblNumero
            PuestoFiscal = lblPuesto
        End If
        
        
    Else
        cn.Open
        Set rs = cn.Execute("SELECT MAX(numero) + 1 AS NuevoP FROM VENTAS WHERE pedido=1")
        If rs!NuevoP > 0 Then
            NumeroComprobante = rs!NuevoP
        Else
            NumeroComprobante = 1
        End If
        rs.Close
        Set rs = Nothing
        cn.Close
        Letra = "P"
        PuestoFiscal = 0
        Pedido = 1
    End If
    
    'posiciones de la grilla de articulos
    '0=cantidad
    '1=descripcion
    '2=PrecioVenta
    '3=idrticulo
    '4=costo total
    '5=impuesto total
    '6=neto 21 total
    '7=iva 21 total
    '10=neto 105 total
    '9=iva 105 total
    
    
    Impuestos = 0
      
    cn.Open
    cn.Execute ("AgregarVenta '" & Format(dateFechaFactura, "yyyy/mm/dd") & "','" & Letra & "'," & PuestoFiscal & "," & NumeroComprobante & "," & Replace(NetoAcumulado, ",", ".") & "," & Replace(NetoAcumuladoMitad, ",", ".") & "," & Replace(IvaAcumulado, ",", ".") & "," & Replace(IvaAcumuladoMitad, ",", ".") & "," & Replace(TotalAcumulado, ",", ".") & "," & Val(lblIdCliente) & "," & Pedido & ",'" & TipoComprobante & "'," & Replace(ImpuestoAcumulado, ",", ".") & ",'" & lblCondicion & "'," & chkElectronica.Value & ",'" & CAE & "','" & CaeVencimiento & "', '" & txtComentario & "'")
    Set rs = cn.Execute("SELECT MAX(idventa) AS Nuevoid FROM VENTAS")
    NuevoID = rs!NuevoID
    idComprobante = NuevoID
    With grDetalle
    For i = 0 To grDetalle.Rows - 1
    'MsgBox ("AgregaDetalleVenta " & Replace(Val(.TextMatrix(i, 6)), ",", ".") & ", " & Replace(Val(.TextMatrix(i, 10)), ",", ".") & "," & Replace(.TextMatrix(i, 0), ",", ".") & "," & Replace(.TextMatrix(i, 4), ",", ".") & "," & NuevoID & "," & .TextMatrix(i, 3) & "," & Replace(.TextMatrix(i, 7), ",", ".") & "," & Replace(.TextMatrix(i, 9), ",", ".") & "," & Replace(.TextMatrix(i, 2), ",", "."))
        cn.Execute ("AgregaDetalleVenta " & Replace(Val(.TextMatrix(i, 6)), ",", ".") & ", " & Replace(Val(.TextMatrix(i, 10)), ",", ".") & "," & Replace(.TextMatrix(i, 0), ",", ".") & "," & Replace(.TextMatrix(i, 4), ",", ".") & "," & NuevoID & "," & .TextMatrix(i, 3) & "," & Replace(.TextMatrix(i, 7), ",", ".") & "," & Replace(.TextMatrix(i, 9), ",", ".") & "," & Replace(.TextMatrix(i, 2), ",", "."))
        If TipoComprobante = "Nota de Crédito" Then
            multiplicador = -1
        Else
            multiplicador = 1
        End If
        cn.Execute ("DescuentaStock " & .TextMatrix(i, 3) & "," & Replace(.TextMatrix(i, 0), ",", ".") * multiplicador)
    Next i
    End With
    
    'agrego el detalle de tubos vendidos
    With grDetalleTubos
    For i = 0 To grDetalleTubos.Rows - 1
        cn.Execute ("INSERT INTO DetalleTubosVendidos(Importe, idVenta, idTubo) VALUES (" & Replace(.TextMatrix(i, 6), ",", ".") & " , " & NuevoID & ", " & .TextMatrix(i, 0) & ")")
        cn.Execute ("UPDATE Tubos set idEstadoTubos=1, ClienteActual= " & lblIdCliente & ", UltimoMovimiento= '" & Format(dateFechaFactura, "dd/mm/yyyy") & "', DetalleUltimo='" & TipoComprobante & " " & Letra & Format(PuestoFiscal, "0000") & "-" & Format(NumeroComprobante, "00000000") & "' where idTubo=" & .TextMatrix(i, 0))
        cn.Execute ("INSERT INTO MovimientosTubos (Fecha, Detalle, idTubo) VALUES ('" & Format(dateFechaFactura, "dd/mm/yyyy") & "', 'OXIJUNIN A CLIENTE " & TipoComprobante & " " & Letra & Format(PuestoFiscal, "0000") & "-" & Format(NumeroComprobante, "00000000") & "'," & .TextMatrix(i, 0) & ")")
    Next i
    End With

    'agrego el detalle de tubos devueltos
    With grTubosDevueltos
    For i = 0 To grTubosDevueltos.Rows - 1
        cn.Execute ("INSERT INTO DetalleTubosDevueltos(idVenta, idTubo) VALUES (" & NuevoID & ", " & .TextMatrix(i, 0) & ")")
        cn.Execute ("UPDATE Tubos set idEstadoTubos=2, ClienteActual= 1, UltimoMovimiento= '" & Format(dateFechaFactura, "dd/mm/yyyy") & "', DetalleUltimo='" & TipoComprobante & " " & Letra & Format(PuestoFiscal, "0000") & "-" & Format(NumeroComprobante, "00000000") & "' where idTubo=" & .TextMatrix(i, 0))
        cn.Execute ("INSERT INTO MovimientosTubos (Fecha, Detalle, idTubo) VALUES ('" & Format(dateFechaFactura, "dd/mm/yyyy") & "', 'CLIENTE A OXIJUNIN " & TipoComprobante & " " & Letra & Format(PuestoFiscal, "0000") & "-" & Format(NumeroComprobante, "00000000") & "'," & .TextMatrix(i, 0) & ")")
    Next i
    End With
    
    
    
    cn.Execute ("AgregarDetalleCaja " & lblCaja & ",'" & Format(dateFechaFactura, "yyyy/mm/dd") & "'," & Replace(Efectivo - Vuelto, ",", ".") & "," & Replace(Debito, ",", ".") & "," & Replace(Credito, ",", ".") & ",1,'Factura " & txtNumero & "'")

    If lblCondicion = "CUENTA CORRIENTE" Then
        If TipoComprobante = "Nota de Crédito" Then
            cn.Execute ("AgregarCuentaCorriente '" & Format(dateFechaFactura, "yyyy/mm/dd") & "','" & TipoComprobante & " " & Letra & Format(PuestoFiscal, "0000") & "-" & Format(NumeroComprobante, "00000000") & "',0," & Replace(TotalAcumulado, ",", ".") & "," & Val(lblIdCliente) & ",'Ven'," & NuevoID)
        Else
            cn.Execute ("AgregarCuentaCorriente '" & Format(dateFechaFactura, "yyyy/mm/dd") & "','" & TipoComprobante & " " & Letra & Format(PuestoFiscal, "0000") & "-" & Format(NumeroComprobante, "00000000") & "'," & Replace(TotalAcumulado, ",", ".") & "," & "0," & Val(lblIdCliente) & ",'Ven'," & NuevoID)
        End If
    End If
    
    cn.Close
    If ComprobanteFiscal <> "NO" Then
        cn.Open
              
        If chkElectronica.Value = 0 Then 'es una factura manual asique actualizo numeración
            If lblLetra = "A" Then
                cn.Execute ("UPDATE Parametros set NumeroA = NumeroA + 1")
            Else
                cn.Execute ("UPDATE Parametros set NumeroB = NumeroB + 1")
            End If
        End If
        cn.Close
        If chkVencimiento.Value = 1 Then
            frmImprimeFactura.lblVencimiento = dateFechaVencimiento.Value
        Else
            frmImprimeFactura.lblVencimiento = ""
        End If
        frmImprimeFactura.lblRemito = txtRemito
        frmImprimeFactura.lblComentario = txtComentario
        If lblLetra = "A" And chkElectronica.Value = 0 Then
            MsgBox ("Prepare la hoja preimpresa para emitir: " & TipoComprobante & " " & Letra & Format(PuestoFiscal, "0000") & "-" & Format(NumeroComprobante, "00000000"))
            frmImprimeFactura.PrintForm
            MsgBox ("Prepare la copia preimpresa para emitir: " & TipoComprobante & " " & Letra & Format(PuestoFiscal, "0000") & "-" & Format(NumeroComprobante, "00000000"))
            frmImprimeFactura.PrintForm
            Unload frmImprimeFactura
        End If
        If chkElectronica.Value = 1 Then
            'frmImprimeFacturaElectronica.lblVencimiento = "Vencimiento: " & dateFechaVencimiento.Value
            'frmImprimeFacturaElectronica.lblComentario = txtComentario
            'frmImprimeFacturaElectronica.lblRemito = txtRemito
            'frmImprimeFacturaElectronica.Show 1
            'frmImprimeFacturaElectronica.lblfacturaOriginal = "Original"
            'frmImprimeFacturaElectronica.PrintForm
            'Unload frmImprimeFacturaElectronica
            'frmImprimeFacturaElectronica.lblfacturaOriginal = "Duplicado"
            'frmImprimeFacturaElectronica.PrintForm
            'Unload frmImprimeFacturaElectronica
            'MsgBox ("Imprime ORIGINAL")
            cn.Open
            condicionComprobante = "Original"
            ImprimeFacturaElectronica.WindowState = 2
            ImprimeFacturaElectronica.PrintReport
            Unload ImprimeFacturaElectronica
            'MsgBox ("Imprime DUPLICADO")
            condicionComprobante = "Duplicado"
            ImprimeFacturaElectronica.WindowState = 2
            ImprimeFacturaElectronica.PrintReport
            Unload ImprimeFacturaElectronica
            cn.Close
         End If
        'frmImprimeFactura.Show 1
        
     End If


    grDetalle.Rows = 0
    grDetalleTubos.Rows = 0
    grTubosDevueltos.Rows = 0
    CalcularTotales
    'Total = 0
    'txtTotal = ""
    'txtEfectivo = ""
    'txtDebito = ""
    'txtCredito = ""
    'txtVuelto = ""
    lblIdCliente = 1
    lblCliente = "CONSUMIDOR FINAL"
    lblCategoria = "Consumidor Final"
    lblTipoDocumento = ""
    lblNumeroDocumento = ""
    optFactura.Value = True
    lblCondicion = "CUENTA CORRIENTE"
    chkConIva.Value = 1
    lblCondicion.ForeColor = vbBlue
    lblUltimoUsoFacturado = ""
    lblCantidadTubosEnCliente = ""
    optFactura.Value = True
    dateFechaFactura = Date
    
    cn.Open
    Set rs = cn.Execute("SELECT (NumeroB + 1) as Numero from Parametros")
    lblLetra = "B"
    frmFacturador.lblPuesto = "0002"
    frmFacturador.lblNumero = Format(rs!numero, "00000000")
    cn.Close
    
    txtBarras.SetFocus
    Exit Sub
    
'impresora_apag:
'    If MsgBox("Error Impresora:" & Err.Description, vbRetryCancel, "Errores") = vbRetry Then
'        Resume Imprimir
'    End If
End Sub
Sub ImpresionFiscal()
        'On Error GoTo impresora_apag
'Imprimir:
        Select Case lblCategoria
            Case "Consumidor Final"
                Letra = "B"
                If lblCliente = "CONSUMIDOR FINAL" And Total <= 1000 Then
                    'es un consumidor final y no supera los $1000. Sale ticket
                    Nombre = ""
                    NumeroDoc = ""
                    TipoDoc = TIPO_NINGUNO
                    Comprobante = TICKET_C
                    Responsable = CONSUMIDOR_FINAL
                Else
                    'es un cliente seleccionado consumidor final o un consumidor final pero con importe mayor a $1000
                    Nombre = Mid(lblCliente, 1, 40)
                    NumeroDoc = lblNumeroDocumento
                    TipoDoc = TIPO_NINGUNO
                    Comprobante = TICKET_FACTURA_B
                    Responsable = CONSUMIDOR_FINAL
                End If
            Case "Monotributo"
                'es un cliente seleccionado consumidor final o un consumidor final pero con importe mayor a $1000
                Nombre = Mid(lblCliente, 1, 40)
                NumeroDoc = lblNumeroDocumento
                TipoDoc = TIPO_NINGUNO
                Comprobante = TICKET_FACTURA_B
                Responsable = CONSUMIDOR_FINAL
                Letra = "B"
            Case "Responsable Inscripto"
                Nombre = Mid(lblCliente, 1, 40)
                NumeroDoc = lblNumeroDocumento
                TipoDoc = TIPO_CUIT
                Comprobante = TICKET_FACTURA_A
                Responsable = RESPONSABLE_INSCRIPTO
                Letra = "A"
        End Select
        HASAR1.Encabezado(1) = Chr(244) & "     S  U  M  A"
        HASAR1.Encabezado(2) = "     R.S. Peña 245 - Junìn (Bs. As.)"
        HASAR1.Encabezado(3) = "          Te: (0236) 4443018        "
        
        HASAR1.DatosCliente Nombre, NumeroDoc, TipoDoc, Responsable

        'HayError = PFiscal.HuboErrorFiscal Or PFiscal.HuboErrorMecanico Or PFiscal.HuboFaltaPapel
        'If HayError Then MsgBox "Los datos del cliente son incorrectos o el cuit es inválido": Exit Sub
        
        
        HASAR1.AbrirComprobanteFiscal Comprobante
        'HayError = PFiscal.HuboErrorFiscal Or PFiscal.HuboErrorMecanico Or PFiscal.HuboFaltaPapel
        'If HayError Then MsgBox "No se puede abrir el comprobante fiscal. Se realizó cierre Z?": Exit Sub
    
        'HASAR1.ImprimirTextoFiscal "Texto Fiscal..."
        With grDetalle
        For i = 0 To grDetalle.Rows - 1
            HASAR1.ImprimirItem Mid(.TextMatrix(i, 1), 1, 20), .TextMatrix(i, 0), Val(.TextMatrix(i, 2)) / Val(.TextMatrix(i, 0)), 21, 0
        Next i
        End With
        'HASAR1.ImprimirPago "Efectivo", Val(Total)
        HASAR1.CerrarComprobanteFiscal
        If lblCategoria = "Responsable Inscripto" Then
            NumeroComprobante = HASAR1.UltimoDocumentoFiscalA
        Else
            NumeroComprobante = HASAR1.UltimoDocumentoFiscalBC
        End If
End Sub
Sub EntregarMercaderia()
    If (txtEfectivo - txtVuelto + txtDebito + txtCredito) <> Total Then
        MsgBox ("El detalle de pago no coincide con el total")
        Exit Sub
    End If
    Respuesta = MsgBox("¿Confirma la entrega de mercadería?", vbYesNo, "")
    If Respuesta = vbNo Then Exit Sub
    'Confirmo comprobante
    cn.Open
    Dim rs As ADODB.Recordset
    Set rs = cn.Execute("SELECT MAX(numero) + 1 AS NuevoP FROM VENTAS WHERE pedido=1")
    If rs!NuevoP > 0 Then
        NumeroComprobante = rs!NuevoP
    Else
        NumeroComprobante = 1
    End If
    rs.Close
    Set rs = Nothing
    cn.Execute ("AgregarVenta '" & Format(Date, "mm/dd/yyyy") & "','P',0," & NumeroComprobante & "," & Replace((Total / 1.21), ",", ".") & "," & Replace(((txtTotal / 1.21) * 0.21), ",", ".") & "," & Replace(Total, ",", ".") & "," & Val(lblIdCliente) & ",1")
    Set rs = cn.Execute("SELECT MAX(idventa) AS Nuevoid FROM VENTAS")
    NuevoID = rs!NuevoID
    With grDetalle
    For i = 0 To grDetalle.Rows - 1
        cn.Execute ("AgregaDetalleVenta " & Replace(Val(.TextMatrix(i, 6)), ",", ".") & "," & Replace(Val(.TextMatrix(i, 10)), ",", ".") & ", " & Replace(.TextMatrix(i, 0), ",", ".") & "," & Replace(.TextMatrix(i, 4), ",", ".") & "," & NuevoID & "," & .TextMatrix(i, 3) & ", " & Replace(.TextMatrix(i, 7), ",", ".") & ", " & Replace(.TextMatrix(i, 9), ",", "."))
        cn.Execute ("DescuentaStock " & .TextMatrix(i, 3) & "," & .TextMatrix(i, 0))
    Next i
    End With
    cn.Execute ("AgregarDetalleCaja " & lblCaja & ",'" & Format(Date, "mm/dd/yyyy") & "'," & Replace(txtEfectivo - txtVuelto, ",", ".") & "," & Replace(txtDebito, ",", ".") & "," & Replace(Credito, ",", ".") & ",1,'Factura " & txtNumero & "'")
    cn.Close
    grDetalle.Rows = 0
    Total = 0
    txtTotal = ""
    txtEfectivo = ""
    txtDebito = ""
    txtCredito = ""
    txtVuelto = ""
    lblIdCliente = 1
    lblCliente = "CONSUMIDOR FINAL"
    lblCategoria = "Consumidor Final"
    lblTipoDocumento = ""
    lblNumeroDocumento = ""

    txtBarras.SetFocus
    Exit Sub
End Sub

Private Sub cmdAgregar_Click()
    buscarTubosPara = "Vender"
    frmBuscarTubos.Show 1
End Sub

Private Sub cmdAgregarDevuelto_Click()
    buscarTubosPara = "Devolver"
    frmBuscarTubos.Show 1
End Sub

Private Sub cmdBuscar_Click()
    If lblCondicion.Caption = "FORMA DE PAGO" Then MsgBox ("Debe especificar la forma de Pago"): Exit Sub
    frmBuscaArticulos.Show 1
    If Len(txtBarras.Text) > 0 Then CargarDetalle
    txtCantidad.SetFocus
    txtCantidad.SelStart = 0
    txtCantidad.SelLength = Len(txtCantidad.Text)
End Sub

Private Sub cmdBuscarCliente_Click()
    EligiendoCliente = 1
    frmClientes.Show 1
    EligiendoCliente = 0
    grDetalle.Rows = 0
    CalcularTotales
    'txtCantidad.SetFocus
End Sub

Private Sub cmdCaja_Click()
    CerroCaja = 0
    frmCajaCierre.Show 1
    If CerroCaja = 1 Then Unload Me
End Sub

Private Sub cmdQuitar_Click()
    If grDetalleTubos.Rows = 0 Then Exit Sub
    Respuesta = MsgBox("¿Está seguro de quitar el Tubo?", vbYesNo, "Borrar")
    If Respuesta = vbNo Then Exit Sub
    'Total = Total - grDetalle.TextMatrix(grDetalle.Row, 2)
    'txtTotal = Total
    'txtEfectivo = Total
    Cantidad = grDetalleTubos.TextMatrix(grDetalleTubos.Row, 4)
    idArticulo = grDetalleTubos.TextMatrix(grDetalleTubos.Row, 1)
    Total = grDetalleTubos.TextMatrix(grDetalleTubos.Row, 6)
    If grDetalleTubos.Rows > 1 Then
        grDetalleTubos.RemoveItem (grDetalleTubos.Row)
    Else
        grDetalleTubos.Rows = 0
        'Total = 0
        'txtTotal = Total
        'txtEfectivo = Total
    End If
    
    'quito el articulo de la grilla de articulos
    For x = 0 To grDetalle.Rows - 1
        If grDetalle.TextMatrix(x, 3) = idArticulo And Val(grDetalle.TextMatrix(x, 2)) = Val(Total) And Val(grDetalle.TextMatrix(x, 0)) = Val(Cantidad) Then
            If grDetalle.Rows = 1 Then
                grDetalle.Rows = 0
            Else
                grDetalle.RemoveItem (x)
            End If
            CalcularTotales
            Exit Sub
        End If
    Next x
End Sub

Private Sub cmdQuitarDevuelto_Click()
    If grTubosDevueltos.Rows = 0 Then Exit Sub
    Respuesta = MsgBox("¿Está seguro de quitar el Tubo?", vbYesNo, "Borrar")
    If Respuesta = vbNo Then Exit Sub
    If grTubosDevueltos.Rows > 1 Then
        grTubosDevueltos.RemoveItem (grTubosDevueltos.Row)
    Else
        grTubosDevueltos.Rows = 0
        'Total = 0
        'txtTotal = Total
        'txtEfectivo = Total
    End If
End Sub

Private Sub cmdSalir_Click()
    Unload Me
End Sub



Private Sub Command1_Click()
MsgBox (Efectivo & " - " & Debito & " - " & Credito & " - " & Vuelto)
End Sub



Private Sub Command2_Click()

End Sub

Private Sub Command3_Click()

End Sub

Private Sub Form_Activate()
    If PidiendoPrecio = 0 Then ControlarCaja
End Sub

Private Sub Form_Load()
    'inicializo grilla de articulos
    grDetalle.Cols = 11
    grDetalle.ColWidth(0) = 600  'cantidad
    grDetalle.ColWidth(1) = 3600 'detalle
    grDetalle.ColWidth(2) = 2000 'precio total venta
    grDetalle.ColWidth(3) = 0 'id articulo
    grDetalle.ColWidth(4) = 0 'precio costo
    grDetalle.ColWidth(5) = 0 'impuesto
    grDetalle.ColWidth(6) = 0         'neto
    grDetalle.ColWidth(7) = 0         'iva
    grDetalle.ColWidth(8) = 0 'Tipo: gas/articulo
    grDetalle.ColWidth(9) = 0   'mitad de iva
    grDetalle.ColWidth(10) = 0   'mitad de neto

    'inicializo grilla de tubos a entregar
    grDetalleTubos.Cols = 7
    grDetalleTubos.ColWidth(0) = 0  'idTubo
    grDetalleTubos.ColWidth(1) = 0 'idArticulo
    grDetalleTubos.ColWidth(2) = 1500 'tubo
    grDetalleTubos.ColWidth(3) = 2500 'Gas
    grDetalleTubos.ColWidth(4) = 1000 'capacidad
    grDetalleTubos.ColWidth(5) = 1500 'unidad
    grDetalleTubos.ColWidth(6) = 1500 'total

    'inicializo grilla de tubos a devolver
    grTubosDevueltos.Cols = 6
    grTubosDevueltos.ColWidth(0) = 0  'idTubo
    grTubosDevueltos.ColWidth(1) = 0 'idArticulo
    grTubosDevueltos.ColWidth(2) = 1500 'tubo
    grTubosDevueltos.ColWidth(3) = 2500 'Gas
    grTubosDevueltos.ColWidth(4) = 1000 'capacidad
    grDetalleTubos.ColWidth(5) = 1500 'unidad
    'grDetalleTubos.ColWidth(6) = 1500 'total


    txtCantidad = 1
    lblCajero = Usuario
    lblIdCliente = 1
    lblCliente = "CONSUMIDOR FINAL"
    lblCategoria = "Consumidor Final"
    lblTipoDocumento = ""
    lblNumeroDocumento = ""
    optFactura.Value = True
    dateFechaFactura = Date
    dateFechaVencimiento = Date + 30
    cn.Open
    Dim rs As ADODB.Recordset
    Set rs = cn.Execute("SELECT (NumeroB + 1) as Numero, PorcentajeCtaCte from Parametros")
    lblLetra = "B"
    frmFacturador.lblPuesto = "0002"
    frmFacturador.lblNumero = Format(rs!numero, "00000000")
    PorcentajeCtaCte = rs!PorcentajeCtaCte
    lblSaldo = "Saldo: $ 0.00"
    'lblCondicion = "CUENTA CORRIENTE"
    'lblCondicion.ForeColor = vbBlue
    lblUltimoUsoFacturado = ""
    lblCantidadTubosEnCliente = ""
    cn.Close
    ControloScanner
    cmdBuscarCliente_Click
    

On Error GoTo impresora_apag
Procesar:

    HASAR1.Puerto = portfiscal
    HASAR1.Modelo = MODELO_PR4
    HASAR1.Comenzar
    HASAR1.PrecioBase = False
    HASAR1.TratarDeCancelarTodo
    HASAR1.ObtenerDatosDeInicializacion Cuit, razon, serie, fechainicio, Puesto, fechainicio, codiibb, Categoria
    PuestoFiscal = Puesto
    Exit Sub

impresora_apag:

    'If MsgBox("Error Impresora:" & Err.Description, vbRetryCancel, "Errores") = vbRetry Then
    '    Resume Procesar
    'End If
    
End Sub
Sub ControloScanner()
On Error GoTo NoScanner
    MSComm1.CommPort = portscan
    MSComm1.PortOpen = True
    Exit Sub
NoScanner:
    'MsgBox ("No se localizó ningún scanner")
    'HayScanner = "no"
    'Resume Next
End Sub
Sub ControlarCaja()
    cn.Open
    Dim rs As New ADODB.Recordset
    Set rs = cn.Execute("BuscarCajaAbierta " & idUsuario)
    If rs.EOF = True Then 'no existe ninguna caja abierta para este usuario
        Set rs = Nothing
        Set rs = cn.Execute("VerUltimaCaja " & idUsuario)
        frmCajaApertura.lblUsuario = Usuario
        If rs.EOF = True Then 'este usuario nunca tuvo una caja
            frmCajaApertura.txtEfectivo = 0
            frmCajaApertura.txtDebito = 0
            frmCajaApertura.txtCredito = 0
            frmCajaApertura.txtSaldoApertura = 0
        Else
            frmCajaApertura.txtEfectivo = Format(rs!EfectivoFinal, "0.00")
            frmCajaApertura.txtDebito = Format(rs!DebitoFinal, "0.00")
            frmCajaApertura.txtCredito = Format(rs!CreditoFinal, "0.00")
            frmCajaApertura.txtSaldoApertura = Format(rs!EfectivoFinal + rs!DebitoFinal + rs!CreditoFinal, "0.00")
        End If
        frmCajaApertura.Show 1
    Else
        lblCaja = rs!idCaja
    End If
    rs.Close
    Set rs = Nothing
    cn.Close
    If Val(lblCaja) = 0 Then Unload Me: Exit Sub
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If HayScanner <> "no" Then
        'MSComm1.PortOpen = False
    End If
    HASAR1.Finalizar
    Total = 0
End Sub

Private Sub grDetalle_KeyDown(KeyCode As Integer, Shift As Integer)
    If grDetalle.Rows = 0 Then Exit Sub
    If KeyCode = 46 Then
        Respuesta = MsgBox("¿Está seguro de borrar el artículo?", vbYesNo, "Borrar")
        If Respuesta = vbNo Then Exit Sub
        Total = Total - grDetalle.TextMatrix(grDetalle.Row, 2)
        txtTotal = Total
        txtEfectivo = Total
        If grDetalle.Rows > 1 Then
            grDetalle.RemoveItem (grDetalle.Row)
        Else
            grDetalle.Rows = 0
            Total = 0
            txtTotal = Total
            txtEfectivo = Total
        End If
    End If
    CalcularTotales
    ControlTeclas (KeyCode)
End Sub



Private Sub lblCondicion_Click()
    If grDetalle.Rows > 0 Then
        Respuesta = MsgBox("Si cambia la condición se quitarán los artículos seleccionados", vbOKCancel, "Atención!")
        If Respuesta = vbCancel Then Exit Sub
        grDetalle.Rows = 0
        CalcularTotales
    End If


    If lblCondicion = "CONTADO" Then
        lblCondicion = "CUENTA CORRIENTE"
        lblCondicion.ForeColor = vbBlue
    Else
        lblCondicion = "CONTADO"
        lblCondicion.ForeColor = vbGreen
    End If
End Sub

Private Sub MSComm1_OnComm()
    If MSComm1.CommEvent = comEvReceive And MSComm1.InBufferCount > 0 Then
        Buffer = Buffer & MSComm1.Input
        If Asc(Mid(Buffer, Len(Buffer), 1)) = 10 Then
            txtBarras = Mid(Buffer, 1, Len(Buffer) - 2)
            Buffer = ""
            CargarDetalle
        End If
    End If
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then txtBarras.SetFocus
    If InStr(1, "0123456789." & Chr(13) & Chr(8), Chr(KeyAscii)) = 0 Then
        KeyAscii = 0
    End If
End Sub

Private Sub txtBarras_GotFocus()
    txtBarras = ""
End Sub

Private Sub txtBarras_KeyDown(KeyCode As Integer, Shift As Integer)
    ControlTeclas (KeyCode)
End Sub

Private Sub txtBarras_KeyPress(KeyAscii As Integer)
    If txtBarras.Text = "" Then Exit Sub
    If InStr(1, "0123456789" & Chr(13), Chr(KeyAscii)) = 0 Then
        KeyAscii = 0
    End If
    If KeyAscii <> 13 Then Exit Sub
    CargarDetalle
End Sub
Sub CargarDetalle()
    If lblCondicion.Caption = "FORMA DE PAGO" Then MsgBox ("Debe especificar la forma de Pago"): Exit Sub
    If grDetalle.Rows = 13 Then MsgBox ("Se llegó al límite máximo de items"): Exit Sub
    If IsNumeric(txtCantidad) = False Then MsgBox ("La cantidad no es valida"): txtCantidad.SetFocus: Exit Sub
    cn.Open
    Dim rs As ADODB.Recordset
    Barra = txtBarras.Text
    Set rs = cn.Execute("SELECT idArticulo,Descripcion ,CodBar,Venta, VentaRevendedor, Costo, Impuesto, NoTomarPrecioCtaCte, ivamitad FROM Articulos where CodBar='" & Barra & "'")
    If rs.EOF = True Then
        encontro = "no"
       'busco quitando el ultimo digito
       Barra = Mid(RTrim(txtBarras), 1, Len(RTrim(txtBarras)) - 1)
       Set rs = cn.Execute("SELECT idArticulo,Descripcion ,CodBar,Venta, Costo, Impuesto, NoTomarPrecioCtaCte, ivamitad FROM Articulos where CodBar='" & Barra & "'")
       If rs.EOF = False Then
            encontro = "si"
       End If
    Else
        encontro = "si"
    End If

    If encontro = "si" Then

       PrecioDeVenta = 0
       If lblTipoPrecio = "Revendedor" Then
            PrecioDeVenta = Format(rs!VentaRevendedor, "0.0000")
       Else
            PrecioDeVenta = Format(rs!Venta, "0.0000")
       End If
       
      
       If rs!NoTomarPrecioCtaCte = 0 Then
            If lblCondicion = "CUENTA CORRIENTE" Then PrecioDeVenta = PrecioDeVenta + ((PrecioDeVenta * PorcentajeCtaCte) / 100)
       End If
        
       'grDetalle.TextMatrix(grDetalle.Rows - 1, 2) = Format(PrecioDeVenta, "0.0000")
       frmPedirPrecio.txtPrecio = Format(PrecioDeVenta, "0.0000")
       
       'If rs!Venta = 0 Or rs!Descripcion = "FOTOCOPIAS" Or rs!Descripcion = "LIBRERIA" Or rs!Descripcion = "VARIOS" Then
            frmPedirPrecio.lblDescripcion = rs!Descripcion
            PidiendoPrecio = 1
            frmPedirPrecio.Show 1
            If frmPedirPrecio.txtPrecio = "" Then rs.Close: cn.Close: Exit Sub

            
       grDetalle.Rows = grDetalle.Rows + 1
       'grDetalle.AddItem ("hola")
       grDetalle.TextMatrix(grDetalle.Rows - 1, 0) = txtCantidad
       grDetalle.TextMatrix(grDetalle.Rows - 1, 1) = rs!Descripcion
       PrecioDeVenta = Format(frmPedirPrecio.txtPrecio * txtCantidad, "0.00")
       grDetalle.TextMatrix(grDetalle.Rows - 1, 2) = PrecioDeVenta
       '     cn.Execute ("ActualizarPrecioVenta " & rs!idArticulo & "," & Replace(Val(frmPedirPrecio.txtPrecio), ",", "."))
       'End If
       grDetalle.TextMatrix(grDetalle.Rows - 1, 3) = rs!idArticulo
       'total costo
       grDetalle.TextMatrix(grDetalle.Rows - 1, 4) = Format(rs!Costo * txtCantidad, "0.00")
       'total impuesto
       grDetalle.TextMatrix(grDetalle.Rows - 1, 5) = Format(rs!Impuesto * txtCantidad, "0.00")
       If rs!IvaMitad = 0 Then
            'neto
            grDetalle.TextMatrix(grDetalle.Rows - 1, 6) = Format(((PrecioDeVenta) - (rs!Impuesto * txtCantidad)) / (1.21), "0.00")
            'neto a la mitad
            grDetalle.TextMatrix(grDetalle.Rows - 1, 10) = Format(0, "0.00")
            'total iva
            grDetalle.TextMatrix(grDetalle.Rows - 1, 7) = Format(grDetalle.TextMatrix(grDetalle.Rows - 1, 6) * (0.21), "0.00")
            'total iva mitad
            grDetalle.TextMatrix(grDetalle.Rows - 1, 9) = Format(0, "0.00")
       Else
            'neto a la mitad
            grDetalle.TextMatrix(grDetalle.Rows - 1, 10) = Format(((PrecioDeVenta) - (rs!Impuesto * txtCantidad)) / (1.105), "0.00")
            'neto al 21%
            grDetalle.TextMatrix(grDetalle.Rows - 1, 6) = Format(0, "0.00")
            'total iva
            grDetalle.TextMatrix(grDetalle.Rows - 1, 7) = Format(0, "0.00")
            'total iva mitad
            grDetalle.TextMatrix(grDetalle.Rows - 1, 9) = Format(grDetalle.TextMatrix(grDetalle.Rows - 1, 10) * (0.105), "0.00")
       End If
       If chkConIva.Value = 0 Then
            'neto
            grDetalle.TextMatrix(grDetalle.Rows - 1, 6) = PrecioDeVenta
            'total iva
            grDetalle.TextMatrix(grDetalle.Rows - 1, 7) = Format(0, "0.00")
            grDetalle.TextMatrix(grDetalle.Rows - 1, 9) = Format(0, "0.00")
       End If
        
    Else
       If HayScanner <> "no" Then
            'MSComm1.PortOpen = False
       End If
       MsgBox ("No se encontro el articulo")
       If HayScanner <> "no" Then
            'MSComm1.PortOpen = True
       End If
       txtBarras = "": txtBarras.SetFocus
    End If
    rs.Close
    Set rs = Nothing
    cn.Close
    txtBarras = ""
    txtCantidad = 1
    PidiendoPrecio = 0
    txtCantidad.SetFocus
    CalcularTotales
End Sub

Private Sub txtCantidad_GotFocus()
    txtCantidad.SelStart = 0
    txtCantidad.SelLength = Len(txtCantidad.Text)
End Sub

Private Sub txtCantidad_KeyDown(KeyCode As Integer, Shift As Integer)
    ControlTeclas (KeyCode)
End Sub

Private Sub txtCantidad_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then txtBarras.SetFocus
    If InStr(1, "0123456789," & Chr(13) & Chr(8), Chr(KeyAscii)) = 0 Then
        KeyAscii = 0
    End If
End Sub

Private Sub txtCredito_Change()
    Credito = Val(txtCredito)
End Sub

Private Sub txtCredito_KeyDown(KeyCode As Integer, Shift As Integer)
    ControlTeclas (KeyCode)
End Sub

Private Sub txtDebito_Change()
    Debito = Val(txtDebito)
End Sub

Private Sub txtDebito_GotFocus()
    'txtDebito = txtTotal
    txtDebito.SelStart = 0
    txtDebito.SelLength = Len(txtDebito.Text)
End Sub

Private Sub txtDebito_KeyDown(KeyCode As Integer, Shift As Integer)
    ControlTeclas (KeyCode)
End Sub

Private Sub txtDebito_KeyPress(KeyAscii As Integer)
    If InStr(1, "0123456789.", Chr(KeyAscii)) = 0 Then
        KeyAscii = 0
    End If
End Sub
Private Sub txtCredito_GotFocus()
    'txtCredito = txtTotal
    txtCredito.SelStart = 0
    txtCredito.SelLength = Len(txtCredito.Text)
End Sub

Private Sub txtCredito_KeyPress(KeyAscii As Integer)
    If InStr(1, "0123456789.", Chr(KeyAscii)) = 0 Then
        KeyAscii = 0
    End If
End Sub

Private Sub txtEfectivo_Change()
    If Val(txtEfectivo) + Debito + Credito > Total Then
        Efectivo = Val(txtEfectivo)
        Vuelto = Efectivo + Debito + Credito - Total
        txtVuelto = Format(Vuelto, "0.00")
    Else
        txtVuelto = "0.00"
        Vuelto = 0
    End If
    Efectivo = Val(txtEfectivo)
    End Sub

Private Sub txtEfectivo_GotFocus()
    'txtEfectivo = Total
    txtEfectivo.SelStart = 0
    txtEfectivo.SelLength = Len(txtEfectivo.Text)
End Sub

Private Sub txtEfectivo_KeyDown(KeyCode As Integer, Shift As Integer)
    ControlTeclas (KeyCode)
End Sub

Private Sub txtEfectivo_KeyPress(KeyAscii As Integer)
    If InStr(1, "0123456789.", Chr(KeyAscii)) = 0 Then
        KeyAscii = 0
    End If
End Sub

Sub ControlTeclas(Tecla As Integer)
    Select Case Tecla
        Case 112
            cmdBuscar_Click
        Case 113
            cmdAceptar_Click
        Case 114
            cmdCaja_Click
        Case 115
            cmdBuscarCliente_Click
        Case 123
            ComprobanteFiscal = "NO"
            cmdAceptar_Click
            ComprobanteFiscal = "SI"
        Case 27
            Respuesta = MsgBox("¿Cancela el comprobante?", vbYesNo, "")
            If Respuesta = vbYes Then cmdSalir_Click
    End Select
End Sub

Private Sub txtTotal_KeyDown(KeyCode As Integer, Shift As Integer)
    ControlTeclas (KeyCode)
End Sub

Private Sub txtVuelto_KeyDown(KeyCode As Integer, Shift As Integer)
    ControlTeclas (KeyCode)
End Sub

Private Sub PasarTotalesAtxt()
    txtTotal = Total
    txtEfectivo = Efectivo
    txtDebito = Debito
    txtCredito = Credito
    txtVuelto = Vuelto
End Sub

Private Sub CalcularTotales()
    TotalAcumulado = 0
    NetoAcumulado = 0
    NetoAcumuladoMitad = 0
    ImpuestoAcumulado = 0
    IvaAcumulado = 0
    IvaAcumuladoMitad = 0
    For x = 0 To grDetalle.Rows - 1
        TotalAcumulado = TotalAcumulado + grDetalle.TextMatrix(x, 2)
        NetoAcumulado = NetoAcumulado + grDetalle.TextMatrix(x, 6)
        ImpuestoAcumulado = ImpuestoAcumulado + grDetalle.TextMatrix(x, 5)
        IvaAcumulado = IvaAcumulado + grDetalle.TextMatrix(x, 7)
        IvaAcumuladoMitad = IvaAcumuladoMitad + grDetalle.TextMatrix(x, 9)
        NetoAcumuladoMitad = NetoAcumuladoMitad + grDetalle.TextMatrix(x, 10)
        
    Next x
    Efectivo = TotalAcumulado
    Debito = 0
    Credito = 0
    
    txtTotal = Format(TotalAcumulado, "0.00")
    txtNeto = Format(NetoAcumulado + NetoAcumuladoMitad, "0.00")
    txtImpuestos = Format(ImpuestoAcumulado, "0.00")
    txtSubtotal = Format(NetoAcumulado + NetoAcumuladoMitad + ImpuestoAcumulado, "0.00")
    txtIva = Format(IvaAcumulado, "0.00")
    texIvaMitad = Format(IvaAcumuladoMitad, "0.00")
    
    txtEfectivo = Efectivo
    txtDebito = Debito
    txtCredito = Credito
    txtVuelto = Vuelto
End Sub
