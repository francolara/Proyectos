VERSION 5.00
Object = "{6A24B331-7634-11D3-A5B0-0050044A7E1A}#1.5#0"; "DXDBGrid.dll"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Object = "{F41D1D30-7878-4923-8CB3-6CCACDC9C9DE}#1.0#0"; "CATControls.ocx"
Begin VB.Form frmMantClientes_Consul 
   Appearance      =   0  'Flat
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Mantenimiento de Clientes"
   ClientHeight    =   9090
   ClientLeft      =   1500
   ClientTop       =   1245
   ClientWidth     =   12630
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   9090
   ScaleWidth      =   12630
   ShowInTaskbar   =   0   'False
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   660
      Left            =   0
      TabIndex        =   22
      Top             =   0
      Width           =   12630
      _ExtentX        =   22278
      _ExtentY        =   1164
      ButtonWidth     =   1614
      ButtonHeight    =   1005
      Appearance      =   1
      ImageList       =   "imgDocVentas"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   9
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
            Caption         =   "      Nuevo      "
            Object.ToolTipText     =   "Nuevo"
            ImageIndex      =   1
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
            Caption         =   "Grabar"
            Object.ToolTipText     =   "Grabar"
            ImageIndex      =   4
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
            Caption         =   "Modificar"
            Object.ToolTipText     =   "Modificar"
            ImageIndex      =   10
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
            Caption         =   "Cancelar"
            Object.ToolTipText     =   "Cancelar"
            ImageIndex      =   9
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
            Caption         =   "Eliminar"
            ImageIndex      =   6
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
            Caption         =   "Imprimir"
            Object.ToolTipText     =   "Imprimir"
            ImageIndex      =   11
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "     Lista     "
            Object.ToolTipText     =   "Lista"
            ImageIndex      =   12
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Excel"
            ImageIndex      =   13
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Salir"
            Object.ToolTipText     =   "Salir"
            ImageIndex      =   2
         EndProperty
      EndProperty
      BorderStyle     =   1
   End
   Begin VB.Frame fraGeneral 
      Appearance      =   0  'Flat
      ForeColor       =   &H80000008&
      Height          =   8340
      Left            =   30
      TabIndex        =   23
      Top             =   675
      Width           =   12480
      Begin TabDlg.SSTab SSTab2 
         Height          =   7800
         Left            =   270
         TabIndex        =   24
         Top             =   360
         Width           =   11925
         _ExtentX        =   21034
         _ExtentY        =   13758
         _Version        =   393216
         Tabs            =   7
         TabsPerRow      =   7
         TabHeight       =   706
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         TabCaption(0)   =   "Datos Generales"
         TabPicture(0)   =   "frmMantClientes_Consul.frx":0000
         Tab(0).ControlEnabled=   -1  'True
         Tab(0).Control(0)=   "fraDatos"
         Tab(0).Control(0).Enabled=   0   'False
         Tab(0).ControlCount=   1
         TabCaption(1)   =   "Contactos"
         TabPicture(1)   =   "frmMantClientes_Consul.frx":001C
         Tab(1).ControlEnabled=   0   'False
         Tab(1).Control(0)=   "Frame3"
         Tab(1).ControlCount=   1
         TabCaption(2)   =   "Tiendas"
         TabPicture(2)   =   "frmMantClientes_Consul.frx":0038
         Tab(2).ControlEnabled=   0   'False
         Tab(2).Control(0)=   "Frame2"
         Tab(2).ControlCount=   1
         TabCaption(3)   =   "Línea de Crédito"
         TabPicture(3)   =   "frmMantClientes_Consul.frx":0054
         Tab(3).ControlEnabled=   0   'False
         Tab(3).Control(0)=   "Frame5"
         Tab(3).Control(1)=   "FraRegistro"
         Tab(3).Control(2)=   "TxtCodMoneda"
         Tab(3).Control(3)=   "TxtGlsMoneda"
         Tab(3).Control(4)=   "TxtGlsEstado"
         Tab(3).Control(5)=   "Label23"
         Tab(3).Control(6)=   "Label22"
         Tab(3).ControlCount=   7
         TabCaption(4)   =   "Forma de Pago"
         TabPicture(4)   =   "frmMantClientes_Consul.frx":0070
         Tab(4).ControlEnabled=   0   'False
         Tab(4).Control(0)=   "Frame4"
         Tab(4).ControlCount=   1
         TabCaption(5)   =   "Productos"
         TabPicture(5)   =   "frmMantClientes_Consul.frx":008C
         Tab(5).ControlEnabled=   0   'False
         Tab(5).Control(0)=   "Frame6"
         Tab(5).ControlCount=   1
         TabCaption(6)   =   "Observaciones"
         TabPicture(6)   =   "frmMantClientes_Consul.frx":00A8
         Tab(6).ControlEnabled=   0   'False
         Tab(6).Control(0)=   "FraObs"
         Tab(6).ControlCount=   1
         Begin VB.Frame Frame6 
            Appearance      =   0  'Flat
            ForeColor       =   &H00000000&
            Height          =   6510
            Left            =   -74595
            TabIndex        =   92
            Top             =   765
            Width           =   11010
            Begin DXDBGRIDLibCtl.dxDBGrid gProductos 
               Height          =   6120
               Left            =   135
               OleObjectBlob   =   "frmMantClientes_Consul.frx":00C4
               TabIndex        =   93
               Top             =   225
               Width           =   10740
            End
         End
         Begin VB.Frame FraObs 
            Appearance      =   0  'Flat
            ForeColor       =   &H00000000&
            Height          =   5025
            Left            =   -74640
            TabIndex        =   90
            Top             =   1440
            Width           =   11115
            Begin VB.TextBox txtGls_Obs 
               Height          =   4530
               Left            =   180
               MultiLine       =   -1  'True
               ScrollBars      =   2  'Vertical
               TabIndex        =   91
               Tag             =   "TGlsObservacion"
               Top             =   255
               Width           =   10815
            End
         End
         Begin VB.Frame Frame5 
            Appearance      =   0  'Flat
            ForeColor       =   &H80000008&
            Height          =   1500
            Left            =   -71220
            TabIndex        =   75
            Top             =   6525
            Width           =   4605
            Begin VB.TextBox TxtSaldo 
               Alignment       =   1  'Right Justify
               Appearance      =   0  'Flat
               BackColor       =   &H00C0E0FF&
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   330
               Left            =   2070
               Locked          =   -1  'True
               TabIndex        =   76
               Text            =   "0"
               Top             =   1080
               Width           =   1680
            End
            Begin CATControls.CATTextBox TxtLineaAprobada 
               Height          =   315
               Left            =   2070
               TabIndex        =   77
               Top             =   180
               Width           =   1680
               _ExtentX        =   2963
               _ExtentY        =   556
               BackColor       =   12640511
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Alignment       =   1
               FontName        =   "Arial"
               FontSize        =   8.25
               ForeColor       =   -2147483640
               Locked          =   -1  'True
               Container       =   "frmMantClientes_Consul.frx":268B
               Text            =   "0.00"
               Decimales       =   2
               Estilo          =   4
               Vacio           =   -1  'True
               EnterTab        =   -1  'True
            End
            Begin CATControls.CATTextBox TxtDeuda 
               Height          =   315
               Left            =   2070
               TabIndex        =   78
               Top             =   585
               Width           =   1680
               _ExtentX        =   2963
               _ExtentY        =   556
               BackColor       =   12640511
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Alignment       =   1
               FontName        =   "Arial"
               FontSize        =   8.25
               ForeColor       =   -2147483640
               Locked          =   -1  'True
               Container       =   "frmMantClientes_Consul.frx":26A7
               Text            =   "0.00"
               Decimales       =   2
               Estilo          =   4
               Vacio           =   -1  'True
               EnterTab        =   -1  'True
            End
            Begin VB.Label Label21 
               Appearance      =   0  'Flat
               Caption         =   "__________________________"
               ForeColor       =   &H80000007&
               Height          =   240
               Left            =   1890
               TabIndex        =   82
               Top             =   810
               Width           =   2145
            End
            Begin VB.Label Label20 
               Caption         =   "Línea Aprobada"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   195
               Left            =   135
               TabIndex        =   81
               Top             =   225
               Width           =   1455
            End
            Begin VB.Label Label18 
               Caption         =   "Deuda Actual"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   195
               Left            =   135
               TabIndex        =   80
               Top             =   630
               Width           =   1635
            End
            Begin VB.Label Label11 
               Caption         =   "Saldo Disponible"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   195
               Left            =   135
               TabIndex        =   79
               Top             =   1110
               Width           =   1635
            End
         End
         Begin VB.Frame FraRegistro 
            Appearance      =   0  'Flat
            ForeColor       =   &H80000008&
            Height          =   4740
            Left            =   -74820
            TabIndex        =   72
            Top             =   1740
            Width           =   11535
            Begin TabDlg.SSTab SSTab1 
               Height          =   4470
               Left            =   90
               TabIndex        =   73
               Top             =   180
               Width           =   11355
               _ExtentX        =   20029
               _ExtentY        =   7885
               _Version        =   393216
               TabHeight       =   520
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               TabCaption(0)   =   "Documentos por Cobrar"
               TabPicture(0)   =   "frmMantClientes_Consul.frx":26C3
               Tab(0).ControlEnabled=   -1  'True
               Tab(0).Control(0)=   "GDocumentos"
               Tab(0).Control(0).Enabled=   0   'False
               Tab(0).ControlCount=   1
               TabCaption(1)   =   "Guías por Facturar"
               TabPicture(1)   =   "frmMantClientes_Consul.frx":26DF
               Tab(1).ControlEnabled=   0   'False
               Tab(1).Control(0)=   "GGuiasNF"
               Tab(1).ControlCount=   1
               TabCaption(2)   =   "Documentos Generados"
               TabPicture(2)   =   "frmMantClientes_Consul.frx":26FB
               Tab(2).ControlEnabled=   0   'False
               Tab(2).Control(0)=   "GDocumentosGen"
               Tab(2).ControlCount=   1
               Begin DXDBGRIDLibCtl.dxDBGrid GDocumentosGen 
                  Height          =   3960
                  Left            =   -74910
                  OleObjectBlob   =   "frmMantClientes_Consul.frx":2717
                  TabIndex        =   74
                  Top             =   405
                  Width           =   11115
               End
               Begin DXDBGRIDLibCtl.dxDBGrid GDocumentos 
                  Height          =   3960
                  Left            =   90
                  OleObjectBlob   =   "frmMantClientes_Consul.frx":50BF
                  TabIndex        =   83
                  Top             =   405
                  Width           =   11115
               End
               Begin DXDBGRIDLibCtl.dxDBGrid GGuiasNF 
                  Height          =   3960
                  Left            =   -74910
                  OleObjectBlob   =   "frmMantClientes_Consul.frx":7A67
                  TabIndex        =   84
                  Top             =   405
                  Width           =   11115
               End
            End
         End
         Begin VB.Frame fraDatos 
            Appearance      =   0  'Flat
            Enabled         =   0   'False
            ForeColor       =   &H80000008&
            Height          =   6975
            Left            =   405
            TabIndex        =   30
            Top             =   960
            Width           =   11115
            Begin VB.CheckBox ChkEspecial 
               Appearance      =   0  'Flat
               Caption         =   "Cliente Especial"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H80000008&
               Height          =   315
               Left            =   8655
               TabIndex        =   71
               Tag             =   "NIndEspecial"
               Top             =   6210
               Visible         =   0   'False
               Width           =   1995
            End
            Begin VB.CommandButton cmbAyudaPersona 
               Height          =   315
               Left            =   10485
               Picture         =   "frmMantClientes_Consul.frx":A40F
               Style           =   1  'Graphical
               TabIndex        =   35
               Top             =   345
               Width           =   390
            End
            Begin VB.CommandButton cmbAyudaVendedorCampo 
               Height          =   315
               Left            =   10485
               Picture         =   "frmMantClientes_Consul.frx":A799
               Style           =   1  'Graphical
               TabIndex        =   34
               Top             =   3330
               Width           =   390
            End
            Begin VB.CommandButton cmbAyudaEmpTrans 
               Height          =   315
               Left            =   10485
               Picture         =   "frmMantClientes_Consul.frx":AB23
               Style           =   1  'Graphical
               TabIndex        =   33
               Top             =   3765
               Width           =   390
            End
            Begin VB.CommandButton cmbAyudaLista 
               Height          =   315
               Left            =   10485
               Picture         =   "frmMantClientes_Consul.frx":AEAD
               Style           =   1  'Graphical
               TabIndex        =   32
               Top             =   4155
               Width           =   390
            End
            Begin VB.CheckBox chkAgenteRetencion 
               Appearance      =   0  'Flat
               Caption         =   "Agente de Retención"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H80000008&
               Height          =   315
               Left            =   8655
               TabIndex        =   12
               Tag             =   "NindAgenteRetencion"
               Top             =   4980
               Width           =   1860
            End
            Begin VB.CommandButton cmbAyudaCanal 
               Height          =   315
               Left            =   10485
               Picture         =   "frmMantClientes_Consul.frx":B237
               Style           =   1  'Graphical
               TabIndex        =   31
               Top             =   4545
               Width           =   390
            End
            Begin VB.CheckBox chktercros 
               Appearance      =   0  'Flat
               Caption         =   "Ventas a Terceros"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H80000008&
               Height          =   315
               Left            =   8655
               TabIndex        =   13
               Tag             =   "Nindventasterceros"
               Top             =   5265
               Width           =   1860
            End
            Begin VB.CheckBox chkComision 
               Appearance      =   0  'Flat
               Caption         =   "Comisión"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H80000008&
               Height          =   315
               Left            =   8655
               TabIndex        =   14
               Tag             =   "NindComision"
               Top             =   5580
               Width           =   1815
            End
            Begin VB.CheckBox ChkAnula 
               Appearance      =   0  'Flat
               Caption         =   "Anula"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H80000008&
               Height          =   315
               Left            =   8655
               TabIndex        =   15
               Tag             =   "NIndAnula"
               Top             =   5895
               Visible         =   0   'False
               Width           =   915
            End
            Begin CATControls.CATTextBox txtCod_Persona 
               Height          =   315
               Left            =   2610
               TabIndex        =   2
               Tag             =   "TidCliente"
               Top             =   345
               Width           =   915
               _ExtentX        =   1614
               _ExtentY        =   556
               BackColor       =   16777215
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               FontName        =   "Arial"
               FontSize        =   8.25
               ForeColor       =   -2147483640
               MaxLength       =   8
               Container       =   "frmMantClientes_Consul.frx":B5C1
               Estilo          =   1
               EnterTab        =   -1  'True
            End
            Begin CATControls.CATTextBox txtGls_Persona 
               Height          =   315
               Left            =   3570
               TabIndex        =   36
               Top             =   345
               Width           =   6870
               _ExtentX        =   12118
               _ExtentY        =   556
               BackColor       =   16777152
               Enabled         =   0   'False
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               FontName        =   "Arial"
               FontSize        =   8.25
               ForeColor       =   -2147483640
               Container       =   "frmMantClientes_Consul.frx":B5DD
            End
            Begin CATControls.CATTextBox txtGls_Direccion 
               Height          =   315
               Left            =   2610
               TabIndex        =   37
               Top             =   2610
               Width           =   7830
               _ExtentX        =   13811
               _ExtentY        =   556
               BackColor       =   16777152
               Enabled         =   0   'False
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               FontName        =   "Arial"
               FontSize        =   8.25
               ForeColor       =   -2147483640
               Container       =   "frmMantClientes_Consul.frx":B5F9
               Estilo          =   1
               Vacio           =   -1  'True
               EnterTab        =   -1  'True
            End
            Begin CATControls.CATTextBox txt_RUC 
               Height          =   315
               Left            =   2610
               TabIndex        =   38
               Top             =   2985
               Width           =   2115
               _ExtentX        =   3731
               _ExtentY        =   556
               BackColor       =   16777152
               Enabled         =   0   'False
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               FontName        =   "Arial"
               FontSize        =   8.25
               ForeColor       =   -2147483640
               Container       =   "frmMantClientes_Consul.frx":B615
               Estilo          =   1
               Vacio           =   -1  'True
               EnterTab        =   -1  'True
            End
            Begin CATControls.CATTextBox txtCod_Pais 
               Height          =   315
               Left            =   2610
               TabIndex        =   39
               Top             =   1095
               Width           =   915
               _ExtentX        =   1614
               _ExtentY        =   556
               BackColor       =   16777152
               Enabled         =   0   'False
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               FontName        =   "Arial"
               FontSize        =   8.25
               ForeColor       =   -2147483640
               Container       =   "frmMantClientes_Consul.frx":B631
               Estilo          =   1
               Vacio           =   -1  'True
               EnterTab        =   -1  'True
            End
            Begin CATControls.CATTextBox txtGls_Pais 
               Height          =   315
               Left            =   3570
               TabIndex        =   40
               Top             =   1095
               Width           =   6870
               _ExtentX        =   12118
               _ExtentY        =   556
               BackColor       =   16777152
               Enabled         =   0   'False
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               FontName        =   "Arial"
               FontSize        =   8.25
               ForeColor       =   -2147483640
               Container       =   "frmMantClientes_Consul.frx":B64D
               Vacio           =   -1  'True
            End
            Begin CATControls.CATTextBox txtCod_Depa 
               Height          =   315
               Left            =   2610
               TabIndex        =   41
               Top             =   1470
               Width           =   915
               _ExtentX        =   1614
               _ExtentY        =   556
               BackColor       =   16777152
               Enabled         =   0   'False
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               FontName        =   "Arial"
               FontSize        =   8.25
               ForeColor       =   -2147483640
               Container       =   "frmMantClientes_Consul.frx":B669
               Estilo          =   1
               Vacio           =   -1  'True
               EnterTab        =   -1  'True
            End
            Begin CATControls.CATTextBox txtGls_Depa 
               Height          =   315
               Left            =   3570
               TabIndex        =   42
               Top             =   1470
               Width           =   6870
               _ExtentX        =   12118
               _ExtentY        =   556
               BackColor       =   16777152
               Enabled         =   0   'False
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               FontName        =   "Arial"
               FontSize        =   8.25
               ForeColor       =   -2147483640
               Container       =   "frmMantClientes_Consul.frx":B685
               Vacio           =   -1  'True
            End
            Begin CATControls.CATTextBox txtCod_Prov 
               Height          =   315
               Left            =   2610
               TabIndex        =   43
               Top             =   1845
               Width           =   915
               _ExtentX        =   1614
               _ExtentY        =   556
               BackColor       =   16777152
               Enabled         =   0   'False
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               FontName        =   "Arial"
               FontSize        =   8.25
               ForeColor       =   -2147483640
               Container       =   "frmMantClientes_Consul.frx":B6A1
               Estilo          =   1
               Vacio           =   -1  'True
               EnterTab        =   -1  'True
            End
            Begin CATControls.CATTextBox txtGls_Prov 
               Height          =   315
               Left            =   3570
               TabIndex        =   44
               Top             =   1845
               Width           =   6870
               _ExtentX        =   12118
               _ExtentY        =   556
               BackColor       =   16777152
               Enabled         =   0   'False
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               FontName        =   "Arial"
               FontSize        =   8.25
               ForeColor       =   -2147483640
               Container       =   "frmMantClientes_Consul.frx":B6BD
               Vacio           =   -1  'True
            End
            Begin CATControls.CATTextBox txtCod_Distrito 
               Height          =   315
               Left            =   2610
               TabIndex        =   45
               Top             =   2220
               Width           =   915
               _ExtentX        =   1614
               _ExtentY        =   556
               BackColor       =   16777152
               Enabled         =   0   'False
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               FontName        =   "Arial"
               FontSize        =   8.25
               ForeColor       =   -2147483640
               Container       =   "frmMantClientes_Consul.frx":B6D9
               Estilo          =   1
               Vacio           =   -1  'True
               EnterTab        =   -1  'True
            End
            Begin CATControls.CATTextBox txtGls_Distrito 
               Height          =   315
               Left            =   3570
               TabIndex        =   46
               Top             =   2220
               Width           =   6870
               _ExtentX        =   12118
               _ExtentY        =   556
               BackColor       =   16777152
               Enabled         =   0   'False
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               FontName        =   "Arial"
               FontSize        =   8.25
               ForeColor       =   -2147483640
               Container       =   "frmMantClientes_Consul.frx":B6F5
               Vacio           =   -1  'True
            End
            Begin CATControls.CATTextBox txtCod_TipoPersona 
               Height          =   315
               Left            =   2610
               TabIndex        =   47
               Top             =   720
               Width           =   915
               _ExtentX        =   1614
               _ExtentY        =   556
               BackColor       =   16777152
               Enabled         =   0   'False
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               FontName        =   "Arial"
               FontSize        =   8.25
               ForeColor       =   -2147483640
               Container       =   "frmMantClientes_Consul.frx":B711
               Estilo          =   1
               Vacio           =   -1  'True
               EnterTab        =   -1  'True
            End
            Begin CATControls.CATTextBox txtGls_TipoPersona 
               Height          =   315
               Left            =   3570
               TabIndex        =   48
               Top             =   720
               Width           =   6870
               _ExtentX        =   12118
               _ExtentY        =   556
               BackColor       =   16777152
               Enabled         =   0   'False
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               FontName        =   "Arial"
               FontSize        =   8.25
               ForeColor       =   -2147483640
               Container       =   "frmMantClientes_Consul.frx":B72D
               Vacio           =   -1  'True
            End
            Begin CATControls.CATTextBox txtCod_VendedorCampo 
               Height          =   315
               Left            =   2610
               TabIndex        =   4
               Tag             =   "TidVendedorCampo"
               Top             =   3360
               Width           =   915
               _ExtentX        =   1614
               _ExtentY        =   556
               BackColor       =   16777215
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               FontName        =   "Arial"
               FontSize        =   8.25
               ForeColor       =   -2147483640
               MaxLength       =   8
               Container       =   "frmMantClientes_Consul.frx":B749
               Estilo          =   1
               Vacio           =   -1  'True
               EnterTab        =   -1  'True
            End
            Begin CATControls.CATTextBox txtGls_VendedorCampo 
               Height          =   315
               Left            =   3570
               TabIndex        =   49
               Top             =   3360
               Width           =   6870
               _ExtentX        =   12118
               _ExtentY        =   556
               BackColor       =   16777152
               Enabled         =   0   'False
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               FontName        =   "Arial"
               FontSize        =   8.25
               ForeColor       =   -2147483640
               Container       =   "frmMantClientes_Consul.frx":B765
               Vacio           =   -1  'True
            End
            Begin CATControls.CATTextBox txtCod_EmpTrans 
               Height          =   315
               Left            =   2610
               TabIndex        =   5
               Tag             =   "TidEmpTrans"
               Top             =   3765
               Width           =   915
               _ExtentX        =   1614
               _ExtentY        =   556
               BackColor       =   16777215
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               FontName        =   "Arial"
               FontSize        =   8.25
               ForeColor       =   -2147483640
               MaxLength       =   8
               Container       =   "frmMantClientes_Consul.frx":B781
               Estilo          =   1
               Vacio           =   -1  'True
               EnterTab        =   -1  'True
            End
            Begin CATControls.CATTextBox txtGls_EmpTrans 
               Height          =   315
               Left            =   3570
               TabIndex        =   50
               Top             =   3765
               Width           =   6870
               _ExtentX        =   12118
               _ExtentY        =   556
               BackColor       =   16777152
               Enabled         =   0   'False
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               FontName        =   "Arial"
               FontSize        =   8.25
               ForeColor       =   -2147483640
               Container       =   "frmMantClientes_Consul.frx":B79D
               Vacio           =   -1  'True
            End
            Begin CATControls.CATTextBox txtCod_Lista 
               Height          =   315
               Left            =   2610
               TabIndex        =   6
               Tag             =   "TidLista"
               Top             =   4170
               Width           =   915
               _ExtentX        =   1614
               _ExtentY        =   556
               BackColor       =   16777215
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               FontName        =   "Arial"
               FontSize        =   8.25
               ForeColor       =   -2147483640
               MaxLength       =   8
               Container       =   "frmMantClientes_Consul.frx":B7B9
               Estilo          =   1
               Vacio           =   -1  'True
               EnterTab        =   -1  'True
            End
            Begin CATControls.CATTextBox txtGls_Lista 
               Height          =   315
               Left            =   3570
               TabIndex        =   51
               Top             =   4170
               Width           =   6870
               _ExtentX        =   12118
               _ExtentY        =   556
               BackColor       =   16777152
               Enabled         =   0   'False
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               FontName        =   "Arial"
               FontSize        =   8.25
               ForeColor       =   -2147483640
               Container       =   "frmMantClientes_Consul.frx":B7D5
               Vacio           =   -1  'True
            End
            Begin CATControls.CATTextBox txtCod_ClienteInterno 
               Height          =   315
               Left            =   2610
               TabIndex        =   8
               Tag             =   "TidClienteInterno"
               Top             =   4950
               Width           =   2115
               _ExtentX        =   3731
               _ExtentY        =   556
               BackColor       =   16777215
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               FontName        =   "Arial"
               FontSize        =   8.25
               ForeColor       =   -2147483640
               MaxLength       =   30
               Container       =   "frmMantClientes_Consul.frx":B7F1
               Estilo          =   1
               Vacio           =   -1  'True
               EnterTab        =   -1  'True
            End
            Begin CATControls.CATTextBox txtVal_Dcto 
               Height          =   315
               Left            =   2610
               TabIndex        =   9
               Tag             =   "NVal_Dscto"
               Top             =   5310
               Width           =   1290
               _ExtentX        =   2275
               _ExtentY        =   556
               BackColor       =   16777215
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Alignment       =   1
               FontName        =   "Arial"
               FontSize        =   8.25
               ForeColor       =   -2147483640
               Container       =   "frmMantClientes_Consul.frx":B80D
               Text            =   "0.00"
               Decimales       =   2
               Estilo          =   4
               Vacio           =   -1  'True
               EnterTab        =   -1  'True
            End
            Begin CATControls.CATTextBox txtCod_Canal 
               Height          =   315
               Left            =   2610
               TabIndex        =   7
               Tag             =   "TidCanal"
               Top             =   4545
               Width           =   915
               _ExtentX        =   1614
               _ExtentY        =   556
               BackColor       =   16777215
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               FontName        =   "Arial"
               FontSize        =   8.25
               ForeColor       =   -2147483640
               MaxLength       =   8
               Container       =   "frmMantClientes_Consul.frx":B829
               Estilo          =   1
               Vacio           =   -1  'True
               EnterTab        =   -1  'True
            End
            Begin CATControls.CATTextBox txtGls_Canal 
               Height          =   315
               Left            =   3570
               TabIndex        =   52
               Top             =   4545
               Width           =   6870
               _ExtentX        =   12118
               _ExtentY        =   556
               BackColor       =   16777152
               Enabled         =   0   'False
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               FontName        =   "Arial"
               FontSize        =   8.25
               ForeColor       =   -2147483640
               Container       =   "frmMantClientes_Consul.frx":B845
               Vacio           =   -1  'True
            End
            Begin CATControls.CATTextBox txt_Telefono 
               Height          =   315
               Left            =   6840
               TabIndex        =   3
               Tag             =   "TTelefonos"
               Top             =   2970
               Width           =   3585
               _ExtentX        =   6324
               _ExtentY        =   556
               BackColor       =   16777152
               Enabled         =   0   'False
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               FontName        =   "Arial"
               FontSize        =   8.25
               ForeColor       =   -2147483640
               MaxLength       =   85
               Container       =   "frmMantClientes_Consul.frx":B861
               Estilo          =   1
               Vacio           =   -1  'True
               EnterTab        =   -1  'True
            End
            Begin CATControls.CATTextBox txt_Email 
               Height          =   315
               Left            =   2610
               TabIndex        =   10
               Tag             =   "Tmail"
               Top             =   5670
               Width           =   4980
               _ExtentX        =   8784
               _ExtentY        =   556
               BackColor       =   16777152
               Enabled         =   0   'False
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               FontName        =   "Arial"
               FontSize        =   8.25
               ForeColor       =   -2147483640
               MaxLength       =   80
               Container       =   "frmMantClientes_Consul.frx":B87D
               Estilo          =   1
               Vacio           =   -1  'True
               EnterTab        =   -1  'True
            End
            Begin CATControls.CATTextBox txtidProvCliente 
               Height          =   315
               Left            =   2610
               TabIndex        =   11
               Tag             =   "TidProvCliente"
               Top             =   6030
               Width           =   2115
               _ExtentX        =   3731
               _ExtentY        =   556
               BackColor       =   16777215
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               FontName        =   "Arial"
               FontSize        =   8.25
               ForeColor       =   -2147483640
               MaxLength       =   30
               Container       =   "frmMantClientes_Consul.frx":B899
               Estilo          =   1
               Vacio           =   -1  'True
               EnterTab        =   -1  'True
            End
            Begin VB.Label Label1 
               Appearance      =   0  'Flat
               AutoSize        =   -1  'True
               Caption         =   "Entidad"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H80000008&
               Height          =   210
               Left            =   555
               TabIndex        =   70
               Top             =   375
               Width           =   525
            End
            Begin VB.Label Label2 
               Appearance      =   0  'Flat
               AutoSize        =   -1  'True
               Caption         =   "Distrito"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H80000008&
               Height          =   210
               Left            =   555
               TabIndex        =   69
               Top             =   2265
               Width           =   495
            End
            Begin VB.Label Label5 
               Appearance      =   0  'Flat
               AutoSize        =   -1  'True
               Caption         =   "País"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H80000008&
               Height          =   210
               Left            =   555
               TabIndex        =   68
               Top             =   1125
               Width           =   300
            End
            Begin VB.Label Label7 
               Appearance      =   0  'Flat
               AutoSize        =   -1  'True
               Caption         =   "Departamento"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H80000008&
               Height          =   210
               Left            =   555
               TabIndex        =   67
               Top             =   1500
               Width           =   1005
            End
            Begin VB.Label Label8 
               Appearance      =   0  'Flat
               AutoSize        =   -1  'True
               Caption         =   "Provincia"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H80000008&
               Height          =   210
               Left            =   555
               TabIndex        =   66
               Top             =   1875
               Width           =   660
            End
            Begin VB.Label Label10 
               Appearance      =   0  'Flat
               AutoSize        =   -1  'True
               Caption         =   "Dirección"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H80000008&
               Height          =   210
               Left            =   555
               TabIndex        =   65
               Top             =   2640
               Width           =   675
            End
            Begin VB.Label Label14 
               Appearance      =   0  'Flat
               AutoSize        =   -1  'True
               Caption         =   "R.U.C."
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H80000008&
               Height          =   210
               Left            =   555
               TabIndex        =   64
               Top             =   3030
               Width           =   450
            End
            Begin VB.Label Label4 
               Appearance      =   0  'Flat
               AutoSize        =   -1  'True
               Caption         =   "Tipo de Entidad"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H80000008&
               Height          =   210
               Left            =   555
               TabIndex        =   63
               Top             =   750
               Width           =   1095
            End
            Begin VB.Label Label6 
               Appearance      =   0  'Flat
               AutoSize        =   -1  'True
               Caption         =   "Vendedor Campo"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H80000008&
               Height          =   210
               Left            =   555
               TabIndex        =   62
               Top             =   3390
               Width           =   1260
            End
            Begin VB.Label lbl_EmpTrans 
               Appearance      =   0  'Flat
               AutoSize        =   -1  'True
               Caption         =   "Empresa de Transporte"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H80000007&
               Height          =   210
               Left            =   555
               TabIndex        =   61
               Top             =   3795
               Width           =   1695
            End
            Begin VB.Label lbl_Lista 
               Appearance      =   0  'Flat
               AutoSize        =   -1  'True
               Caption         =   "Lista de Precios"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H80000007&
               Height          =   210
               Left            =   540
               TabIndex        =   60
               Top             =   4230
               Width           =   1155
            End
            Begin VB.Label Label12 
               Appearance      =   0  'Flat
               AutoSize        =   -1  'True
               Caption         =   "Código Interno"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H80000008&
               Height          =   210
               Left            =   555
               TabIndex        =   59
               Top             =   5040
               Width           =   1035
            End
            Begin VB.Label Label13 
               Appearance      =   0  'Flat
               AutoSize        =   -1  'True
               Caption         =   "% Descuento"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H80000008&
               Height          =   210
               Left            =   555
               TabIndex        =   58
               Top             =   5400
               Width           =   975
            End
            Begin VB.Label Label15 
               Appearance      =   0  'Flat
               AutoSize        =   -1  'True
               Caption         =   "* Este descuento será asignado para todos los productos"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H80000008&
               Height          =   210
               Left            =   3990
               TabIndex        =   57
               Top             =   5355
               Width           =   4170
            End
            Begin VB.Label Label16 
               Appearance      =   0  'Flat
               AutoSize        =   -1  'True
               Caption         =   "Canal"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H80000007&
               Height          =   210
               Left            =   555
               TabIndex        =   56
               Top             =   4650
               Width           =   405
            End
            Begin VB.Label Label17 
               Appearance      =   0  'Flat
               AutoSize        =   -1  'True
               Caption         =   "Teléfonos"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H80000008&
               Height          =   210
               Left            =   6030
               TabIndex        =   55
               Top             =   3015
               Width           =   720
            End
            Begin VB.Label Label19 
               Appearance      =   0  'Flat
               AutoSize        =   -1  'True
               Caption         =   "E Mail"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H80000008&
               Height          =   210
               Left            =   555
               TabIndex        =   54
               Top             =   5715
               Width           =   405
            End
            Begin VB.Label Label9 
               Appearance      =   0  'Flat
               AutoSize        =   -1  'True
               Caption         =   "Código Proveedor"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H80000008&
               Height          =   210
               Left            =   555
               TabIndex        =   53
               Top             =   6075
               Width           =   1290
            End
         End
         Begin VB.Frame Frame4 
            Appearance      =   0  'Flat
            ForeColor       =   &H00000000&
            Height          =   5565
            Left            =   -74640
            TabIndex        =   29
            Top             =   1620
            Width           =   11115
            Begin DXDBGRIDLibCtl.dxDBGrid gFormaPago 
               Height          =   5145
               Left            =   135
               OleObjectBlob   =   "frmMantClientes_Consul.frx":B8B5
               TabIndex        =   18
               Top             =   270
               Width           =   10875
            End
         End
         Begin VB.Frame Frame2 
            Appearance      =   0  'Flat
            ForeColor       =   &H00000000&
            Height          =   5565
            Left            =   -74640
            TabIndex        =   26
            Top             =   1620
            Width           =   11115
            Begin DXDBGRIDLibCtl.dxDBGrid gTiendas 
               Height          =   5100
               Left            =   180
               OleObjectBlob   =   "frmMantClientes_Consul.frx":ED7B
               TabIndex        =   17
               Top             =   315
               Width           =   10785
            End
            Begin CATControls.CATTextBox txtDes 
               Height          =   315
               Left            =   10800
               TabIndex        =   27
               Top             =   4860
               Visible         =   0   'False
               Width           =   375
               _ExtentX        =   661
               _ExtentY        =   556
               BackColor       =   16777215
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               FontName        =   "MS Sans Serif"
               FontSize        =   8.25
               ForeColor       =   -2147483640
               Container       =   "frmMantClientes_Consul.frx":13667
               Estilo          =   1
               EnterTab        =   -1  'True
            End
            Begin CATControls.CATTextBox txtCodigo 
               Height          =   315
               Left            =   10800
               TabIndex        =   28
               Top             =   4500
               Visible         =   0   'False
               Width           =   375
               _ExtentX        =   661
               _ExtentY        =   556
               BackColor       =   16777215
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               FontName        =   "MS Sans Serif"
               FontSize        =   8.25
               ForeColor       =   -2147483640
               MaxLength       =   8
               Container       =   "frmMantClientes_Consul.frx":13683
               Estilo          =   1
               EnterTab        =   -1  'True
            End
         End
         Begin VB.Frame Frame3 
            Appearance      =   0  'Flat
            ForeColor       =   &H00000000&
            Height          =   5565
            Left            =   -74610
            TabIndex        =   25
            Top             =   1620
            Width           =   11115
            Begin DXDBGRIDLibCtl.dxDBGrid gContactos 
               Height          =   5130
               Left            =   135
               OleObjectBlob   =   "frmMantClientes_Consul.frx":1369F
               TabIndex        =   16
               Top             =   270
               Width           =   10830
            End
         End
         Begin CATControls.CATTextBox TxtCodMoneda 
            Height          =   315
            Left            =   -72570
            TabIndex        =   85
            Top             =   1395
            Width           =   1020
            _ExtentX        =   1799
            _ExtentY        =   556
            BackColor       =   16777152
            Enabled         =   0   'False
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            FontName        =   "Arial"
            FontSize        =   8.25
            ForeColor       =   -2147483640
            Container       =   "frmMantClientes_Consul.frx":160E5
            Vacio           =   -1  'True
         End
         Begin CATControls.CATTextBox TxtGlsMoneda 
            Height          =   315
            Left            =   -71490
            TabIndex        =   86
            Top             =   1395
            Width           =   6735
            _ExtentX        =   11880
            _ExtentY        =   556
            BackColor       =   16777152
            Enabled         =   0   'False
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            FontName        =   "Arial"
            FontSize        =   8.25
            ForeColor       =   -2147483640
            Container       =   "frmMantClientes_Consul.frx":16101
            Vacio           =   -1  'True
         End
         Begin CATControls.CATTextBox TxtGlsEstado 
            Height          =   315
            Left            =   -72570
            TabIndex        =   88
            Top             =   1035
            Width           =   7815
            _ExtentX        =   13785
            _ExtentY        =   556
            BackColor       =   16777152
            Enabled         =   0   'False
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            FontName        =   "Arial"
            FontSize        =   8.25
            ForeColor       =   -2147483640
            Container       =   "frmMantClientes_Consul.frx":1611D
            Vacio           =   -1  'True
         End
         Begin VB.Label Label23 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            Caption         =   "Estado"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000007&
            Height          =   210
            Left            =   -73920
            TabIndex        =   89
            Top             =   1080
            Width           =   495
         End
         Begin VB.Label Label22 
            AutoSize        =   -1  'True
            Caption         =   "Moneda Línea"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Left            =   -73920
            TabIndex        =   87
            Top             =   1440
            Width           =   1005
         End
      End
   End
   Begin VB.Frame fraListado 
      Appearance      =   0  'Flat
      ForeColor       =   &H00000000&
      Height          =   8370
      Left            =   60
      TabIndex        =   19
      Top             =   660
      Width           =   12495
      Begin MSComctlLib.ImageList imgDocVentas 
         Left            =   2280
         Top             =   3480
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   -2147483643
         ImageWidth      =   16
         ImageHeight     =   16
         MaskColor       =   12632256
         _Version        =   393216
         BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
            NumListImages   =   16
            BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMantClientes_Consul.frx":16139
               Key             =   ""
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMantClientes_Consul.frx":164D3
               Key             =   ""
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMantClientes_Consul.frx":16925
               Key             =   ""
            EndProperty
            BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMantClientes_Consul.frx":16CBF
               Key             =   ""
            EndProperty
            BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMantClientes_Consul.frx":17059
               Key             =   ""
            EndProperty
            BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMantClientes_Consul.frx":173F3
               Key             =   ""
            EndProperty
            BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMantClientes_Consul.frx":1778D
               Key             =   ""
            EndProperty
            BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMantClientes_Consul.frx":17B27
               Key             =   ""
            EndProperty
            BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMantClientes_Consul.frx":17EC1
               Key             =   ""
            EndProperty
            BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMantClientes_Consul.frx":1825B
               Key             =   ""
            EndProperty
            BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMantClientes_Consul.frx":185F5
               Key             =   ""
            EndProperty
            BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMantClientes_Consul.frx":192B7
               Key             =   ""
            EndProperty
            BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMantClientes_Consul.frx":19651
               Key             =   ""
            EndProperty
            BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMantClientes_Consul.frx":19AA3
               Key             =   ""
            EndProperty
            BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMantClientes_Consul.frx":19E3D
               Key             =   ""
            EndProperty
            BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMantClientes_Consul.frx":1A84F
               Key             =   ""
            EndProperty
         EndProperty
      End
      Begin VB.Frame Frame1 
         Appearance      =   0  'Flat
         ForeColor       =   &H80000008&
         Height          =   660
         Left            =   165
         TabIndex        =   20
         Top             =   165
         Width           =   12180
         Begin CATControls.CATTextBox txt_TextoBuscar 
            Height          =   315
            Left            =   1035
            TabIndex        =   0
            Top             =   210
            Width           =   10995
            _ExtentX        =   19394
            _ExtentY        =   556
            BackColor       =   16777215
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            FontName        =   "Arial"
            FontSize        =   8.25
            ForeColor       =   -2147483640
            Container       =   "frmMantClientes_Consul.frx":1AF21
            Estilo          =   1
            Vacio           =   -1  'True
            EnterTab        =   -1  'True
         End
         Begin VB.Label Label3 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            Caption         =   "Búsqueda"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   210
            Left            =   165
            TabIndex        =   21
            Top             =   270
            Width           =   735
         End
      End
      Begin DXDBGRIDLibCtl.dxDBGrid gLista 
         Height          =   7335
         Left            =   165
         OleObjectBlob   =   "frmMantClientes_Consul.frx":1AF3D
         TabIndex        =   1
         Top             =   900
         Width           =   12195
      End
   End
End
Attribute VB_Name = "frmMantClientes_Consul"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim indBoton                            As Integer
Dim CTmpDocumentos                      As String
Dim CTmpGuiasNF                         As String
Dim CTmpDocumentosGen                   As String

Private Sub cmbAyudaCanal_Click()

    mostrarAyuda "CANALES", txtCod_Canal, txtGls_Canal
    
End Sub

Private Sub cmbAyudaEmpTrans_Click()

    mostrarAyuda "EMPTRANS", txtCod_EmpTrans, txtGls_EmpTrans
    If txtCod_EmpTrans.Text <> "" Then SendKeys "{tab}"
    
End Sub

Private Sub cmbAyudaLista_Click()

    mostrarAyuda "LISTAPRECIOS", txtCod_Lista, txtGls_Lista
    If txtCod_Lista.Text <> "" Then SendKeys "{tab}"
    
End Sub

Private Sub cmbAyudaPersona_Click()

    mostrarAyuda "PERSONACLIENTE", txtCod_Persona, txtGls_Persona
    If txtCod_Persona.Text <> "" Then mostrarDatosPersona
    
End Sub

Private Sub mostrarDatosPersona()
Dim rst As New ADODB.Recordset

    csql = "SELECT idPersona,GlsPersona," & _
            "tipoPersona , ruc, direccion, p.iddistrito, u.idDpto, u.idProv ,p.Telefonos, p.Mail,p.GlsContacto,p.idPais " & _
            "FROM personas p,ubigeo u " & _
            "WHERE p.iddistrito = u.iddistrito  and p.idPais = u.idPais and idpersona = '" & Trim(txtCod_Persona.Text) & "' "
               
    rst.Open csql, Cn, adOpenStatic, adLockReadOnly
    If Not rst.EOF Then
        txtCod_TipoPersona.Text = "" & rst.Fields("tipoPersona")
        txtCod_Pais.Text = "" & rst.Fields("idPais")
        txtCod_Depa.Text = "" & rst.Fields("idDpto")
        txtCod_Prov.Text = "" & rst.Fields("idProv")
        txtCod_Distrito.Text = "" & rst.Fields("iddistrito")
        txtGls_Direccion.Text = "" & rst.Fields("direccion")
        txt_RUC.Text = "" & rst.Fields("ruc")
        txt_Email.Text = "" & rst.Fields("mail")
        txt_Telefono.Text = "" & rst.Fields("Telefonos")
    Else
        txtCod_TipoPersona.Text = ""
        txtCod_Pais.Text = ""
        txtCod_Depa.Text = ""
        txtCod_Prov.Text = ""
        txtCod_Distrito.Text = ""
        txtGls_Direccion.Text = ""
        txt_RUC.Text = ""
        txt_Email.Text = ""
        txt_Telefono.Text = ""
    End If
    rst.Close: Set rst = Nothing

End Sub

Private Sub cmbAyudaVendedorCampo_Click()
    
    mostrarAyuda "VENDEDOR", txtCod_VendedorCampo, txtGls_VendedorCampo

End Sub

Private Sub Form_Load()
On Error GoTo Err
Dim StrMsgError As String

    Me.left = 0
    Me.top = 0
    
    CTmpDocumentos = ""
    CTmpDocumentosGen = ""
    CTmpGuiasNF = ""
    
    ConfGrid GDocumentos, False, True, False, False
    ConfGrid GGuiasNF, False, True, False, False
    ConfGrid GDocumentosGen, False, True, False, False
    
    ConfGrid gLista, False, False, False, False
    ConfGrid gTiendas, True, False, False, False
    ConfGrid gContactos, True, False, False, False
    ConfGrid gFormaPago, True, False, False, False
    ConfGrid gProductos, True, False, False, False
    
    fraListado.Visible = True
    fraGeneral.Visible = False
    habilitaBotones 7
    nuevo StrMsgError
    If StrMsgError <> "" Then GoTo Err
    
    
    
    Exit Sub

Err:
    If StrMsgError = "" Then StrMsgError = Err.Description
    MsgBox StrMsgError, vbInformation, App.Title

End Sub

Private Sub Form_Unload(Cancel As Integer)
On Error GoTo Err
Dim StrMsgError                         As String
    
    If Len(Trim(CTmpDocumentos)) > 0 Then
        
        Cn.Execute "Drop Table If Exists " & CTmpDocumentos
    
    End If
    
    If Len(Trim(CTmpGuiasNF)) > 0 Then
        
        Cn.Execute "Drop Table If Exists " & CTmpGuiasNF
    
    End If
    
    If Len(Trim(CTmpDocumentosGen)) > 0 Then
        
        Cn.Execute "Drop Table If Exists " & CTmpDocumentosGen
    
    End If
    
    Exit Sub
Err:
    If StrMsgError = "" Then StrMsgError = Err.Description
    MsgBox StrMsgError, vbInformation, App.Title
End Sub

Private Sub gContactos_OnAfterDatasetAction(ByVal Action As DXDBGRIDLibCtl.ExDatasetAction)
Dim i As Integer
    
    If Action = daInsert Then
        gContactos.Columns.ColumnByFieldName("item").Value = gContactos.Count
        gContactos.Dataset.Post
    End If

End Sub

Private Sub gContactos_OnBeforeDatasetAction(ByVal Action As DXDBGRIDLibCtl.ExDatasetAction, Allow As Boolean)
    
    If Action = daInsert Then
        If gContactos.Columns.ColumnByFieldName("GlsPersona").Value = "" Then
            Allow = False
        Else
            gContactos.Columns.FocusedIndex = gContactos.Columns.ColumnByFieldName("idPersona").Index
        End If
    End If
    
End Sub

Private Sub gContactos_OnEditButtonClick(ByVal Column As DXDBGRIDLibCtl.IdxGridColumn, ByVal Node As DXDBGRIDLibCtl.IdxGridNode)
On Error GoTo Err
Dim A           As String
Dim B           As String
Dim StrMsgError As String

    Select Case Column.Index
        Case gContactos.Columns.ColumnByFieldName("IdPersona").Index
            mostrarAyudaTexto "PERSONAUSUARIO", A, B
            If Len(Trim(A)) > 0 Then
                mostrarDatosPersona_prov StrMsgError, A
                If StrMsgError <> "" Then GoTo Err
            End If
    End Select
    
    Exit Sub
    
Err:
    If StrMsgError = "" Then StrMsgError = Err.Description
    MsgBox StrMsgError, vbInformation, App.Title

End Sub

Private Sub mostrarDatosPersona_prov(StrMsgError As String, PIdPersona As String)
On Error GoTo Err
Dim CSqlC                   As String
Dim rst                     As New ADODB.Recordset

    CSqlC = "Select IdPersona,GlsPersona,Telefonos,Mail " & _
            "From Personas " & _
            "Where IdPersona = '" & PIdPersona & "'"
    rst.Open CSqlC, Cn, adOpenStatic, adLockReadOnly
    If Not rst.EOF Then
        gContactos.Dataset.Edit
        gContactos.Columns.ColumnByFieldName("IdPersona").Value = "" & rst.Fields("IdPersona")
        gContactos.Columns.ColumnByFieldName("GlsPersona").Value = "" & rst.Fields("GlsPersona")
        gContactos.Columns.ColumnByFieldName("Telefonos").Value = "" & rst.Fields("Telefonos")
        gContactos.Columns.ColumnByFieldName("Mail").Value = "" & rst.Fields("Mail")
        gContactos.Columns.ColumnByFieldName("IdPersonaAux").Value = "" & rst.Fields("IdPersona")
        
    Else
        gContactos.Dataset.Edit
        gContactos.Columns.ColumnByFieldName("GlsPersona").Value = ""
        gContactos.Columns.ColumnByFieldName("Telefonos").Value = ""
        gContactos.Columns.ColumnByFieldName("Mail").Value = ""
        gContactos.Columns.ColumnByFieldName("IdPersonaAux").Value = ""
    End If
    rst.Close: Set rst = Nothing
    
    Exit Sub
    
Err:
    If StrMsgError = "" Then StrMsgError = Err.Description

End Sub

Private Sub gContactos_OnEdited(ByVal Node As DXDBGRIDLibCtl.IdxGridNode)
On Error GoTo Err
Dim StrMsgError                     As String
    
    If Not gContactos.Dataset.Modified Then Exit Sub
    
    Select Case gContactos.Columns.FocusedColumn.Index
        Case gContactos.Columns.ColumnByFieldName("IdPersona").Index
            If Len(Trim("" & gContactos.Columns.ColumnByFieldName("IdPersona").Value)) > 0 Then
                mostrarDatosPersona_prov StrMsgError, "" & gContactos.Columns.ColumnByFieldName("IdPersona").Value
                If StrMsgError <> "" Then GoTo Err
            End If
    End Select
    
    Exit Sub
    
Err:
    If StrMsgError = "" Then StrMsgError = Err.Description
    MsgBox StrMsgError, vbInformation, App.Title

End Sub

Private Sub gContactos_OnKeyDown(KeyCode As Integer, ByVal Shift As Long)
Dim i As Integer
    
    If KeyCode = 46 Then
        If gContactos.Count > 0 Then
            If MsgBox("Está seguro(a) de eliminar el registro?", vbInformation + vbYesNo, App.Title) = vbYes Then
                If gContactos.Count = 1 Then
                    gContactos.Dataset.Edit
                    gContactos.Columns.ColumnByFieldName("Item").Value = 1
                    gContactos.Columns.ColumnByFieldName("IDPERSONA").Value = ""
                    gContactos.Columns.ColumnByFieldName("GLSPERSONA").Value = ""
                    gContactos.Dataset.Post
                
                Else
                    gContactos.Dataset.Delete
                    gContactos.Dataset.First
                    Do While Not gContactos.Dataset.EOF
                        i = i + 1
                        gContactos.Dataset.Edit
                        gContactos.Columns.ColumnByFieldName("Item").Value = i
                        gContactos.Dataset.Post
                        gContactos.Dataset.Next
                    Loop
                    If gContactos.Dataset.State = dsEdit Or gContactos.Dataset.State = dsInsert Then
                        gContactos.Dataset.Post
                    End If
                End If
            End If
        End If
    End If
    
    If KeyCode = 13 Then
        If gContactos.Dataset.State = dsEdit Or gContactos.Dataset.State = dsInsert Then
              gContactos.Dataset.Post
        End If
    End If

End Sub

Private Sub gFormaPago_OnAfterDatasetAction(ByVal Action As DXDBGRIDLibCtl.ExDatasetAction)
Dim i As Integer

    If Action = daInsert Then
        gFormaPago.Columns.ColumnByFieldName("item").Value = gFormaPago.Count
        gFormaPago.Columns.ColumnByFieldName("GlsCliente").Value = Trim("" & txtGls_Persona.Text)
        gFormaPago.Columns.ColumnByFieldName("idUsuario").Value = Trim("" & glsUser)
        gFormaPago.Dataset.Post
    End If

End Sub

Private Sub gFormaPago_OnBeforeDatasetAction(ByVal Action As DXDBGRIDLibCtl.ExDatasetAction, Allow As Boolean)

    If Action = daInsert Then
        If gFormaPago.Columns.ColumnByFieldName("GlsFormaPago").Value = "" Then
            Allow = False
        Else
            gFormaPago.Columns.FocusedIndex = gFormaPago.Columns.ColumnByFieldName("GlsFormaPago").Index
        End If
    End If

End Sub

Private Sub gFormaPago_OnEditButtonClick(ByVal Column As DXDBGRIDLibCtl.IdxGridColumn, ByVal Node As DXDBGRIDLibCtl.IdxGridNode)
Dim strCod  As String
Dim strDes  As String
Dim IndExi  As Boolean
Dim intFila As Integer
    
    intFila = Node.Index + 1
    IndExi = False
    Select Case Column.Index
        Case gFormaPago.Columns.ColumnByFieldName("GlsFormaPago").Index
            mostrarAyudaTexto "FORMASPAGO", strCod, strDes
            gFormaPago.Dataset.First
            If Not gFormaPago.Dataset.EOF Then
                Do While Not gFormaPago.Dataset.EOF
                    If Trim("" & gFormaPago.Columns.ColumnByFieldName("idFormaPago").Value) = Trim("" & strCod) Then
                        IndExi = True
                    End If
                    gFormaPago.Dataset.Next
                Loop
            End If
                
            gFormaPago.Dataset.RecNo = intFila
            If IndExi = True Then
                MsgBox ("La Forma de Pago Seleccionada ya Esta en la Lista Verifique."), vbInformation, App.Title
                gFormaPago.Dataset.Edit
                gFormaPago.Columns.ColumnByFieldName("idFormaPago").Value = ""
                gFormaPago.Columns.ColumnByFieldName("GlsFormaPago").Value = ""
                gFormaPago.Dataset.Post
                            
            Else
                gFormaPago.Dataset.Edit
                gFormaPago.Columns.ColumnByFieldName("idFormaPago").Value = strCod
                gFormaPago.Columns.ColumnByFieldName("GlsFormaPago").Value = strDes
                gFormaPago.Dataset.Post
            End If
    End Select
    
End Sub

Private Sub gFormaPago_OnKeyDown(KeyCode As Integer, ByVal Shift As Long)
Dim i As Integer
    
    If KeyCode = 46 Then
        If gFormaPago.Count > 0 Then
            If MsgBox("Está seguro(a) de eliminar el registro?", vbInformation + vbYesNo, App.Title) = vbYes Then
                If gFormaPago.Count = 1 Then
                    gFormaPago.Dataset.Edit
                    gFormaPago.Columns.ColumnByFieldName("Item").Value = 1
                    gFormaPago.Columns.ColumnByFieldName("IdCliente").Value = ""
                    gFormaPago.Columns.ColumnByFieldName("idFormaPago").Value = ""
                    gFormaPago.Columns.ColumnByFieldName("GlsCliente").Value = Trim("" & txtGls_Persona.Text)
                    gFormaPago.Columns.ColumnByFieldName("GlsFormaPago").Value = ""
                    gFormaPago.Columns.ColumnByFieldName("indEstado").Value = 1
                    gFormaPago.Columns.ColumnByFieldName("idUsuario").Value = Trim("" & glsUser)
                    gFormaPago.Columns.ColumnByFieldName("FecRegistro").Value = ""
                    gFormaPago.Dataset.Post
                
                Else
                    gFormaPago.Dataset.Delete
                    gFormaPago.Dataset.First
                    Do While Not gFormaPago.Dataset.EOF
                        i = i + 1
                        gFormaPago.Dataset.Edit
                        gFormaPago.Columns.ColumnByFieldName("Item").Value = i
                        gFormaPago.Dataset.Post
                        gFormaPago.Dataset.Next
                    Loop
                    If gFormaPago.Dataset.State = dsEdit Or gFormaPago.Dataset.State = dsInsert Then
                        gFormaPago.Dataset.Post
                    End If
                End If
            End If
        End If
    End If
    
    If KeyCode = 13 Then
        If gFormaPago.Dataset.State = dsEdit Or gFormaPago.Dataset.State = dsInsert Then
              gFormaPago.Dataset.Post
        End If
    End If
    
End Sub

Private Sub gProductos_OnAfterDatasetAction(ByVal Action As DXDBGRIDLibCtl.ExDatasetAction)
Dim i As Integer
    
    If Action = daInsert Then
        gProductos.Columns.ColumnByFieldName("item").Value = gProductos.Count
        gProductos.Dataset.Post
    End If

End Sub

Private Sub gProductos_OnBeforeDatasetAction(ByVal Action As DXDBGRIDLibCtl.ExDatasetAction, Allow As Boolean)
    
    If Action = daInsert Then
        If gProductos.Columns.ColumnByFieldName("idProducto").Value = "" Then
            Allow = False
        Else
            gProductos.Columns.FocusedIndex = gProductos.Columns.ColumnByFieldName("idProducto").Index
        End If
    End If

End Sub

Private Sub gProductos_OnEditButtonClick(ByVal Column As DXDBGRIDLibCtl.IdxGridColumn, ByVal Node As DXDBGRIDLibCtl.IdxGridNode)
Dim strCod As String
Dim strDes As String
    
    Select Case Column.Index
        Case gProductos.Columns.ColumnByFieldName("idProducto").Index
            strCod = gProductos.Columns.ColumnByFieldName("idProducto").Value
            strDes = gProductos.Columns.ColumnByFieldName("GlsProducto").Value
            
            mostrarAyudaTexto "PRODUCTOS", strCod, strDes
            
            If existeEnGrilla(gProductos, "idProducto", strCod) = False Then
                gProductos.Dataset.Edit
                gProductos.Columns.ColumnByFieldName("idProducto").Value = strCod
                gProductos.Columns.ColumnByFieldName("GlsProducto").Value = strDes
                gProductos.Dataset.Post
            Else
                MsgBox "El Producto ya fue ingresado.", vbInformation, App.Title
            End If
    End Select
    
End Sub

Private Sub gProductos_OnKeyDown(KeyCode As Integer, ByVal Shift As Long)
Dim i As Integer

    If KeyCode = 46 Then
        If gProductos.Count > 0 Then
            If MsgBox("Está seguro(a) de eliminar el registro?", vbInformation + vbYesNo, App.Title) = vbYes Then
                If gProductos.Count = 1 Then
                    gProductos.Dataset.Edit
                    gProductos.Columns.ColumnByFieldName("Item").Value = 1
                    gProductos.Columns.ColumnByFieldName("idProducto").Value = ""
                    gProductos.Columns.ColumnByFieldName("GlsProducto").Value = ""
                    gProductos.Dataset.Post
                
                Else
                    gProductos.Dataset.Delete
                    gProductos.Dataset.First
                    Do While Not gProductos.Dataset.EOF
                        i = i + 1
                        gProductos.Dataset.Edit
                        gProductos.Columns.ColumnByFieldName("Item").Value = i
                        gProductos.Dataset.Post
                        gProductos.Dataset.Next
                    Loop
                    If gProductos.Dataset.State = dsEdit Or gProductos.Dataset.State = dsInsert Then
                        gProductos.Dataset.Post
                    End If
                End If
            End If
        End If
    End If
    
    If KeyCode = 13 Then
        If gProductos.Dataset.State = dsEdit Or gProductos.Dataset.State = dsInsert Then
              gProductos.Dataset.Post
        End If
    End If
    
End Sub

Private Sub gProductos_OnKeyPress(Key As Integer)
Dim strCod As String
Dim strDes As String

    Select Case gProductos.Columns.FocusedColumn.Index
        Case gProductos.Columns.ColumnByFieldName("idProducto").Index
            strCod = gProductos.Columns.ColumnByFieldName("idProducto").Value
            strDes = gProductos.Columns.ColumnByFieldName("GlsProducto").Value
            
            mostrarAyudaKeyasciiTexto Key, "PRODUCTOS", strCod, strDes
            Key = 0
            If existeEnGrilla(gProductos, "idProducto", strCod) = False Then
                gProductos.Dataset.Edit
                gProductos.Columns.ColumnByFieldName("idProducto").Value = strCod
                gProductos.Columns.ColumnByFieldName("GlsProducto").Value = strDes
                gProductos.Dataset.Post
                gProductos.SetFocus
            Else
                MsgBox "El Producto ya fue ingresado.", vbInformation, App.Title
            End If
    End Select
    
End Sub

Private Sub gTiendas_OnAfterDatasetAction(ByVal Action As DXDBGRIDLibCtl.ExDatasetAction)
Dim i As Integer

    If Action = daInsert Then
        gTiendas.Columns.ColumnByFieldName("item").Value = gTiendas.Count
        gTiendas.Dataset.Post
    End If

End Sub

Private Sub gTiendas_OnBeforeDatasetAction(ByVal Action As DXDBGRIDLibCtl.ExDatasetAction, Allow As Boolean)

    If Action = daInsert Then
        If gTiendas.Columns.ColumnByFieldName("GlsDireccion").Value = "" Then
            Allow = False
        Else
            gTiendas.Columns.FocusedIndex = gTiendas.Columns.ColumnByFieldName("GlsDireccion").Index
        End If
    End If

End Sub

Private Sub gLista_OnDblClick()
 On Error GoTo Err
Dim StrMsgError As String

    mostrarCliente gLista.Columns.ColumnByName("idCliente").Value, StrMsgError
    If StrMsgError <> "" Then GoTo Err
    fraListado.Visible = False
    fraGeneral.Visible = True
    fraGeneral.Enabled = False
    txtCod_Persona.Enabled = False
    cmbAyudaPersona.Enabled = False
    habilitaBotones 2
    
    SSTab2.Tab = 0
    SSTab1.Tab = 0
    
    fraListado.Visible = False
    fraGeneral.Visible = True
    fraGeneral.Enabled = True
    If traerCampo("usuarios", "indJefe", "idUsuario", glsUser, True) = "0" Then
        Frame4.Enabled = True
    End If
            
    
Exit Sub
Err:
    If StrMsgError = "" Then StrMsgError = Err.Description
    MsgBox StrMsgError, vbInformation, App.Title
End Sub

Private Sub gTiendas_OnEditButtonClick(ByVal Column As DXDBGRIDLibCtl.IdxGridColumn, ByVal Node As DXDBGRIDLibCtl.IdxGridNode)
Dim strCod  As String
Dim strDes  As String
Dim IndExi  As Boolean
Dim intFila As Integer
    
    intFila = Node.Index + 1
    IndExi = False
    
    Select Case Column.Index
        Case gTiendas.Columns.ColumnByFieldName("idPais").Index
            mostrarAyuda "PAIS", txtCodigo, txtDes
            gTiendas.Dataset.Edit
            gTiendas.Columns.ColumnByFieldName("idPais").Value = Trim("" & txtCodigo.Text)
            gTiendas.Columns.ColumnByFieldName("iddpto").Value = ""
            gTiendas.Columns.ColumnByFieldName("idprov").Value = ""
            gTiendas.Columns.ColumnByFieldName("idDistrito").Value = ""
            gTiendas.Dataset.Post
                               
        Case gTiendas.Columns.ColumnByFieldName("iddpto").Index
            mostrarAyuda "DEPARTAMENTO", txtCodigo, txtDes, " AND idPais = '" & Trim("" & gTiendas.Columns.ColumnByFieldName("idPais").Value) & "'"
            gTiendas.Dataset.Edit
            gTiendas.Columns.ColumnByFieldName("iddpto").Value = Trim("" & txtCodigo.Text)
            gTiendas.Columns.ColumnByFieldName("idprov").Value = ""
            gTiendas.Columns.ColumnByFieldName("idDistrito").Value = ""
            gTiendas.Dataset.Post
                
        Case gTiendas.Columns.ColumnByFieldName("idprov").Index
            mostrarAyuda "PROVINCIA", txtCodigo, txtDes, "AND idPais = '" & Trim("" & gTiendas.Columns.ColumnByFieldName("idPais").Value) & "' AND idDpto = '" & Trim("" & gTiendas.Columns.ColumnByFieldName("iddpto").Value) + "'"
            gTiendas.Dataset.Edit
            gTiendas.Columns.ColumnByFieldName("idprov").Value = Trim("" & txtCodigo.Text)
            gTiendas.Columns.ColumnByFieldName("idDistrito").Value = ""
            gTiendas.Dataset.Post
                
        Case gTiendas.Columns.ColumnByFieldName("idDistrito").Index
            mostrarAyuda "DISTRITO", txtCodigo, txtDes, "AND idPais = '" & Trim("" & gTiendas.Columns.ColumnByFieldName("idPais").Value) & "' AND idDpto = '" & Trim("" & gTiendas.Columns.ColumnByFieldName("iddpto").Value) & "' and idProv = '" + Trim("" & gTiendas.Columns.ColumnByFieldName("idprov").Value) + "'"
            
            gTiendas.Dataset.Edit
            gTiendas.Columns.ColumnByFieldName("idDistrito").Value = Trim("" & txtCodigo.Text)
            gTiendas.Dataset.Post
            
        Case gTiendas.Columns.ColumnByFieldName("idVendedor").Index
            mostrarAyuda "VENDEDOR", txtCodigo, txtDes
            gTiendas.Dataset.Edit
            gTiendas.Columns.ColumnByFieldName("idVendedor").Value = Trim("" & txtCodigo.Text)
            gTiendas.Columns.ColumnByFieldName("GlsVendedor").Value = Trim("" & txtDes.Text)
            gTiendas.Dataset.Post
            
            
    End Select

End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
On Error GoTo Err
Dim StrMsgError As String

    Select Case Button.Index
        Case 1 'Nuevo
            indBoton = 0
            nuevo StrMsgError
            If StrMsgError <> "" Then GoTo Err
            fraListado.Visible = False
            fraGeneral.Visible = True
            fraGeneral.Enabled = True
            If traerCampo("usuarios", "indJefe", "idUsuario", glsUser, True) = "0" Then
                Frame4.Enabled = True
            End If
            
        Case 2 'Grabar
            Grabar StrMsgError
            If StrMsgError <> "" Then GoTo Err
        
        Case 3 'Modificar
            
            indBoton = 1
            fraGeneral.Enabled = True
            If traerCampo("usuarios", "indJefe", "idUsuario", glsUser, True) = "0" Then
                Frame4.Enabled = False
            End If
            
        Case 4, 7 'Cancelar
            fraListado.Visible = True
            fraGeneral.Visible = False
            fraGeneral.Enabled = False
        
        Case 5 'Eliminar
            eliminar StrMsgError
            If StrMsgError <> "" Then GoTo Err
        
        Case 6 'Imprimir
        
        Case 8
            gLista.m.ExportToXLS App.Path & "\Temporales\ListadoClientes.xls"
            ShellEx App.Path & "\Temporales\ListadoClientes.xls", essSW_MAXIMIZE, , , "open", Me.hwnd
        
        Case 9 'Salir
            Unload Me
    End Select
    habilitaBotones Button.Index
    
    Exit Sub

Err:
    MsgBox StrMsgError, vbInformation, App.Title
    Exit Sub
    Resume
End Sub

Private Sub txt_TextoBuscar_Change()
On Error GoTo Err
Dim StrMsgError As String
    
    If glsEnterAyudaClientes = False Then
        listaCliente StrMsgError
        If StrMsgError <> "" Then GoTo Err
    End If
    
    Exit Sub
    
Err:
    If StrMsgError = "" Then StrMsgError = Err.Description
    MsgBox StrMsgError, vbInformation, App.Title

End Sub

Private Sub txt_TextoBuscar_KeyPress(KeyAscii As Integer)
On Error GoTo Err
Dim StrMsgError As String

    If KeyAscii = 13 Then
        listaCliente StrMsgError
        If StrMsgError <> "" Then GoTo Err
    End If

    Exit Sub
    
Err:
    If StrMsgError = "" Then StrMsgError = Err.Description
    MsgBox StrMsgError, vbInformation, App.Title

End Sub

Private Sub txtCod_Canal_Change()
    
    txtGls_Canal.Text = traerCampo("canal", "GlsCanal", "idCanal", txtCod_Canal.Text, True)

End Sub

Private Sub txtCod_EmpTrans_Change()
    
    txtGls_EmpTrans.Text = traerCampo("personas", "GlsPersona", "idPersona", txtCod_EmpTrans.Text, False)

End Sub

Private Sub txtCod_EmpTrans_KeyPress(KeyAscii As Integer)

    If KeyAscii <> 13 Then
        mostrarAyudaKeyascii KeyAscii, "EMPTRANS", txtCod_EmpTrans, txtGls_EmpTrans
        KeyAscii = 0
        If txtCod_EmpTrans.Text <> "" Then SendKeys "{tab}"
    End If

End Sub

Private Sub txtCod_Lista_Change()
    
    txtGls_Lista.Text = traerCampo("listaprecios", "GlsLista", "idLista", txtCod_Lista.Text, True)

End Sub

Private Sub txtCod_Lista_KeyDown(KeyCode As Integer, Shift As Integer)

    If KeyCode = 8 Then
        txtCod_Lista.Text = ""
    End If

End Sub

Private Sub txtCod_Lista_KeyPress(KeyAscii As Integer)

    If KeyAscii <> 13 And KeyAscii <> 8 Then
        mostrarAyudaKeyascii KeyAscii, "LISTAPRECIOS", txtCod_Lista, txtGls_Lista
        KeyAscii = 0
        If txtCod_Lista.Text <> "" Then SendKeys "{tab}"
    End If

End Sub

Private Sub txtCod_Persona_Change()
    
    txtGls_Persona.Text = traerCampo("personas", "GlsPersona", "idPersona", txtCod_Persona.Text, False)
    gFormaPago.Dataset.Edit
    gFormaPago.Columns.ColumnByFieldName("idCliente").Value = Trim("" & txtCod_Persona.Text)
    gFormaPago.Columns.ColumnByFieldName("GlsCliente").Value = Trim("" & txtGls_Persona.Text)
    gFormaPago.Columns.ColumnByFieldName("idUsuario").Value = Trim("" & glsUser)
    gFormaPago.Dataset.Post

End Sub

Private Sub txtCod_Persona_KeyPress(KeyAscii As Integer)

    If KeyAscii <> 13 Then
        mostrarAyudaKeyascii KeyAscii, "PERSONACLIENTE", txtCod_Persona, txtGls_Persona
        KeyAscii = 0
        If txtCod_Persona.Text <> "" Then
            mostrarDatosPersona
            SendKeys "{tab}"
        End If
    End If

End Sub

Private Sub txtCod_TipoPersona_Change()
    
    txtGls_TipoPersona.Text = traerCampo("Datos", "GlsDato", "idDato", txtCod_TipoPersona.Text, False)

End Sub

Private Sub txtCod_Depa_Change()
    
    txtGls_Depa.Text = traerCampo("Ubigeo", "GlsUbigeo", "idDpto", txtCod_Depa.Text, False, " idProv = '00' And idPais = '" & txtCod_Pais.Text & "' ")

End Sub

Private Sub txtCod_Distrito_Change()
    
    txtGls_Distrito.Text = traerCampo("Ubigeo", "GlsUbigeo", "idDistrito", txtCod_Distrito.Text, False, " idPais = '" & txtCod_Pais.Text & "' ")

End Sub

Private Sub txtCod_Pais_Change()
    
    txtGls_Pais.Text = traerCampo("Datos", "GlsDato", "idDato", txtCod_Pais.Text, False)

End Sub

Private Sub txtCod_Prov_Change()
    
    txtGls_Prov.Text = traerCampo("Ubigeo", "GlsUbigeo", "idProv", txtCod_Prov.Text, False, " idDpto = '" & txtCod_Depa.Text & "' and idProv <> '00' and idDist = '00' And idPais = '" & txtCod_Pais.Text & "'  ")

End Sub

Private Sub habilitaBotones(indexBoton As Integer)
Dim indHabilitar As Boolean

    Select Case indexBoton
        Case 1, 2, 3 'Nuevo, Grabar, Modificar
            'If indexBoton = 2 Then indHabilitar = True
            'Toolbar1.Buttons(1).Visible = indHabilitar 'Nuevo
            'Toolbar1.Buttons(2).Visible = Not indHabilitar 'Grabar
            'Toolbar1.Buttons(3).Visible = indHabilitar 'Modificar
            'Toolbar1.Buttons(4).Visible = Not indHabilitar 'Cancelar
            'Toolbar1.Buttons(5).Visible = indHabilitar 'Imprimir
            'Toolbar1.Buttons(6).Visible = indHabilitar 'Imprimir
            Toolbar1.Buttons(7).Visible = True 'Lista
            Toolbar1.Buttons(8).Visible = True 'Excel
        Case 4, 7 'Cancelar, Lista
            'Toolbar1.Buttons(1).Visible = True
            'Toolbar1.Buttons(2).Visible = False
            'Toolbar1.Buttons(3).Visible = False
            'Toolbar1.Buttons(4).Visible = False
            'Toolbar1.Buttons(5).Visible = False
            'Toolbar1.Buttons(6).Visible = False
            Toolbar1.Buttons(7).Visible = True
            Toolbar1.Buttons(8).Visible = True
    End Select

End Sub

Private Sub Grabar(ByRef StrMsgError As String)
On Error GoTo Err
Dim strCodigo                   As String
Dim strMsg                      As String
Dim CSqlC                       As String
Dim CIdPersona                  As String
Dim CGlsPersona                 As String
Dim i                           As Long
Dim indTrans                    As Boolean
Dim CodFormaPago                As String
    
    indTrans = False
    validaFormSQL Me, StrMsgError
    If StrMsgError <> "" Then GoTo Err
    If txtCod_ClienteInterno.Text <> "" Then
        validaHomonimia "clientes", "idClienteInterno", "idCliente", txtCod_ClienteInterno.Text, txtCod_Persona.Text, True, StrMsgError, "", "El codigo interno ya se encuentra registrado"
        If StrMsgError <> "" Then GoTo Err
    End If
    
    txt_Email.Tag = ""
    txt_Telefono.Tag = ""
    glsobservacioncliente = txtGls_Obs.Text

    gTiendas.Dataset.First
    If Not gTiendas.Dataset.EOF Then
        Do While Not gTiendas.Dataset.EOF
            gTiendas.Dataset.Edit
            gTiendas.Columns.ColumnByFieldName("idtdacli").Value = Trim(glsEmpresa & Trim("" & txtCod_Persona.Text) & gTiendas.Columns.ColumnByFieldName("item").Value)
            gTiendas.Dataset.Post
            gTiendas.Dataset.Next
        Loop
    End If

    Cn.BeginTrans
    indTrans = True
    If indBoton = 0 Then 'graba
        EjecutaSQLFormTrans Me, 0, True, "Clientes", StrMsgError, False, , gTiendas, "TiendasCliente", "IdPersona", txtCod_Persona.Text
        If StrMsgError <> "" Then GoTo Err
        strMsg = "Grabó"
        
    Else 'modifica
        EjecutaSQLFormTrans Me, 1, True, "Clientes", StrMsgError, False, "IdCliente", gTiendas, "TiendasCliente", "IdPersona", txtCod_Persona.Text
        If StrMsgError <> "" Then GoTo Err
        strMsg = "Modificó"
    End If
    
    '--- Graba Contactos
    If gContactos.Count > 0 Then
        gContactos.Dataset.First
        Do While Not gContactos.Dataset.EOF
            If Len(Trim("" & gContactos.Columns.ColumnByFieldName("IdPersona").Value)) = 0 And Len(Trim("" & gContactos.Columns.ColumnByFieldName("GlsPersona").Value)) = 0 Then
                If gContactos.Count = 1 Then
                    gContactos.Dataset.Edit
                    gContactos.Columns.ColumnByFieldName("Item").Value = 1
                    gContactos.Columns.ColumnByFieldName("IdPersona").Value = ""
                    gContactos.Columns.ColumnByFieldName("GlsPersona").Value = ""
                    gContactos.Columns.ColumnByFieldName("Telefonos").Value = ""
                    gContactos.Columns.ColumnByFieldName("Mail").Value = ""
                    gContactos.Columns.ColumnByFieldName("GlsPersonaAux").Value = ""
                    gContactos.Dataset.Post
                    
                Else
                    gContactos.Dataset.Delete
                End If
            End If
            gContactos.Dataset.Next
        Loop
        gContactos.Dataset.First
        
        Do While Not gContactos.Dataset.EOF
            i = i + 1
            gContactos.Dataset.Edit
            gContactos.Columns.ColumnByFieldName("Item").Value = i
            gContactos.Dataset.Post
            gContactos.Dataset.Next
        Loop
        
        If gContactos.Dataset.State = dsEdit Or gContactos.Dataset.State = dsInsert Then gContactos.Dataset.Post
        CSqlC = "Delete From ContactosClientes " & _
                "Where IdEmpresa='" & glsEmpresa & "' And IdCliente = '" & txtCod_Persona.Text & "'"
        Cn.Execute CSqlC
        
        gContactos.Dataset.First
        Do While Not gContactos.Dataset.EOF
            CGlsPersona = "" & gContactos.Columns.ColumnByFieldName("GlsPersona").Value
            If Len(Trim("" & gContactos.Columns.ColumnByFieldName("IdPersona").Value)) = 0 Then
                If Len(Trim("" & gContactos.Columns.ColumnByFieldName("GlsPersona").Value)) > 0 Then
                    validaHomonimia "Personas", "GlsPersona", "IdPersona", CGlsPersona, "", False, StrMsgError
                    If StrMsgError <> "" Then GoTo Err
                    CIdPersona = GeneraCorrelativoAnoMes("Personas", "IdPersona", False)
                    CSqlC = "Insert Into Personas(IdPersona,GlsPersona,ApellidoPaterno,ApellidoMaterno,Nombres,TipoPersona,Ruc,IdDistrito,Direccion," & _
                            "FechaNacimiento,Telefonos,Mail,DireccionEntrega,GlsContacto,IdPais,Linea_Credito)" & _
                            "Select '" & CIdPersona & "','" & CGlsPersona & "','" & CGlsPersona & "','" & CGlsPersona & "','" & CGlsPersona & "'," & _
                            "'01001','" & CIdPersona & "',IdDistrito,Direccion,FechaNacimiento," & _
                            "'" & "" & gContactos.Columns.ColumnByFieldName("Telefonos").Value & "'," & _
                            "'" & "" & gContactos.Columns.ColumnByFieldName("Mail").Value & "',DireccionEntrega,'',IdPais,Linea_Credito " & _
                            "From Personas " & _
                            "Where IdPersona = '" & txtCod_Persona.Text & "'"
                    Cn.Execute CSqlC
                    
                    gContactos.Dataset.Edit
                    gContactos.Columns.ColumnByFieldName("IdPersonaAux").Value = CIdPersona
                    gContactos.Dataset.Post
                End If
            
            Else
                CIdPersona = "" & gContactos.Columns.ColumnByFieldName("IdPersona").Value
                validaHomonimia "Personas", "GlsPersona", "IdPersona", CGlsPersona, CIdPersona, False, StrMsgError
                If StrMsgError <> "" Then GoTo Err
                    
                CSqlC = "Update Personas " & _
                        "Set GlsPersona = '" & CGlsPersona & "',ApellidoPaterno = '" & CGlsPersona & "',ApellidoMaterno = '" & CGlsPersona & "'," & _
                        "Nombres = '" & CGlsPersona & "',Telefonos = '" & "" & gContactos.Columns.ColumnByFieldName("Telefonos").Value & "'," & _
                        "Mail = '" & "" & gContactos.Columns.ColumnByFieldName("Mail").Value & "' " & _
                        "Where IdPersona = '" & CIdPersona & "'"
                Cn.Execute CSqlC
            End If
            
            If Len(Trim("" & gContactos.Columns.ColumnByFieldName("IdPersonaAux").Value)) > 0 Then
                CSqlC = "Insert Into ContactosClientes(IdCliente,IdContacto,IdEmpresa)" & _
                        "Values('" & txtCod_Persona.Text & "','" & gContactos.Columns.ColumnByFieldName("IdPersonaAux").Value & "','" & glsEmpresa & "')"
                Cn.Execute CSqlC
            End If
            gContactos.Dataset.Next
        Loop
    End If

    CSqlC = "Delete From ClientesFormaPagos " & _
            "Where IdCliente = '" & txtCod_Persona.Text & "' And IdEmpresa = '" & glsEmpresa & "'"
    Cn.Execute CSqlC
                             
    gFormaPago.Dataset.First
    If Not gFormaPago.Dataset.EOF Then
        Do While Not gFormaPago.Dataset.EOF
        
            CSqlC = "Insert Into ClientesFormaPagos(IdEmpresa,IdCliente,IdFormaPago,FecRegistro,IdUsuario,IndEstado)" & _
                    "Values('" & glsEmpresa & "','" & txtCod_Persona.Text & "','" & gFormaPago.Columns.ColumnByFieldName("idFormaPago").Value & "'," & _
                    "'" & Format(Trim("" & getFechaHoraSistema), "yyyy-mm-dd hh:mm:ss") & "'," & _
                    "'" & gFormaPago.Columns.ColumnByFieldName("idUsuario").Value & "'," & gFormaPago.Columns.ColumnByFieldName("indEstado").Value & ")"
                    
            Cn.Execute CSqlC
            
            If gFormaPago.Columns.ColumnByFieldName("indEstado").Value = 1 Then
                CodFormaPago = Trim("" & gFormaPago.Columns.ColumnByFieldName("idFormaPago").Value)
            End If
            
            gFormaPago.Dataset.Next
            
        Loop
    End If
    
    CSqlC = "Update Clientes set IdFormaPago  = '" & CodFormaPago & "' where idcliente  = '" & txtCod_Persona.Text & "' and idempresa = '" & glsEmpresa & "' "
    Cn.Execute CSqlC
    
    CSqlC = "Delete From ProductosClientes " & _
            "Where IdCliente = '" & txtCod_Persona.Text & "' And IdEmpresa = '" & glsEmpresa & "'"
    Cn.Execute CSqlC
                             
    gProductos.Dataset.First
    If Not gProductos.Dataset.EOF Then
        Do While Not gProductos.Dataset.EOF
        
            CSqlC = "Insert Into ProductosClientes(IdEmpresa,IdCliente,IdProducto,Codigo,CodigoBarra)" & _
                    "Values('" & glsEmpresa & "','" & txtCod_Persona.Text & "','" & gProductos.Columns.ColumnByFieldName("IdProducto").Value & "'," & _
                    "'" & gProductos.Columns.ColumnByFieldName("Codigo").Value & "','" & gProductos.Columns.ColumnByFieldName("CodigoBarra").Value & "')"
                    
            Cn.Execute CSqlC
            
            gProductos.Dataset.Next
            
        Loop
    End If
    
    Cn.CommitTrans
    indTrans = False
    
    gContactos.Dataset.Edit
    gContactos.Columns.ColumnByFieldName("IdPersona").Value = gContactos.Columns.ColumnByFieldName("IdPersonaAux").Value
    gContactos.Dataset.Post
                    
    MsgBox "Se " & strMsg & " Satisfactoriamente", vbInformation, App.Title
    
    fraGeneral.Enabled = False
    
    listaCliente StrMsgError
    If StrMsgError <> "" Then GoTo Err
    
    Exit Sub
    
Err:
    If indTrans Then Cn.RollbackTrans
    If StrMsgError = "" Then StrMsgError = Err.Description
    Exit Sub
    Resume
End Sub

Private Sub nuevo(ByRef StrMsgError As String)
On Error GoTo Err
Dim rst As New ADODB.Recordset
Dim RsC As New ADODB.Recordset
Dim rsl As New ADODB.Recordset
Dim rsF As New ADODB.Recordset
Dim rsp As New ADODB.Recordset

    limpiaForm Me
    txtCod_Persona.Enabled = True
    cmbAyudaPersona.Enabled = True
    
    rst.Fields.Append "Item", adInteger, , adFldRowID
    rst.Fields.Append "GlsNombre", adVarChar, 120, adFldIsNullable
    rst.Fields.Append "GlsDireccion", adVarChar, 255, adFldIsNullable
    rst.Fields.Append "GlsTelefonos", adVarChar, 85, adFldIsNullable
    rst.Fields.Append "GlsContacto", adVarChar, 120, adFldIsNullable
    rst.Fields.Append "idtdacli", adVarChar, 120, adFldIsNullable
    rst.Fields.Append "idPais", adVarChar, 85, adFldIsNullable
    rst.Fields.Append "iddpto", adVarChar, 120, adFldIsNullable
    rst.Fields.Append "idprov", adVarChar, 120, adFldIsNullable
    rst.Fields.Append "idDistrito", adVarChar, 120, adFldIsNullable
    rst.Fields.Append "idVendedor", adVarChar, 8, adFldIsNullable
    rst.Fields.Append "GlsVendedor", adVarChar, 120, adFldIsNullable
    
    rst.Open
    
    rst.AddNew
    rst.Fields("Item") = 1
    rst.Fields("GlsNombre") = ""
    rst.Fields("GlsDireccion") = ""
    rst.Fields("GlsTelefonos") = ""
    rst.Fields("GlsContacto") = ""
    rst.Fields("idtdacli") = ""
    rst.Fields("idPais") = ""
    rst.Fields("iddpto") = ""
    rst.Fields("idprov") = ""
    rst.Fields("idDistrito") = ""
    rst.Fields("idVendedor") = ""
    rst.Fields("GlsVendedor") = ""
    
    mostrarDatosGridSQL gTiendas, rst, StrMsgError
    If StrMsgError <> "" Then GoTo Err
    
    gTiendas.Columns.FocusedIndex = gTiendas.Columns.ColumnByFieldName("GlsNombre").Index
    RsC.Fields.Append "Item", adInteger, , adFldRowID
    RsC.Fields.Append "IdPersona", adVarChar, 8, adFldIsNullable
    RsC.Fields.Append "GlsPersona", adVarChar, 250, adFldIsNullable
    RsC.Fields.Append "Telefonos", adVarChar, 85, adFldIsNullable
    RsC.Fields.Append "Mail", adVarChar, 85, adFldIsNullable
    RsC.Fields.Append "IdPersonaAux", adVarChar, 8, adFldIsNullable
    RsC.Open
    
    RsC.AddNew
    RsC.Fields("Item") = 1
    RsC.Fields("IdPersona") = ""
    RsC.Fields("GlsPersona") = ""
    RsC.Fields("Telefonos") = ""
    RsC.Fields("Mail") = ""
    RsC.Fields("IdPersonaAux") = ""
    
    mostrarDatosGridSQL gContactos, RsC, StrMsgError
    If StrMsgError <> "" Then GoTo Err
    
    gContactos.Columns.FocusedIndex = gContactos.Columns.ColumnByFieldName("IdPersona").Index
    
    rsF.Fields.Append "Item", adInteger, , adFldRowID
    rsF.Fields.Append "IdCliente", adVarChar, 8, adFldIsNullable
    rsF.Fields.Append "glsCliente", adVarChar, 255, adFldIsNullable
    rsF.Fields.Append "idFormaPago", adVarChar, 8, adFldIsNullable
    rsF.Fields.Append "GlsFormaPago", adVarChar, 100, adFldIsNullable
    rsF.Fields.Append "indEstado", adInteger, 4, adFldIsNullable
    rsF.Fields.Append "idUsuario", adVarChar, 8, adFldIsNullable
    rsF.Fields.Append "FecRegistro", adVarChar, 40, adFldIsNullable
    rsF.Open
    
    rsF.AddNew
    rsF.Fields("Item") = 1
    rsF.Fields("IdCliente") = ""
    rsF.Fields("idFormaPago") = ""
    rsF.Fields("GlsCliente") = ""
    rsF.Fields("GlsFormaPago") = ""
    rsF.Fields("indEstado") = 1
    rsF.Fields("idUsuario") = ""
    rsF.Fields("FecRegistro") = ""
    
    mostrarDatosGridSQL gFormaPago, rsF, StrMsgError
    If StrMsgError <> "" Then GoTo Err
    
    rsp.Fields.Append "Item", adInteger, , adFldRowID
    rsp.Fields.Append "IdProducto", adVarChar, 8, adFldIsNullable
    rsp.Fields.Append "GlsProducto", adVarChar, 255, adFldIsNullable
    rsp.Fields.Append "Codigo", adVarChar, 100, adFldIsNullable
    rsp.Fields.Append "CodigoBarra", adVarChar, 100, adFldIsNullable
    rsp.Open
        
    rsp.AddNew
    rsp.Fields("Item") = 1
    rsp.Fields("IdProducto") = ""
    rsp.Fields("GlsProducto") = ""
    rsp.Fields("Codigo") = ""
    rsp.Fields("CodigoBarra") = ""
        
    mostrarDatosGridSQL gProductos, rsp, StrMsgError
    If StrMsgError <> "" Then GoTo Err
    
    If traerCampo("Parametros", "ValParametro", "GlsParametro", "CLIENTE_ANULA", True) = "S" Then
        ChkAnula.Visible = True
    Else
        ChkAnula.Visible = False
    End If
        
    If traerCampo("Parametros", "ValParametro", "GlsParametro", "CLIENTE_ESPECIAL", True) = "S" Then
        ChkEspecial.Visible = True
    Else
        ChkEspecial.Visible = False
    End If
    
    chkAgenteRetencion.Value = 0
    chktercros.Value = 0
    chkComision.Value = 0
    ChkAnula.Value = 0
    ChkEspecial.Value = 0
    
Exit Sub
Err:
    If StrMsgError = "" Then StrMsgError = Err.Description

End Sub

Private Sub listaCliente(ByRef StrMsgError As String)
On Error GoTo Err
Dim strCond As String

    strCond = ""
    If Trim(txt_TextoBuscar.Text) <> "" Then
        strCond = Trim(txt_TextoBuscar.Text)
        strCond = "AND (GlsPersona LIKE '%" & strCond & "%' Or idCliente LIKE '%" & strCond & "%' Or ruc LIKE '%" & strCond & "%') "
    End If
    
    csql = "SELECT c.idCliente ,p.GlsPersona ," & _
            "if(p.tipoPersona = '01001','Natural','Juridica') as TipoPersona,p.ruc,concat(p.direccion,', ',ifnull(u.glsUbigeo, '')) as Direccion, " & _
            "c.idClienteInterno, c.Val_Dscto " & _
            "FROM clientes c inner join personas p " & _
            "on c.idCliente = p.idPersona left join ubigeo u on p.iddistrito = u.iddistrito and p.idPais = u.idPais WHERE c.idEmpresa = '" & glsEmpresa & "'" & strCond & _
            "ORDER BY idCliente"
    With gLista
        .DefaultFields = False
        .Dataset.ADODataset.ConnectionString = strcn
        .Dataset.ADODataset.CursorLocation = clUseClient
        .Dataset.Active = False
        .Dataset.ADODataset.CommandText = csql
        .Dataset.DisableControls
        .Dataset.Active = True
        .KeyField = "idCliente"
    End With
    Me.Refresh
    
    Exit Sub

Err:
    If StrMsgError = "" Then StrMsgError = Err.Description

End Sub

Private Sub mostrarCliente(strCodCli As String, ByRef StrMsgError As String)
On Error GoTo Err
Dim rst         As New ADODB.Recordset
Dim rsg         As New ADODB.Recordset
Dim RsC         As New ADODB.Recordset
Dim rsl         As New ADODB.Recordset
Dim rsF         As New ADODB.Recordset
Dim CSqlC       As String
Dim cont        As Integer
Dim item        As Integer
Dim rsp         As New ADODB.Recordset

    limpiaForm Me

    CSqlC = "SELECT c.idProvCliente,c.idCliente, c.idVendedorCampo, c.idEmpTrans, c.idLista, c.idFormaPago, c.Val_LineaCredito, c.indAgenteRetencion, " & _
            "c.idMonedaLineaCredito, c.idClienteInterno, c.Val_Dscto, c.idCanal,c.indventasterceros, " & _
            "p.Telefonos, p.Mail,p.GlsContacto,c.GlsObservacion,c.indComision,c.indAnula,C.IndEspecial " & _
            "FROM Clientes c  Inner Join Personas p On c.idCliente=p.idPersona " & _
            "WHERE c.idCliente = '" & strCodCli & "' AND idEmpresa = '" & glsEmpresa & "'"
            rst.Open CSqlC, Cn, adOpenStatic, adLockReadOnly
    
    mostrarDatosFormSQL Me, rst, StrMsgError
    If StrMsgError <> "" Then GoTo Err
    
    mostrarDatosPersona
    
    '-- TRAE EL LISTADO DE TIENDAS Y LO ALMACENA EN UN RECORSET
    If rst.State = 1 Then rst.Close
    CSqlC = "SELECT p.idPais,p.iddistrito, u.idDpto, u.idProv,p.idtdacli,p.item, p.GlsNombre, p.GlsDireccion, p.GlsTelefonos, p.GlsContacto, p.idVendedor, pp.GlsPersona As GlsVendedor " & _
            "FROM TiendasCliente p " & _
            "Left Join Ubigeo u " & _
                "On p.iddistrito = u.iddistrito and p.idPais = u.idPais " & _
            "Left Join  Personas PP " & _
                "On p.idVendedor = PP.idPersona " & _
            "WHERE p.idEmpresa = '" & glsEmpresa & "' " & _
            "AND p.idPersona = '" & strCodCli & "'"
            
    rst.Open CSqlC, Cn, adOpenKeyset, adLockOptimistic
    rsg.Fields.Append "Item", adInteger, , adFldRowID
    rsg.Fields.Append "GlsNombre", adVarChar, 120, adFldIsNullable
    rsg.Fields.Append "GlsDireccion", adVarChar, 255, adFldIsNullable
    rsg.Fields.Append "GlsTelefonos", adVarChar, 85, adFldIsNullable
    rsg.Fields.Append "GlsContacto", adVarChar, 120, adFldIsNullable
    rsg.Fields.Append "idtdacli", adVarChar, 120, adFldIsNullable
    rsg.Fields.Append "idPais", adVarChar, 85, adFldIsNullable
    rsg.Fields.Append "iddpto", adVarChar, 120, adFldIsNullable
    rsg.Fields.Append "idprov", adVarChar, 120, adFldIsNullable
    rsg.Fields.Append "idDistrito", adVarChar, 120, adFldIsNullable
    rsg.Fields.Append "idVendedor", adVarChar, 8, adFldIsNullable
    rsg.Fields.Append "GlsVendedor", adVarChar, 120, adFldIsNullable
    
    rsg.Open
    
    If rst.RecordCount = 0 Then
        rsg.AddNew
        rsg.Fields("Item") = 1
        rsg.Fields("GlsNombre") = ""
        rsg.Fields("GlsDireccion") = ""
        rsg.Fields("GlsTelefonos") = ""
        rsg.Fields("GlsContacto") = ""
        rsg.Fields("idtdacli") = ""
        rsg.Fields("idPais") = ""
        rsg.Fields("iddpto") = ""
        rsg.Fields("idprov") = ""
        rsg.Fields("idDistrito") = ""
        rsg.Fields("idVendedor") = ""
        rsg.Fields("GlsVendedor") = ""
        
    Else
        cont = 0
        Do While Not rst.EOF
            rsg.AddNew
            cont = cont + 1
            rsg.Fields("Item") = cont
            rsg.Fields("GlsNombre") = "" & rst.Fields("GlsNombre")
            rsg.Fields("GlsDireccion") = "" & rst.Fields("GlsDireccion")
            rsg.Fields("GlsTelefonos") = "" & rst.Fields("GlsTelefonos")
            rsg.Fields("GlsContacto") = "" & rst.Fields("GlsContacto")
            rsg.Fields("idtdacli") = Trim("" & rst.Fields("idtdacli"))
            rsg.Fields("idPais") = Trim("" & rst.Fields("idPais"))
            rsg.Fields("iddpto") = Trim("" & rst.Fields("idDpto"))
            rsg.Fields("idprov") = Trim("" & rst.Fields("idProv"))
            rsg.Fields("idDistrito") = Trim("" & rst.Fields("iddistrito"))
            rsg.Fields("idVendedor") = Trim("" & rst.Fields("idVendedor"))
            rsg.Fields("GlsVendedor") = Trim("" & rst.Fields("GlsVendedor"))
            rst.MoveNext
        Loop
    End If
    rst.Close: Set rst = Nothing
    
    mostrarDatosGridSQL gTiendas, rsg, StrMsgError
    If StrMsgError <> "" Then GoTo Err
    
    '--- TRAE EL LISTADO DE CONTACTOS
    CSqlC = "Select P.IdPersona,P.GlsPersona,P.Telefonos,P.Mail " & _
            "From ContactosClientes C " & _
            "Inner Join Personas P " & _
            "On C.IdContacto = P.IdPersona " & _
            "Where C.IdEmpresa = '" & glsEmpresa & "' And C.IdCliente = '" & txtCod_Persona.Text & "' " & _
            "Order By P.GlsPersona"
    rst.Open CSqlC, Cn, adOpenKeyset, adLockOptimistic
    RsC.Fields.Append "Item", adInteger, , adFldRowID
    RsC.Fields.Append "IdPersona", adVarChar, 8, adFldIsNullable
    RsC.Fields.Append "GlsPersona", adVarChar, 250, adFldIsNullable
    RsC.Fields.Append "Telefonos", adVarChar, 85, adFldIsNullable
    RsC.Fields.Append "Mail", adVarChar, 85, adFldIsNullable
    RsC.Fields.Append "IdPersonaAux", adVarChar, 8, adFldIsNullable
    RsC.Open
    
    If rst.RecordCount = 0 Then
        RsC.AddNew
        RsC.Fields("Item") = 1
        RsC.Fields("IdPersona") = ""
        RsC.Fields("GlsPersona") = ""
        RsC.Fields("Telefonos") = ""
        RsC.Fields("Mail") = ""
        RsC.Fields("IdPersonaAux") = ""
        
    Else
        item = 0
        Do While Not rst.EOF
            item = item + 1
            RsC.AddNew
            RsC.Fields("Item") = item
            RsC.Fields("IdPersona") = rst.Fields("IdPersona") & ""
            RsC.Fields("GlsPersona") = rst.Fields("GlsPersona") & ""
            RsC.Fields("Telefonos") = rst.Fields("Telefonos") & ""
            RsC.Fields("Mail") = rst.Fields("Mail") & ""
            RsC.Fields("IdPersonaAux") = rst.Fields("IdPersona") & ""
            rst.MoveNext
        Loop
    End If
    rst.Close: Set rst = Nothing
    
    mostrarDatosGridSQL gContactos, RsC, StrMsgError
    If StrMsgError <> "" Then GoTo Err
    
    TxtCodMoneda.Text = Trim("" & traerCampo("ControlLineaCredito", "IdMoneda", "IdCliente", Trim("" & txtCod_Persona.Text), True))
    
    CargaLineaCredito StrMsgError, txtCod_Persona.Text
    If StrMsgError <> "" Then GoTo Err
    
    '--- LISTA FORMA DE PAGO
    CSqlC = "Select p.GlsPersona, c.idFormaPago, C.indEstado, C.FecRegistro, C.idUsuario " & _
             "from clientesformapagos c, personas p " & _
             "where c.idcliente = p.idPersona " & _
             "and c.idcliente = '" & txtCod_Persona.Text & "' " & _
             "and c.idEmpresa = '" & glsEmpresa & "' " & _
             "order by c.FecRegistro desc "
    rst.Open CSqlC, Cn, adOpenKeyset, adLockOptimistic
    rsF.Fields.Append "Item", adInteger, , adFldRowID
    rsF.Fields.Append "IdCliente", adVarChar, 8, adFldIsNullable
    rsF.Fields.Append "glsCliente", adVarChar, 255, adFldIsNullable
    rsF.Fields.Append "idFormaPago", adVarChar, 8, adFldIsNullable
    rsF.Fields.Append "GlsFormaPago", adVarChar, 100, adFldIsNullable
    rsF.Fields.Append "indEstado", adInteger, 4, adFldIsNullable
    rsF.Fields.Append "idUsuario", adVarChar, 8, adFldIsNullable
    rsF.Fields.Append "FecRegistro", adVarChar, 40, adFldIsNullable
    rsF.Open
        
    If rst.RecordCount = 0 Then
        rsF.AddNew
        rsF.Fields("Item") = 1
        rsF.Fields("IdCliente") = ""
        rsF.Fields("idFormaPago") = ""
        rsF.Fields("GlsCliente") = ""
        rsF.Fields("GlsFormaPago") = ""
        rsF.Fields("indEstado") = 1
        rsF.Fields("idUsuario") = ""
        rsF.Fields("FecRegistro") = ""

    Else
        item = 0
        Do While Not rst.EOF
            item = item + 1
            rsF.AddNew
            rsF.Fields("Item") = item
            rsF.Fields("IdCliente") = Trim("" & txtCod_Persona.Text)
            rsF.Fields("idFormaPago") = Trim("" & rst.Fields("idFormaPago"))
            rsF.Fields("GlsCliente") = Trim("" & rst.Fields("GlsPersona"))
            rsF.Fields("GlsFormaPago") = Trim("" & traerCampo("formaspagos", "GlsFormaPago", "idFormaPago", Trim("" & rst.Fields("idFormaPago")), True))
            rsF.Fields("indEstado") = Val(rst.Fields("indEstado"))
            rsF.Fields("idUsuario") = Trim("" & rst.Fields("idUsuario"))
            rsF.Fields("FecRegistro") = Format(Trim("" & rst.Fields("FecRegistro")), "dd/mm/yyyy") & " " & Format(Format(Trim("" & rst.Fields("FecRegistro")), "hh:mm:ss "), "Medium Time")
            rst.MoveNext
        Loop
    End If
    rst.Close: Set rst = Nothing
    
    mostrarDatosGridSQL gFormaPago, rsF, StrMsgError
    If StrMsgError <> "" Then GoTo Err
    
    '--- LISTA PRODUCTOS
    CSqlC = "Select A.IdEmpresa,A.IdProducto,B.GlsProducto,A.Codigo,A.CodigoBarra " & _
            "From ProductosClientes A " & _
            "Inner Join Productos B " & _
                "On A.IdEmpresa = B.IdEmpresa And A.IdProducto = B.IdProducto " & _
            "Where A.IdEmpresa = '" & glsEmpresa & "' And A.IdCliente = '" & txtCod_Persona.Text & "' " & _
            "Order By A.IdProducto"
            
    rst.Open CSqlC, Cn, adOpenKeyset, adLockOptimistic
    rsp.Fields.Append "Item", adInteger, , adFldRowID
    rsp.Fields.Append "IdProducto", adVarChar, 8, adFldIsNullable
    rsp.Fields.Append "GlsProducto", adVarChar, 255, adFldIsNullable
    rsp.Fields.Append "Codigo", adVarChar, 100, adFldIsNullable
    rsp.Fields.Append "CodigoBarra", adVarChar, 100, adFldIsNullable
    rsp.Open
        
    If rst.RecordCount = 0 Then
        rsp.AddNew
        rsp.Fields("Item") = 1
        rsp.Fields("IdProducto") = ""
        rsp.Fields("GlsProducto") = ""
        rsp.Fields("Codigo") = ""
        rsp.Fields("CodigoBarra") = ""

    Else
        item = 0
        Do While Not rst.EOF
            item = item + 1
            rsp.AddNew
            rsp.Fields("Item") = item
            rsp.Fields("IdProducto") = Trim("" & rst.Fields("IdProducto"))
            rsp.Fields("GlsProducto") = Trim("" & rst.Fields("GlsProducto"))
            rsp.Fields("Codigo") = Trim("" & rst.Fields("Codigo"))
            rsp.Fields("CodigoBarra") = Trim("" & rst.Fields("CodigoBarra"))
            rst.MoveNext
        Loop
    End If
    rst.Close: Set rst = Nothing
    
    mostrarDatosGridSQL gProductos, rsp, StrMsgError
    If StrMsgError <> "" Then GoTo Err
    
    Me.Refresh
    
    Exit Sub
Err:
    If StrMsgError = "" Then StrMsgError = Err.Description
End Sub

Private Sub txtCod_VendedorCampo_Change()
    
    txtGls_VendedorCampo.Text = traerCampo("personas", "GlsPersona", "idPersona", txtCod_VendedorCampo.Text, False)

End Sub

Private Sub txtCod_VendedorCampo_KeyPress(KeyAscii As Integer)

    If KeyAscii <> 13 Then
        mostrarAyudaKeyascii KeyAscii, "VENDEDOR", txtCod_VendedorCampo, txtGls_VendedorCampo
        KeyAscii = 0
    End If

End Sub

Private Sub eliminar(ByRef StrMsgError As String)
On Error GoTo Err
Dim indTrans As Boolean
Dim strCodigo As String
Dim rsValida As New ADODB.Recordset

    If MsgBox("Está seguro(a) de eliminar el registro?" & vbCrLf & "Se eliminarán todas sus dependencias.", vbQuestion + vbYesNo, App.Title) = vbNo Then Exit Sub
    
    strCodigo = Trim(txtCod_Persona.Text)
    '--- VALIDANDO
    csql = "SELECT idDocVentas FROM docventas WHERE idPerCliente = '" & strCodigo & "' AND idEmpresa = '" & glsEmpresa & "'"
    rsValida.Open csql, Cn, adOpenForwardOnly, adLockReadOnly
    If Not rsValida.EOF Then
        StrMsgError = "No se puede eliminar el registro, el registro se encuentra en uso (Ventas)."
        GoTo Err
    End If
    
    csql = "SELECT idValesCab FROM valescab WHERE idProvCliente = '" & strCodigo & "' AND idEmpresa = '" & glsEmpresa & "'"
    If rsValida.State = 1 Then rsValida.Close
    rsValida.Open csql, Cn, adOpenForwardOnly, adLockReadOnly
    If Not rsValida.EOF Then
        StrMsgError = "No se puede eliminar el registro, el registro se encuentra en uso (Vales)."
        GoTo Err
    End If
    
    Cn.BeginTrans
    indTrans = True
    
    '--- Eliminando el registro
    csql = "DELETE FROM clientes WHERE idCliente = '" & strCodigo & "' AND idEmpresa = '" & glsEmpresa & "'"
    Cn.Execute csql
    
    Cn.CommitTrans
    
    '--- Nuevo
    Toolbar1_ButtonClick Toolbar1.Buttons(1)
    
    MsgBox "Registro eliminado satisfactoriamente", vbInformation, App.Title
    
    If rsValida.State = 1 Then rsValida.Close: Set rsValida = Nothing
    
    Exit Sub
    
Err:
    If rsValida.State = 1 Then rsValida.Close: Set rsValida = Nothing
    If indTrans Then Cn.RollbackTrans
    If StrMsgError = "" Then StrMsgError = Err.Description

End Sub

Private Sub CargaLineaCredito(StrMsgError As String, PIdCliente As String)
On Error GoTo Err
Dim CSqlC                               As String
Dim RsC                                 As New ADODB.Recordset
Dim CPC                                 As String
    
    'Linea Actual
    CSqlC = "Select A.Linea_Actual,A.Linea_Usada,A.Saldo,A.IndSuspension,B.GlsDato " & _
            "From ControlLineaCredito A " & _
            "Inner Join Datos B " & _
                "On A.Estado = B.IdDato " & _
            "Where A.IdEmpresa = '" & glsEmpresa & "' And A.IdCliente = '" & PIdCliente & "'"
    RsC.Open CSqlC, Cn, adOpenStatic, adLockReadOnly
    If Not RsC.EOF Then
        
        If Trim("" & RsC.Fields("IndSuspension")) = "1" Then
            
            TxtGlsEstado.Text = "Suspendido"
        
        Else
            
            TxtGlsEstado.Text = "" & RsC.Fields("GlsDato")
            
        End If
        
        TxtLineaAprobada.Text = Val("" & RsC.Fields("Linea_Actual"))
        TxtDeuda.Text = Val("" & RsC.Fields("Linea_Usada"))
        TxtSaldo.Text = Format(Val("" & RsC.Fields("Saldo")), "#,###,##0.00")
        
    Else
        
        TxtCodMoneda.Text = traerCampo("Parametros", "ValParametro", "GlsParametro", "MONEDA_LINEA_CREDITO", True)
        
        TxtLineaAprobada.Text = "0"
        TxtDeuda.Text = "0"
        TxtSaldo.Text = Format(Val("0"), "#,###,##0.00")
        
    End If
    
    RsC.Close: Set RsC = Nothing
    
    'Documentos
    If Len(Trim(CTmpDocumentos)) = 0 Then
        
        CPC = ComputerName
        CPC = Replace(CPC, "-", "")
        CPC = Trim(CPC)
        
        CTmpDocumentos = "TmpDocumentos" & Trim(Format(Now, "ddmmyyy") & Format(Now, "hhMMss") & glsUser & CPC)
        
    End If
    
    CSqlC = "Call Spu_VerificaMorosos('" & glsEmpresa & "','" & PIdCliente & "','1','" & CTmpDocumentos & "','','')"
    
    With GDocumentos
        .DefaultFields = False
        .Dataset.ADODataset.ConnectionString = strcn
        .Dataset.ADODataset.CursorLocation = clUseClient
        .Dataset.Active = False
        .Dataset.ADODataset.CommandText = CSqlC
        .Dataset.DisableControls
        .Dataset.Active = True
        .KeyField = "Nro_Comp"
    End With
    
    'Guias Por Facturar
    If Len(Trim(CTmpGuiasNF)) = 0 Then
        
        CPC = ComputerName
        CPC = Replace(CPC, "-", "")
        CPC = Trim(CPC)
        
        CTmpGuiasNF = "TmpGuiasNF" & Trim(Format(Now, "ddmmyyy") & Format(Now, "hhMMss") & glsUser & CPC)
        
    End If
    
    CSqlC = "Call Spu_VerificaMorosos('" & glsEmpresa & "','" & PIdCliente & "','2','','" & CTmpGuiasNF & "','')"
    
    With GGuiasNF
        .DefaultFields = False
        .Dataset.ADODataset.ConnectionString = strcn
        .Dataset.ADODataset.CursorLocation = clUseClient
        .Dataset.Active = False
        .Dataset.ADODataset.CommandText = CSqlC
        .Dataset.DisableControls
        .Dataset.Active = True
        .KeyField = "Nro_Comp"
    End With
    
    'Docuentos Generados
    If Len(Trim(CTmpDocumentosGen)) = 0 Then
        
        CPC = ComputerName
        CPC = Replace(CPC, "-", "")
        CPC = Trim(CPC)
        
        CTmpDocumentosGen = "TmpDocumentosGen" & Trim(Format(Now, "ddmmyyy") & Format(Now, "hhMMss") & glsUser & CPC)
        
    End If
    
    CSqlC = "Call Spu_VerificaMorosos('" & glsEmpresa & "','" & PIdCliente & "','3','','','" & CTmpDocumentosGen & "')"
    
    With GDocumentosGen
        .DefaultFields = False
        .Dataset.ADODataset.ConnectionString = strcn
        .Dataset.ADODataset.CursorLocation = clUseClient
        .Dataset.Active = False
        .Dataset.ADODataset.CommandText = CSqlC
        .Dataset.DisableControls
        .Dataset.Active = True
        .KeyField = "Nro_Comp"
    End With
    
    Me.Refresh
        
    Exit Sub
Err:
    If StrMsgError = "" Then StrMsgError = Err.Description
End Sub

Private Sub TxtCodMoneda_Change()
On Error GoTo Err
Dim StrMsgError
    
    TxtGlsMoneda.Text = traerCampo("Monedas", "GlsMoneda", "IdMoneda", TxtCodMoneda.Text, False)
    
    Exit Sub
Err:
    If StrMsgError = "" Then StrMsgError = Err.Description
    MsgBox StrMsgError, vbInformation, App.Title
End Sub

Private Sub GDocumentos_OnCustomDrawCell(ByVal hDC As Long, ByVal left As Single, ByVal top As Single, ByVal right As Single, ByVal bottom As Single, ByVal Node As DXDBGRIDLibCtl.IdxGridNode, ByVal Column As DXDBGRIDLibCtl.IdxGridColumn, ByVal Selected As Boolean, ByVal Focused As Boolean, ByVal NewItemRow As Boolean, Text As String, Color As Long, ByVal Font As stdole.IFontDisp, FontColor As Long, Alignment As DXDBGRIDLibCtl.ExAlignment, Done As Boolean)

    If Column.FieldName = "ValTotal" Or Column.FieldName = "SaldoPorCobrar" Then
        Text = Format(Text, "###,###,#0.00")
    End If

End Sub

Private Sub GDocumentos_OnCustomDrawFooter(ByVal hDC As Long, ByVal left As Single, ByVal top As Single, ByVal right As Single, ByVal bottom As Single, ByVal Node As DXDBGRIDLibCtl.IdxGridNode, ByVal Column As DXDBGRIDLibCtl.IdxGridColumn, Text As String, Color As Long, ByVal Font As stdole.IFontDisp, FontColor As Long, Alignment As DXDBGRIDLibCtl.ExAlignment, Done As Boolean)
    
    If Column.FieldName = "ValTotal" Or Column.FieldName = "SaldoPorCobrar" Then
        Text = Format(Text, "###,###,#0.00")
    End If
    
End Sub

Private Sub GDocumentos_OnCustomDrawFooterNode(ByVal hDC As Long, ByVal left As Single, ByVal top As Single, ByVal right As Single, ByVal bottom As Single, ByVal Node As DXDBGRIDLibCtl.IdxGridNode, ByVal Column As DXDBGRIDLibCtl.IdxGridColumn, ByVal FooterIndex As Integer, Text As String, Color As Long, ByVal Font As stdole.IFontDisp, FontColor As Long, Alignment As DXDBGRIDLibCtl.ExAlignment, Done As Boolean)
    
    If Column.FieldName = "ValTotal" Or Column.FieldName = "SaldoPorCobrar" Then
        Text = Format(Text, "###,###,#0.00")
    End If
    
End Sub

Private Sub GGuiasNF_OnCustomDrawCell(ByVal hDC As Long, ByVal left As Single, ByVal top As Single, ByVal right As Single, ByVal bottom As Single, ByVal Node As DXDBGRIDLibCtl.IdxGridNode, ByVal Column As DXDBGRIDLibCtl.IdxGridColumn, ByVal Selected As Boolean, ByVal Focused As Boolean, ByVal NewItemRow As Boolean, Text As String, Color As Long, ByVal Font As stdole.IFontDisp, FontColor As Long, Alignment As DXDBGRIDLibCtl.ExAlignment, Done As Boolean)

    If Column.FieldName = "ValTotal" Or Column.FieldName = "SaldoPorCobrar" Then
        Text = Format(Text, "###,###,#0.00")
    End If

End Sub

Private Sub GGuiasNF_OnCustomDrawFooter(ByVal hDC As Long, ByVal left As Single, ByVal top As Single, ByVal right As Single, ByVal bottom As Single, ByVal Node As DXDBGRIDLibCtl.IdxGridNode, ByVal Column As DXDBGRIDLibCtl.IdxGridColumn, Text As String, Color As Long, ByVal Font As stdole.IFontDisp, FontColor As Long, Alignment As DXDBGRIDLibCtl.ExAlignment, Done As Boolean)
    
    If Column.FieldName = "ValTotal" Or Column.FieldName = "SaldoPorCobrar" Then
        Text = Format(Text, "###,###,#0.00")
    End If
    
End Sub

Private Sub GGuiasNF_OnCustomDrawFooterNode(ByVal hDC As Long, ByVal left As Single, ByVal top As Single, ByVal right As Single, ByVal bottom As Single, ByVal Node As DXDBGRIDLibCtl.IdxGridNode, ByVal Column As DXDBGRIDLibCtl.IdxGridColumn, ByVal FooterIndex As Integer, Text As String, Color As Long, ByVal Font As stdole.IFontDisp, FontColor As Long, Alignment As DXDBGRIDLibCtl.ExAlignment, Done As Boolean)
    
    If Column.FieldName = "ValTotal" Or Column.FieldName = "SaldoPorCobrar" Then
        Text = Format(Text, "###,###,#0.00")
    End If
    
End Sub

Private Sub GDocumentosGen_OnCustomDrawCell(ByVal hDC As Long, ByVal left As Single, ByVal top As Single, ByVal right As Single, ByVal bottom As Single, ByVal Node As DXDBGRIDLibCtl.IdxGridNode, ByVal Column As DXDBGRIDLibCtl.IdxGridColumn, ByVal Selected As Boolean, ByVal Focused As Boolean, ByVal NewItemRow As Boolean, Text As String, Color As Long, ByVal Font As stdole.IFontDisp, FontColor As Long, Alignment As DXDBGRIDLibCtl.ExAlignment, Done As Boolean)

    If Column.FieldName = "ValTotal" Or Column.FieldName = "SaldoPorCobrar" Then
        Text = Format(Text, "###,###,#0.00")
    End If

End Sub

Private Sub GDocumentosGen_OnCustomDrawFooter(ByVal hDC As Long, ByVal left As Single, ByVal top As Single, ByVal right As Single, ByVal bottom As Single, ByVal Node As DXDBGRIDLibCtl.IdxGridNode, ByVal Column As DXDBGRIDLibCtl.IdxGridColumn, Text As String, Color As Long, ByVal Font As stdole.IFontDisp, FontColor As Long, Alignment As DXDBGRIDLibCtl.ExAlignment, Done As Boolean)
    
    If Column.FieldName = "ValTotal" Or Column.FieldName = "SaldoPorCobrar" Then
        Text = Format(Text, "###,###,#0.00")
    End If
    
End Sub

Private Sub GDocumentosGen_OnCustomDrawFooterNode(ByVal hDC As Long, ByVal left As Single, ByVal top As Single, ByVal right As Single, ByVal bottom As Single, ByVal Node As DXDBGRIDLibCtl.IdxGridNode, ByVal Column As DXDBGRIDLibCtl.IdxGridColumn, ByVal FooterIndex As Integer, Text As String, Color As Long, ByVal Font As stdole.IFontDisp, FontColor As Long, Alignment As DXDBGRIDLibCtl.ExAlignment, Done As Boolean)
    
    If Column.FieldName = "ValTotal" Or Column.FieldName = "SaldoPorCobrar" Then
        Text = Format(Text, "###,###,#0.00")
    End If
    
End Sub
