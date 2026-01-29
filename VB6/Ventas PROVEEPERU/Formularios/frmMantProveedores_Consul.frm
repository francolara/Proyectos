VERSION 5.00
Object = "{6A24B331-7634-11D3-A5B0-0050044A7E1A}#1.5#0"; "DXDBGrid.dll"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TabCtl32.Ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "mscomctl.OCX"
Object = "{F41D1D30-7878-4923-8CB3-6CCACDC9C9DE}#1.0#0"; "CATControls.ocx"
Begin VB.Form frmMantProveedores_Consul 
   Appearance      =   0  'Flat
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Mantenimiento de Proveedores"
   ClientHeight    =   9075
   ClientLeft      =   1755
   ClientTop       =   1065
   ClientWidth     =   12000
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   9075
   ScaleWidth      =   12000
   ShowInTaskbar   =   0   'False
   Begin MSComctlLib.ImageList imgDocVentas 
      Left            =   6360
      Top             =   0
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   12
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMantProveedores_Consul.frx":0000
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMantProveedores_Consul.frx":039A
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMantProveedores_Consul.frx":07EC
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMantProveedores_Consul.frx":0B86
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMantProveedores_Consul.frx":0F20
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMantProveedores_Consul.frx":12BA
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMantProveedores_Consul.frx":1654
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMantProveedores_Consul.frx":19EE
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMantProveedores_Consul.frx":1D88
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMantProveedores_Consul.frx":2122
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMantProveedores_Consul.frx":24BC
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMantProveedores_Consul.frx":317E
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   660
      Left            =   0
      TabIndex        =   9
      Top             =   0
      Width           =   12000
      _ExtentX        =   21167
      _ExtentY        =   1164
      ButtonWidth     =   1931
      ButtonHeight    =   1005
      Appearance      =   1
      ImageList       =   "imgDocVentas"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   8
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
            Description     =   "6"
            ImageIndex      =   6
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
            Caption         =   "Imprimir"
            Object.ToolTipText     =   "Imprimir"
            ImageIndex      =   11
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "       Lista       "
            Object.ToolTipText     =   "Lista"
            ImageIndex      =   12
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Salir"
            Object.ToolTipText     =   "Salir"
            ImageIndex      =   2
         EndProperty
      EndProperty
      BorderStyle     =   1
   End
   Begin VB.Frame fraListado 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   8355
      Left            =   90
      TabIndex        =   6
      Top             =   630
      Width           =   11820
      Begin VB.Frame Frame1 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   795
         Left            =   135
         TabIndex        =   7
         Top             =   120
         Width           =   11550
         Begin CATControls.CATTextBox txt_TextoBuscar 
            Height          =   315
            Left            =   1035
            TabIndex        =   0
            Top             =   300
            Width           =   10380
            _ExtentX        =   18309
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
            MaxLength       =   255
            Container       =   "frmMantProveedores_Consul.frx":3518
            Estilo          =   1
            Vacio           =   -1  'True
            EnterTab        =   -1  'True
         End
         Begin VB.Label Label3 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            Caption         =   "Búsqueda"
            ForeColor       =   &H80000008&
            Height          =   210
            Left            =   180
            TabIndex        =   8
            Top             =   345
            Width           =   735
         End
      End
      Begin DXDBGRIDLibCtl.dxDBGrid gLista 
         Height          =   7245
         Left            =   150
         OleObjectBlob   =   "frmMantProveedores_Consul.frx":3534
         TabIndex        =   1
         Top             =   990
         Width           =   11520
      End
   End
   Begin VB.Frame fraGeneral 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   8385
      Left            =   90
      TabIndex        =   10
      Top             =   630
      Width           =   11850
      Begin TabDlg.SSTab SSTab2 
         Height          =   7575
         Left            =   180
         TabIndex        =   11
         Top             =   450
         Width           =   11445
         _ExtentX        =   20188
         _ExtentY        =   13361
         _Version        =   393216
         Tabs            =   4
         TabsPerRow      =   4
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
         TabCaption(0)   =   "Datos Generales"
         TabPicture(0)   =   "frmMantProveedores_Consul.frx":5F12
         Tab(0).ControlEnabled=   -1  'True
         Tab(0).Control(0)=   "fraDatos"
         Tab(0).Control(0).Enabled=   0   'False
         Tab(0).ControlCount=   1
         TabCaption(1)   =   "Productos"
         TabPicture(1)   =   "frmMantProveedores_Consul.frx":5F2E
         Tab(1).ControlEnabled=   0   'False
         Tab(1).Control(0)=   "Frame2"
         Tab(1).ControlCount=   1
         TabCaption(2)   =   "Contactos"
         TabPicture(2)   =   "frmMantProveedores_Consul.frx":5F4A
         Tab(2).ControlEnabled=   0   'False
         Tab(2).Control(0)=   "Frame3"
         Tab(2).ControlCount=   1
         TabCaption(3)   =   "Observaciones"
         TabPicture(3)   =   "frmMantProveedores_Consul.frx":5F66
         Tab(3).ControlEnabled=   0   'False
         Tab(3).Control(0)=   "FraObs"
         Tab(3).ControlCount=   1
         Begin VB.Frame fraDatos 
            Appearance      =   0  'Flat
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
            ForeColor       =   &H80000008&
            Height          =   5625
            Left            =   405
            TabIndex        =   15
            Top             =   945
            Width           =   10665
            Begin VB.CheckBox chkretencion 
               Caption         =   "Agente de Retención"
               Height          =   330
               Left            =   495
               TabIndex        =   40
               Top             =   3555
               Width           =   1950
            End
            Begin VB.CheckBox chkpercepcion 
               Caption         =   "Agente de Percepción"
               Height          =   330
               Left            =   495
               TabIndex        =   39
               Top             =   3870
               Width           =   1950
            End
            Begin VB.CheckBox chkbuenc 
               Caption         =   "Buen Contribuyente"
               Height          =   330
               Left            =   495
               TabIndex        =   38
               Top             =   4185
               Width           =   1950
            End
            Begin VB.CommandButton cmbAyudaPersona 
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
               Height          =   315
               Left            =   9990
               Picture         =   "frmMantProveedores_Consul.frx":5F82
               Style           =   1  'Graphical
               TabIndex        =   16
               Top             =   390
               Width           =   390
            End
            Begin CATControls.CATTextBox txtCod_Persona 
               Height          =   315
               Left            =   2000
               TabIndex        =   2
               Tag             =   "TidProveedor"
               Top             =   405
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
               Container       =   "frmMantProveedores_Consul.frx":630C
               Estilo          =   1
               EnterTab        =   -1  'True
            End
            Begin CATControls.CATTextBox txtGls_Persona 
               Height          =   315
               Left            =   2970
               TabIndex        =   17
               Top             =   405
               Width           =   6970
               _ExtentX        =   12303
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
               Container       =   "frmMantProveedores_Consul.frx":6328
            End
            Begin CATControls.CATTextBox txtGls_Direccion 
               Height          =   315
               Left            =   1995
               TabIndex        =   18
               Top             =   2655
               Width           =   7965
               _ExtentX        =   14049
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
               Container       =   "frmMantProveedores_Consul.frx":6344
               Estilo          =   1
               Vacio           =   -1  'True
               EnterTab        =   -1  'True
            End
            Begin CATControls.CATTextBox txt_RUC 
               Height          =   315
               Left            =   1995
               TabIndex        =   19
               Top             =   3030
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
               Container       =   "frmMantProveedores_Consul.frx":6360
               Estilo          =   1
               Vacio           =   -1  'True
               EnterTab        =   -1  'True
            End
            Begin CATControls.CATTextBox txtCod_Pais 
               Height          =   315
               Left            =   2000
               TabIndex        =   20
               Top             =   1155
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
               Container       =   "frmMantProveedores_Consul.frx":637C
               Estilo          =   1
               Vacio           =   -1  'True
               EnterTab        =   -1  'True
            End
            Begin CATControls.CATTextBox txtGls_Pais 
               Height          =   315
               Left            =   2970
               TabIndex        =   21
               Top             =   1155
               Width           =   6970
               _ExtentX        =   12303
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
               Container       =   "frmMantProveedores_Consul.frx":6398
               Vacio           =   -1  'True
            End
            Begin CATControls.CATTextBox txtCod_Depa 
               Height          =   315
               Left            =   2000
               TabIndex        =   22
               Top             =   1530
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
               Container       =   "frmMantProveedores_Consul.frx":63B4
               Estilo          =   1
               Vacio           =   -1  'True
               EnterTab        =   -1  'True
            End
            Begin CATControls.CATTextBox txtGls_Depa 
               Height          =   315
               Left            =   2970
               TabIndex        =   23
               Top             =   1530
               Width           =   6970
               _ExtentX        =   12303
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
               Container       =   "frmMantProveedores_Consul.frx":63D0
               Vacio           =   -1  'True
            End
            Begin CATControls.CATTextBox txtCod_Prov 
               Height          =   315
               Left            =   2000
               TabIndex        =   24
               Top             =   1905
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
               Container       =   "frmMantProveedores_Consul.frx":63EC
               Estilo          =   1
               Vacio           =   -1  'True
               EnterTab        =   -1  'True
            End
            Begin CATControls.CATTextBox txtGls_Prov 
               Height          =   315
               Left            =   2970
               TabIndex        =   25
               Top             =   1905
               Width           =   6970
               _ExtentX        =   12303
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
               Container       =   "frmMantProveedores_Consul.frx":6408
               Vacio           =   -1  'True
            End
            Begin CATControls.CATTextBox txtCod_Distrito 
               Height          =   315
               Left            =   1980
               TabIndex        =   26
               Top             =   2280
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
               Container       =   "frmMantProveedores_Consul.frx":6424
               Estilo          =   1
               Vacio           =   -1  'True
               EnterTab        =   -1  'True
            End
            Begin CATControls.CATTextBox txtGls_Distrito 
               Height          =   315
               Left            =   2970
               TabIndex        =   27
               Top             =   2280
               Width           =   6970
               _ExtentX        =   12303
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
               Container       =   "frmMantProveedores_Consul.frx":6440
               Vacio           =   -1  'True
            End
            Begin CATControls.CATTextBox txtCod_TipoPersona 
               Height          =   315
               Left            =   2000
               TabIndex        =   28
               Top             =   780
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
               Container       =   "frmMantProveedores_Consul.frx":645C
               Estilo          =   1
               Vacio           =   -1  'True
               EnterTab        =   -1  'True
            End
            Begin CATControls.CATTextBox txtGls_TipoPersona 
               Height          =   315
               Left            =   2970
               TabIndex        =   29
               Top             =   780
               Width           =   6970
               _ExtentX        =   12303
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
               Container       =   "frmMantProveedores_Consul.frx":6478
               Vacio           =   -1  'True
            End
            Begin CATControls.CATTextBox txtagenteret 
               Height          =   315
               Left            =   2940
               TabIndex        =   41
               Tag             =   "NIndRetencion"
               Top             =   3555
               Visible         =   0   'False
               Width           =   330
               _ExtentX        =   582
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
               Container       =   "frmMantProveedores_Consul.frx":6494
               Estilo          =   1
               EnterTab        =   -1  'True
            End
            Begin CATControls.CATTextBox txtagenteper 
               Height          =   315
               Left            =   2940
               TabIndex        =   42
               Tag             =   "NIndPercepcion"
               Top             =   3870
               Visible         =   0   'False
               Width           =   330
               _ExtentX        =   582
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
               Container       =   "frmMantProveedores_Consul.frx":64B0
               Estilo          =   1
               EnterTab        =   -1  'True
            End
            Begin CATControls.CATTextBox txtbuenc 
               Height          =   315
               Left            =   2940
               TabIndex        =   43
               Tag             =   "NIndBnContribuyente"
               Top             =   4185
               Visible         =   0   'False
               Width           =   330
               _ExtentX        =   582
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
               Container       =   "frmMantProveedores_Consul.frx":64CC
               Estilo          =   1
               EnterTab        =   -1  'True
            End
            Begin VB.Label Label4 
               Appearance      =   0  'Flat
               AutoSize        =   -1  'True
               Caption         =   "Tipo de Entidad"
               ForeColor       =   &H80000008&
               Height          =   210
               Left            =   500
               TabIndex        =   37
               Top             =   795
               Width           =   1095
            End
            Begin VB.Label Label14 
               Appearance      =   0  'Flat
               AutoSize        =   -1  'True
               Caption         =   "R.U.C."
               ForeColor       =   &H80000008&
               Height          =   210
               Left            =   495
               TabIndex        =   36
               Top             =   3105
               Width           =   450
            End
            Begin VB.Label Label10 
               Appearance      =   0  'Flat
               AutoSize        =   -1  'True
               Caption         =   "Dirección"
               ForeColor       =   &H80000008&
               Height          =   210
               Left            =   495
               TabIndex        =   35
               Top             =   2700
               Width           =   675
            End
            Begin VB.Label Label8 
               Appearance      =   0  'Flat
               AutoSize        =   -1  'True
               Caption         =   "Provincia"
               ForeColor       =   &H80000008&
               Height          =   210
               Left            =   500
               TabIndex        =   34
               Top             =   1920
               Width           =   660
            End
            Begin VB.Label Label7 
               Appearance      =   0  'Flat
               AutoSize        =   -1  'True
               Caption         =   "Departamento"
               ForeColor       =   &H80000008&
               Height          =   210
               Left            =   500
               TabIndex        =   33
               Top             =   1545
               Width           =   1005
            End
            Begin VB.Label Label5 
               Appearance      =   0  'Flat
               AutoSize        =   -1  'True
               Caption         =   "País"
               ForeColor       =   &H80000008&
               Height          =   210
               Left            =   500
               TabIndex        =   32
               Top             =   1170
               Width           =   300
            End
            Begin VB.Label Label2 
               Appearance      =   0  'Flat
               AutoSize        =   -1  'True
               Caption         =   "Distrito"
               ForeColor       =   &H80000008&
               Height          =   210
               Left            =   500
               TabIndex        =   31
               Top             =   2325
               Width           =   495
            End
            Begin VB.Label Label1 
               Appearance      =   0  'Flat
               AutoSize        =   -1  'True
               Caption         =   "Entidad"
               ForeColor       =   &H80000008&
               Height          =   210
               Left            =   500
               TabIndex        =   30
               Top             =   420
               Width           =   525
            End
         End
         Begin VB.Frame FraObs 
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   3225
            Left            =   -74730
            TabIndex        =   14
            Top             =   945
            Width           =   10830
            Begin VB.TextBox txtGls_Obs 
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   2775
               Left            =   180
               MultiLine       =   -1  'True
               ScrollBars      =   2  'Vertical
               TabIndex        =   5
               Tag             =   "TGlsObservacion"
               Top             =   300
               Width           =   10545
            End
         End
         Begin VB.Frame Frame3 
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   5700
            Left            =   -74730
            TabIndex        =   13
            Top             =   945
            Width           =   10830
            Begin DXDBGRIDLibCtl.dxDBGrid gContactos 
               Height          =   5340
               Left            =   135
               OleObjectBlob   =   "frmMantProveedores_Consul.frx":64E8
               TabIndex        =   4
               Top             =   225
               Width           =   10560
            End
         End
         Begin VB.Frame Frame2 
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   6510
            Left            =   -74820
            TabIndex        =   12
            Top             =   675
            Width           =   11010
            Begin DXDBGRIDLibCtl.dxDBGrid gProductos 
               Height          =   6120
               Left            =   135
               OleObjectBlob   =   "frmMantProveedores_Consul.frx":812D
               TabIndex        =   3
               Top             =   225
               Width           =   10740
            End
         End
      End
   End
End
Attribute VB_Name = "frmMantProveedores_Consul"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim indBoton As Integer

Private Sub cmbAyudaPersona_Click()

    mostrarAyuda "PERSONAPROVEEDOR", txtCod_Persona, txtGls_Persona
    If txtCod_Persona.Text <> "" Then mostrarDatosPersona

End Sub

Private Sub mostrarDatosPersona()
Dim rst As New ADODB.Recordset

    csql = "SELECT idPersona,GlsPersona," & _
            "tipoPersona , ruc, direccion, p.iddistrito, u.idDpto, u.idProv,p.idPais " & _
            "FROM personas p,ubigeo u " & _
            "WHERE p.iddistrito = u.iddistrito and p.idPais = u.idPais  and idpersona = '" & Trim(txtCod_Persona.Text) & "'"
               
    rst.Open csql, Cn, adOpenStatic, adLockReadOnly
    If Not rst.EOF Then
        txtCod_TipoPersona.Text = "" & rst.Fields("tipoPersona")
        txtCod_Pais.Text = "" & rst.Fields("idPais")
        txtCod_Depa.Text = "" & rst.Fields("idDpto")
        txtCod_Prov.Text = "" & rst.Fields("idProv")
        txtCod_Distrito.Text = "" & rst.Fields("iddistrito")
        txtGls_Direccion.Text = "" & rst.Fields("direccion")
        txt_RUC.Text = "" & rst.Fields("ruc")
    
    Else
        txtCod_TipoPersona.Text = ""
        txtCod_Pais.Text = ""
        txtCod_Depa.Text = ""
        txtCod_Prov.Text = ""
        txtCod_Distrito.Text = ""
        txtGls_Direccion.Text = ""
        txt_RUC.Text = ""
    End If
    rst.Close: Set rst = Nothing
    
End Sub

Private Sub Form_Load()
On Error GoTo Err
Dim StrMsgError As String

    ConfGrid gLista, False, False, False, False
    ConfGrid gProductos, True, False, False, False
    ConfGrid gContactos, True, False, False, False
    fraListado.Visible = True
    fraGeneral.Visible = False
    habilitaBotones 7
    nuevo

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
        If gContactos.Columns.ColumnByFieldName("idPersona").Value = "" Then
            Allow = False
        Else
            gContactos.Columns.FocusedIndex = gContactos.Columns.ColumnByFieldName("idPersona").Index
        End If
    End If

End Sub

Private Sub gContactos_OnEditButtonClick(ByVal Column As DXDBGRIDLibCtl.IdxGridColumn, ByVal Node As DXDBGRIDLibCtl.IdxGridNode)
Dim A As String
Dim B As String
    
    Select Case Column.Index
        Case gContactos.Columns.ColumnByFieldName("idPersona").Index
            A = txtCod_Persona.Text
            B = txtGls_Persona.Text
            mostrarAyuda "PERSONAUSUARIO", txtCod_Persona, txtGls_Persona
            If txtCod_Persona.Text <> "" Then mostrarDatosPersona_prov
            txtCod_Persona.Text = A
            txtGls_Persona.Text = B
    End Select

End Sub

Private Sub mostrarDatosPersona_prov()
Dim rst As New ADODB.Recordset

    csql = "SELECT idPersona, GlsPersona FROM PERSONAS WHERE IDPERSONA= '" & Trim(txtCod_Persona.Text) & "'"
    rst.Open csql, Cn, adOpenStatic, adLockReadOnly
    If Not rst.EOF Then
        gContactos.Dataset.Edit
        gContactos.Columns.ColumnByFieldName("idPersona").Value = "" & rst.Fields("idpersona")
        gContactos.Columns.ColumnByFieldName("glspersona").Value = "" & rst.Fields("glspersona")
    Else
        gContactos.Columns.ColumnByFieldName("idPersona").Value = ""
        gContactos.Columns.ColumnByFieldName("glspersona").Value = ""
    End If
    rst.Close: Set rst = Nothing

End Sub

Private Sub gLista_OnDblClick()
On Error GoTo Err
Dim StrMsgError As String

    mostrarProveedor gLista.Columns.ColumnByName("idProveedor").Value, StrMsgError
    If StrMsgError <> "" Then GoTo Err
    fraListado.Visible = False
    fraGeneral.Visible = True
    fraGeneral.Enabled = False
    txtCod_Persona.Enabled = False
    cmbAyudaPersona.Enabled = False
    habilitaBotones 2
    fraGeneral.Enabled = True
    
Exit Sub
Err:
    MsgBox StrMsgError, vbInformation, App.Title
End Sub

Private Sub gProductos_OnEdited(ByVal Node As DXDBGRIDLibCtl.IdxGridNode)
On Error GoTo Err
Dim StrMsgError As String
    
    If gProductos.Dataset.Modified = False Then Exit Sub
    Select Case gProductos.Columns.FocusedColumn.Index
        Case gProductos.Columns.ColumnByFieldName("ValorVenta").Index
            gProductos.Dataset.Edit
            gProductos.Columns.ColumnByFieldName("IGVVenta").Value = Val(Format(gProductos.Columns.ColumnByFieldName("ValorVenta").Value * glsIGV, "0.00"))
            gProductos.Columns.ColumnByFieldName("PrecioVenta").Value = Val(Format(gProductos.Columns.ColumnByFieldName("ValorVenta").Value * (glsIGV + 1), "0.00"))
            gProductos.Dataset.Post
        
        Case gProductos.Columns.ColumnByFieldName("PrecioVenta").Index
            gProductos.Dataset.Edit
            gProductos.Columns.ColumnByFieldName("ValorVenta").Value = Val(Format(gProductos.Columns.ColumnByFieldName("ValorVenta").Value / (glsIGV + 1), "0.00"))
            gProductos.Columns.ColumnByFieldName("IGVVenta").Value = Val(Format(Val(Format(gProductos.Columns.ColumnByFieldName("PrecioVenta").Value, "0.00")) - Val(Format(gProductos.Columns.ColumnByFieldName("ValorVenta").Value, "0.00")), "0.00"))
            gProductos.Dataset.Post
    End Select

    Exit Sub
    
Err:
    If StrMsgError = "" Then StrMsgError = Err.Description
    MsgBox StrMsgError, vbInformation, App.Title
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
On Error GoTo Err
Dim StrMsgError As String

    Select Case Button.Index
        Case 1 '--- Nuevo
            indBoton = 0
            nuevo
            fraListado.Visible = False
            fraGeneral.Visible = True
            fraGeneral.Enabled = True
            
        Case 2 '--- Grabar
            Grabar StrMsgError
            If StrMsgError <> "" Then GoTo Err
        
        Case 3 '--- Modificar
            indBoton = 1
            fraGeneral.Enabled = True
        
        Case 4, 7 '--- Cancelar
            fraListado.Visible = True
            fraGeneral.Visible = False
            fraGeneral.Enabled = False
        
        Case 5 '--- Eliminar
            eliminar StrMsgError
            If StrMsgError <> "" Then GoTo Err
        
        Case 6 'Imprimir
        
        Case 8 'Salir
            Unload Me
    End Select
    habilitaBotones Button.Index

    Exit Sub
Err:
    MsgBox StrMsgError, vbInformation, App.Title
End Sub

Private Sub txt_TextoBuscar_Change()
On Error GoTo Err
Dim StrMsgError As String
    
    If glsEnterAyudaClientes = False Then
        listaProveedor StrMsgError
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
        listaProveedor StrMsgError
        If StrMsgError <> "" Then GoTo Err
    End If

    Exit Sub
    
Err:
    If StrMsgError = "" Then StrMsgError = Err.Description
    MsgBox StrMsgError, vbInformation, App.Title
End Sub

Private Sub txtCod_Persona_Change()
    
    txtGls_Persona.Text = traerCampo("personas", "GlsPersona", "idPersona", txtCod_Persona.Text, False)

End Sub

Private Sub txtCod_Persona_KeyPress(KeyAscii As Integer)
    
    If KeyAscii <> 13 Then
        mostrarAyudaKeyascii KeyAscii, "PERSONAPROVEEDOR", txtCod_Persona, txtGls_Persona
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
    
    txtGls_Prov.Text = traerCampo("Ubigeo", "GlsUbigeo", "idProv", txtCod_Prov.Text, False, " idDpto = '" & txtCod_Depa.Text & "' and idProv <> '00' and idDist = '00' And idPais = '" & txtCod_Pais.Text & "' ")

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
            'Toolbar1.Buttons(5).Visible = indHabilitar 'Eliminar
            'Toolbar1.Buttons(6).Visible = indHabilitar 'Imprimir
            Toolbar1.Buttons(7).Visible = True 'Lista
        Case 4, 7 'Cancelar, Lista
            'Toolbar1.Buttons(1).Visible = True
            'Toolbar1.Buttons(2).Visible = False
            'Toolbar1.Buttons(3).Visible = False
            'Toolbar1.Buttons(4).Visible = False
            'Toolbar1.Buttons(5).Visible = False
            'Toolbar1.Buttons(6).Visible = False
            Toolbar1.Buttons(7).Visible = True
    End Select

End Sub

Private Sub Grabar(ByRef StrMsgError As String)
On Error GoTo Err
Dim strCodigo As String
Dim strMsg      As String

    glsobservacioncliente = txtGls_Obs.Text

    validaFormSQL Me, StrMsgError
    If StrMsgError <> "" Then GoTo Err
    
    If indBoton = 0 Then  '--- graba
        EjecutaSQLForm Me, 0, True, "proveedores", StrMsgError, "", , , "idProveedor", txtCod_Persona.Text
        If StrMsgError <> "" Then GoTo Err
        strMsg = "Grabó"
    
    Else '--- modifica
        EjecutaSQLForm Me, 1, True, "proveedores", StrMsgError, "idProveedor", , , "idProveedor", txtCod_Persona.Text
        If StrMsgError <> "" Then GoTo Err
        strMsg = "Modificó"
    End If
    
    If gContactos.Count > 0 Then
        csql = "delete from contactosproveedores " & _
                "where  idProveedor='" & txtCod_Persona.Text & "' and idempresa='" & glsEmpresa & "'"
        Cn.Execute csql
        gContactos.Dataset.First
        Do While Not gContactos.Dataset.EOF
            csql = "insert into contactosproveedores(idProveedor, idContacto, idEmpresa) " & _
                    "values ('" & txtCod_Persona.Text & "','" & gContactos.Columns.ColumnByFieldName("idpErsona").Value & "','" & glsEmpresa & "')"
            Cn.Execute (csql)
            gContactos.Dataset.Next
        Loop
    End If
    MsgBox "Se " & strMsg & " Satisfactoriamente", vbInformation, App.Title
    fraGeneral.Enabled = False
    listaProveedor StrMsgError
    If StrMsgError <> "" Then GoTo Err

    Exit Sub
    
Err:
    If StrMsgError = "" Then StrMsgError = Err.Description
End Sub

Private Sub nuevo()
On Error GoTo Err
Dim StrMsgError As String
Dim rst As New ADODB.Recordset
Dim RsC As New ADODB.Recordset

    limpiaForm Me
    txtCod_Persona.Enabled = True
    cmbAyudaPersona.Enabled = True
    
     '--- PRODUCTOS
    rst.Fields.Append "Item", adInteger, , adFldRowID
    rst.Fields.Append "idProducto", adVarChar, 8, adFldIsNullable
    rst.Fields.Append "GlsProducto", adVarChar, 250, adFldIsNullable
    rst.Open
    
    rst.AddNew
    rst.Fields("Item") = 1
    rst.Fields("idProducto") = ""
    rst.Fields("GlsProducto") = ""
    
    mostrarDatosGridSQL gProductos, rst, StrMsgError
    If StrMsgError <> "" Then GoTo Err
    gProductos.Columns.FocusedIndex = gProductos.Columns.ColumnByFieldName("idProducto").Index
    
    '--- CONTACTOS
    RsC.Fields.Append "Item", adInteger, , adFldRowID
    RsC.Fields.Append "idPersona", adVarChar, 8, adFldIsNullable
    RsC.Fields.Append "GlsPersona", adVarChar, 250, adFldIsNullable
    RsC.Open
    
    RsC.AddNew
    RsC.Fields("Item") = 1
    RsC.Fields("idPersona") = ""
    RsC.Fields("GlsPersona") = ""
    
    mostrarDatosGridSQL gContactos, RsC, StrMsgError
    If StrMsgError <> "" Then GoTo Err
    
    gContactos.Columns.FocusedIndex = gContactos.Columns.ColumnByFieldName("idPersona").Index
    
    Exit Sub
    
Err:
    If StrMsgError = "" Then StrMsgError = Err.Description
    MsgBox StrMsgError
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
Dim StrDes As String
    
    Select Case Column.Index
        Case gProductos.Columns.ColumnByFieldName("idProducto").Index
            strCod = gProductos.Columns.ColumnByFieldName("idProducto").Value
            StrDes = gProductos.Columns.ColumnByFieldName("GlsProducto").Value
            
            mostrarAyudaTexto "PRODUCTOS", strCod, StrDes
            
            If existeEnGrilla(gProductos, "idProducto", strCod) = False Then
                gProductos.Dataset.Edit
                gProductos.Columns.ColumnByFieldName("idProducto").Value = strCod
                gProductos.Columns.ColumnByFieldName("GlsProducto").Value = StrDes
                gProductos.Dataset.Post
            Else
                MsgBox "El Almacén ya fue ingresado.", vbInformation, App.Title
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
Dim StrDes As String

    Select Case gProductos.Columns.FocusedColumn.Index
        Case gProductos.Columns.ColumnByFieldName("idProducto").Index
            strCod = gProductos.Columns.ColumnByFieldName("idProducto").Value
            StrDes = gProductos.Columns.ColumnByFieldName("GlsProducto").Value
            
            mostrarAyudaKeyasciiTexto Key, "PRODUCTOS", strCod, StrDes
            Key = 0
            If existeEnGrilla(gProductos, "idProducto", strCod) = False Then
                gProductos.Dataset.Edit
                gProductos.Columns.ColumnByFieldName("idProducto").Value = strCod
                gProductos.Columns.ColumnByFieldName("GlsProducto").Value = StrDes
                gProductos.Dataset.Post
                gProductos.SetFocus
            Else
                MsgBox "El Producto ya fue ingresado.", vbInformation, App.Title
            End If
    End Select
    
End Sub

Private Sub listaProveedor(ByRef StrMsgError As String)
On Error GoTo Err
Dim strCond As String
    
    strCond = ""
    If Trim(txt_TextoBuscar.Text) <> "" Then
        strCond = Trim(txt_TextoBuscar.Text)
        strCond = " And (idProveedor LIKE '%" & strCond & "%'  OR  GlsPersona LIKE '%" & strCond & "%' OR GlsObservacion LIKE '%" & strCond & "%' OR ruc LIKE '%" & strCond & "%') "
    End If
    
    csql = "SELECT c.idProveedor ,p.GlsPersona ," & _
            "if(p.tipoPersona = '01001','Natural','Juridica') as TipoPersona,p.ruc,concat(p.direccion,', ',u.GlsUbigeo) as Direccion,GlsObservacion " & _
            "FROM proveedores c " & _
            "inner join personas p " & _
            "on c.idProveedor = p.idPersona " & _
            "left join ubigeo u " & _
            "on p.iddistrito = u.iddistrito  and p.idPais = u.idPais  " & _
            "WHERE c.idEmpresa = '" & glsEmpresa & "'" & strCond & _
            "ORDER BY c.idProveedor"
    With gLista
        .DefaultFields = False
        .Dataset.ADODataset.ConnectionString = strcn
        .Dataset.ADODataset.CursorLocation = clUseClient
        .Dataset.Active = False
        .Dataset.ADODataset.CommandText = csql
        .Dataset.DisableControls
        .Dataset.Active = True
        .KeyField = "idProveedor"
    End With
    Me.Refresh

    Exit Sub
    
Err:
    If StrMsgError = "" Then StrMsgError = Err.Description
End Sub

Private Sub mostrarProveedor(strCodProv As String, ByRef StrMsgError As String)
On Error GoTo Err
Dim rst As New ADODB.Recordset
Dim rsg As New ADODB.Recordset
Dim RsC As New ADODB.Recordset
Dim item As Integer

    csql = "SELECT c.idProveedor, c.GlsObservacion, c.indretencion, c.indpercepcion, c.IndBnContribuyente " & _
           "FROM Proveedores c " & _
           "WHERE c.idProveedor = '" & strCodProv & "' and idempresa = '" & glsEmpresa & "'"
    rst.Open csql, Cn, adOpenStatic, adLockReadOnly
    
    If Not rst.EOF Then
        If Val(rst.Fields("indretencion").Value & "") = 1 Then
            chkretencion.Value = 1
        Else
            chkretencion.Value = 0
        End If
        
        If Val(rst.Fields("indpercepcion").Value & "") = 1 Then
            chkpercepcion.Value = 1
        Else
            chkpercepcion.Value = 0
        End If
        
        If Val(rst.Fields("IndBnContribuyente").Value & "") = 1 Then
            chkbuenc.Value = 1
        Else
            chkbuenc.Value = 0
        End If
    End If
    
    mostrarDatosFormSQL Me, rst, StrMsgError
    If StrMsgError <> "" Then GoTo Err
    
    mostrarDatosPersona

    '--- TRAE EL LISTADO DE PRODUCTOS Y LO ALMACENA EN UN RECORSET
    csql = "SELECT p.item,p.idProducto,a.glsProducto,p.ValorVenta,p.IGVVenta,p.PrecioVenta " & _
            "FROM productosproveedor p, productos a " & _
            "WHERE p.idProducto = a.idProducto and p.idempresa = '" & glsEmpresa & "' " & _
            "AND p.idProveedor = '" & strCodProv & "' and p.idempresa = a.idempresa " & _
            "order by p.item"
    rst.Open csql, Cn, adOpenKeyset, adLockOptimistic
    
    rsg.Fields.Append "Item", adInteger, , adFldRowID
    rsg.Fields.Append "idProducto", adVarChar, 8, adFldIsNullable
    rsg.Fields.Append "GlsProducto", adVarChar, 250, adFldIsNullable
    rsg.Fields.Append "ValorVenta", adDouble, 14, adFldIsNullable
    rsg.Fields.Append "IGVVenta", adDouble, 14, adFldIsNullable
    rsg.Fields.Append "PrecioVenta", adDouble, 14, adFldIsNullable
    rsg.Open
    
    If rst.RecordCount = 0 Then
        rsg.AddNew
        rsg.Fields("Item") = 1
        rsg.Fields("idProducto") = ""
        rsg.Fields("GlsProducto") = ""
        rsg.Fields("ValorVenta") = 0
        rsg.Fields("IGVVenta") = 0
        rsg.Fields("PrecioVenta") = 0
    Else
        Do While Not rst.EOF
            rsg.AddNew
            rsg.Fields("Item") = Val("" & rst.Fields("Item"))
            rsg.Fields("idProducto") = rst.Fields("idProducto") & ""
            rsg.Fields("GlsProducto") = rst.Fields("GlsProducto") & ""
            rsg.Fields("ValorVenta") = Val(rst.Fields("ValorVenta") & "")
            rsg.Fields("IGVVenta") = Val(rst.Fields("IGVVenta") & "")
            rsg.Fields("PrecioVenta") = Val(rst.Fields("PrecioVenta") & "")
            rst.MoveNext
        Loop
    End If
    rst.Close: Set rst = Nothing
    
    mostrarDatosGridSQL gProductos, rsg, StrMsgError
    If StrMsgError <> "" Then GoTo Err
    
    '--- TRAE EL LISTADO DE CONTACTOS
    csql = "select p.idPersona, p.glsPersona " & _
             "from contactosproveedores c, personas p " & _
             "where c.idContacto = p.idPersona " & _
             "and c.idProveedor = '" & strCodProv & "' " & _
             "and c.idEmpresa = '" & glsEmpresa & "' " & _
             "order by p.glsPersona "
    rst.Open csql, Cn, adOpenKeyset, adLockOptimistic
    
    RsC.Fields.Append "Item", adInteger, , adFldRowID
    RsC.Fields.Append "idPersona", adVarChar, 8, adFldIsNullable
    RsC.Fields.Append "GlsPersona", adVarChar, 250, adFldIsNullable
    RsC.Open
    
    If rst.RecordCount = 0 Then
        RsC.AddNew
        RsC.Fields("Item") = 1
        RsC.Fields("idPersona") = ""
        RsC.Fields("GlsPersona") = ""
    Else
        item = 0
        Do While Not rst.EOF
            item = item + 1
            RsC.AddNew
            RsC.Fields("Item") = item
            RsC.Fields("idPersona") = rst.Fields("idPersona") & ""
            RsC.Fields("GlsPersona") = rst.Fields("GlsPersona") & ""
            rst.MoveNext
        Loop
    End If
    rst.Close: Set rst = Nothing
    
    mostrarDatosGridSQL gContactos, RsC, StrMsgError
    If StrMsgError <> "" Then GoTo Err
    Me.Refresh
    
    Exit Sub

Err:
    If StrMsgError = "" Then StrMsgError = Err.Description
End Sub

Private Sub eliminar(ByRef StrMsgError As String)
On Error GoTo Err
Dim indTrans As Boolean
Dim strCodigo As String
Dim rsValida As New ADODB.Recordset

    If MsgBox("Está seguro(a) de eliminar el registro?" & vbCrLf & "Se eliminarán todas sus dependencias.", vbQuestion + vbYesNo, App.Title) = vbNo Then Exit Sub
    strCodigo = Trim(txtCod_Persona.Text)
    
    Cn.BeginTrans
    indTrans = True
    
    '--- Eliminando productosalmacen
    csql = "DELETE FROM productosproveedor WHERE idProveedor = '" & strCodigo & "' AND idEmpresa = '" & glsEmpresa & "'"
    Cn.Execute csql
    
    '--- Eliminando contactosproveedores
    csql = "DELETE FROM contactosproveedores WHERE idProveedor = '" & strCodigo & "' AND idEmpresa = '" & glsEmpresa & "'"
    Cn.Execute csql
    
    '--- Eliminando el registro
    csql = "DELETE FROM proveedores WHERE idProveedor = '" & strCodigo & "' AND idEmpresa = '" & glsEmpresa & "'"
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

Private Sub chkbuenc_Click()

    If chkbuenc.Value = 1 Then
        txtbuenc.Text = "1"
    Else
        txtbuenc.Text = "0"
    End If
    
End Sub

Private Sub chkpercepcion_Click()

    If chkpercepcion.Value = 1 Then
        txtagenteper.Text = "1"
    Else
        txtagenteper.Text = "0"
    End If
    
End Sub

Private Sub chkretencion_Click()

    If chkretencion.Value = 1 Then
        txtagenteret.Text = "1"
    Else
        txtagenteret.Text = "0"
    End If

End Sub
