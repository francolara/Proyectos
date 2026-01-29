VERSION 5.00
Object = "{6A24B331-7634-11D3-A5B0-0050044A7E1A}#1.5#0"; "DXDBGrid.dll"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TabCtl32.Ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "mscomctl.OCX"
Object = "{F41D1D30-7878-4923-8CB3-6CCACDC9C9DE}#1.0#0"; "CATControls.ocx"
Begin VB.Form frmMantVendedores 
   Appearance      =   0  'Flat
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Mantenimiento de Vendedores"
   ClientHeight    =   7950
   ClientLeft      =   4290
   ClientTop       =   4020
   ClientWidth     =   11580
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7950
   ScaleWidth      =   11580
   ShowInTaskbar   =   0   'False
   Begin MSComctlLib.ImageList imgDocVentas 
      Left            =   6120
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
            Picture         =   "frmMantVendedores.frx":0000
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMantVendedores.frx":039A
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMantVendedores.frx":07EC
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMantVendedores.frx":0B86
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMantVendedores.frx":0F20
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMantVendedores.frx":12BA
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMantVendedores.frx":1654
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMantVendedores.frx":19EE
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMantVendedores.frx":1D88
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMantVendedores.frx":2122
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMantVendedores.frx":24BC
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMantVendedores.frx":317E
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   660
      Left            =   0
      TabIndex        =   5
      Top             =   0
      Width           =   11580
      _ExtentX        =   20426
      _ExtentY        =   1164
      ButtonWidth     =   2381
      ButtonHeight    =   1005
      Appearance      =   1
      ImageList       =   "imgDocVentas"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   8
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "        Nuevo        "
            Object.ToolTipText     =   "     Nuevo     "
            ImageIndex      =   1
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Grabar"
            Object.ToolTipText     =   "Grabar"
            ImageIndex      =   4
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Modificar"
            Object.ToolTipText     =   "Modificar"
            ImageIndex      =   10
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Cancelar"
            Object.ToolTipText     =   "Cancelar"
            ImageIndex      =   9
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Eliminar"
            ImageIndex      =   6
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Imprimir"
            Object.ToolTipText     =   "Imprimir"
            ImageIndex      =   11
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Lista"
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
   Begin VB.Frame FraGeneral 
      Appearance      =   0  'Flat
      ForeColor       =   &H00000000&
      Height          =   7200
      Left            =   -45
      TabIndex        =   6
      Top             =   675
      Width           =   11595
      Begin TabDlg.SSTab SSTab1 
         Height          =   6810
         Left            =   135
         TabIndex        =   7
         Top             =   225
         Width           =   11265
         _ExtentX        =   19870
         _ExtentY        =   12012
         _Version        =   393216
         Tabs            =   2
         TabsPerRow      =   2
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
         TabPicture(0)   =   "frmMantVendedores.frx":3518
         Tab(0).ControlEnabled=   -1  'True
         Tab(0).Control(0)=   "FraDatosGenerales"
         Tab(0).Control(0).Enabled=   0   'False
         Tab(0).ControlCount=   1
         TabCaption(1)   =   "Comisión"
         TabPicture(1)   =   "frmMantVendedores.frx":3534
         Tab(1).ControlEnabled=   0   'False
         Tab(1).Control(0)=   "Frame6"
         Tab(1).Control(0).Enabled=   0   'False
         Tab(1).ControlCount=   1
         Begin VB.Frame FraDatosGenerales 
            Appearance      =   0  'Flat
            ForeColor       =   &H80000008&
            Height          =   6105
            Left            =   180
            TabIndex        =   10
            Top             =   405
            Width           =   10800
            Begin VB.Frame FraTipoVenta 
               Appearance      =   0  'Flat
               Caption         =   "Tipo Venta"
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
               Height          =   825
               Left            =   3195
               TabIndex        =   50
               Top             =   5085
               Width           =   4560
               Begin VB.OptionButton OptTipoVenta 
                  Caption         =   "Avícola"
                  BeginProperty Font 
                     Name            =   "Arial"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   285
                  Index           =   1
                  Left            =   3240
                  TabIndex        =   52
                  Top             =   360
                  Width           =   915
               End
               Begin VB.OptionButton OptTipoVenta 
                  Caption         =   "Porcina"
                  BeginProperty Font 
                     Name            =   "Arial"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   285
                  Index           =   0
                  Left            =   495
                  TabIndex        =   51
                  Top             =   360
                  Value           =   -1  'True
                  Width           =   915
               End
               Begin CATControls.CATTextBox Txt_TipoVenta 
                  Height          =   300
                  Left            =   2070
                  TabIndex        =   53
                  Tag             =   "TIndTipoVenta"
                  Top             =   270
                  Visible         =   0   'False
                  Width           =   525
                  _ExtentX        =   926
                  _ExtentY        =   529
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
                  MaxLength       =   200
                  Container       =   "frmMantVendedores.frx":3550
                  Estilo          =   1
                  Vacio           =   -1  'True
                  EnterTab        =   -1  'True
               End
            End
            Begin VB.CheckBox ChkJefe 
               Caption         =   "Jefe"
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
               Left            =   765
               TabIndex        =   48
               Top             =   4950
               Width           =   825
            End
            Begin VB.CommandButton CmbAyudaJefe 
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   315
               Left            =   9915
               Picture         =   "frmMantVendedores.frx":356C
               Style           =   1  'Graphical
               TabIndex        =   44
               Top             =   3105
               Width           =   390
            End
            Begin VB.CommandButton cmbAyudaPersona 
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   315
               Left            =   9930
               Picture         =   "frmMantVendedores.frx":38F6
               Style           =   1  'Graphical
               TabIndex        =   16
               Top             =   495
               Width           =   390
            End
            Begin VB.CommandButton cmbAyudaListaPrecios 
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   315
               Left            =   9915
               Picture         =   "frmMantVendedores.frx":3C80
               Style           =   1  'Graphical
               TabIndex        =   15
               Top             =   2720
               Width           =   390
            End
            Begin VB.Frame framCobranzas 
               Appearance      =   0  'Flat
               Caption         =   " Datos de Cobranzas "
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
               Height          =   870
               Left            =   720
               TabIndex        =   12
               Top             =   3945
               Width           =   9645
               Begin VB.CheckBox chkVendedor 
                  Caption         =   "Vendedor "
                  BeginProperty Font 
                     Name            =   "Arial"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   420
                  Left            =   1125
                  TabIndex        =   14
                  Top             =   315
                  Width           =   1545
               End
               Begin VB.CheckBox chkResponsable 
                  Caption         =   "Responsable"
                  BeginProperty Font 
                     Name            =   "Arial"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   420
                  Left            =   7695
                  TabIndex        =   13
                  Top             =   315
                  Width           =   1455
               End
            End
            Begin VB.CheckBox chkVisualizaclientes 
               Caption         =   "Visualiza Clientes"
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
               Left            =   8730
               TabIndex        =   11
               Top             =   4950
               Width           =   1590
            End
            Begin CATControls.CATTextBox txtCod_Persona 
               Height          =   315
               Left            =   2220
               TabIndex        =   17
               Tag             =   "TidVendedor"
               Top             =   495
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
               Container       =   "frmMantVendedores.frx":400A
               Estilo          =   1
               EnterTab        =   -1  'True
            End
            Begin CATControls.CATTextBox txtGls_Persona 
               Height          =   315
               Left            =   3180
               TabIndex        =   18
               Top             =   495
               Width           =   6690
               _ExtentX        =   11800
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
               Container       =   "frmMantVendedores.frx":4026
            End
            Begin CATControls.CATTextBox txtGls_Direccion 
               Height          =   315
               Left            =   2220
               TabIndex        =   19
               Top             =   2360
               Width           =   8100
               _ExtentX        =   14288
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
               Container       =   "frmMantVendedores.frx":4042
               Estilo          =   1
               Vacio           =   -1  'True
               EnterTab        =   -1  'True
            End
            Begin CATControls.CATTextBox txtCod_Pais 
               Height          =   315
               Left            =   2220
               TabIndex        =   20
               Top             =   870
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
               Container       =   "frmMantVendedores.frx":405E
               Estilo          =   1
               Vacio           =   -1  'True
               EnterTab        =   -1  'True
            End
            Begin CATControls.CATTextBox txtGls_Pais 
               Height          =   315
               Left            =   3195
               TabIndex        =   21
               Top             =   870
               Width           =   7140
               _ExtentX        =   12594
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
               Container       =   "frmMantVendedores.frx":407A
               Vacio           =   -1  'True
            End
            Begin CATControls.CATTextBox txtCod_Depa 
               Height          =   315
               Left            =   2220
               TabIndex        =   22
               Top             =   1245
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
               Container       =   "frmMantVendedores.frx":4096
               Estilo          =   1
               Vacio           =   -1  'True
               EnterTab        =   -1  'True
            End
            Begin CATControls.CATTextBox txtGls_Depa 
               Height          =   315
               Left            =   3195
               TabIndex        =   23
               Top             =   1245
               Width           =   7140
               _ExtentX        =   12594
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
               Container       =   "frmMantVendedores.frx":40B2
               Vacio           =   -1  'True
            End
            Begin CATControls.CATTextBox txtCod_Prov 
               Height          =   315
               Left            =   2220
               TabIndex        =   24
               Top             =   1620
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
               Container       =   "frmMantVendedores.frx":40CE
               Estilo          =   1
               Vacio           =   -1  'True
               EnterTab        =   -1  'True
            End
            Begin CATControls.CATTextBox txtGls_Prov 
               Height          =   315
               Left            =   3195
               TabIndex        =   25
               Top             =   1620
               Width           =   7140
               _ExtentX        =   12594
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
               Container       =   "frmMantVendedores.frx":40EA
               Vacio           =   -1  'True
            End
            Begin CATControls.CATTextBox txtCod_Distrito 
               Height          =   315
               Left            =   2220
               TabIndex        =   26
               Top             =   1995
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
               Container       =   "frmMantVendedores.frx":4106
               Estilo          =   1
               Vacio           =   -1  'True
               EnterTab        =   -1  'True
            End
            Begin CATControls.CATTextBox txtGls_Distrito 
               Height          =   315
               Left            =   3195
               TabIndex        =   27
               Top             =   1995
               Width           =   7140
               _ExtentX        =   12594
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
               Container       =   "frmMantVendedores.frx":4122
               Vacio           =   -1  'True
            End
            Begin CATControls.CATTextBox txtCod_ListaPrecio 
               Height          =   315
               Left            =   2220
               TabIndex        =   28
               Tag             =   "TidLista"
               Top             =   2720
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
               Container       =   "frmMantVendedores.frx":413E
               Estilo          =   1
               Vacio           =   -1  'True
               EnterTab        =   -1  'True
            End
            Begin CATControls.CATTextBox txtGls_ListaPrecio 
               Height          =   315
               Left            =   3195
               TabIndex        =   29
               Top             =   2720
               Width           =   6690
               _ExtentX        =   11800
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
               Container       =   "frmMantVendedores.frx":415A
               Vacio           =   -1  'True
            End
            Begin CATControls.CATTextBox TxtIndVendedor 
               Height          =   285
               Left            =   855
               TabIndex        =   30
               Tag             =   "TIndVendedor"
               Top             =   6075
               Visible         =   0   'False
               Width           =   915
               _ExtentX        =   1614
               _ExtentY        =   503
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
               Container       =   "frmMantVendedores.frx":4176
               Estilo          =   1
               Vacio           =   -1  'True
               EnterTab        =   -1  'True
            End
            Begin CATControls.CATTextBox txtIndResponsable 
               Height          =   285
               Left            =   1890
               TabIndex        =   31
               Tag             =   "TindResponsable"
               Top             =   6120
               Visible         =   0   'False
               Width           =   780
               _ExtentX        =   1376
               _ExtentY        =   503
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
               Container       =   "frmMantVendedores.frx":4192
               Estilo          =   1
               Vacio           =   -1  'True
               EnterTab        =   -1  'True
            End
            Begin CATControls.CATTextBox txtIndVisualizaclientes 
               Height          =   285
               Left            =   2880
               TabIndex        =   32
               Tag             =   "TindVisualizaclientes"
               Top             =   6480
               Visible         =   0   'False
               Width           =   645
               _ExtentX        =   1138
               _ExtentY        =   503
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
               Container       =   "frmMantVendedores.frx":41AE
               Estilo          =   1
               Vacio           =   -1  'True
               EnterTab        =   -1  'True
            End
            Begin CATControls.CATTextBox txt_rmp 
               Height          =   315
               Left            =   8580
               TabIndex        =   33
               Tag             =   "TRpm"
               Top             =   3495
               Width           =   1725
               _ExtentX        =   3043
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
               MaxLength       =   200
               Container       =   "frmMantVendedores.frx":41CA
               Estilo          =   1
               Vacio           =   -1  'True
               EnterTab        =   -1  'True
            End
            Begin CATControls.CATTextBox txt_Nextel 
               Height          =   300
               Left            =   2220
               TabIndex        =   34
               Tag             =   "TNextel"
               Top             =   3495
               Width           =   1830
               _ExtentX        =   3228
               _ExtentY        =   529
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
               MaxLength       =   200
               Container       =   "frmMantVendedores.frx":41E6
               Estilo          =   1
               Vacio           =   -1  'True
               EnterTab        =   -1  'True
            End
            Begin CATControls.CATTextBox TxtCod_Jefe 
               Height          =   315
               Left            =   2220
               TabIndex        =   45
               Tag             =   "TidJefe"
               Top             =   3105
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
               Container       =   "frmMantVendedores.frx":4202
               Estilo          =   1
               Vacio           =   -1  'True
               EnterTab        =   -1  'True
            End
            Begin CATControls.CATTextBox TxtGls_Jefe 
               Height          =   315
               Left            =   3195
               TabIndex        =   46
               Top             =   3105
               Width           =   6690
               _ExtentX        =   11800
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
               Container       =   "frmMantVendedores.frx":421E
               Vacio           =   -1  'True
            End
            Begin CATControls.CATTextBox TxtInd_Jefe 
               Height          =   285
               Left            =   3825
               TabIndex        =   49
               Tag             =   "TindJefe"
               Top             =   6390
               Visible         =   0   'False
               Width           =   780
               _ExtentX        =   1376
               _ExtentY        =   503
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
               Container       =   "frmMantVendedores.frx":423A
               Estilo          =   1
               Vacio           =   -1  'True
               EnterTab        =   -1  'True
            End
            Begin VB.Label Label9 
               Appearance      =   0  'Flat
               AutoSize        =   -1  'True
               Caption         =   "Jefe"
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
               Left            =   720
               TabIndex        =   47
               Top             =   3150
               Width           =   315
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
               Left            =   720
               TabIndex        =   43
               Top             =   2445
               Width           =   675
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
               Left            =   720
               TabIndex        =   42
               Top             =   1635
               Width           =   660
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
               Left            =   720
               TabIndex        =   41
               Top             =   1260
               Width           =   1005
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
               Left            =   720
               TabIndex        =   40
               Top             =   885
               Width           =   300
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
               Left            =   720
               TabIndex        =   39
               Top             =   2085
               Width           =   495
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
               Left            =   720
               TabIndex        =   38
               Top             =   510
               Width           =   525
            End
            Begin VB.Label Label4 
               Appearance      =   0  'Flat
               AutoSize        =   -1  'True
               Caption         =   "Lista de Precio"
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
               Left            =   720
               TabIndex        =   37
               Top             =   2805
               Width           =   1065
            End
            Begin VB.Label Label13 
               Appearance      =   0  'Flat
               AutoSize        =   -1  'True
               Caption         =   "RPM"
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
               Left            =   8100
               TabIndex        =   36
               Top             =   3570
               Width           =   315
            End
            Begin VB.Label Label6 
               Appearance      =   0  'Flat
               AutoSize        =   -1  'True
               Caption         =   "Nextel"
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
               Left            =   720
               TabIndex        =   35
               Top             =   3570
               Width           =   450
            End
         End
         Begin VB.Frame Frame6 
            Appearance      =   0  'Flat
            ForeColor       =   &H00000000&
            Height          =   5565
            Left            =   -74865
            TabIndex        =   8
            Top             =   495
            Width           =   10980
            Begin DXDBGRIDLibCtl.dxDBGrid gComision 
               Height          =   5130
               Left            =   135
               OleObjectBlob   =   "frmMantVendedores.frx":4256
               TabIndex        =   9
               Top             =   270
               Width           =   10830
            End
         End
      End
   End
   Begin VB.Frame fraListado 
      Appearance      =   0  'Flat
      ForeColor       =   &H00000000&
      Height          =   7200
      Left            =   -45
      TabIndex        =   2
      Top             =   675
      Width           =   11595
      Begin VB.Frame Frame1 
         Appearance      =   0  'Flat
         ForeColor       =   &H80000008&
         Height          =   705
         Left            =   90
         TabIndex        =   3
         Top             =   180
         Width           =   11280
         Begin CATControls.CATTextBox txt_TextoBuscar 
            Height          =   315
            Left            =   990
            TabIndex        =   0
            Top             =   255
            Width           =   10110
            _ExtentX        =   17833
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
            Container       =   "frmMantVendedores.frx":7672
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
            Left            =   120
            TabIndex        =   4
            Top             =   300
            Width           =   735
         End
      End
      Begin DXDBGRIDLibCtl.dxDBGrid gLista 
         Height          =   6045
         Left            =   150
         OleObjectBlob   =   "frmMantVendedores.frx":768E
         TabIndex        =   1
         Top             =   945
         Width           =   11250
      End
   End
End
Attribute VB_Name = "frmMantVendedores"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim indBoton As Integer

Private Sub ChkJefe_Click()
    If ChkJefe = 1 Then
        TxtInd_Jefe.Text = "1"
    Else
        TxtInd_Jefe.Text = "0"
    End If
End Sub

Private Sub chkVisualizaclientes_Click()

    If chkVisualizaclientes = 1 Then
        txtIndVisualizaclientes.Text = "1"
    Else
        txtIndVisualizaclientes.Text = "0"
    End If

End Sub

Private Sub CmbAyudaJefe_Click()

    mostrarAyuda "VENDEDORJEFE", TxtCod_Jefe, TxtGls_Jefe
    
End Sub

Private Sub cmbAyudaListaPrecios_Click()
    
    mostrarAyuda "LISTAPRECIOS", txtCod_ListaPrecio, txtGls_ListaPrecio

End Sub

Private Sub cmbAyudaPersona_Click()
    
    mostrarAyuda "PERSONAVENDEDOR", txtCod_Persona, txtGls_Persona
    If txtCod_Persona.Text <> "" Then mostrarDatosPersona

End Sub

Private Sub mostrarDatosPersona()
Dim rst As New ADODB.Recordset

    csql = "SELECT idPersona,GlsPersona," & _
                       "direccion, p.iddistrito, u.idDpto, u.idProv " & _
               "FROM personas p,ubigeo u " & _
               "WHERE p.iddistrito = u.iddistrito and idpersona = '" & Trim(txtCod_Persona.Text) & "'"
               
    rst.Open csql, Cn, adOpenStatic, adLockReadOnly
    If Not rst.EOF Then
        txtCod_Pais.Text = "02001"
        txtCod_Depa.Text = "" & rst.Fields("idDpto")
        txtCod_Prov.Text = "" & rst.Fields("idProv")
        txtCod_Distrito.Text = "" & rst.Fields("iddistrito")
        txtGls_Direccion.Text = "" & rst.Fields("direccion")
    Else
        txtCod_Pais.Text = ""
        txtCod_Depa.Text = ""
        txtCod_Prov.Text = ""
        txtCod_Distrito.Text = ""
        txtGls_Direccion.Text = ""
    End If
    
    rst.Close: Set rst = Nothing

End Sub

Private Sub chkResponsable_Click()

    If chkResponsable.Value = 1 Then
        txtIndResponsable.Text = "S"
    Else
        txtIndResponsable.Text = ""
    End If

End Sub

Private Sub chkVendedor_Click()

    If chkVendedor = 1 Then
        TxtIndVendedor.Text = "S"
    Else
        TxtIndVendedor.Text = ""
    End If

End Sub

Private Sub Form_Load()
On Error GoTo Err
Dim StrMsgError     As String
Dim strRUC          As String

    Me.left = 0
    Me.top = 0
    
    strRUC = traerCampo("Empresas", "Ruc", "idEmpresa", glsEmpresa, False)
    
    ConfGrid gLista, False, False, False, False
    ConfGrid gComision, True, False, False, False
     
    listaVendedor StrMsgError
    If StrMsgError <> "" Then GoTo Err
    
    fraListado.Visible = True
    fraGeneral.Visible = False
    habilitaBotones 7
    
    'Solo Para Apimas
    If (strRUC = "20305948277" Or strRUC = "20987898989" Or strRUC = "20544632192") Then
        SSTab1.TabVisible(1) = True
        ChkJefe.Visible = True
    Else
        SSTab1.TabVisible(1) = False
        ChkJefe.Visible = False
    End If
    
    nuevo
    
    If leeParametro("VISUALIZA_TIPOVENTA") = "S" Then
        
        FraTipoVenta.Visible = True
    
    Else
        
        FraTipoVenta.Visible = False
    
    End If
    
    Exit Sub

Err:
    If StrMsgError = "" Then StrMsgError = Err.Description
    MsgBox StrMsgError, vbInformation, App.Title

End Sub
 

Private Sub gLista_OnDblClick()
On Error GoTo Err
Dim StrMsgError As String

    mostrarVendedor gLista.Columns.ColumnByName("idVendedor").Value, StrMsgError
    If StrMsgError <> "" Then GoTo Err
    
    fraListado.Visible = False
    fraGeneral.Visible = True
    fraGeneral.Enabled = False
    
    txtCod_Persona.Enabled = False
    cmbAyudaPersona.Enabled = False
    habilitaBotones 2
    
    Exit Sub

Err:
    MsgBox StrMsgError, vbInformation, App.Title

End Sub

Private Sub OptTipoVenta_Click(Index As Integer)
On Error GoTo Err
Dim StrMsgError As String
    
    If FraTipoVenta.Visible Then
    
        If OptTipoVenta(0).Value Then
            
            Txt_TipoVenta.Text = "P"
        
        Else
            
            Txt_TipoVenta.Text = "A"
        
        End If
    
    Else
    
        Txt_TipoVenta.Text = ""
    
    End If
    
    Exit Sub

Err:
    If StrMsgError = "" Then StrMsgError = Err.Description
    MsgBox StrMsgError, vbInformation, App.Title
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
On Error GoTo Err
Dim StrMsgError As String

    Select Case Button.Index
        Case 1 'Nuevo
            indBoton = 0
            nuevo
            fraListado.Visible = False
            fraGeneral.Visible = True
            fraGeneral.Enabled = True
        Case 2 'Grabar
            Grabar StrMsgError
            If StrMsgError <> "" Then GoTo Err
        Case 3 'Modificar
            indBoton = 1
            fraGeneral.Enabled = True
        Case 4, 7 'Cancelar
            fraListado.Visible = True
            fraGeneral.Visible = False
            fraGeneral.Enabled = False
        Case 5 'Eliminar
            eliminar StrMsgError
            If StrMsgError <> "" Then GoTo Err
        Case 6 'Imprimir
            gLista.m.ExportToXLS App.Path & "\Temporales\Mantenimiento_Vendedores.xls"
            ShellEx App.Path & "\Temporales\Mantenimiento_Vendedores.xls", essSW_MAXIMIZE, , , "open", Me.hwnd
        Case 8 'Salir
            Unload Me
    End Select
    habilitaBotones Button.Index
    
    Exit Sub

Err:
    If StrMsgError = "" Then StrMsgError = Err.Description
    MsgBox StrMsgError, vbInformation, App.Title
End Sub

Private Sub txt_TextoBuscar_Change()
On Error GoTo Err
Dim StrMsgError As String
    
    If glsEnterAyudaClientes = False Then
        listaVendedor StrMsgError
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
        listaVendedor StrMsgError
        If StrMsgError <> "" Then GoTo Err
    End If

    Exit Sub
    
Err:
    If StrMsgError = "" Then StrMsgError = Err.Description
    MsgBox StrMsgError, vbInformation, App.Title

End Sub

Private Sub TxtCod_Jefe_Change()

    TxtGls_Jefe.Text = traerCampo("Personas", "GlsPersona", "idPersona", TxtCod_Jefe.Text, False)
    
End Sub

Private Sub txtCod_ListaPrecio_Change()
    
    txtGls_ListaPrecio.Text = traerCampo("listaprecios", "GlsLista", "idLista", txtCod_ListaPrecio.Text, True)

End Sub

Private Sub txtCod_ListaPrecio_KeyPress(KeyAscii As Integer)

    If KeyAscii <> 13 Then
        mostrarAyudaKeyascii KeyAscii, "LISTAPRECIOS", txtCod_ListaPrecio, txtGls_ListaPrecio
        KeyAscii = 0
        If txtCod_ListaPrecio.Text <> "" Then SendKeys "{tab}"
    End If

End Sub

Private Sub txtCod_Persona_Change()
    
    txtGls_Persona.Text = traerCampo("personas", "GlsPersona", "idPersona", txtCod_Persona.Text, False)

End Sub

Private Sub txtCod_Depa_Change()
    
    txtGls_Depa.Text = traerCampo("Ubigeo", "GlsUbigeo", "idDpto", txtCod_Depa.Text, False, " idProv = '00'")

End Sub

Private Sub txtCod_Distrito_Change()
    
    txtGls_Distrito.Text = traerCampo("Ubigeo", "GlsUbigeo", "idDistrito", txtCod_Distrito.Text, False)

End Sub

Private Sub txtCod_Pais_Change()
    
    txtGls_Pais.Text = traerCampo("Datos", "GlsDato", "idDato", txtCod_Pais.Text, False)

End Sub

Private Sub txtCod_Persona_KeyPress(KeyAscii As Integer)

    If KeyAscii <> 13 Then
        mostrarAyudaKeyascii KeyAscii, "PERSONAVENDEDOR", txtCod_Persona, txtGls_Persona
        KeyAscii = 0
        If txtCod_Persona.Text <> "" Then
            mostrarDatosPersona
            SendKeys "{tab}"
        End If
    End If

End Sub

Private Sub txtCod_Prov_Change()
    
    txtGls_Prov.Text = traerCampo("Ubigeo", "GlsUbigeo", "idProv", txtCod_Prov.Text, False, " idDpto = '" & txtCod_Depa.Text & "' and idProv <> '00' and idDist = '00'")

End Sub

Private Sub habilitaBotones(indexBoton As Integer)
Dim indHabilitar As Boolean

    Select Case indexBoton
        Case 1, 2, 3 'Nuevo, Grabar, Modificar
            If indexBoton = 2 Then indHabilitar = True
            Toolbar1.Buttons(1).Visible = indHabilitar 'Nuevo
            Toolbar1.Buttons(2).Visible = Not indHabilitar 'Grabar
            Toolbar1.Buttons(3).Visible = indHabilitar 'Modificar
            Toolbar1.Buttons(4).Visible = Not indHabilitar 'Cancelar
            Toolbar1.Buttons(5).Visible = indHabilitar 'Eliminar
            Toolbar1.Buttons(6).Visible = indHabilitar 'Imprimir
            Toolbar1.Buttons(7).Visible = indHabilitar 'Lista
        Case 4, 7 'Cancelar, Lista
            Toolbar1.Buttons(1).Visible = True
            Toolbar1.Buttons(2).Visible = False
            Toolbar1.Buttons(3).Visible = False
            Toolbar1.Buttons(4).Visible = False
            Toolbar1.Buttons(5).Visible = False
            Toolbar1.Buttons(6).Visible = True
            Toolbar1.Buttons(7).Visible = False
    End Select

End Sub

Private Sub Grabar(ByRef StrMsgError As String)
On Error GoTo Err
Dim strCodigo As String
Dim strMsg      As String

    validaFormSQL Me, StrMsgError
    If StrMsgError <> "" Then GoTo Err
    
    If indBoton = 0 Then 'graba
        EjecutaSQLForm Me, 0, True, "vendedores", StrMsgError
        If StrMsgError <> "" Then GoTo Err
        
        GrabaComisionxVendedor StrMsgError
        If StrMsgError <> "" Then GoTo Err
        
        strMsg = "Grabó"
        
    Else 'modifica
        EjecutaSQLForm Me, 1, True, "vendedores", StrMsgError, "idVendedor"
        If StrMsgError <> "" Then GoTo Err
        
        GrabaComisionxVendedor StrMsgError
        If StrMsgError <> "" Then GoTo Err
        
        strMsg = "Modificó"
    End If
    
    MsgBox "Se " & strMsg & " Satisfactoriamente", vbInformation, App.Title
    
    fraGeneral.Enabled = False
    
    listaVendedor StrMsgError
    If StrMsgError <> "" Then GoTo Err
    
    Exit Sub

Err:
    If StrMsgError = "" Then StrMsgError = Err.Description

End Sub

Private Sub nuevo()
Dim StrMsgError As String
Dim rst As New ADODB.Recordset
On Error GoTo Err
  
    limpiaForm Me
    txtCod_Persona.Enabled = True
    cmbAyudaPersona.Enabled = True
    txtIndVisualizaclientes.Text = "0"
 
  
    rst.Fields.Append "Item", adInteger, , adFldRowID
    rst.Fields.Append "PeriodoIni", adVarChar, 8, adFldIsNullable
    rst.Fields.Append "Monto1", adDouble, 14, adFldIsNullable
    rst.Fields.Append "Monto2", adDouble, 14, adFldIsNullable
    rst.Fields.Append "Porcentaje1", adDouble, 14, adFldIsNullable
    rst.Fields.Append "Porcentaje2", adDouble, 14, adFldIsNullable
    rst.Fields.Append "Porcentaje3", adDouble, 14, adFldIsNullable
    rst.Open
    
    rst.AddNew
    rst.Fields("Item") = 1
    rst.Fields("PeriodoIni") = ""
    rst.Fields("Monto1") = 0
    rst.Fields("Monto2") = 0
    rst.Fields("Porcentaje1") = 0
    rst.Fields("Porcentaje2") = 0
    rst.Fields("Porcentaje3") = 0
    
    TxtInd_Jefe.Text = "0"
    
    mostrarDatosGridSQL gComision, rst, StrMsgError
    If StrMsgError <> "" Then GoTo Err
    
    SSTab1.Tab = 0
    
    OptTipoVenta(0).Value = False
    OptTipoVenta(1).Value = False
    
    Txt_TipoVenta.Text = ""
    
    
    Exit Sub
Err:
If StrMsgError = "" Then StrMsgError = Err.Description
MsgBox StrMsgError, vbInformation, App.Title
End Sub

Private Sub listaVendedor(ByRef StrMsgError As String)
Dim rsdatos                     As New ADODB.Recordset
On Error GoTo Err
Dim strCond As String

    strCond = ""
    If Trim(txt_TextoBuscar.Text) <> "" Then
        strCond = Trim(txt_TextoBuscar.Text)
        strCond = " AND GlsPersona LIKE '%" & strCond & "%'"
    End If
    
    csql = "SELECT v.idVendedor ,p.GlsPersona ," & _
            "(p.direccion + ', ' + isnull(u.glsUbigeo, '')) as Direccion " & _
            "FROM vendedores v inner join personas p on v.idVendedor = p.idPersona " & _
            "left join ubigeo u on p.iddistrito = u.iddistrito and p.idPais = u.idPais WHERE v.idEmpresa = '" & glsEmpresa & "'" & strCond & _
            "ORDER BY v.idVendedor"
            
 If rsdatos.State = 1 Then rsdatos.Close: Set rsdatos = Nothing
 rsdatos.Open csql, Cn, adOpenStatic, adLockOptimistic
            
 Set gLista.DataSource = rsdatos
'    With gLista
'        .DefaultFields = False
'        .Dataset.ADODataset.ConnectionString = strcn
'        .Dataset.ADODataset.CursorLocation = clUseClient
'        .Dataset.Active = False
'        .Dataset.ADODataset.CommandText = csql
'        .Dataset.DisableControls
'        .Dataset.Active = True
'        .KeyField = "idVendedor"
'    End With
    
    Me.Refresh
    
    Exit Sub

Err:
    If StrMsgError = "" Then StrMsgError = Err.Description

End Sub

Private Sub mostrarVendedor(strCodVen As String, ByRef StrMsgError As String)
On Error GoTo Err
Dim rst As New ADODB.Recordset
Dim RsC As New ADODB.Recordset
Dim RsD As New ADODB.Recordset
   
    csql = "SELECT v.idVendedor,v.idLista,v.IndVendedor, v.indResponsable,v.indVisualizaclientes,v.Rpm,v.Nextel,v.indJefe,v.idJefe,V.IndTipoVenta " & _
           "FROM vendedores v " & _
           "WHERE v.idVendedor = '" & strCodVen & "' AND IDEMPRESA='" & glsEmpresa & "'"
    rst.Open csql, Cn, adOpenStatic, adLockReadOnly
    
    If Not rst.EOF Then
        
        If Trim("" & rst.Fields("IndTipoVenta")) = "P" Then
            
            OptTipoVenta(0).Value = True
        
        ElseIf Trim("" & rst.Fields("IndTipoVenta")) = "A" Then
            
            OptTipoVenta(1).Value = True
        
        Else
            
            OptTipoVenta(0).Value = False
            OptTipoVenta(1).Value = False
            
        End If
            
        mostrarDatosFormSQL Me, rst, StrMsgError
        If StrMsgError <> "" Then GoTo Err
    
    End If
    
    txtIndResponsable_Change
    TxtIndVendedor_Change
    txtIndVisualizaclientes_Change
    TxtInd_Jefe_Change
    mostrarDatosPersona
    
    
    csql = "Select item, PeriodoIni, Monto1, Monto2, Porcentaje1, Porcentaje2, Porcentaje3, indJefe " & _
            "From comisionxvendedor Where idEmpresa  = '" & glsEmpresa & "' And idVendedor = '" & txtCod_Persona.Text & "'"
    RsC.Open csql, Cn, adOpenStatic, adLockReadOnly
    
    
    RsD.Fields.Append "Item", adInteger, , adFldRowID
    RsD.Fields.Append "PeriodoIni", adVarChar, 8, adFldIsNullable
    RsD.Fields.Append "Monto1", adDouble, 14, adFldIsNullable
    RsD.Fields.Append "Monto2", adDouble, 14, adFldIsNullable
    RsD.Fields.Append "Porcentaje1", adDouble, 14, adFldIsNullable
    RsD.Fields.Append "Porcentaje2", adDouble, 14, adFldIsNullable
    RsD.Fields.Append "Porcentaje3", adDouble, 14, adFldIsNullable
    RsD.Fields.Append "indJefe", adChar, 1, adFldIsNullable
    RsD.Open
    
    If RsC.RecordCount = 0 Then
        RsD.AddNew
        RsD.Fields("Item") = 1
        RsD.Fields("PeriodoIni") = ""
        RsD.Fields("Monto1") = 0
        RsD.Fields("Monto2") = 0
        RsD.Fields("Porcentaje1") = 0
        RsD.Fields("Porcentaje2") = 0
        RsD.Fields("Porcentaje3") = 0
        RsD.Fields("indJefe") = ""
    Else
        If Not RsC.EOF Then
            RsC.MoveFirst
            Do While Not RsC.EOF
                RsD.AddNew
                RsD.Fields("Item") = RsC.Fields("Item")
                RsD.Fields("PeriodoIni") = RsC.Fields("PeriodoIni")
                RsD.Fields("Monto1") = RsC.Fields("Monto1")
                RsD.Fields("Monto2") = RsC.Fields("Monto2")
                RsD.Fields("Porcentaje1") = RsC.Fields("Porcentaje1")
                RsD.Fields("Porcentaje2") = RsC.Fields("Porcentaje2")
                RsD.Fields("Porcentaje3") = RsC.Fields("Porcentaje3")
                RsD.Fields("indJefe") = RsC.Fields("indJefe")
                
                RsC.MoveNext
            Loop
        End If
    End If
    mostrarDatosGridSQL gComision, RsD, StrMsgError
    If StrMsgError <> "" Then GoTo Err
    
    If RsC.State = 0 Then RsC.Close: Set RsC = Nothing
    SSTab1.Tab = 0
    
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

    If MsgBox("¿Seguro de eliminar el registro?" & vbCrLf & "Se eliminaran todas sus dependencias.", vbQuestion + vbYesNo, App.Title) = vbNo Then Exit Sub
    
    strCodigo = Trim(txtCod_Persona.Text)

    Cn.BeginTrans
    indTrans = True

    'Eliminando el registro
    csql = "DELETE FROM vendedores WHERE idVendedor = '" & strCodigo & "' AND idEmpresa = '" & glsEmpresa & "'"
    Cn.Execute csql
    
    Cn.CommitTrans
    
    'Nuevo
    Toolbar1_ButtonClick Toolbar1.Buttons(1)
    
    MsgBox "Registro eliminado satisfactoriamente", vbInformation, App.Title
    listaVendedor StrMsgError
    If rsValida.State = 1 Then rsValida.Close: Set rsValida = Nothing
    
    Exit Sub

Err:
    If rsValida.State = 1 Then rsValida.Close: Set rsValida = Nothing
    If indTrans Then Cn.RollbackTrans
    If StrMsgError = "" Then StrMsgError = Err.Description

End Sub

Private Sub txtIndResponsable_Change()

    If txtIndResponsable.Text = "" Then
        chkResponsable.Value = 0
    Else
        chkResponsable.Value = 1
    End If

End Sub

Private Sub TxtIndVendedor_Change()

    If TxtIndVendedor.Text = "" Then
        chkVendedor.Value = 0
    Else
        chkVendedor.Value = 1
    End If

End Sub

Private Sub txtIndVisualizaclientes_Change()

    If (txtIndVisualizaclientes.Text = "" Or txtIndVisualizaclientes.Text = "0") Then
        chkVisualizaclientes.Value = 0
    Else
        chkVisualizaclientes.Value = 1
    End If

End Sub

Private Sub TxtInd_Jefe_Change()

    If TxtInd_Jefe.Text = "0" Then
        ChkJefe.Value = 0
    Else
        ChkJefe.Value = 1
    End If

End Sub

Private Sub gComision_OnKeyDown(KeyCode As Integer, ByVal Shift As Long)
Dim i As Integer
    
    If KeyCode = 46 Then
        If gComision.Count > 0 Then
            If MsgBox("Está seguro(a) de eliminar el registro?", vbInformation + vbYesNo, App.Title) = vbYes Then
                If gComision.Count = 1 Then
                    gComision.Dataset.Edit
                    gComision.Columns.ColumnByFieldName("Item").Value = 1
                    gComision.Columns.ColumnByFieldName("PeriodoIni").Value = ""
                    gComision.Columns.ColumnByFieldName("Monto1").Value = 0
                    gComision.Columns.ColumnByFieldName("Monto2").Value = 0
                    gComision.Columns.ColumnByFieldName("Porcentaje1").Value = 0
                    gComision.Columns.ColumnByFieldName("Porcentaje2").Value = 0
                    gComision.Columns.ColumnByFieldName("Porcentaje3").Value = 0
                    gComision.Dataset.Post
                
                Else
                    gComision.Dataset.Delete
                    gComision.Dataset.First
                    Do While Not gComision.Dataset.EOF
                        i = i + 1
                        gComision.Dataset.Edit
                        gComision.Columns.ColumnByFieldName("Item").Value = i
                        gComision.Dataset.Post
                        gComision.Dataset.Next
                    Loop
                    If gComision.Dataset.State = dsEdit Or gComision.Dataset.State = dsInsert Then
                        gComision.Dataset.Post
                    End If
                End If
            End If
        End If
    End If
    
    If KeyCode = 13 Then
        If gComision.Dataset.State = dsEdit Or gComision.Dataset.State = dsInsert Then
              gComision.Dataset.Post
        End If
    End If

End Sub

Private Sub gComision_OnAfterDatasetAction(ByVal Action As DXDBGRIDLibCtl.ExDatasetAction)
Dim i As Integer
    
    If Action = daInsert Then
        gComision.Columns.ColumnByFieldName("item").Value = gComision.Count
        gComision.Dataset.Post
    End If

End Sub

Private Sub gComision_OnBeforeDatasetAction(ByVal Action As DXDBGRIDLibCtl.ExDatasetAction, Allow As Boolean)
    
    If Action = daInsert Then
        If gComision.Columns.ColumnByFieldName("PeriodoIni").Value = "" Then
            Allow = False
        Else
            gComision.Columns.FocusedIndex = gComision.Columns.ColumnByFieldName("PeriodoIni").Index
        End If
    End If
    
End Sub

Private Sub GrabaComisionxVendedor(StrMsgError As String)
On Error GoTo Err
Dim Cadmysql As String

    Cadmysql = "Delete From ComisionxVendedor Where idEmpresa = '" & glsEmpresa & "' And idVendedor = '" & txtCod_Persona.Text & "'"
    Cn.Execute (Cadmysql)
    
    Cadmysql = ""

  
     With gComision
         .Dataset.First
         If Not .Dataset.EOF Then
             Cadmysql = "Insert Into ComisionxVendedor(item, idEmpresa, idVendedor, PeriodoIni, Monto1, Monto2, Porcentaje1, Porcentaje2, Porcentaje3,indJefe) Values "
             Do While Not .Dataset.EOF
                  Cadmysql = Cadmysql & "('" & .Columns.ColumnByFieldName("item").Value & "','" & glsEmpresa & "','" & txtCod_Persona.Text & "','" & .Columns.ColumnByFieldName("PeriodoIni").Value & "', " & _
                                "'" & .Columns.ColumnByFieldName("Monto1").Value & "', " & _
                                "'" & .Columns.ColumnByFieldName("Monto2").Value & "','" & .Columns.ColumnByFieldName("Porcentaje1").Value & "','" & .Columns.ColumnByFieldName("Porcentaje2").Value & "','" & .Columns.ColumnByFieldName("Porcentaje3").Value & "','" & .Columns.ColumnByFieldName("indJefe").Value & "' )" & ","
                 .Dataset.Next
             Loop
         End If
     End With
     If Cadmysql <> "" Then
         Cadmysql = left(Cadmysql, Len(Cadmysql) - 1)
         Cn.Execute (Cadmysql)
     End If
 
   
     
    Exit Sub
Err:
If StrMsgError = "" Then StrMsgError = Err.Description
End Sub

