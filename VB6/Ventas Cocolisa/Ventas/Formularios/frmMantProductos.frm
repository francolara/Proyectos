VERSION 5.00
Object = "{6A24B331-7634-11D3-A5B0-0050044A7E1A}#1.5#0"; "DXDBGrid.dll"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TabCtl32.Ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "ComDlg32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "mscomctl.OCX"
Object = "{F41D1D30-7878-4923-8CB3-6CCACDC9C9DE}#1.0#0"; "CATControls.ocx"
Begin VB.Form frmMantProductos 
   Appearance      =   0  'Flat
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Mantenimiento de Productos"
   ClientHeight    =   9420
   ClientLeft      =   3750
   ClientTop       =   2940
   ClientWidth     =   12810
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   9420
   ScaleWidth      =   12810
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmbAyudaTallaPeso 
      Height          =   315
      Left            =   6150
      Picture         =   "frmMantProductos.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   100
      Top             =   0
      Visible         =   0   'False
      Width           =   390
   End
   Begin MSComctlLib.ImageList imgDocVentas 
      Left            =   11025
      Top             =   225
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
            Picture         =   "frmMantProductos.frx":038A
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMantProductos.frx":0724
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMantProductos.frx":0B76
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMantProductos.frx":0F10
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMantProductos.frx":12AA
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMantProductos.frx":1644
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMantProductos.frx":19DE
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMantProductos.frx":1D78
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMantProductos.frx":2112
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMantProductos.frx":24AC
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMantProductos.frx":2846
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMantProductos.frx":3508
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin CATControls.CATTextBox txtGls_TallaPeso 
      Height          =   315
      Left            =   0
      TabIndex        =   101
      Top             =   15
      Visible         =   0   'False
      Width           =   6135
      _ExtentX        =   10821
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
      Container       =   "frmMantProductos.frx":38A2
      Vacio           =   -1  'True
   End
   Begin VB.Frame fraGeneral 
      Height          =   8700
      Left            =   30
      TabIndex        =   27
      Top             =   675
      Width           =   12705
      Begin TabDlg.SSTab SSTab1 
         Height          =   8430
         Left            =   105
         TabIndex        =   28
         Top             =   135
         Width           =   12315
         _ExtentX        =   21722
         _ExtentY        =   14870
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
         TabCaption(0)   =   "Datos Generales"
         TabPicture(0)   =   "frmMantProductos.frx":38BE
         Tab(0).ControlEnabled=   -1  'True
         Tab(0).Control(0)=   "fraDatos"
         Tab(0).Control(0).Enabled=   0   'False
         Tab(0).ControlCount=   1
         TabCaption(1)   =   "Almacenes"
         TabPicture(1)   =   "frmMantProductos.frx":38DA
         Tab(1).ControlEnabled=   0   'False
         Tab(1).Control(0)=   "Frame2"
         Tab(1).ControlCount=   1
         TabCaption(2)   =   "Presentaciones"
         TabPicture(2)   =   "frmMantProductos.frx":38F6
         Tab(2).ControlEnabled=   0   'False
         Tab(2).Control(0)=   "Frame3"
         Tab(2).ControlCount=   1
         Begin VB.Frame fraDatos 
            Appearance      =   0  'Flat
            ForeColor       =   &H80000008&
            Height          =   7590
            Left            =   1080
            TabIndex        =   31
            Top             =   525
            Width           =   9885
            Begin VB.Frame fraNivel 
               Appearance      =   0  'Flat
               BorderStyle     =   0  'None
               ForeColor       =   &H80000008&
               Height          =   360
               Left            =   360
               TabIndex        =   70
               Top             =   675
               Width           =   8895
               Begin VB.CommandButton cmbAyudaNivel 
                  Height          =   315
                  Index           =   4
                  Left            =   8400
                  Picture         =   "frmMantProductos.frx":3912
                  Style           =   1  'Graphical
                  TabIndex        =   75
                  Top             =   1440
                  Width           =   390
               End
               Begin VB.CommandButton cmbAyudaNivel 
                  Height          =   315
                  Index           =   3
                  Left            =   8400
                  Picture         =   "frmMantProductos.frx":3C9C
                  Style           =   1  'Graphical
                  TabIndex        =   74
                  Top             =   1080
                  Width           =   390
               End
               Begin VB.CommandButton cmbAyudaNivel 
                  Height          =   315
                  Index           =   2
                  Left            =   8400
                  Picture         =   "frmMantProductos.frx":4026
                  Style           =   1  'Graphical
                  TabIndex        =   73
                  Top             =   720
                  Width           =   390
               End
               Begin VB.CommandButton cmbAyudaNivel 
                  Height          =   315
                  Index           =   1
                  Left            =   8400
                  Picture         =   "frmMantProductos.frx":43B0
                  Style           =   1  'Graphical
                  TabIndex        =   72
                  Top             =   360
                  Width           =   390
               End
               Begin VB.CommandButton cmbAyudaNivel 
                  Height          =   315
                  Index           =   0
                  Left            =   8400
                  Picture         =   "frmMantProductos.frx":473A
                  Style           =   1  'Graphical
                  TabIndex        =   71
                  Top             =   30
                  Width           =   390
               End
               Begin CATControls.CATTextBox txtCod_Nivel 
                  Height          =   315
                  Index           =   0
                  Left            =   1305
                  TabIndex        =   76
                  Tag             =   "TidNivelPred"
                  Top             =   30
                  Visible         =   0   'False
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
                  Container       =   "frmMantProductos.frx":4AC4
                  Estilo          =   1
                  EnterTab        =   -1  'True
               End
               Begin CATControls.CATTextBox txtGls_Nivel 
                  Height          =   315
                  Index           =   0
                  Left            =   2280
                  TabIndex        =   77
                  Top             =   30
                  Width           =   6090
                  _ExtentX        =   10742
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
                  Container       =   "frmMantProductos.frx":4AE0
                  Vacio           =   -1  'True
               End
               Begin CATControls.CATTextBox txtCod_Nivel 
                  Height          =   285
                  Index           =   1
                  Left            =   1305
                  TabIndex        =   78
                  Tag             =   "TidNivelPred"
                  Top             =   390
                  Visible         =   0   'False
                  Width           =   915
                  _ExtentX        =   1614
                  _ExtentY        =   503
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
                  Container       =   "frmMantProductos.frx":4AFC
                  Estilo          =   1
                  EnterTab        =   -1  'True
               End
               Begin CATControls.CATTextBox txtGls_Nivel 
                  Height          =   285
                  Index           =   1
                  Left            =   2280
                  TabIndex        =   79
                  Top             =   390
                  Width           =   6090
                  _ExtentX        =   10742
                  _ExtentY        =   503
                  BackColor       =   12648447
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
                  Container       =   "frmMantProductos.frx":4B18
                  Vacio           =   -1  'True
               End
               Begin CATControls.CATTextBox txtCod_Nivel 
                  Height          =   285
                  Index           =   2
                  Left            =   1305
                  TabIndex        =   80
                  Tag             =   "TidNivelPred"
                  Top             =   720
                  Visible         =   0   'False
                  Width           =   915
                  _ExtentX        =   1614
                  _ExtentY        =   503
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
                  Container       =   "frmMantProductos.frx":4B34
                  Estilo          =   1
                  EnterTab        =   -1  'True
               End
               Begin CATControls.CATTextBox txtGls_Nivel 
                  Height          =   285
                  Index           =   2
                  Left            =   2280
                  TabIndex        =   81
                  Top             =   720
                  Width           =   6090
                  _ExtentX        =   10742
                  _ExtentY        =   503
                  BackColor       =   12648447
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
                  Container       =   "frmMantProductos.frx":4B50
                  Vacio           =   -1  'True
               End
               Begin CATControls.CATTextBox txtCod_Nivel 
                  Height          =   285
                  Index           =   3
                  Left            =   1305
                  TabIndex        =   82
                  Tag             =   "TidNivelPred"
                  Top             =   1110
                  Visible         =   0   'False
                  Width           =   915
                  _ExtentX        =   1614
                  _ExtentY        =   503
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
                  Container       =   "frmMantProductos.frx":4B6C
                  Estilo          =   1
                  EnterTab        =   -1  'True
               End
               Begin CATControls.CATTextBox txtGls_Nivel 
                  Height          =   285
                  Index           =   3
                  Left            =   2280
                  TabIndex        =   83
                  Top             =   1110
                  Width           =   6090
                  _ExtentX        =   10742
                  _ExtentY        =   503
                  BackColor       =   12648447
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
                  Container       =   "frmMantProductos.frx":4B88
                  Vacio           =   -1  'True
               End
               Begin CATControls.CATTextBox txtCod_Nivel 
                  Height          =   285
                  Index           =   4
                  Left            =   1305
                  TabIndex        =   84
                  Tag             =   "TidNivelPred"
                  Top             =   1470
                  Visible         =   0   'False
                  Width           =   915
                  _ExtentX        =   1614
                  _ExtentY        =   503
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
                  Container       =   "frmMantProductos.frx":4BA4
                  Estilo          =   1
                  EnterTab        =   -1  'True
               End
               Begin CATControls.CATTextBox txtGls_Nivel 
                  Height          =   285
                  Index           =   4
                  Left            =   2280
                  TabIndex        =   85
                  Top             =   1470
                  Width           =   6090
                  _ExtentX        =   10742
                  _ExtentY        =   503
                  BackColor       =   12648447
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
                  Container       =   "frmMantProductos.frx":4BC0
                  Vacio           =   -1  'True
               End
               Begin VB.Label lblNivel 
                  Appearance      =   0  'Flat
                  AutoSize        =   -1  'True
                  Caption         =   "Nivel"
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
                  Index           =   4
                  Left            =   135
                  TabIndex        =   90
                  Top             =   1485
                  Width           =   345
               End
               Begin VB.Label lblNivel 
                  Appearance      =   0  'Flat
                  AutoSize        =   -1  'True
                  Caption         =   "Nivel"
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
                  Index           =   3
                  Left            =   135
                  TabIndex        =   89
                  Top             =   1125
                  Width           =   345
               End
               Begin VB.Label lblNivel 
                  Appearance      =   0  'Flat
                  AutoSize        =   -1  'True
                  Caption         =   "Nivel"
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
                  Index           =   2
                  Left            =   135
                  TabIndex        =   88
                  Top             =   765
                  Width           =   345
               End
               Begin VB.Label lblNivel 
                  Appearance      =   0  'Flat
                  AutoSize        =   -1  'True
                  Caption         =   "Nivel"
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
                  Index           =   1
                  Left            =   135
                  TabIndex        =   87
                  Top             =   405
                  Width           =   345
               End
               Begin VB.Label lblNivel 
                  Appearance      =   0  'Flat
                  AutoSize        =   -1  'True
                  Caption         =   "Nivel"
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
                  Index           =   0
                  Left            =   135
                  TabIndex        =   86
                  Top             =   90
                  Width           =   345
               End
            End
            Begin VB.Frame fraContenido 
               Appearance      =   0  'Flat
               BorderStyle     =   0  'None
               ForeColor       =   &H80000008&
               Height          =   6105
               Left            =   360
               TabIndex        =   34
               Top             =   1260
               Width           =   9300
               Begin VB.CheckBox ChkAfectoIVAP 
                  Appearance      =   0  'Flat
                  Caption         =   "Afecto al IVAP"
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
                  Left            =   480
                  TabIndex        =   99
                  Tag             =   "NIndAfectoIVAP"
                  Top             =   4260
                  Visible         =   0   'False
                  Width           =   1440
               End
               Begin VB.CheckBox ChkIndRptGL 
                  Appearance      =   0  'Flat
                  Caption         =   "Reporte Genética Líquida"
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
                  Left            =   6390
                  TabIndex        =   96
                  Tag             =   "TIndRptGL"
                  Top             =   4560
                  Visible         =   0   'False
                  Width           =   2340
               End
               Begin VB.CommandButton CmdAyudaDetraccion 
                  Height          =   315
                  Left            =   8820
                  Picture         =   "frmMantProductos.frx":4BDC
                  Style           =   1  'Graphical
                  TabIndex        =   91
                  Top             =   4350
                  Visible         =   0   'False
                  Width           =   390
               End
               Begin VB.CheckBox chkAfecto 
                  Alignment       =   1  'Right Justify
                  Appearance      =   0  'Flat
                  Caption         =   "Afecto al IGV"
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
                  Left            =   7350
                  TabIndex        =   9
                  Tag             =   "NafectoIGV"
                  Top             =   2460
                  Width           =   1440
               End
               Begin VB.CommandButton cmbAyudaUMVenta 
                  Height          =   315
                  Left            =   8400
                  Picture         =   "frmMantProductos.frx":4F66
                  Style           =   1  'Graphical
                  TabIndex        =   48
                  Top             =   2085
                  Width           =   390
               End
               Begin VB.CommandButton cmbAyudaUMCompra 
                  Height          =   315
                  Left            =   8400
                  Picture         =   "frmMantProductos.frx":52F0
                  Style           =   1  'Graphical
                  TabIndex        =   47
                  Top             =   1710
                  Width           =   390
               End
               Begin VB.CommandButton cmbAyudaMoneda 
                  Height          =   315
                  Left            =   8400
                  Picture         =   "frmMantProductos.frx":567A
                  Style           =   1  'Graphical
                  TabIndex        =   46
                  Top             =   1335
                  Width           =   390
               End
               Begin VB.CommandButton cmbAyudaMarca 
                  Height          =   315
                  Left            =   8400
                  Picture         =   "frmMantProductos.frx":5A04
                  Style           =   1  'Graphical
                  TabIndex        =   45
                  Top             =   960
                  Width           =   390
               End
               Begin VB.CommandButton cmbAyudaTipoProd 
                  Height          =   315
                  Left            =   8400
                  Picture         =   "frmMantProductos.frx":5D8E
                  Style           =   1  'Graphical
                  TabIndex        =   44
                  Top             =   585
                  Width           =   390
               End
               Begin VB.CommandButton cmbImagen 
                  Height          =   315
                  Left            =   4470
                  Picture         =   "frmMantProductos.frx":6118
                  Style           =   1  'Graphical
                  TabIndex        =   43
                  Top             =   2430
                  Width           =   390
               End
               Begin VB.CommandButton cmbMostrarImagen 
                  DownPicture     =   "frmMantProductos.frx":64A2
                  BeginProperty Font 
                     Name            =   "Arial"
                     Size            =   9
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   315
                  Left            =   5470
                  Picture         =   "frmMantProductos.frx":684D
                  Style           =   1  'Graphical
                  TabIndex        =   42
                  Top             =   2430
                  Width           =   345
               End
               Begin VB.CommandButton cmbAyudaGrupo 
                  Height          =   315
                  Left            =   8400
                  Picture         =   "frmMantProductos.frx":6BF8
                  Style           =   1  'Graphical
                  TabIndex        =   41
                  Top             =   2835
                  Visible         =   0   'False
                  Width           =   390
               End
               Begin VB.Frame fraPrecios 
                  Appearance      =   0  'Flat
                  ForeColor       =   &H80000008&
                  Height          =   795
                  Left            =   45
                  TabIndex        =   36
                  Top             =   5220
                  Visible         =   0   'False
                  Width           =   8790
                  Begin CATControls.CATTextBox txtVal_VV 
                     Height          =   315
                     Left            =   990
                     TabIndex        =   17
                     Top             =   280
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
                     Container       =   "frmMantProductos.frx":6F82
                     Text            =   "0.00"
                     Decimales       =   2
                     Estilo          =   4
                     Vacio           =   -1  'True
                     EnterTab        =   -1  'True
                  End
                  Begin CATControls.CATTextBox txtVal_IGV 
                     Height          =   315
                     Left            =   3285
                     TabIndex        =   18
                     Top             =   280
                     Width           =   1290
                     _ExtentX        =   2275
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
                     Alignment       =   1
                     FontName        =   "Arial"
                     FontSize        =   8.25
                     ForeColor       =   -2147483640
                     Container       =   "frmMantProductos.frx":6F9E
                     Text            =   "0.00"
                     Decimales       =   2
                     Estilo          =   4
                     Vacio           =   -1  'True
                     EnterTab        =   -1  'True
                  End
                  Begin CATControls.CATTextBox txtVal_PV 
                     Height          =   315
                     Left            =   5445
                     TabIndex        =   19
                     Top             =   280
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
                     Container       =   "frmMantProductos.frx":6FBA
                     Text            =   "0.00"
                     Decimales       =   2
                     Estilo          =   4
                     Vacio           =   -1  'True
                     EnterTab        =   -1  'True
                  End
                  Begin CATControls.CATTextBox TxtDctoListaPrec 
                     Height          =   315
                     Left            =   7830
                     TabIndex        =   20
                     Top             =   280
                     Width           =   465
                     _ExtentX        =   820
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
                     MaxLength       =   8
                     Container       =   "frmMantProductos.frx":6FD6
                     Estilo          =   3
                     Vacio           =   -1  'True
                     EnterTab        =   -1  'True
                  End
                  Begin VB.Label Label18 
                     Appearance      =   0  'Flat
                     AutoSize        =   -1  'True
                     Caption         =   "V.V. Unit."
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
                     TabIndex        =   40
                     Top             =   345
                     Width           =   690
                  End
                  Begin VB.Label Label19 
                     Appearance      =   0  'Flat
                     AutoSize        =   -1  'True
                     Caption         =   "I.G.V. Unit."
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
                     Left            =   2430
                     TabIndex        =   39
                     Top             =   345
                     Width           =   765
                  End
                  Begin VB.Label Label20 
                     Appearance      =   0  'Flat
                     AutoSize        =   -1  'True
                     Caption         =   "P.V. Unit."
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
                     Left            =   4680
                     TabIndex        =   38
                     Top             =   345
                     Width           =   660
                  End
                  Begin VB.Label Label15 
                     Appearance      =   0  'Flat
                     AutoSize        =   -1  'True
                     Caption         =   "Max % Dcto"
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
                     Left            =   6885
                     TabIndex        =   37
                     Top             =   345
                     Width           =   870
                  End
               End
               Begin VB.CheckBox chkInsertaPrecios 
                  Appearance      =   0  'Flat
                  Caption         =   "Insertar Precios Lista Principal"
                  BeginProperty Font 
                     Name            =   "Arial"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H00000000&
                  Height          =   240
                  Left            =   3240
                  TabIndex        =   15
                  Tag             =   "NindInsertaPrecioLista"
                  Top             =   4950
                  Visible         =   0   'False
                  Width           =   2640
               End
               Begin VB.Frame Frame4 
                  Appearance      =   0  'Flat
                  Caption         =   " Estado "
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
                  Height          =   750
                  Left            =   120
                  TabIndex        =   35
                  Top             =   3210
                  Width           =   8790
                  Begin VB.OptionButton OptActivo 
                     Caption         =   "Activo"
                     BeginProperty Font 
                        Name            =   "Arial"
                        Size            =   8.25
                        Charset         =   0
                        Weight          =   400
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     Height          =   240
                     Left            =   1200
                     TabIndex        =   13
                     Top             =   360
                     Value           =   -1  'True
                     Width           =   930
                  End
                  Begin VB.OptionButton OptInactivo 
                     Caption         =   "Inactivo"
                     BeginProperty Font 
                        Name            =   "Arial"
                        Size            =   8.25
                        Charset         =   0
                        Weight          =   400
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     Height          =   240
                     Left            =   6225
                     TabIndex        =   14
                     Top             =   360
                     Width           =   1005
                  End
               End
               Begin VB.CheckBox chkDctoEspecial 
                  Appearance      =   0  'Flat
                  Caption         =   "Afecto al Dscto. Especial"
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
                  Left            =   6615
                  TabIndex        =   16
                  Tag             =   "NafectoDctoEspecial"
                  Top             =   4905
                  Visible         =   0   'False
                  Width           =   2160
               End
               Begin CATControls.CATTextBox txtGls_Producto 
                  Height          =   315
                  Left            =   1300
                  TabIndex        =   2
                  Tag             =   "TGlsProducto"
                  Top             =   210
                  Width           =   7500
                  _ExtentX        =   13229
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
                  Container       =   "frmMantProductos.frx":6FF2
                  EnterTab        =   -1  'True
               End
               Begin CATControls.CATTextBox txtCod_TipoProd 
                  Height          =   315
                  Left            =   1300
                  TabIndex        =   3
                  Tag             =   "TidTipoProducto"
                  Top             =   585
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
                  Container       =   "frmMantProductos.frx":700E
                  Estilo          =   1
                  EnterTab        =   -1  'True
               End
               Begin CATControls.CATTextBox txtGls_TipoProd 
                  Height          =   315
                  Left            =   2235
                  TabIndex        =   49
                  Top             =   585
                  Width           =   6135
                  _ExtentX        =   10821
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
                  Container       =   "frmMantProductos.frx":702A
                  Vacio           =   -1  'True
               End
               Begin CATControls.CATTextBox txtCod_Marca 
                  Height          =   315
                  Left            =   1300
                  TabIndex        =   4
                  Tag             =   "TidMarca"
                  Top             =   960
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
                  Container       =   "frmMantProductos.frx":7046
                  Estilo          =   1
                  EnterTab        =   -1  'True
               End
               Begin CATControls.CATTextBox txtGls_Marca 
                  Height          =   315
                  Left            =   2235
                  TabIndex        =   50
                  Top             =   960
                  Width           =   6135
                  _ExtentX        =   10821
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
                  Container       =   "frmMantProductos.frx":7062
                  Vacio           =   -1  'True
               End
               Begin CATControls.CATTextBox txtCod_Moneda 
                  Height          =   315
                  Left            =   1300
                  TabIndex        =   5
                  Tag             =   "TidMoneda"
                  Top             =   1335
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
                  Container       =   "frmMantProductos.frx":707E
                  Estilo          =   1
                  EnterTab        =   -1  'True
               End
               Begin CATControls.CATTextBox txtGls_Moneda 
                  Height          =   315
                  Left            =   2235
                  TabIndex        =   51
                  Top             =   1335
                  Width           =   6135
                  _ExtentX        =   10821
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
                  Container       =   "frmMantProductos.frx":709A
                  Vacio           =   -1  'True
               End
               Begin CATControls.CATTextBox txtCod_UMCompra 
                  Height          =   315
                  Left            =   1300
                  TabIndex        =   6
                  Tag             =   "TidUMCompra"
                  Top             =   1710
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
                  Container       =   "frmMantProductos.frx":70B6
                  Estilo          =   1
                  EnterTab        =   -1  'True
               End
               Begin CATControls.CATTextBox txtGls_UMCompra 
                  Height          =   315
                  Left            =   2235
                  TabIndex        =   52
                  Top             =   1710
                  Width           =   6135
                  _ExtentX        =   10821
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
                  Container       =   "frmMantProductos.frx":70D2
                  Vacio           =   -1  'True
               End
               Begin CATControls.CATTextBox txtCod_UMVenta 
                  Height          =   315
                  Left            =   1300
                  TabIndex        =   7
                  Tag             =   "TidUMVenta"
                  Top             =   2085
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
                  Container       =   "frmMantProductos.frx":70EE
                  Estilo          =   1
                  EnterTab        =   -1  'True
               End
               Begin CATControls.CATTextBox txtGls_UMVenta 
                  Height          =   315
                  Left            =   2235
                  TabIndex        =   53
                  Top             =   2085
                  Width           =   6135
                  _ExtentX        =   10821
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
                  Container       =   "frmMantProductos.frx":710A
                  Vacio           =   -1  'True
               End
               Begin CATControls.CATTextBox txt_CodFabricante 
                  Height          =   315
                  Left            =   1300
                  TabIndex        =   8
                  Tag             =   "TidFabricante"
                  Top             =   2445
                  Width           =   2040
                  _ExtentX        =   3598
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
                  Container       =   "frmMantProductos.frx":7126
                  Estilo          =   1
                  Vacio           =   -1  'True
                  EnterTab        =   -1  'True
               End
               Begin CATControls.CATTextBox txtCod_Rapido 
                  Height          =   315
                  Left            =   1300
                  TabIndex        =   10
                  Tag             =   "TCodigoRapido"
                  Top             =   2835
                  Width           =   2040
                  _ExtentX        =   3598
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
                  MaxLength       =   16
                  Container       =   "frmMantProductos.frx":7142
                  Estilo          =   1
                  Vacio           =   -1  'True
                  EnterTab        =   -1  'True
               End
               Begin CATControls.CATTextBox TxtCod_Grupo 
                  Height          =   315
                  Left            =   4440
                  TabIndex        =   11
                  Tag             =   "TidGrupo"
                  Top             =   2835
                  Visible         =   0   'False
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
                  Container       =   "frmMantProductos.frx":715E
                  Estilo          =   1
                  Vacio           =   -1  'True
                  EnterTab        =   -1  'True
               End
               Begin CATControls.CATTextBox TxtGls_Grupo 
                  Height          =   315
                  Left            =   5400
                  TabIndex        =   54
                  Top             =   2835
                  Visible         =   0   'False
                  Width           =   2970
                  _ExtentX        =   5239
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
                  Container       =   "frmMantProductos.frx":717A
                  Vacio           =   -1  'True
               End
               Begin CATControls.CATTextBox txtCod_TallaPeso 
                  Height          =   315
                  Left            =   1305
                  TabIndex        =   12
                  Tag             =   "TidTallaPeso"
                  Top             =   3195
                  Visible         =   0   'False
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
                  Container       =   "frmMantProductos.frx":7196
                  Decimales       =   5
                  Estilo          =   4
                  Vacio           =   -1  'True
                  EnterTab        =   -1  'True
               End
               Begin CATControls.CATTextBox txtRep_Stock 
                  Height          =   315
                  Left            =   1665
                  TabIndex        =   68
                  Tag             =   "TTiemporepinv"
                  Top             =   4860
                  Visible         =   0   'False
                  Width           =   735
                  _ExtentX        =   1296
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
                  Container       =   "frmMantProductos.frx":71B2
                  Estilo          =   1
                  Vacio           =   -1  'True
                  EnterTab        =   -1  'True
               End
               Begin CATControls.CATTextBox TxtCodDetraccion 
                  Height          =   315
                  Left            =   930
                  TabIndex        =   92
                  Tag             =   "TIdConceptoDetraccion"
                  Top             =   4980
                  Visible         =   0   'False
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
                  Container       =   "frmMantProductos.frx":71CE
                  Vacio           =   -1  'True
                  EnterTab        =   -1  'True
               End
               Begin CATControls.CATTextBox TxtGlsDetraccion 
                  Height          =   315
                  Left            =   2820
                  TabIndex        =   93
                  Top             =   4410
                  Visible         =   0   'False
                  Width           =   5640
                  _ExtentX        =   9948
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
                  Container       =   "frmMantProductos.frx":71EA
                  Vacio           =   -1  'True
               End
               Begin CATControls.CATTextBox TxtPorcentajeDetraccion 
                  Height          =   315
                  Left            =   8280
                  TabIndex        =   95
                  Top             =   4350
                  Visible         =   0   'False
                  Width           =   465
                  _ExtentX        =   820
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
                  Alignment       =   1
                  FontName        =   "Arial"
                  FontSize        =   8.25
                  ForeColor       =   -2147483640
                  Container       =   "frmMantProductos.frx":7206
                  Estilo          =   3
                  Vacio           =   -1  'True
                  EnterTab        =   -1  'True
               End
               Begin CATControls.CATTextBox TxtStockMinimo 
                  Height          =   315
                  Left            =   4635
                  TabIndex        =   97
                  Tag             =   "NStockMinimo"
                  Top             =   3195
                  Visible         =   0   'False
                  Width           =   735
                  _ExtentX        =   1296
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
                  Container       =   "frmMantProductos.frx":7222
                  Text            =   "0"
                  Estilo          =   4
                  EnterTab        =   -1  'True
               End
               Begin VB.Label LblStockMinimo 
                  Appearance      =   0  'Flat
                  AutoSize        =   -1  'True
                  Caption         =   "Stock Mínimo"
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
                  Left            =   3555
                  TabIndex        =   98
                  Top             =   3240
                  Visible         =   0   'False
                  Width           =   930
               End
               Begin VB.Label Label21 
                  Appearance      =   0  'Flat
                  AutoSize        =   -1  'True
                  Caption         =   "Conc. Detrac."
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
                  Left            =   -60
                  TabIndex        =   94
                  Top             =   4950
                  Visible         =   0   'False
                  Width           =   990
               End
               Begin VB.Label Label17 
                  Appearance      =   0  'Flat
                  AutoSize        =   -1  'True
                  Caption         =   "T. Reposición Stock"
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
                  Left            =   135
                  TabIndex        =   69
                  Top             =   4905
                  Visible         =   0   'False
                  Width           =   1425
               End
               Begin VB.Label Label1 
                  Appearance      =   0  'Flat
                  AutoSize        =   -1  'True
                  Caption         =   "Descripción"
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
                  TabIndex        =   66
                  Top             =   270
                  Width           =   855
               End
               Begin VB.Label Label2 
                  Appearance      =   0  'Flat
                  AutoSize        =   -1  'True
                  Caption         =   "Tipo Producto"
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
                  TabIndex        =   65
                  Top             =   630
                  Width           =   990
               End
               Begin VB.Label Label4 
                  Appearance      =   0  'Flat
                  AutoSize        =   -1  'True
                  Caption         =   "Marca"
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
                  TabIndex        =   64
                  Top             =   1005
                  Width           =   450
               End
               Begin VB.Label Label5 
                  Appearance      =   0  'Flat
                  AutoSize        =   -1  'True
                  Caption         =   "Moneda"
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
                  TabIndex        =   63
                  Top             =   1380
                  Width           =   570
               End
               Begin VB.Label Label7 
                  Appearance      =   0  'Flat
                  AutoSize        =   -1  'True
                  Caption         =   "U.M. Compra"
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
                  TabIndex        =   62
                  Top             =   1755
                  Width           =   915
               End
               Begin VB.Label Label8 
                  Appearance      =   0  'Flat
                  AutoSize        =   -1  'True
                  Caption         =   "U.M. Venta"
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
                  TabIndex        =   61
                  Top             =   2130
                  Width           =   795
               End
               Begin VB.Label Label10 
                  Appearance      =   0  'Flat
                  AutoSize        =   -1  'True
                  Caption         =   "Cod. Fab."
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
                  TabIndex        =   60
                  Top             =   2490
                  Width           =   690
               End
               Begin VB.Label Label9 
                  Appearance      =   0  'Flat
                  AutoSize        =   -1  'True
                  Caption         =   "Imagen"
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
                  Left            =   3570
                  TabIndex        =   59
                  Top             =   2490
                  Width           =   510
               End
               Begin VB.Label lbl_Imagen 
                  Appearance      =   0  'Flat
                  BackColor       =   &H003161DD&
                  BeginProperty Font 
                     Name            =   "Arial"
                     Size            =   1.5
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H80000008&
                  Height          =   315
                  Left            =   4890
                  TabIndex        =   58
                  Top             =   2430
                  Width           =   630
               End
               Begin VB.Label Label11 
                  Appearance      =   0  'Flat
                  AutoSize        =   -1  'True
                  Caption         =   "Cod. Rápido"
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
                  TabIndex        =   57
                  Top             =   2880
                  Width           =   870
               End
               Begin VB.Label Label12 
                  Appearance      =   0  'Flat
                  AutoSize        =   -1  'True
                  Caption         =   "Grupo"
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
                  Left            =   3570
                  TabIndex        =   56
                  Top             =   2880
                  Visible         =   0   'False
                  Width           =   450
               End
               Begin VB.Label Label13 
                  Appearance      =   0  'Flat
                  AutoSize        =   -1  'True
                  Caption         =   "Peso"
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
                  TabIndex        =   55
                  Top             =   3255
                  Visible         =   0   'False
                  Width           =   360
               End
            End
            Begin VB.CheckBox ChkConservarDatos 
               Appearance      =   0  'Flat
               Caption         =   "Conservar Datos"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000000&
               Height          =   240
               Left            =   450
               TabIndex        =   32
               Top             =   250
               Width           =   1695
            End
            Begin MSComDlg.CommonDialog cd 
               Left            =   6525
               Top             =   135
               _ExtentX        =   847
               _ExtentY        =   847
               _Version        =   393216
            End
            Begin CATControls.CATTextBox txtCod_Producto 
               Height          =   315
               Left            =   8265
               TabIndex        =   33
               Tag             =   "TidProducto"
               Top             =   250
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
               Container       =   "frmMantProductos.frx":723E
               Estilo          =   1
               Vacio           =   -1  'True
               EnterTab        =   -1  'True
            End
            Begin VB.Label Label6 
               Appearance      =   0  'Flat
               AutoSize        =   -1  'True
               Caption         =   "Código"
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
               Left            =   7605
               TabIndex        =   67
               Top             =   300
               Width           =   495
            End
         End
         Begin VB.Frame Frame3 
            Appearance      =   0  'Flat
            ForeColor       =   &H00000000&
            Height          =   5085
            Left            =   -74640
            TabIndex        =   30
            Top             =   645
            Width           =   11400
            Begin DXDBGRIDLibCtl.dxDBGrid gPresentaciones 
               Height          =   4665
               Left            =   165
               OleObjectBlob   =   "frmMantProductos.frx":725A
               TabIndex        =   22
               Top             =   270
               Width           =   11130
            End
         End
         Begin VB.Frame Frame2 
            Appearance      =   0  'Flat
            ForeColor       =   &H00000000&
            Height          =   7155
            Left            =   -74775
            TabIndex        =   29
            Top             =   420
            Width           =   11760
            Begin DXDBGRIDLibCtl.dxDBGrid gAlmacenes 
               Height          =   6690
               Left            =   180
               OleObjectBlob   =   "frmMantProductos.frx":9320
               TabIndex        =   21
               Top             =   270
               Width           =   11400
            End
         End
      End
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   1230
      Left            =   0
      TabIndex        =   26
      Top             =   0
      Width           =   12810
      _ExtentX        =   22595
      _ExtentY        =   2170
      ButtonWidth     =   2540
      ButtonHeight    =   1005
      Appearance      =   1
      ImageList       =   "imgDocVentas"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   9
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "         Nuevo         "
            Object.ToolTipText     =   "Nuevo"
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
            Object.ToolTipText     =   "Eliminar"
            ImageIndex      =   6
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Imprimir"
            Object.ToolTipText     =   "Imprimir"
            ImageIndex      =   11
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Codigo de Barra"
            Object.ToolTipText     =   "Imprimir codigo de barras"
            ImageIndex      =   11
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Lista"
            Object.ToolTipText     =   "Lista"
            ImageIndex      =   12
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Salir"
            Object.ToolTipText     =   "Salir"
            ImageIndex      =   2
         EndProperty
      EndProperty
      BorderStyle     =   1
   End
   Begin VB.Frame fraListado 
      Appearance      =   0  'Flat
      ForeColor       =   &H00000000&
      Height          =   8685
      Left            =   45
      TabIndex        =   23
      Top             =   675
      Width           =   12705
      Begin VB.Frame Frame1 
         Appearance      =   0  'Flat
         ForeColor       =   &H80000008&
         Height          =   705
         Left            =   120
         TabIndex        =   24
         Top             =   150
         Width           =   12465
         Begin CATControls.CATTextBox txt_TextoBuscar 
            Height          =   315
            Left            =   1035
            TabIndex        =   1
            Top             =   255
            Width           =   11280
            _ExtentX        =   19897
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
            Container       =   "frmMantProductos.frx":BEDC
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
            TabIndex        =   0
            Top             =   300
            Width           =   735
         End
      End
      Begin DXDBGRIDLibCtl.dxDBGrid gLista 
         Height          =   7545
         Left            =   120
         OleObjectBlob   =   "frmMantProductos.frx":BEF8
         TabIndex        =   25
         Top             =   990
         Width           =   12480
      End
   End
End
Attribute VB_Name = "frmMantProductos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim numPesos As Integer
Dim indCargando As Boolean
Dim indIngresoImagen As Boolean
Dim indCalculando As Boolean

Private Sub chkAfecto_Click()
On Error GoTo Err
Dim StrMsgError                                 As String
    
    If chkAfecto.Value = 1 Then
        ChkAfectoIVAP.Value = 0
    End If
    
    Exit Sub
Err:
    If StrMsgError = "" Then StrMsgError = Err.Description
    MsgBox StrMsgError, vbInformation, App.Title
End Sub

Private Sub ChkAfectoIVAP_Click()
On Error GoTo Err
Dim StrMsgError                                 As String
    
    If ChkAfectoIVAP.Value = 1 Then
        chkAfecto.Value = 0
    End If
    
    Exit Sub
Err:
    If StrMsgError = "" Then StrMsgError = Err.Description
    MsgBox StrMsgError, vbInformation, App.Title
End Sub

Private Sub chkInsertaPrecios_Click()
    
    fraPrecios.Enabled = chkInsertaPrecios.Value
    txtVal_VV.Text = 0
    txtVal_IGV.Text = 0
    txtVal_PV.Text = 0
    
    If chkInsertaPrecios.Value Then
        txtVal_VV.Vacio = True
        txtVal_IGV.Vacio = True
        txtVal_PV.Vacio = True
    Else
        txtVal_VV.Vacio = True
        txtVal_IGV.Vacio = True
        txtVal_PV.Vacio = True
    End If

End Sub

Private Sub cmbAyudaGrupo_Click()
    
    mostrarAyuda "GRUPOSPRODUCTO", TxtCod_Grupo, TxtGls_Grupo

End Sub

Private Sub CmbAyudaMarca_Click()
    
    mostrarAyuda "MARCA", txtCod_Marca, txtGls_Marca

End Sub

Private Sub cmbAyudaMoneda_Click()
    
    mostrarAyuda "MONEDA", txtCod_Moneda, txtGls_Moneda

End Sub

Private Sub cmbAyudaNivel_Click(Index As Integer)
Dim peso As Integer
Dim strCodTipoNivel As String
Dim strCondPred As String
    
    peso = Index + 1
    strCodTipoNivel = traerCampo("tiposniveles", "idTipoNivel", "peso", CStr(peso), True)
    strCondPred = ""
    If peso > 1 Then
        strCondPred = " AND idNivelPred = '" & txtCod_Nivel(Index - 1).Text & "'"
    End If
    mostrarAyuda "NIVEL", txtCod_Nivel(Index), txtGls_Nivel(Index), " AND idTipoNivel = '" & strCodTipoNivel & "'" & strCondPred
    
End Sub

Private Sub cmbAyudaTallaPeso_Click()
    
    mostrarAyuda "TALLASPESOS", txtCod_TallaPeso, txtGls_TallaPeso

End Sub

Private Sub cmbAyudaTipoProd_Click()
    
    mostrarAyuda "TIPOPRODUCTO", txtCod_TipoProd, txtGls_TipoProd

End Sub

Private Sub cmbAyudaUMCompra_Click()
On Error GoTo Err

    If Len(Trim(traerCampo("ValesDet d Inner Join Valescab c On d.idValescab = c.idValescab And d.tipoVale = c.tipoVale And d.idEmpresa = c.idEmpresa And d.idSucursal = c.idSucursal", "idProducto", "idProducto", txtCod_Producto.Text, False, "c.idEmpresa='" & glsEmpresa & "' And c.idSucursal = '" & glsSucursal & "' And c.estValeCab<> 'ANU' "))) > 0 Then
        StrMsgError = "No se puede Modificar La Unidad de Medida del Producto se encuentra en Uso Inventarios ": Err
    Else
       mostrarAyuda "UMGLOSA", txtCod_UMCompra, txtGls_UMCompra
    End If
    If StrMsgError <> "" Then GoTo Err
    
    Exit Sub

Err:
    If StrMsgError = "" Then StrMsgError = Err.Description
    MsgBox StrMsgError, vbInformation, App.Title
End Sub

Private Sub cmbAyudaUMVenta_Click()
On Error GoTo Err

    If Len(Trim(traerCampo("ValesDet d Inner Join Valescab c On d.idValescab = c.idValescab And d.tipoVale = c.tipoVale And d.idEmpresa = c.idEmpresa And d.idSucursal = c.idSucursal", "idProducto", "idProducto", txtCod_Producto.Text, False, "c.idEmpresa='" & glsEmpresa & "' And c.idSucursal = '" & glsSucursal & "' And c.estValeCab<> 'ANU' "))) > 0 Then
        StrMsgError = "No se puede Modificar La Unidad de Medida del Producto se encuentra en Uso Inventarios": Err
    Else
        mostrarAyuda "UMGLOSA", txtCod_UMVenta, txtGls_UMVenta
    End If

    If StrMsgError <> "" Then GoTo Err
    
    Exit Sub

Err:
    If StrMsgError = "" Then StrMsgError = Err.Description
    MsgBox StrMsgError, vbInformation, App.Title
End Sub

Private Sub cmbImagen_Click()

    cd.Filter = "Imagenes (*.jpg)|*.jpg"
    cd.ShowOpen
    
    If cd.FileName <> "" Then
        lbl_Imagen.Caption = cd.FileName
        indIngresoImagen = True
    End If

End Sub

Private Sub cmbMostrarImagen_Click()

    If lbl_Imagen.Caption <> "" Then
        frmMuestraImagen.MostrarForm lbl_Imagen.Caption
    Else
        MsgBox "No existe una imagen asignada", vbInformation, App.Title
    End If

End Sub

'Private Sub cmdayuda1_Click()
'On Error GoTo Err
'Dim strCod As String
'Dim StrDes As String
'Dim StrMsgError As String
'
'    mostrarAyudaTextoPlanCuentas strcnConta, "", strCod, StrDes, "", IIf(Trim("" & Year(Date)) <= "2010", "2010", "2011")
'    If StrMsgError <> "" Then GoTo Err
'    txt_ctacontable1.Text = "" & strCod
'    txtdescta1.Text = "" & StrDes
'
'    Exit Sub
'
'Err:
'    If StrMsgError = "" Then StrMsgError = Err.Description
'    MsgBox StrMsgError, vbInformation, App.Title
'End Sub

'Private Sub cmdayuda2_Click()
'On Error GoTo Err
'Dim strCod As String
'Dim StrDes As String
'Dim StrMsgError As String
'
'    mostrarAyudaTextoPlanCuentas strcnConta, "", strCod, StrDes, "", IIf(Trim("" & Year(Date)) <= "2010", "2010", "2011")
'    If StrMsgError <> "" Then GoTo Err
'    txt_ctacontable2.Text = "" & strCod
'    txtdescta2.Text = "" & StrDes
'
'    Exit Sub
'
'Err:
'    If StrMsgError = "" Then StrMsgError = Err.Description
'    MsgBox StrMsgError, vbInformation, App.Title
'End Sub

'Private Sub cmdayuda3_Click()
'On Error GoTo Err
'Dim strCod As String
'Dim StrDes As String
'Dim StrMsgError As String
'
'    mostrarAyudaTextoPlanCuentas strcnConta, "", strCod, StrDes, "", IIf(Trim("" & Year(Date)) <= "2010", "2010", "2011")
'    If StrMsgError <> "" Then GoTo Err
'    txtctacontableRelacionada.Text = "" & strCod
'    txtdescta3.Text = "" & StrDes
'
'    Exit Sub
'
'Err:
'    If StrMsgError = "" Then StrMsgError = Err.Description
'    MsgBox StrMsgError, vbInformation, App.Title
'End Sub
'
Private Sub CmdAyudaConceptoCosteo_Click()
'On Error GoTo Err
'Dim StrMsgError                     As String
'
'    mostrarAyuda "CONCEPTOSCOSTEO", TxtIdConceptoCosteo, TxtGlsConceptoCosteo
'
'Exit Sub
'Err:
'    If StrMsgError = "" Then StrMsgError = Err.Description
'    MsgBox StrMsgError, vbInformation, App.Title
End Sub

Private Sub CmdAyudaDetraccion_Click()
On Error GoTo Err
Dim StrMsgError                     As String
Dim CIdDetraccion                   As String

    ayuda_detraccion.MostrarForm StrMsgError, CIdDetraccion
    If StrMsgError <> "" Then GoTo Err
    
    If CIdDetraccion <> "" Then
            
        TxtCodDetraccion.Text = CIdDetraccion
        
    End If
    
    Exit Sub
Err:
    If StrMsgError = "" Then StrMsgError = Err.Description
    MsgBox StrMsgError, vbInformation, App.Title
End Sub

Private Sub Form_Load()
On Error GoTo Err
Dim StrMsgError As String

    ChkConservarDatos.Value = 0
    chkInsertaPrecios.Value = 0
    
    If leeParametro("VISUALIZA_STOCK_MINIMO") = "S" Then
        
        LblStockMinimo.Visible = True
        TxtStockMinimo.Visible = True
    
    Else
        
        LblStockMinimo.Visible = False
        TxtStockMinimo.Visible = False
        
    End If
    
    If leeParametro("VISUALIZA_REPORTE_GL") = "S" Then
        
        ChkIndRptGL.Visible = True
    
    Else
        
        ChkIndRptGL.Visible = False
        
    End If
    
    Me.top = 0
    Me.left = 0
    
    ConfGrid gLista, False, False, False, False
    ConfGrid gAlmacenes, True, False, False, False
    ConfGrid gPresentaciones, True, False, False, False
    
    If leeParametro("VIZUALIZA_CODIGO_RAPIDO") = "S" Then
        gLista.Columns.ColumnByFieldName("IdProducto").Visible = False
        gLista.Columns.ColumnByFieldName("CodigoRapido").Visible = True
    Else
        gLista.Columns.ColumnByFieldName("IdProducto").Visible = True
        gLista.Columns.ColumnByFieldName("CodigoRapido").Visible = False
    End If

    mostrarNiveles StrMsgError
    If StrMsgError <> "" Then GoTo Err
    fraListado.Visible = True
    fraGeneral.Visible = False
    habilitaBotones 8
    nuevo
    
    Exit Sub
Err:
    If StrMsgError = "" Then StrMsgError = Err.Description
    MsgBox StrMsgError, vbInformation, App.Title
End Sub

Private Sub Grabar(ByRef StrMsgError As String)
On Error GoTo Err
Dim strCodigo       As String
Dim strMsg          As String
Dim rsj             As New ADODB.Recordset
Dim i               As Integer
Dim numPesos        As Integer
Dim xCodigoRapido   As String
    
    xCodigoRapido = ""
    
    validaFormSQL Me, StrMsgError
    If StrMsgError <> "" Then GoTo Err
    
    validaHomonimia "productos", "GlsProducto", "idProducto", txtGls_Producto.Text, txtCod_Producto.Text, True, StrMsgError, " idMarca = '" & txtCod_Marca.Text & "'"
    If StrMsgError <> "" Then GoTo Err
    
    If traerCampo("Parametros", "Valparametro", "GlsParametro", "VALIDA_CODIGO_FABRICANTE", True) = "S" Then
        validaHomonimia "productos", "IdFabricante", "idProducto", txt_CodFabricante.Text, txtCod_Producto.Text, True, StrMsgError
        If StrMsgError <> "" Then GoTo Err
    End If
    
    If txtCod_Producto.Text = "" Then 'graba
        txtCod_Producto.Text = GeneraCorrelativoAnoMes("productos", "idProducto")
        EjecutaSQLFormProducto 0, True, "productos", StrMsgError, ""
        If StrMsgError <> "" Then GoTo Err
        
        copiaImagen StrMsgError
        If StrMsgError <> "" Then GoTo Err
        
        If OptActivo.Value = True Then
            csql = "update productos set estProducto = 'A' where idProducto = '" & txtCod_Producto.Text & "' and idEmpresa = '" & glsEmpresa & "' "
            Cn.Execute csql
        Else
            csql = "update productos set estProducto = 'I' where idProducto = '" & txtCod_Producto.Text & "' and idEmpresa = '" & glsEmpresa & "' "
            Cn.Execute csql
        End If
            
        xCodigoRapido = ""
        If Trim("" & traerCampo("Parametros", "ValParametro", "GlsParametro", "VIZUALIZA_CODIGO_RAPIDO", True)) = "S" Then
        
            If rsj.State = 1 Then rsj.Close
            Set rsj = Nothing
            
            rsj.Open "SELECT GlsTipoNivel FROM tiposniveles WHERE idEmpresa = '" & glsEmpresa & "' Order BY Peso ASC", Cn, adOpenForwardOnly, adLockReadOnly
            numPesos = Val("" & rsj.RecordCount)
            
            For i = 0 To numPesos - 1
                xCodigoRapido = xCodigoRapido & Trim("" & traerCampo("Niveles", "GlsAbreviatura", "idnivel", txtCod_Nivel(i).Text, True))
            Next
            
            xCodigoRapido = xCodigoRapido & Format(Val("" & traerCampo("Productos", "Cast(Right(ConCat('0',IfNull(CodigoRapido,'0')),4) As Unsigned)", "IdNivel", txtCod_Nivel(numPesos - 1).Text, False, "IdEmpresa = '" & glsEmpresa & "' Order By 1 Desc Limit 1")) + 1, "0000")
            
            'If xCodigoRapido = Trim("" & traerCampo("Productos", "Left(CodigoRapido,Length(CodigoRapido) - 4)", "Left(CodigoRapido,Length(CodigoRapido) - 4)", xCodigoRapido, False, "IdEmpresa = '" & glsEmpresa & "' Order By Right(Trim(IfNull(CodigoRapido,'')),4) Desc Limit 1")) Then
            '    xCodigoRapido = xCodigoRapido & Format(Val(Trim("" & traerCampo("Productos", "right(trim(ifnull(CodigoRapido,'')),4)", "left(CodigoRapido,length(CodigoRapido) - 4)", xCodigoRapido, False, " idempresa = '" & glsEmpresa & "' order by right(trim(ifnull(CodigoRapido,'')),4) desc limit 1")) + 1), "0000")
            'Else
            '    xCodigoRapido = xCodigoRapido & "0001"
            'End If
            
            csql = "Update productos set CodigoRapido = '" & xCodigoRapido & "' where idProducto = '" & txtCod_Producto.Text & "' and idEmpresa = '" & glsEmpresa & "' "
            Cn.Execute csql
            
            txtCod_Rapido.Text = xCodigoRapido
        
        End If
        
        strMsg = "Grabó"
    
    Else 'modifica
    
        EjecutaSQLFormProducto 1, True, "productos", StrMsgError, "idProducto"
        If StrMsgError <> "" Then GoTo Err
        
        copiaImagen StrMsgError
        If StrMsgError <> "" Then GoTo Err
        
        If OptActivo.Value = True Then
            csql = "update productos set estProducto = 'A' where idProducto = '" & txtCod_Producto.Text & "' "
            Cn.Execute csql
        Else
            csql = "update productos set estProducto = 'I' where idProducto = '" & txtCod_Producto.Text & "' "
            Cn.Execute csql
        End If
        
        strMsg = "Modificó"
        
    End If
    
    MsgBox "Se " & strMsg & " Satisfactoriamente", vbInformation, App.Title
    If StrMsgError <> "" Then GoTo Err
    
    Exit Sub

Err:
    If StrMsgError = "" Then StrMsgError = Err.Description
    Exit Sub
    Resume
End Sub

Private Sub nuevo()
Dim rst As New ADODB.Recordset
Dim rsu As New ADODB.Recordset
Dim StrMsgError As String
    
    limpiaForm Me
    chkAfecto.Value = 1
    ChkIndRptGL.Value = 0
    indIngresoImagen = False
    lbl_Imagen.Caption = ""
    
    '--- ALMACEN
    rst.Fields.Append "Item", adInteger, , adFldRowID
    rst.Fields.Append "idSucursal", adVarChar, 8, adFldIsNullable
    rst.Fields.Append "GlsSucursal", adVarChar, 180, adFldIsNullable
    rst.Fields.Append "idAlmacen", adVarChar, 8, adFldIsNullable
    rst.Fields.Append "GlsAlmacen", adVarChar, 180, adFldIsNullable
    rst.Fields.Append "idUbicacion", adVarChar, 20, adFldIsNullable
    rst.Open
    
    rst.AddNew
    rst.Fields("Item") = 1
    rst.Fields("idSucursal") = ""
    rst.Fields("GlsSucursal") = ""
    rst.Fields("idAlmacen") = ""
    rst.Fields("GlsAlmacen") = ""
    rst.Fields("idUbicacion") = ""
    
    mostrarDatosGridSQL gAlmacenes, rst, StrMsgError
    If StrMsgError <> "" Then GoTo Err
    
    gAlmacenes.Columns.FocusedIndex = gAlmacenes.Columns.ColumnByFieldName("idAlmacen").Index
   
    '--- PRESENTACIONES
    rsu.Fields.Append "Item", adInteger, , adFldRowID
    rsu.Fields.Append "idUM", adVarChar, 8, adFldIsNullable
    rsu.Fields.Append "GlsUM", adVarChar, 250, adFldIsNullable
    rsu.Fields.Append "Factor", adDouble, 14, adFldIsNullable
    rsu.Open
    
    rsu.AddNew
    rsu.Fields("Item") = 1
    rsu.Fields("idUM") = ""
    rsu.Fields("GlsUM") = ""
    rsu.Fields("Factor") = 0
    
    mostrarDatosGridSQL gPresentaciones, rsu, StrMsgError
    If StrMsgError <> "" Then GoTo Err
    
    Exit Sub
    
Err:
    If StrMsgError = "" Then StrMsgError = Err.Description
End Sub

Private Sub gAlmacenes_OnAfterDatasetAction(ByVal Action As DXDBGRIDLibCtl.ExDatasetAction)
Dim i As Integer

    If Action = daInsert Then
        gAlmacenes.Columns.ColumnByFieldName("item").Value = gAlmacenes.Count
        gAlmacenes.Dataset.Post
    End If

End Sub

Private Sub gAlmacenes_OnBeforeDatasetAction(ByVal Action As DXDBGRIDLibCtl.ExDatasetAction, Allow As Boolean)

    If Action = daInsert Then
        If gAlmacenes.Columns.ColumnByFieldName("idSucursal").Value = "" Or gAlmacenes.Columns.ColumnByFieldName("idAlmacen").Value = "" Then
            Allow = False
        Else
            gAlmacenes.Columns.FocusedIndex = gAlmacenes.Columns.ColumnByFieldName("idSucursal").Index
        End If
    End If

End Sub

Private Sub gAlmacenes_OnEditButtonClick(ByVal Column As DXDBGRIDLibCtl.IdxGridColumn, ByVal Node As DXDBGRIDLibCtl.IdxGridNode)
Dim strCod As String
Dim StrDes As String
    
    Select Case Column.Index
        Case gAlmacenes.Columns.ColumnByFieldName("idSucursal").Index
            strCod = gAlmacenes.Columns.ColumnByFieldName("idSucursal").Value
            StrDes = gAlmacenes.Columns.ColumnByFieldName("GlsSucursal").Value
            mostrarAyudaTexto "SUCURSAL", strCod, StrDes
            gAlmacenes.Dataset.Edit
            gAlmacenes.Columns.ColumnByFieldName("idSucursal").Value = strCod
            gAlmacenes.Columns.ColumnByFieldName("GlsSucursal").Value = StrDes
            gAlmacenes.Dataset.Post
        
        Case gAlmacenes.Columns.ColumnByFieldName("idAlmacen").Index
            strCod = gAlmacenes.Columns.ColumnByFieldName("idAlmacen").Value
            StrDes = gAlmacenes.Columns.ColumnByFieldName("GlsAlmacen").Value
            mostrarAyudaTexto "ALMACENVTA", strCod, StrDes, " AND idSucursal = '" & gAlmacenes.Columns.ColumnByFieldName("idSucursal").Value & "'"
            If existeEnGrilla(gAlmacenes, "idAlmacen", strCod) = False Then
                gAlmacenes.Dataset.Edit
                gAlmacenes.Columns.ColumnByFieldName("idAlmacen").Value = strCod
                gAlmacenes.Columns.ColumnByFieldName("GlsAlmacen").Value = StrDes
                gAlmacenes.Dataset.Post
            Else
                MsgBox "El Almacén ya fue ingresado.", vbInformation, App.Title
            End If
                    
        Case gAlmacenes.Columns.ColumnByFieldName("idUbicacion").Index
            strCod = gAlmacenes.Columns.ColumnByFieldName("idUbicacion").Value
            mostrarAyudaTexto "UBICACIONES", strCod, StrDes, " AND idalmacen = '" & gAlmacenes.Columns.ColumnByFieldName("idAlmacen").Value & "'"
            
            If Len(Trim("" & strCod)) > 0 Then
            
                If existeEnGrilla(gAlmacenes, "idUbicacion", strCod) = False Then
                    gAlmacenes.Dataset.Edit
                    gAlmacenes.Columns.ColumnByFieldName("idUbicacion").Value = strCod
                    gAlmacenes.Dataset.Post
                Else
                    MsgBox "La Ubicacion ya fue ingresada.", vbInformation, App.Title
                End If
            End If
    End Select
    
End Sub

Private Sub gAlmacenes_OnKeyDown(KeyCode As Integer, ByVal Shift As Long)
Dim i As Integer

    If KeyCode = 46 Then
        If gAlmacenes.Count > 0 Then
            If MsgBox("Está seguro(a) de eliminar el registro?", vbInformation + vbYesNo, App.Title) = vbYes Then
                If gAlmacenes.Count = 1 Then
                    gAlmacenes.Dataset.Edit
                    gAlmacenes.Columns.ColumnByFieldName("Item").Value = 1
                    gAlmacenes.Columns.ColumnByFieldName("idSucursal").Value = ""
                    gAlmacenes.Columns.ColumnByFieldName("GlsSucursal").Value = ""
                    gAlmacenes.Columns.ColumnByFieldName("idAlmacen").Value = ""
                    gAlmacenes.Columns.ColumnByFieldName("GlsAlmacen").Value = ""
                    gAlmacenes.Dataset.Post
                
                Else
                    gAlmacenes.Dataset.Delete
                    gAlmacenes.Dataset.First
                    Do While Not gAlmacenes.Dataset.EOF
                        i = i + 1
                        gAlmacenes.Dataset.Edit
                        gAlmacenes.Columns.ColumnByFieldName("Item").Value = i
                        gAlmacenes.Dataset.Post
                        gAlmacenes.Dataset.Next
                    Loop
                    If gAlmacenes.Dataset.State = dsEdit Or gAlmacenes.Dataset.State = dsInsert Then
                        gAlmacenes.Dataset.Post
                    End If
                End If
            End If
        End If
    End If
    
    If KeyCode = 13 Then
        If gAlmacenes.Dataset.State = dsEdit Or gAlmacenes.Dataset.State = dsInsert Then
              gAlmacenes.Dataset.Post
        End If
    End If

End Sub

Private Sub gAlmacenes_OnKeyPress(Key As Integer)
Dim strCod As String
Dim StrDes As String

    Select Case gAlmacenes.Columns.FocusedColumn.Index
        Case gAlmacenes.Columns.ColumnByFieldName("idSucursal").Index
            strCod = gAlmacenes.Columns.ColumnByFieldName("idSucursal").Value
            StrDes = gAlmacenes.Columns.ColumnByFieldName("GlsSucursal").Value
            
            mostrarAyudaKeyasciiTexto Key, "SUCURSAL", strCod, StrDes
            Key = 0
            gAlmacenes.Dataset.Edit
            gAlmacenes.Columns.ColumnByFieldName("idSucursal").Value = strCod
            gAlmacenes.Columns.ColumnByFieldName("GlsSucursal").Value = StrDes
            gAlmacenes.Dataset.Post
            gAlmacenes.SetFocus
        
        Case gAlmacenes.Columns.ColumnByFieldName("idAlmacen").Index
            strCod = gAlmacenes.Columns.ColumnByFieldName("idAlmacen").Value
            StrDes = gAlmacenes.Columns.ColumnByFieldName("GlsAlmacen").Value
            
            mostrarAyudaKeyasciiTexto Key, "ALMACENVTA", strCod, StrDes, " AND idSucursal = '" & gAlmacenes.Columns.ColumnByFieldName("idSucursal").Value & "'"
            Key = 0
            If existeEnGrilla(gAlmacenes, "idAlmacen", strCod) = False Then
                gAlmacenes.Dataset.Edit
                gAlmacenes.Columns.ColumnByFieldName("idAlmacen").Value = strCod
                gAlmacenes.Columns.ColumnByFieldName("GlsAlmacen").Value = StrDes
                gAlmacenes.Dataset.Post
                gAlmacenes.SetFocus
            Else
                MsgBox "El Almacén ya fue ingresado.", vbInformation, App.Title
            End If
    End Select

End Sub

Private Sub gLista_OnDblClick()
On Error GoTo Err
Dim StrMsgError As String

    mostrarProducto gLista.Columns.ColumnByName("idProducto").Value, StrMsgError
    If StrMsgError <> "" Then GoTo Err
    For i = 1 To numPesos - 1
        txtCod_Nivel((numPesos - i) - 1).Text = traerCampo("niveles", "idNivelPred", "idNivel", txtCod_Nivel(numPesos - i).Text, True)
    Next
    fraListado.Visible = False
    fraGeneral.Visible = True
    fraGeneral.Enabled = False
    habilitaBotones 2
    
    Exit Sub

Err:
    MsgBox StrMsgError, vbInformation, App.Title
End Sub

Private Sub gLista_OnReloadGroupList()
    
    gLista.m.FullExpand

End Sub

Private Sub lbl_Imagen_Change()

    If Trim(lbl_Imagen.Caption) = "" Then
        lbl_Imagen.BackColor = &H3161DD
    Else
        lbl_Imagen.BackColor = &HC000&
    End If

End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
On Error GoTo Err
Dim StrMsgError             As String
Dim strCodUltProd           As String
Dim rsg                     As New ADODB.Recordset
Dim rst                     As New ADODB.Recordset

    Select Case Button.Index
        Case 1 '--- Nuevo
            nuevo
            indCalculando = False
            strCodUltProd = traerCampo("productos", "MAX(idProducto)", "idEmpresa", glsEmpresa, False)
            
            'Luis 20180312
            If Trim("" & traerCampo("Parametros", "Valparametro", "GlsParametro", "MUESTRA_ALM_DEFECTO", True)) = "S" Then

                'TRAE EL LISTADO DE ALMACENES Y LO ALMACENA EN UN RECORSET
                csql = "SELECT @i:=@i+1 item,p.idSucursal,s.glsPersona as GlsSucursal,a.idAlmacen,a.glsAlmacen " & _
                       "FROM Sucursales p,almacenes a,personas s,(Select @i:=0) z " & _
                       "WHERE p.idEmpresa = a.idEmpresa AND p.idSucursal = a.idSucursal " & _
                         "AND p.idSucursal = s.idPersona " & _
                         "AND p.idEmpresa = '" & glsEmpresa & "'"
                         
                rst.Open csql, Cn, adOpenKeyset, adLockOptimistic
                
                
                rsg.Fields.Append "Item", adInteger, , adFldRowID
                rsg.Fields.Append "idSucursal", adVarChar, 8, adFldIsNullable
                rsg.Fields.Append "GlsSucursal", adVarChar, 180, adFldIsNullable
                rsg.Fields.Append "idAlmacen", adVarChar, 8, adFldIsNullable
                rsg.Fields.Append "GlsAlmacen", adVarChar, 180, adFldIsNullable
                
                rsg.Open
                
                If rst.RecordCount = 0 Then
                    rsg.AddNew
                
                    rsg.Fields("Item") = 1
                    rsg.Fields("idSucursal") = ""
                    rsg.Fields("GlsSucursal") = ""
                    rsg.Fields("idAlmacen") = ""
                    rsg.Fields("GlsAlmacen") = ""
                Else
                    Do While Not rst.EOF
                        rsg.AddNew
                        
                        i = i + 1
                        
                        rsg.Fields("Item") = i
                        rsg.Fields("idSucursal") = rst.Fields("idSucursal")
                        rsg.Fields("GlsSucursal") = rst.Fields("GlsSucursal")
                        rsg.Fields("idAlmacen") = rst.Fields("idAlmacen")
                        rsg.Fields("GlsAlmacen") = rst.Fields("GlsAlmacen")
                        rst.MoveNext
                    Loop
                End If
                
                rst.Close
                Set rst = Nothing
            
                mostrarDatosGridSQL gAlmacenes, rsg, StrMsgError
                If StrMsgError <> "" Then GoTo Err
            
            End If
        
            If StrMsgError <> "" Then GoTo Err
            txtCod_Producto.Text = ""
            txtCod_TallaPeso.Text = ""
            txtGls_Producto.Text = ""
            txt_CodFabricante.Text = ""
            TxtDctoListaPrec.Text = 0
            fraListado.Visible = False
            fraGeneral.Visible = True
            fraGeneral.Enabled = True
            habilitaBotones 1
            
        Case 2 '--- Grabar
            If Len(txtCod_UMCompra.Text) = 0 Or Len(txtCod_UMVenta.Text) = 0 Then
                MsgBox "Falta Unidad de Medida. Verifique", vbInformation, App.Title
                Exit Sub
            End If
            If Len(Trim(gAlmacenes.Columns.ColumnByFieldName("idalmacen").Value)) = 0 And txtCod_TipoProd.Text <> "06002" Then
                MsgBox "Falta Colocar almacen. Verifique", vbInformation, App.Title
                Exit Sub
            End If
            If Len(Trim(txtGls_Producto.Text)) = 0 Then
                txtGls_Producto.Text = txtGls_Nivel(0).Text & " " & txtGls_Nivel(1).Text & " " & txtGls_Nivel(2).Text
            End If
            'If Len(Trim(txt_ctacontable1.Text)) = 0 Then
            '    StrMsgError = "Debe ingresar la Cuenta Contable.": GoTo Err
            'End If
            
            Grabar StrMsgError
            If StrMsgError <> "" Then GoTo Err
            
            habilitaBotones 2
            
            fraGeneral.Enabled = False
            
            If ChkConservarDatos.Value = 1 Then
                txtCod_Producto.Text = ""
                habilitaBotones 1
                fraListado.Visible = False
                fraGeneral.Visible = True
                If fraGeneral.Enabled = True Then
                Else
                    fraGeneral.Enabled = True
                End If
            End If
            
        Case 3 '--- Modificar
            fraGeneral.Enabled = True
            habilitaBotones 3
        
        Case 4, 8 '--- Cancelar
            fraListado.Visible = True
            fraGeneral.Visible = False
            fraGeneral.Enabled = False
            habilitaBotones 4
            
        Case 5 '--- Eliminar
            eliminar StrMsgError
            If StrMsgError <> "" Then GoTo Err
            habilitaBotones 5
        
        Case 6 '--- Imprimir
            gLista.m.ExportToXLS App.Path & "\Temporales\Mantenimiento_Productos.xls"
            ShellEx App.Path & "\Temporales\Mantenimiento_Productos.xls", essSW_MAXIMIZE, , , "open", Me.hwnd
            habilitaBotones 6
        
        Case 7 '--- Imprimir Codigo de Barras
            If txtCod_Producto.Text <> "" Then
                frmIngresoCantidad.MostrarForm txtCod_Producto.Text, StrMsgError
                If StrMsgError <> "" Then GoTo Err
            End If
            habilitaBotones 7
        
        Case 9 '--- Salir
            Unload Me
    End Select
    
    Exit Sub
    
Err:
    MsgBox StrMsgError, vbInformation, App.Title
    Exit Sub
    Resume
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
            Toolbar1.Buttons(7).Visible = indHabilitar 'Imprimir Codigo de Barras
            Toolbar1.Buttons(8).Visible = indHabilitar 'Lista
        Case 4, 8 'Cancelar, Lista
            Toolbar1.Buttons(1).Visible = True
            Toolbar1.Buttons(2).Visible = False
            Toolbar1.Buttons(3).Visible = False
            Toolbar1.Buttons(4).Visible = False
            Toolbar1.Buttons(5).Visible = False
            Toolbar1.Buttons(6).Visible = True
            Toolbar1.Buttons(7).Visible = False
            Toolbar1.Buttons(8).Visible = False
    End Select

End Sub

Private Sub txt_CodFabricante_LostFocus()
On Error GoTo Err
Dim StrMsgError                 As String
Dim CFiltro                     As String
    
    If traerCampo("Parametros", "Valparametro", "GlsParametro", "VALIDA_CODIGO_FABRICANTE", True) = "S" Then
        If Len(Trim(txt_CodFabricante.Text)) > 0 Then
            CFiltro = IIf(Len(Trim(txtCod_Producto.Text)) = 0, "", "IdProducto <> '" & txtCod_Producto.Text & "'")
            If Len(Trim(traerCampo("Productos", "IdProducto", "IdFabricante", txt_CodFabricante.Text, True, CFiltro))) > 0 Then
                If txt_CodFabricante.Visible And txt_CodFabricante.Enabled Then txt_CodFabricante.SetFocus
                StrMsgError = "Código de Fabricante ya existe.": GoTo Err
            End If
        End If
    End If
    
    Exit Sub
    
Err:
    If StrMsgError = "" Then StrMsgError = Err.Description
    MsgBox StrMsgError, vbInformation, App.Title
End Sub

'Private Sub txt_ctacontable1_Change()
'
'    If Len(txt_ctacontable1.Text) = 0 Then
'        txt_ctacontable1.Text = ""
'        txtdescta1.Text = ""
'    Else
'        txtdescta1.Text = traerCampoConta("Plancuentas", "glsNombreCuenta", "idctacontable", txt_ctacontable1.Text, False, " idanno = '2011' ")
'    End If
'
'End Sub

'Private Sub txt_ctacontable1_GotFocus()
'
'    txt_ctacontable1.SelStart = 0: txt_ctacontable1.SelLength = Len(txt_ctacontable1.Text)
'
'End Sub

'Private Sub txt_ctacontable1_KeyDown(KeyCode As Integer, Shift As Integer)
'
'    If KeyCode = 113 Then
'        cmdayuda1_Click
'    End If
'
'End Sub
'
'Private Sub txt_ctacontable2_Change()
'
'    If Len(txt_ctacontable2.Text) = 0 Then
'        txt_ctacontable2.Text = ""
'        txtdescta2.Text = ""
'    Else
'        txtdescta2.Text = traerCampoConta("Plancuentas", "glsNombreCuenta", "idctacontable", txt_ctacontable2.Text, False, " idanno = '2011' ")
'    End If
'
'End Sub

'Private Sub txt_ctacontable2_GotFocus()
'
'    txt_ctacontable2.SelStart = 0: txt_ctacontable2.SelLength = Len(txt_ctacontable2.Text)
'
'End Sub

'Private Sub txt_ctacontable2_KeyDown(KeyCode As Integer, Shift As Integer)
'
'    If KeyCode = 113 Then
'        cmdayuda2_Click
'    End If
'
'End Sub

Private Sub txt_TextoBuscar_KeyDown(KeyCode As Integer, Shift As Integer)

    If KeyCode = vbKeyDown Then gLista.SetFocus

End Sub

Private Sub listaProducto(ByRef StrMsgError As String)
On Error GoTo Err
Dim rst                             As New ADODB.Recordset
Dim strCond                         As String
Dim intNumNiveles                   As Integer
Dim strTabla                        As String
Dim strWhere                        As String
Dim strCampos                       As String
Dim strTablas                       As String
Dim strTablaAnt                     As String
Dim i                               As Integer
Dim CSqlC                           As String
Dim rsdatos                         As New ADODB.Recordset
 
    rst.Open "SELECT idTipoNivel,GlsTipoNivel,peso " & _
             "FROM tiposniveles WHERE idEmpresa = '" & glsEmpresa & _
             "' ORDER BY peso DESC", Cn, adOpenKeyset, adLockOptimistic
    intNumNiveles = rst.RecordCount
    
    Do While Not rst.EOF
        i = i + 1
        strTabla = "niveles" & Format(i, "00")
        If i = 1 Then
            strWhere = "p.idNivel = " & strTabla & ".idNivel AND " & strTabla & ".idEmpresa = '" & glsEmpresa & "' "
        Else
            strWhere = strTablaAnt & ".idNivelPred = " & strTabla & ".idNivel AND " & strTabla & ".idEmpresa = '" & glsEmpresa & "' "
        End If
        strCampos = strCampos & strTabla & ".idNivel as idNivel" & Format(i, "00") & "," & strTabla & ".GlsNivel as GlsNivel" & Format(i, "00") & ","
        
        '--- Agrupando tablas
        strTablas = strTablas & "INNER JOIN niveles " & strTabla & " ON " & strWhere
        strTablaAnt = strTabla
        
        rst.MoveNext
    Loop
    
    strCond = ""
    If Trim(txt_TextoBuscar.Text) <> "" Then
        strCond = Trim(txt_TextoBuscar.Text)
        strCond = " AND (GlsProducto LIKE '%" & strCond & "%' or CodigoRapido LIKE '%" & strCond & "%' or p.IdFabricante LIKE '%" & strCond & "%' or p.IdProducto LIKE '%" & strCond & "%') "
    End If
    
    CSqlC = "Select " & strCampos & "P.IdProducto,P.GlsProducto,M.GlsMarca,UMC.AbreUM GlsUMC,UMV.AbreUM GlsUMV," & _
            "CASE WHEN P.EstProducto = 'A' THEN 'ACT' ELSE 'INA' END EstProducto,P.IdFabricante,P.CtaContable2,P.CtaContable,P.CtaContable_Relacionada,P.CodigoRapido " & _
            "From Productos P " & _
            "Left Join Marcas M " & _
                "On P.IdEmpresa = M.IdEmpresa And P.IdMarca = M.IdMarca " & _
            "Left Join UnidadMedida UMC " & _
                "On P.IdUMCompra = UMC.IdUM " & _
            "Left Join UnidadMedida UMV " & _
                "On P.IdUMVenta = UMV.IdUM " & strTablas & _
            "Where P.IdEmpresa = '" & glsEmpresa & "'"
            
    If strCond <> "" Then CSqlC = CSqlC + strCond
    
    If leeParametro("VIZUALIZA_CODIGO_RAPIDO") = "S" Then
        CSqlC = CSqlC & " Order By P.CodigoRapido"
    Else
        CSqlC = CSqlC & " Order By P.IdProducto"
    End If
    
If rsdatos.State = 1 Then rsdatos.Close: Set rsdatos = Nothing
rsdatos.Open CSqlC, Cn, adOpenStatic, adLockOptimistic
Set gLista.DataSource = rsdatos

'    With gLista
'        .DefaultFields = False
'        .Dataset.ADODataset.ConnectionString = strcn
'        .Dataset.ADODataset.CursorLocation = clUseClient
'        .Dataset.Active = False
'        .Dataset.ADODataset.CommandText = CSqlC
'        .Dataset.DisableControls
'        .Dataset.Active = True
'        .KeyField = "idProducto"
'    End With
    
    '--- Realizando las agrupaciones dinamicamente
    gLista.m.ClearGroupColumns
    
    If gLista.Ex.GroupColumnCount = 0 Then
        For i = 1 To intNumNiveles
            gLista.Columns.ColumnByName("GlsNivel" & Format(i, "00")).Caption = "Nivel:"
            gLista.Columns.ColumnByName("GlsNivel" & Format(i, "00")).Visible = True
            gLista.Columns.ColumnByName("GlsNivel" & Format(i, "00")).GroupIndex = 0
        Next
    End If
    
    Me.Refresh
    If rst.State = 1 Then rst.Close: Set rst = Nothing
    
    Exit Sub
    
Err:
    If rst.State = 1 Then rst.Close: Set rst = Nothing
    If StrMsgError = "" Then StrMsgError = Err.Description
End Sub

Private Sub mostrarProducto(strCodProd As String, ByRef StrMsgError As String)
On Error GoTo Err
Dim rst                         As New ADODB.Recordset
Dim rsu                         As New ADODB.Recordset
Dim rsg                         As New ADODB.Recordset
Dim strIndImagen                As String
Dim i                           As Integer
Dim CSqlC                       As String

    indCargando = True
    CSqlC = "SELECT p.idProducto,p.GlsProducto,p.idNivel,p.idUMCompra,p.idUMVenta,p.idMarca,p.afectoIGV,p.idMoneda,p.idTipoProducto,p.idFabricante," & _
            "p.CodigoRapido,p.idGrupo,p.idTallaPeso,p.indImagen,p.indInsertaPrecioLista,p.afectoDctoEspecial, p.CtaContable, p.estProducto,p.CtaContable2," & _
            "p.CtaContable_Relacionada,tiemporepinv,p.IdConceptoDetraccion,IsNull(P.IndRptGL,'0') IndRptGL,P.StockMinimo," & _
            "Cast(('0' + IsNull(P.IndAfectoIVAP,0)) As TINYINT) IndAfectoIVAP,P.IdConceptoCosteo " & _
            "FROM productos p " & _
            "WHERE p.idProducto = '" & strCodProd & "' AND p.idEmpresa = '" & glsEmpresa & "'"
    rst.Open CSqlC, Cn, adOpenStatic, adLockReadOnly
    
    strIndImagen = ("" & rst.Fields("indImagen"))
    If rst.Fields("estProducto") = "A" Then OptActivo.Value = True
    If rst.Fields("estProducto") = "I" Then OptInactivo.Value = True
    mostrarDatosFormSQL Me, rst, StrMsgError
    If StrMsgError <> "" Then GoTo Err
    
    lbl_Imagen.Caption = ""
    'If strIndImagen = "S" Then lbl_Imagen.Caption = glsRutaImagenProd & "\" & glsEmpresa & "-" & txtCod_Producto.Text & ".jpg"
    If strIndImagen = "S" Then lbl_Imagen.Caption = gbRutaProductos & glsEmpresa & "-" & txtCod_Producto.Text & ".jpg"
    
    '--- TRAE EL LISTADO DE ALMACENES Y LO ALMACENA EN UN RECORSET
    CSqlC = "SELECT p.idUbicacion,p.item,p.idSucursal,s.glsPersona as GlsSucursal,p.idAlmacen,a.glsAlmacen " & _
            "FROM productosalmacen p,almacenes a,personas s " & _
            "WHERE p.idEmpresa = a.idEmpresa AND p.idAlmacen = a.idAlmacen " & _
            "AND p.idSucursal = s.idPersona " & _
            "AND p.idEmpresa = '" & glsEmpresa & "' AND p.idProducto = '" & strCodProd & "' ORDER BY item"
    rst.Open CSqlC, Cn, adOpenKeyset, adLockOptimistic
    
    rsg.Fields.Append "Item", adInteger, , adFldRowID
    rsg.Fields.Append "idSucursal", adVarChar, 8, adFldIsNullable
    rsg.Fields.Append "GlsSucursal", adVarChar, 180, adFldIsNullable
    rsg.Fields.Append "idAlmacen", adVarChar, 8, adFldIsNullable
    rsg.Fields.Append "GlsAlmacen", adVarChar, 180, adFldIsNullable
    rsg.Fields.Append "idUbicacion", adVarChar, 20, adFldIsNullable
    rsg.Open
    
    If rst.RecordCount = 0 Then
        rsg.AddNew
        rsg.Fields("Item") = 1
        rsg.Fields("idSucursal") = ""
        rsg.Fields("GlsSucursal") = ""
        rsg.Fields("idAlmacen") = ""
        rsg.Fields("GlsAlmacen") = ""
        rsg.Fields("idUbicacion") = ""
    
    Else
        Do While Not rst.EOF
            rsg.AddNew
            i = i + 1
            rsg.Fields("Item") = i
            rsg.Fields("idSucursal") = rst.Fields("idSucursal")
            rsg.Fields("GlsSucursal") = rst.Fields("GlsSucursal")
            rsg.Fields("idAlmacen") = rst.Fields("idAlmacen")
            rsg.Fields("GlsAlmacen") = rst.Fields("GlsAlmacen")
            rsg.Fields("idUbicacion") = rst.Fields("idUbicacion")
            rst.MoveNext
        Loop
    End If
    rst.Close: Set rst = Nothing
    
    mostrarDatosGridSQL gAlmacenes, rsg, StrMsgError
    If StrMsgError <> "" Then GoTo Err
    
    '--- TRAE EL LISTADO DE PRESENTACIONES Y LO ALMACENA EN UN RECORSET
    CSqlC = "SELECT p.item,p.idUM,u.glsUM,p.Factor " & _
            "FROM presentaciones p,unidadmedida u " & _
            "WHERE p.idUM = u.idUM " & _
            "AND p.idProducto = '" & strCodProd & "' AND p.idEmpresa = '" & glsEmpresa & "' ORDER BY p.item ASC"
    If rst.State = 1 Then rst.Close
    rst.Open CSqlC, Cn, adOpenKeyset, adLockOptimistic
    rsu.Fields.Append "Item", adInteger, , adFldRowID
    rsu.Fields.Append "idUM", adVarChar, 8, adFldIsNullable
    rsu.Fields.Append "GlsUM", adVarChar, 250, adFldIsNullable
    rsu.Fields.Append "Factor", adDouble, 14, adFldIsNullable
    rsu.Open
    
    If rst.RecordCount = 0 Then
        rsu.AddNew
        rsu.Fields("Item") = 1
        rsu.Fields("idUM") = ""
        rsu.Fields("GlsUM") = ""
        rsu.Fields("Factor") = 0
    
    Else
        Do While Not rst.EOF
            rsu.AddNew
            rsu.Fields("Item") = rst.Fields("Item")
            rsu.Fields("idUM") = rst.Fields("idUM")
            rsu.Fields("GlsUM") = rst.Fields("GlsUM")
            rsu.Fields("Factor") = rst.Fields("Factor")
            rst.MoveNext
        Loop
    End If
    rst.Close
    
    mostrarDatosGridSQL gPresentaciones, rsu, StrMsgError
    If StrMsgError <> "" Then GoTo Err
    
    chkInsertaPrecios_Click
    
    'Mostrando los precios de la lista de precios principal
    If glsListaVentas <> "" And chkInsertaPrecios.Value Then
        CSqlC = "SELECT VVUnit,IGVUnit,PVUnit,MaxDcto " & _
                "FROM preciosventa " & _
                "WHERE idEmpresa = '" & glsEmpresa & "' " & _
                "AND idLista = '" & glsListaVentas & "' " & _
                "AND idProducto = '" & txtCod_Producto.Text & "' " & _
                "AND idUM = '" & txtCod_UMVenta.Text & "'"
        rst.Open CSqlC, Cn, adOpenForwardOnly, adLockReadOnly
        
        If Not rst.EOF Then
            txtVal_VV.Text = Val("" & rst.Fields("VVUnit"))
            txtVal_IGV.Text = Val("" & rst.Fields("IGVUnit"))
            txtVal_PV.Text = Val("" & rst.Fields("PVUnit"))
            TxtDctoListaPrec.Text = Val("" & rst.Fields("MaxDcto"))
        Else
            txtVal_VV.Text = 0
            txtVal_IGV.Text = 0
            txtVal_PV.Text = 0
            TxtDctoListaPrec.Text = 0
        End If
        rst.Close
    
    Else
        txtVal_VV.Text = 0
        txtVal_IGV.Text = 0
        txtVal_PV.Text = 0
        TxtDctoListaPrec.Text = 0
    End If
    Set rst = Nothing
    indCargando = False
    Me.Refresh

    Exit Sub
    
Err:
    If StrMsgError = "" Then StrMsgError = Err.Description
End Sub

Private Sub txt_TextoBuscar_KeyPress(KeyAscii As Integer)
On Error GoTo Err
Dim StrMsgError As String

    If KeyAscii = 13 Then
        listaProducto StrMsgError
        If StrMsgError <> "" Then GoTo Err
    End If
    
    Exit Sub
    
Err:
    If StrMsgError = "" Then StrMsgError = Err.Description
    MsgBox StrMsgError, vbInformation, App.Title
End Sub

Private Sub TxtCod_Grupo_Change()
    
    TxtGls_Grupo.Text = traerCampo("GruposProducto", "GlsGrupo", "idGrupo", TxtCod_Grupo.Text, True)
 
End Sub

Private Sub TxtCod_Grupo_KeyPress(KeyAscii As Integer)
    
    If KeyAscii <> 13 Then
        mostrarAyudaKeyascii KeyAscii, "GRUPOSPRODUCTO", TxtCod_Grupo, TxtGls_Grupo
        KeyAscii = 0
        If TxtCod_Grupo.Text <> "" Then SendKeys "{tab}"
    End If
    
End Sub

Private Sub txtCod_Marca_Change()
    
    txtGls_Marca.Text = traerCampo("marcas", "GlsMarca", "idMarca", txtCod_Marca.Text, True)

End Sub

Private Sub txtCod_Marca_KeyPress(KeyAscii As Integer)
    
    If KeyAscii <> 13 Then
        mostrarAyudaKeyascii KeyAscii, "MARCA", txtCod_Marca, txtGls_Marca
        KeyAscii = 0
        If txtCod_Marca.Text <> "" Then SendKeys "{tab}"
    End If

End Sub

Private Sub txtCod_Moneda_Change()
    
    txtGls_Moneda.Text = traerCampo("monedas", "GlsMoneda", "idMoneda", txtCod_Moneda.Text, False)

End Sub

Private Sub txtCod_Moneda_KeyPress(KeyAscii As Integer)
    
    If KeyAscii <> 13 Then
        mostrarAyudaKeyascii KeyAscii, "MONEDA", txtCod_Moneda, txtGls_Moneda
        KeyAscii = 0
        If txtCod_Moneda.Text <> "" Then SendKeys "{tab}"
    End If

End Sub

Private Sub txtCod_Nivel_Change(Index As Integer)
On Error GoTo Err
Dim StrMsgError                         As String
Dim peso                                As Integer
Dim strCodTipoNivel                     As String
Dim CSqlC                               As String
Dim RsC                                 As New ADODB.Recordset

    txtGls_Nivel(Index).Text = traerCampo("niveles", "GlsNivel", "idNivel", txtCod_Nivel(Index).Text, True)
    
    peso = Index + 1
    strCodTipoNivel = traerCampo("tiposniveles", "idTipoNivel", "peso", CStr(peso), True)
    
    If txtCod_Producto.Text = "" Then
    
        If strCodTipoNivel = leeParametro("NIVEL_CUENTA_CONTABLE") Then
            
'            CSqlC = "Select IdCtaContableC,IdCtaContableV,IdCtaContableVR " & _
'                    "From Niveles " & _
'                    "Where IdEmpresa = '" & glsEmpresa & "' And IdNivel = '" & txtCod_Nivel(Index).Text & "'"
'            RsC.Open CSqlC, Cn, adOpenStatic, adLockReadOnly
'            If Not RsC.EOF Then
'
'                txt_ctacontable2.Text = "" & RsC.Fields("IdCtaContableC")
'                txt_ctacontable1.Text = "" & RsC.Fields("IdCtaContableV")
'                txtctacontableRelacionada.Text = "" & RsC.Fields("IdCtaContableVR")
'
'            End If
'
'            RsC.Close: Set RsC = Nothing
            
        End If
    
    End If
    
    Exit Sub
Err:
    If RsC.State = 1 Then RsC.Close: Set RsC = Nothing
    If StrMsgError = "" Then StrMsgError = Err.Description
    MsgBox StrMsgError, vbInformation, App.Title
End Sub

Private Sub txtCod_Nivel_KeyPress(Index As Integer, KeyAscii As Integer)
Dim peso As Integer
Dim strCodTipoNivel As String
Dim strCondPred As String
    
    If KeyAscii <> 13 Then
        peso = Index + 1
        strCodJerarquia = traerCampo("tiposniveles", "idTipoNivel", "peso", CStr(peso), True)
        strCondPred = ""
        If peso > 1 Then
            strCondPred = " AND idNivelPred = '" & txtCod_Nivel(Index - 1).Text & "'"
        End If
        mostrarAyudaKeyascii KeyAscii, "NIVEL", txtCod_Nivel(Index), txtGls_Nivel(Index), " AND idTipoNivel = '" & strCodTipoNivel & "'" & strCondPred
    End If

End Sub

Private Sub txtCod_TipoProd_Change()
    
    txtGls_TipoProd.Text = traerCampo("Datos", "GlsDato", "idDato", txtCod_TipoProd.Text, False)
    If Trim(txtCod_TipoProd.Text) = "06002" Then
        txtCod_Marca.Vacio = True
        txtCod_UMCompra.Vacio = True
        txtCod_UMVenta.Vacio = True
    Else
        txtCod_Marca.Vacio = False
        txtCod_UMCompra.Vacio = False
        txtCod_UMCompra.Vacio = False
    End If

End Sub

Private Sub txtCod_TipoProd_KeyPress(KeyAscii As Integer)
    
    If KeyAscii <> 13 Then
        mostrarAyudaKeyascii KeyAscii, "TIPOPRODUCTO", txtCod_TipoProd, txtGls_TipoProd
        KeyAscii = 0
        If txtCod_TipoProd.Text <> "" Then SendKeys "{tab}"
    End If

End Sub

Private Sub txtCod_UMCompra_Change()
    
    txtGls_UMCompra.Text = traerCampo("unidadmedida", "GlsUM", "idUM", txtCod_UMCompra.Text, False)

End Sub

Private Sub txtCod_UMCompra_KeyPress(KeyAscii As Integer)
    
    If KeyAscii <> 13 Then
        mostrarAyudaKeyascii KeyAscii, "UMGLOSA", txtCod_UMCompra, txtGls_UMCompra
        KeyAscii = 0
        If txtCod_UMCompra.Text <> "" Then SendKeys "{tab}"
    End If

End Sub

Private Sub txtCod_UMVenta_Change()
    
    txtGls_UMVenta.Text = traerCampo("unidadmedida", "GlsUM", "idUM", txtCod_UMVenta.Text, False)
    If Trim(txtGls_UMVenta.Text) <> "" And indCargando = False Then
        gPresentaciones.Dataset.Edit
        gPresentaciones.Columns.ColumnByFieldName("idUM").Value = txtCod_UMVenta.Text
        gPresentaciones.Columns.ColumnByFieldName("GlsUM").Value = txtGls_UMVenta.Text
        gPresentaciones.Columns.ColumnByFieldName("Factor").Value = 1
        gPresentaciones.Dataset.Post
    End If
    
End Sub

Private Sub txtCod_UMVenta_KeyPress(KeyAscii As Integer)
    
    If KeyAscii <> 13 Then
        mostrarAyudaKeyascii KeyAscii, "UMGLOSA", txtCod_UMVenta, txtGls_UMVenta
        KeyAscii = 0
        If txtCod_UMVenta.Text <> "" Then SendKeys "{tab}"
    End If

End Sub

Private Sub gPresentaciones_OnAfterDatasetAction(ByVal Action As DXDBGRIDLibCtl.ExDatasetAction)
Dim i As Integer

    If Action = daInsert Then
        gPresentaciones.Columns.ColumnByFieldName("item").Value = gPresentaciones.Count
        gPresentaciones.Dataset.Post
    End If

End Sub

Private Sub gPresentaciones_OnBeforeDatasetAction(ByVal Action As DXDBGRIDLibCtl.ExDatasetAction, Allow As Boolean)

    If Action = daInsert Then
        If gPresentaciones.Columns.ColumnByFieldName("idUM").Value = "" Then
            Allow = False
        Else
            gPresentaciones.Columns.FocusedIndex = gPresentaciones.Columns.ColumnByFieldName("idUM").Index
        End If
    End If

End Sub

Private Sub gPresentaciones_OnEditButtonClick(ByVal Column As DXDBGRIDLibCtl.IdxGridColumn, ByVal Node As DXDBGRIDLibCtl.IdxGridNode)
Dim strCod As String
Dim StrDes As String
    
    Select Case Column.Index
        Case gPresentaciones.Columns.ColumnByFieldName("idUM").Index
            strCod = gPresentaciones.Columns.ColumnByFieldName("idUM").Value
            StrDes = gPresentaciones.Columns.ColumnByFieldName("GlsUM").Value
            
            mostrarAyudaTexto "UMGLOSA", strCod, StrDes
            
            If existeEnGrilla(gPresentaciones, "idUM", strCod) = False Then
                gPresentaciones.Dataset.Edit
                gPresentaciones.Columns.ColumnByFieldName("idUM").Value = strCod
                gPresentaciones.Columns.ColumnByFieldName("GlsUM").Value = StrDes
                gPresentaciones.Columns.ColumnByFieldName("Factor").Value = 1
                gPresentaciones.Dataset.Post
            Else
                MsgBox "La Unidad de Medida ya fue ingresada", vbInformation, App.Title
            End If
    End Select
    
End Sub

Private Sub gPresentaciones_OnKeyDown(KeyCode As Integer, ByVal Shift As Long)
On Error GoTo Err
Dim StrMsgError                     As String
Dim i                               As Integer
Dim RsC                             As New ADODB.Recordset
Dim CSqlC                           As String

    If KeyCode = 46 Then
        If gPresentaciones.Count > 0 Then
            
            If Trim("" & gPresentaciones.Columns.ColumnByFieldName("IdUM").Value) <> "" Then
                
                CSqlC = "Select A.IdUM " & _
                        "From Prod_ParteDiario_MateriaPrima A " & _
                        "Where A.IdEmpresa = '" & glsEmpresa & "' And A.IdProducto = '" & txtCod_Producto.Text & "' " & _
                        "And A.IdUM = '" & Trim("" & gPresentaciones.Columns.ColumnByFieldName("IdUM").Value) & "' " & _
                        "Union All " & _
                        "Select A.IdUM " & _
                        "From Prod_ParteDiario_ProductoDescarte A " & _
                        "Where A.IdEmpresa = '" & glsEmpresa & "' And A.IdProducto = '" & txtCod_Producto.Text & "' " & _
                        "And A.IdUM = '" & Trim("" & gPresentaciones.Columns.ColumnByFieldName("IdUM").Value) & "' " & _
                        "Union All " & _
                        "Select A.IdUM " & _
                        "From Prod_ParteDiario_ProductoTerminado A " & _
                        "Where A.IdEmpresa = '" & glsEmpresa & "' And A.IdProducto = '" & txtCod_Producto.Text & "' " & _
                        "And A.IdUM = '" & Trim("" & gPresentaciones.Columns.ColumnByFieldName("IdUM").Value) & "' " & _
                        "Union All " & _
                        "Select A.IdUM " & _
                        "From Prod_Receta_Actividad_MateriaPrima A " & _
                        "Where A.IdEmpresa = '" & glsEmpresa & "' And A.IdProducto = '" & txtCod_Producto.Text & "' " & _
                        "And A.IdUM = '" & Trim("" & gPresentaciones.Columns.ColumnByFieldName("IdUM").Value) & "' " & _
                        "Union All "
                CSqlC = CSqlC & _
                        "Select A.IdUM " & _
                        "From Prod_Receta_Actividad_ProductoDescarte A " & _
                        "Where A.IdEmpresa = '" & glsEmpresa & "' And A.IdProducto = '" & txtCod_Producto.Text & "' " & _
                        "And A.IdUM = '" & Trim("" & gPresentaciones.Columns.ColumnByFieldName("IdUM").Value) & "' " & _
                        "Union All " & _
                        "Select A.IdUM " & _
                        "From Prod_Receta_Actividad_ProductoTerminado A " & _
                        "Where A.IdEmpresa = '" & glsEmpresa & "' And A.IdProducto = '" & txtCod_Producto.Text & "' " & _
                        "And A.IdUM = '" & Trim("" & gPresentaciones.Columns.ColumnByFieldName("IdUM").Value) & "' " & _
                        "Union All " & _
                        "Select A.IdUM " & _
                        "From Prod_Req_MateriaPrima A " & _
                        "Where A.IdEmpresa = '" & glsEmpresa & "' And A.IdProducto = '" & txtCod_Producto.Text & "' " & _
                        "And A.IdUM = '" & Trim("" & gPresentaciones.Columns.ColumnByFieldName("IdUM").Value) & "' " & _
                        "Union All " & _
                        "Select A.IdUM " & _
                        "From Prod_Req_ProductoDescarte A " & _
                        "Where A.IdEmpresa = '" & glsEmpresa & "' And A.IdProducto = '" & txtCod_Producto.Text & "' " & _
                        "And A.IdUM = '" & Trim("" & gPresentaciones.Columns.ColumnByFieldName("IdUM").Value) & "' " & _
                        "Union All " & _
                        "Select A.IdUM " & _
                        "From Prod_Req_ProductoTerminado A " & _
                        "Where A.IdEmpresa = '" & glsEmpresa & "' And A.IdProducto = '" & txtCod_Producto.Text & "' " & _
                        "And A.IdUM = '" & Trim("" & gPresentaciones.Columns.ColumnByFieldName("IdUM").Value) & "'"
                
                RsC.Open CSqlC, Cn, adOpenStatic, adLockReadOnly
                If Not RsC.EOF Then
                    
                    StrMsgError = "La presentación se encuentra en uso, no puede ser eliminada.": GoTo Err
                
                End If
                
                RsC.Close: Set RsC = Nothing
            
            End If
            
            If MsgBox("Está seguro(a) de eliminar el registro?", vbInformation + vbYesNo, App.Title) = vbYes Then
                If gPresentaciones.Count = 1 Then
                    gPresentaciones.Dataset.Edit
                    gPresentaciones.Columns.ColumnByFieldName("Item").Value = 1
                    gPresentaciones.Columns.ColumnByFieldName("idUM").Value = ""
                    gPresentaciones.Columns.ColumnByFieldName("GlsUM").Value = ""
                    gPresentaciones.Columns.ColumnByFieldName("Factor").Value = 0
                    gPresentaciones.Dataset.Post
                
                Else
                    gPresentaciones.Dataset.Delete
                    gPresentaciones.Dataset.First
                    Do While Not gPresentaciones.Dataset.EOF
                        i = i + 1
                        gPresentaciones.Dataset.Edit
                        gPresentaciones.Columns.ColumnByFieldName("Item").Value = i
                        gPresentaciones.Dataset.Post
                        gPresentaciones.Dataset.Next
                    Loop
                    If gPresentaciones.Dataset.State = dsEdit Or gPresentaciones.Dataset.State = dsInsert Then
                        gPresentaciones.Dataset.Post
                    End If
                End If
            End If
        End If
    End If
    
    If KeyCode = 13 Then
        If gPresentaciones.Dataset.State = dsEdit Or gPresentaciones.Dataset.State = dsInsert Then
            gPresentaciones.Dataset.Post
        End If
    End If
    
    Exit Sub
Err:
    If RsC.State = 1 Then RsC.Close: Set RsC = Nothing
    If StrMsgError = "" Then StrMsgError = Err.Description
    MsgBox StrMsgError, vbInformation, App.Title
End Sub

Private Sub gPresentaciones_OnKeyPress(Key As Integer)
Dim strCod As String
Dim StrDes As String

    Select Case gPresentaciones.Columns.FocusedColumn.Index
        Case gPresentaciones.Columns.ColumnByFieldName("idUM").Index
            strCod = gPresentaciones.Columns.ColumnByFieldName("idUM").Value
            StrDes = gPresentaciones.Columns.ColumnByFieldName("GlsUM").Value
                
            mostrarAyudaKeyasciiTexto Key, "UMGLOSA", strCod, StrDes
            Key = 0
            If existeEnGrilla(gPresentaciones, "idUM", strCod) = False Then
                gPresentaciones.Dataset.Edit
                gPresentaciones.Columns.ColumnByFieldName("idUM").Value = strCod
                gPresentaciones.Columns.ColumnByFieldName("GlsUM").Value = StrDes
                gPresentaciones.Columns.ColumnByFieldName("Factor").Value = 1
                gPresentaciones.Dataset.Post
                gPresentaciones.SetFocus
            Else
                MsgBox "La Unidad de Medida ya fue ingresada.", vbInformation, App.Title
            End If
    End Select
End Sub

Private Sub mostrarNiveles(ByRef StrMsgError As String)
On Error GoTo Err
Dim rsj As New ADODB.Recordset
Dim i As Integer

    '--- Limpiando Tag
    For i = 0 To 4
        txtCod_Nivel(i).Tag = ""
        txtCod_Nivel(i).Visible = False
    Next
    
    '--- Tipos nivel
    rsj.Open "SELECT GlsTipoNivel FROM tiposniveles WHERE idEmpresa = '" & glsEmpresa & "' Order BY Peso ASC", Cn, adOpenForwardOnly, adLockReadOnly
    numPesos = Val("" & rsj.RecordCount)
    For i = 0 To numPesos - 1
        txtCod_Nivel(i).Tag = ""
        txtCod_Nivel(i).Visible = True
    Next
    fraNivel.Height = 355 * numPesos
    
    i = 0
    Do While Not rsj.EOF
        If (i + 1) = numPesos Then
            txtCod_Nivel(i).Tag = "TidNivel"
        End If
        lblNivel(i).Caption = "" & rsj.Fields("GlsTipoNivel")
        rsj.MoveNext
        i = i + 1
    Loop
    
    fraContenido.top = fraNivel.top + fraNivel.Height - 70
    If rsj.State = 1 Then rsj.Close: Set rsj = Nothing
    
    Exit Sub
    
Err:
    If rsj.State = 1 Then rsj.Close: Set rsj = Nothing
    If StrMsgError = "" Then StrMsgError = Err.Description
End Sub

Private Sub copiaImagen(ByRef StrMsgError As String)
On Error GoTo Err
Dim strRuta As String

    If indIngresoImagen = False Then Exit Sub
    strRuta = Trim(lbl_Imagen.Caption)
    If strRuta = "" Then Exit Sub
    If Len(Dir(gbRutaProductos, vbDirectory)) = 0 Then
        MkDir gbRutaProductos
    End If

    FileCopy strRuta, gbRutaProductos & glsEmpresa & "-" & txtCod_Producto.Text & "." & right(strRuta, 3)
    lbl_Imagen.Caption = gbRutaProductos & glsEmpresa & "-" & txtCod_Producto.Text & "." & right(strRuta, 3)
    
    csql = "UPDATE productos SET indImagen = 'S' WHERE idEmpresa = '" & glsEmpresa & "' AND idProducto = '" & txtCod_Producto.Text & "'"
    Cn.Execute csql
    
Exit Sub
Err:
    If StrMsgError = "" Then StrMsgError = Err.Description
    Exit Sub
    Resume
End Sub

Private Sub eliminar(ByRef StrMsgError As String)
On Error GoTo Err
Dim indTrans As Boolean
Dim strCodigo As String
Dim rsValida As New ADODB.Recordset

    If MsgBox("Está seguro(a) de eliminar el registro?" & vbCrLf & "Se eliminaran todas sus dependencias.", vbQuestion + vbYesNo, App.Title) = vbNo Then Exit Sub
    
    strCodigo = Trim(txtCod_Producto.Text)
    
    csql = "SELECT Item FROM docventasdet WHERE idProducto = '" & strCodigo & "' AND idEmpresa = '" & glsEmpresa & "'"
    rsValida.Open csql, Cn, adOpenForwardOnly, adLockReadOnly
    If Not rsValida.EOF Then
        StrMsgError = "No se puede eliminar el registro, el registro se encuentra en uso (Ventas)"
        GoTo Err
    End If
    
    csql = "SELECT Item FROM valesdet WHERE idProducto = '" & strCodigo & "' AND idEmpresa = '" & glsEmpresa & "'"
    If rsValida.State = 1 Then rsValida.Close
    rsValida.Open csql, Cn, adOpenForwardOnly, adLockReadOnly
    If Not rsValida.EOF Then
        StrMsgError = "No se puede eliminar el registro, el registro se encuentra en uso (Vales)"
        GoTo Err
    End If
    
    Cn.BeginTrans
    indTrans = True
    
    '--- Eliminando preciosventa
    csql = "DELETE FROM preciosventa WHERE idProducto = '" & strCodigo & "' AND idEmpresa = '" & glsEmpresa & "'"
    Cn.Execute csql
    
    '--- Eliminando productosalmacen
    csql = "DELETE FROM productosalmacen WHERE idProducto = '" & strCodigo & "' AND idEmpresa = '" & glsEmpresa & "'"
    Cn.Execute csql
    
    '--- Eliminando el registro
    csql = "DELETE FROM productos WHERE idProducto = '" & strCodigo & "' AND idEmpresa = '" & glsEmpresa & "'"
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

'tipoOperacion : 0 = insert     1 = Update
Public Sub EjecutaSQLFormProducto(tipoOperacion As Integer, indEmpresa As Boolean, strTabla As String, ByRef StrMsgError As String, Optional strCampoCod As String, Optional indFechaRegistro As Boolean = False)
On Error GoTo Err
Dim C               As Object
Dim CSqlC            As String
Dim strCampo        As String
Dim strTipoDato     As String
Dim strCampos       As String
Dim strValores      As String
Dim strValCod       As String
Dim strCampoEmpresa As String
Dim strValorEmpresa As String
Dim strCondEmpresa  As String
Dim strCampoFecReg  As String
Dim strValorFecReg  As String
Dim indTrans        As Boolean
Dim CodLote         As String
Dim RsC             As New ADODB.Recordset

    If indEmpresa Then
        strCampoEmpresa = ",idEmpresa"
        strValorEmpresa = ",'" & glsEmpresa & "'"
        strCondEmpresa = " AND idEmpresa = '" & glsEmpresa & "'"
    End If
    
    If indFechaRegistro Then
        strCampoFecReg = ",FecRegistro"
        strValorFecReg = ",sysdate()"
    End If
    
    indTrans = False
    CSqlC = ""
    For Each C In Me.Controls
        If TypeOf C Is CATTextBox Or TypeOf C Is DTPicker Or TypeOf C Is CheckBox Then
            If C.Tag <> "" Then
                strTipoDato = left(C.Tag, 1)
                strCampo = right(C.Tag, Len(C.Tag) - 1)
                Select Case tipoOperacion
                    Case 0 'inserta
                        strCampos = strCampos & strCampo & ","
                        Select Case strTipoDato
                            Case "N"
                                strValores = strValores & C.Value & ","
                            Case "T"
                                strValores = strValores & "'" & Trim(C.Value) & "',"
                            Case "F"
                                strValores = strValores & "'" & Format(C.Value, "yyyy-mm-dd") & "',"
                        End Select
                    
                    Case 1
                        If UCase(strCampoCod) <> UCase(strCampo) Then
                            Select Case strTipoDato
                                Case "N"
                                    strValores = C.Value
                                Case "T"
                                    strValores = "'" & C.Value & "'"
                                Case "F"
                                    strValores = "'" & Format(C.Value, "yyyy-mm-dd") & "'"
                            End Select
                            strCampos = strCampos & strCampo & "=" & strValores & ","
                        Else
                            strValCod = C.Value
                        End If
                End Select
            End If
        End If
    Next
    
    If Len(strCampos) > 1 Then strCampos = left(strCampos, Len(strCampos) - 1)
    Select Case tipoOperacion
        Case 0
            If Len(strValores) > 1 Then strValores = left(strValores, Len(strValores) - 1)
            CSqlC = "INSERT INTO " & strTabla & "(" & strCampos & strCampoEmpresa & strCampoFecReg & ") VALUES(" & strValores & strValorEmpresa & strValorFecReg & ")"
        Case 1
            CSqlC = "UPDATE " & strTabla & " SET " & strCampos & " WHERE " & strCampoCod & " = '" & strValCod & "'" & strCondEmpresa
    End Select
     
    indTrans = True
    Cn.BeginTrans
    
    '--- Graba controles
    If strCampos <> "" Then
        Cn.Execute CSqlC
    End If
    
    '--- Grabando Grilla Presentaciones
    Cn.Execute "DELETE FROM presentaciones WHERE idProducto = '" & txtCod_Producto.Text & "' AND idEmpresa = '" & glsEmpresa & "'"
        
    gPresentaciones.Dataset.First
    Do While Not gPresentaciones.Dataset.EOF
        If gPresentaciones.Columns.ColumnByFieldName("IdUM").Value <> "" And Val("" & gPresentaciones.Columns.ColumnByFieldName("Factor").Value) <> 0 Then
            CSqlC = "INSERT INTO presentaciones(Item,IdProducto,IdUM,Factor,idEmpresa) " & _
                   "VALUES(" & gPresentaciones.Columns.ColumnByFieldName("Item").Value & ",'" & _
                   txtCod_Producto.Text & "','" & _
                   gPresentaciones.Columns.ColumnByFieldName("IdUM").Value & "'," & _
                   gPresentaciones.Columns.ColumnByFieldName("Factor").Value & ",'" & _
                   glsEmpresa & "')"
            Cn.Execute CSqlC
        End If
        gPresentaciones.Dataset.Next
    Loop
    
    gAlmacenes.Dataset.First
    Do While Not gAlmacenes.Dataset.EOF
        If gAlmacenes.Columns.ColumnByFieldName("idAlmacen").Value <> "" Then
            If traerCampo("productosalmacen", "item", "idSucursal", gAlmacenes.Columns.ColumnByFieldName("idSucursal").Value, True, "idAlmacen = '" & gAlmacenes.Columns.ColumnByFieldName("idAlmacen").Value & "' AND idProducto = '" & txtCod_Producto.Text & "' AND idUMCompra = '" & txtCod_UMCompra.Text & "'") = "" Then
                CSqlC = "INSERT INTO productosalmacen(item,idEmpresa,idSucursal,idAlmacen,idProducto,idUMCompra,idUbicacion) " & _
                       "VALUES(" & gAlmacenes.Columns.ColumnByFieldName("Item").Value & ",'" & _
                       glsEmpresa & "','" & _
                       gAlmacenes.Columns.ColumnByFieldName("idSucursal").Value & "','" & _
                       gAlmacenes.Columns.ColumnByFieldName("idAlmacen").Value & "','" & _
                       txtCod_Producto.Text & "','" & _
                       txtCod_UMCompra.Text & "','" & _
                       Trim("" & gAlmacenes.Columns.ColumnByFieldName("idUbicacion").Value) & "')"
                Cn.Execute CSqlC
            Else
                CSqlC = "UPDATE productosalmacen SET idUbicacion = '" & Trim("" & gAlmacenes.Columns.ColumnByFieldName("idUbicacion").Value) & "' " & _
                       "WHERE idempresa = '" & glsEmpresa & "' and idAlmacen = '" & Trim("" & gAlmacenes.Columns.ColumnByFieldName("idAlmacen").Value) & "' "
                Cn.Execute CSqlC
                
            End If
        End If
        gAlmacenes.Dataset.Next
    Loop
    
    CodLote = ""
    gAlmacenes.Dataset.First
    Do While Not gAlmacenes.Dataset.EOF
        If gAlmacenes.Columns.ColumnByFieldName("idAlmacen").Value <> "" Then
            CodLote = traerCampo("Sucursales", "idLote", "idSucursal", Trim("" & gAlmacenes.Columns.ColumnByFieldName("idSucursal").Value), True)
            If traerCampo("productosalmacenporlote", "item", "idSucursal", gAlmacenes.Columns.ColumnByFieldName("idSucursal").Value, True, "idAlmacen = '" & gAlmacenes.Columns.ColumnByFieldName("idAlmacen").Value & "' AND idProducto = '" & txtCod_Producto.Text & "' AND idUMCompra = '" & txtCod_UMCompra.Text & "' and idlote = '" & CodLote & "' ") = "" Then
                CSqlC = "INSERT INTO productosalmacenporlote(idLote,item,idEmpresa,idSucursal,idAlmacen,idProducto,idUMCompra) " & _
                       "VALUES('" & CodLote & "'," & gAlmacenes.Columns.ColumnByFieldName("Item").Value & ",'" & _
                       glsEmpresa & "','" & _
                       gAlmacenes.Columns.ColumnByFieldName("idSucursal").Value & "','" & _
                       gAlmacenes.Columns.ColumnByFieldName("idAlmacen").Value & "','" & _
                       txtCod_Producto.Text & "','" & _
                       txtCod_UMCompra.Text & "')"
                Cn.Execute CSqlC
            End If
        End If
        gAlmacenes.Dataset.Next
    Loop
    
    actualizaPrecios txtCod_Producto.Text, txtCod_UMVenta.Text, txtVal_VV.Value, txtVal_IGV.Value, txtVal_PV.Value, StrMsgError
    If StrMsgError <> "" Then GoTo Err
    
    CSqlC = "Select A.IdSucursal " & _
            "From Sucursales A " & _
            "Where A.IdEmpresa = '" & glsEmpresa & "'"
    RsC.Open CSqlC, Cn, adOpenStatic, adLockReadOnly
    Do While Not RsC.EOF
        
        If Len(Trim(traerCampo("ProductosStock", "IdProducto", "IdProducto", txtCod_Producto.Text, True, "IdSucursal = '" & Trim("" & RsC.Fields("IdSucursal")) & "'"))) = 0 Then
        
            CSqlC = "Insert Into ProductosStock(IdEmpresa,IdSucursal,IdProducto,Stock,Separacion,Disponible)Values" & _
                    "('" & glsEmpresa & "','" & Trim("" & RsC.Fields("IdSucursal")) & "','" & txtCod_Producto.Text & "',0,0,0)"
            
            Cn.Execute CSqlC
            
        End If
    
        If Len(Trim(traerCampo("ProductosStockDisponible", "IdProducto", "IdProducto", txtCod_Producto.Text, True, "IdSucursal = '" & Trim("" & RsC.Fields("IdSucursal")) & "'"))) = 0 Then
        
            CSqlC = "Insert Into ProductosStockDisponible(IdEmpresa,IdSucursal,IdProducto,Stock,Separacion,Disponible)Values" & _
                    "('" & glsEmpresa & "','" & Trim("" & RsC.Fields("IdSucursal")) & "','" & txtCod_Producto.Text & "',0,0,0)"
            
            Cn.Execute CSqlC
            
        End If
        
        RsC.MoveNext
        
    Loop
    
    RsC.Close: Set RsC = Nothing
    
    Cn.CommitTrans
    
    Exit Sub
Err:
    If RsC.State = 1 Then RsC.Close: Set RsC = Nothing
    If StrMsgError = "" Then StrMsgError = Err.Description
    If indTrans Then Cn.RollbackTrans
End Sub

Private Sub TxtCodDetraccion_Change()
On Error GoTo Err
Dim StrMsgError                     As String
Dim CIdDetraccion                   As String

    If TxtCodDetraccion.Text <> "" Then
        
        TxtGlsDetraccion.Text = traerCampo("Tb_Concep_Detrac", "Descripcion", "CodConcepto", TxtCodDetraccion.Text, False)
        TxtPorcentajeDetraccion.Text = Val("" & traerCampo("Tb_Concep_Detrac", "Porcentaje", "CodConcepto", TxtCodDetraccion.Text, False))
    
    Else
        
        TxtGlsDetraccion.Text = ""
        TxtPorcentajeDetraccion.Text = "0"
        
    End If
    
    Exit Sub
Err:
    If StrMsgError = "" Then StrMsgError = Err.Description
    MsgBox StrMsgError, vbInformation, App.Title
End Sub

'Private Sub txtctacontableRelacionada_Change()
'
'    If Len(txtctacontableRelacionada.Text) = 0 Then
'        txtctacontableRelacionada.Text = ""
'        txtdescta3.Text = ""
'    Else
'        txtdescta3.Text = traerCampoConta("Plancuentas", "glsNombreCuenta", "idctacontable", txtctacontableRelacionada.Text, False, " idanno = '2011' ")
'    End If
'
'End Sub

'Private Sub txtctacontableRelacionada_GotFocus()
'
'    txtctacontableRelacionada.SelStart = 0: txtctacontableRelacionada.SelLength = Len(txtctacontableRelacionada.Text)
'
'End Sub

'Private Sub txtctacontableRelacionada_KeyDown(KeyCode As Integer, Shift As Integer)
'
'    If KeyCode = 113 Then
'        cmdayuda3_Click
'    End If
'
'End Sub

Private Sub TxtIdConceptoCosteo_Change()
On Error GoTo Err
Dim StrMsgError                         As String
    
    If TxtIdConceptoCosteo.Text <> "" Then
        TxtGlsConceptoCosteo.Text = "" & traerCampo("Datos", "GlsDato", "IdDato", TxtIdConceptoCosteo.Text, False, "IdTipoDatos = '26'")
    Else
        TxtGlsConceptoCosteo.Text = ""
    End If
    
    Exit Sub
Err:
    If StrMsgError = "" Then StrMsgError = Err.Description
    MsgBox StrMsgError, vbInformation, App.Title
End Sub

Private Sub txtVal_PV_Change()
    
    If indCargando Then Exit Sub
    If indCalculando Then Exit Sub
    
    indCalculando = True
    If chkAfecto.Value Then
        txtVal_VV.Text = Val(txtVal_PV.Value) / (glsIGV + 1)
        txtVal_IGV.Text = Val(txtVal_PV.Value) - Val(txtVal_VV.Value)
    Else
        txtVal_IGV.Text = 0#
        txtVal_VV.Text = txtVal_PV.Value
    End If
    indCalculando = False

End Sub

Private Sub txtVal_VV_Change()
    
    If indCargando Then Exit Sub
    If indCalculando Then Exit Sub
    
    indCalculando = True
    If chkAfecto.Value Then
        txtVal_IGV.Text = Val(txtVal_VV.Value) * glsIGV
        txtVal_PV.Text = Val(txtVal_VV.Value) + Val(txtVal_IGV.Value)
    Else
        txtVal_IGV.Text = 0#
        txtVal_PV.Text = txtVal_VV.Value
    End If
    indCalculando = False

End Sub

Private Sub actualizaPrecios(ByVal strCodProd As String, ByVal strCodUM As String, ByVal dblVVUnit As Double, ByVal dblIGVUnit As Double, ByVal dblPVUnit As Double, ByRef StrMsgError As String)
On Error GoTo Err
Dim rsl As New ADODB.Recordset

    If glsListaVentas <> "" Then
        csql = "SELECT idLista FROM preciosventa " & _
               "WHERE idEmpresa = '" & glsEmpresa & "' " & _
                 "AND idLista = '" & glsListaVentas & "' " & _
                 "AND idProducto = '" & strCodProd & "' " & _
                 "AND idUM = '" & strCodUM & "'"
        rsl.Open csql, Cn, adOpenForwardOnly, adLockReadOnly
        
        If rsl.EOF Then '--- Inserto
            csql = "INSERT INTO preciosventa (idEmpresa,idLista,idProducto,idUM,VVUnit,IGVUnit,PVUnit,MaxDcto) " & _
                   "VALUES('" & glsEmpresa & "','" & glsListaVentas & "','" & strCodProd & "'," & _
                          "'" & strCodUM & "'," & dblVVUnit & "," & dblIGVUnit & "," & dblPVUnit & "," & TxtDctoListaPrec.Text & ")"
        Else
            csql = "UPDATE preciosventa  SET MaxDcto = " & TxtDctoListaPrec.Text & ",VVUnit = " & dblVVUnit & ",IGVUnit = " & dblIGVUnit & ",PVUnit = " & dblPVUnit & _
                   " WHERE idEmpresa = '" & glsEmpresa & "' AND idLista = '" & glsListaVentas & "' AND idProducto = '" & strCodProd & "' " & _
                   " AND idUM = '" & strCodUM & "'"
        End If
        Cn.Execute csql
    End If
    If rsl.State = 1 Then rsl.Close: Set rsl = Nothing
    
    Exit Sub
    
Err:
    If rsl.State = 1 Then rsl.Close: Set rsl = Nothing
    If StrMsgError = "" Then StrMsgError = Err.Description
End Sub
