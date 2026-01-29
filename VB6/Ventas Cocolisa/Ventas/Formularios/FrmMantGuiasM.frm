VERSION 5.00
Object = "{6A24B331-7634-11D3-A5B0-0050044A7E1A}#1.5#0"; "DXDBGrid.dll"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{F41D1D30-7878-4923-8CB3-6CCACDC9C9DE}#1.0#0"; "catcontrols.ocx"
Begin VB.Form FrmMantGuiasM 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Registro Guías Madre"
   ClientHeight    =   9975
   ClientLeft      =   3585
   ClientTop       =   615
   ClientWidth     =   11565
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   9975
   ScaleWidth      =   11565
   Begin MSComctlLib.ImageList imgDocVentas 
      Index           =   0
      Left            =   0
      Top             =   4560
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
            Picture         =   "FrmMantGuiasM.frx":0000
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmMantGuiasM.frx":039A
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmMantGuiasM.frx":07EC
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmMantGuiasM.frx":0B86
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmMantGuiasM.frx":0F20
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmMantGuiasM.frx":12BA
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmMantGuiasM.frx":1654
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmMantGuiasM.frx":19EE
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmMantGuiasM.frx":1D88
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmMantGuiasM.frx":2122
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmMantGuiasM.frx":24BC
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmMantGuiasM.frx":317E
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList imgDocVentas 
      Index           =   1
      Left            =   0
      Top             =   4560
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
            Picture         =   "FrmMantGuiasM.frx":3518
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmMantGuiasM.frx":38B2
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmMantGuiasM.frx":3D04
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmMantGuiasM.frx":409E
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmMantGuiasM.frx":4438
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmMantGuiasM.frx":47D2
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmMantGuiasM.frx":4B6C
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmMantGuiasM.frx":4F06
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmMantGuiasM.frx":52A0
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmMantGuiasM.frx":563A
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmMantGuiasM.frx":59D4
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmMantGuiasM.frx":6696
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   660
      Left            =   0
      TabIndex        =   76
      Top             =   0
      Width           =   11565
      _ExtentX        =   20399
      _ExtentY        =   1164
      ButtonWidth     =   2064
      ButtonHeight    =   1005
      Appearance      =   1
      ImageList       =   "imgDocVentas(0)"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   9
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "      Nuevo      "
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
            ImageIndex      =   6
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Reporte 1"
            Object.ToolTipText     =   "Imprimir"
            ImageIndex      =   11
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Reporte2"
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
   Begin VB.Frame Fra_Lista 
      Appearance      =   0  'Flat
      ForeColor       =   &H80000008&
      Height          =   9285
      Left            =   45
      TabIndex        =   73
      Top             =   630
      Width           =   11445
      Begin VB.Frame Frame6 
         Appearance      =   0  'Flat
         ForeColor       =   &H80000008&
         Height          =   735
         Left            =   120
         TabIndex        =   74
         Top             =   135
         Width           =   11160
         Begin VB.ComboBox cbx_Mes 
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
            ItemData        =   "FrmMantGuiasM.frx":6A30
            Left            =   7605
            List            =   "FrmMantGuiasM.frx":6A58
            Style           =   2  'Dropdown List
            TabIndex        =   1
            Top             =   240
            Width           =   1620
         End
         Begin CATControls.CATTextBox txt_TextoBuscar 
            Height          =   315
            Left            =   1185
            TabIndex        =   0
            Top             =   240
            Width           =   5745
            _ExtentX        =   10134
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
            Container       =   "FrmMantGuiasM.frx":6AC1
            Estilo          =   1
            Vacio           =   -1  'True
            EnterTab        =   -1  'True
         End
         Begin CATControls.CATTextBox txt_Ano 
            Height          =   315
            Left            =   10035
            TabIndex        =   2
            Top             =   240
            Width           =   870
            _ExtentX        =   1535
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
            Container       =   "FrmMantGuiasM.frx":6ADD
            Estilo          =   3
            Vacio           =   -1  'True
            EnterTab        =   -1  'True
         End
         Begin VB.Label Label26 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            Caption         =   "Año"
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
            Left            =   9630
            TabIndex        =   80
            Top             =   285
            Width           =   300
         End
         Begin VB.Label Label24 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            Caption         =   "Mes"
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
            Left            =   7110
            TabIndex        =   79
            Top             =   280
            Width           =   300
         End
         Begin VB.Label Label25 
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
            TabIndex        =   68
            Top             =   280
            Width           =   735
         End
      End
      Begin DXDBGRIDLibCtl.dxDBGrid gLista 
         Height          =   8145
         Left            =   135
         OleObjectBlob   =   "FrmMantGuiasM.frx":6AF9
         TabIndex        =   3
         Top             =   975
         Width           =   11160
      End
   End
   Begin VB.Frame Fra_Registro 
      Appearance      =   0  'Flat
      ForeColor       =   &H80000008&
      Height          =   9285
      Left            =   45
      TabIndex        =   69
      Top             =   630
      Width           =   11445
      Begin VB.Frame Frame3 
         Appearance      =   0  'Flat
         Caption         =   " Producto "
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
         Height          =   4065
         Left            =   450
         TabIndex        =   72
         Top             =   5130
         Width           =   10545
         Begin VB.Frame Frame4 
            Appearance      =   0  'Flat
            Caption         =   "Mortalidad"
            ForeColor       =   &H80000008&
            Height          =   780
            Left            =   270
            TabIndex        =   91
            Top             =   2115
            Width           =   4065
            Begin VB.TextBox TxtMortCantidad 
               Alignment       =   1  'Right Justify
               Appearance      =   0  'Flat
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
               Left            =   945
               TabIndex        =   93
               Top             =   270
               Width           =   825
            End
            Begin VB.TextBox TxtMortPeso 
               Alignment       =   1  'Right Justify
               Appearance      =   0  'Flat
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
               Left            =   3015
               TabIndex        =   92
               Top             =   270
               Width           =   825
            End
            Begin VB.Label Label29 
               AutoSize        =   -1  'True
               Caption         =   "Cantidad"
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
               Left            =   180
               TabIndex        =   95
               Top             =   315
               Width           =   630
            End
            Begin VB.Label Label28 
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
               Height          =   210
               Left            =   2385
               TabIndex        =   94
               Top             =   315
               Width           =   360
            End
         End
         Begin VB.TextBox Txt_ValPeso 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
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
            Left            =   1575
            TabIndex        =   15
            Top             =   675
            Width           =   825
         End
         Begin VB.Frame Frame7 
            Appearance      =   0  'Flat
            Caption         =   "Horario"
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
            Height          =   1050
            Left            =   270
            TabIndex        =   75
            Top             =   2925
            Width           =   9960
            Begin MSComCtl2.DTPicker Dtp_HoraLlegada 
               Height          =   330
               Left            =   1575
               TabIndex        =   24
               Top             =   585
               Width           =   1410
               _ExtentX        =   2487
               _ExtentY        =   582
               _Version        =   393216
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Format          =   103809026
               CurrentDate     =   40718
            End
            Begin MSComCtl2.DTPicker Dtp_HoraAtencion 
               Height          =   330
               Left            =   5085
               TabIndex        =   25
               Top             =   585
               Width           =   1410
               _ExtentX        =   2487
               _ExtentY        =   582
               _Version        =   393216
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Format          =   103809026
               CurrentDate     =   40718
            End
            Begin MSComCtl2.DTPicker Dtp_HoraSalida 
               Height          =   330
               Left            =   8055
               TabIndex        =   26
               Top             =   585
               Width           =   1410
               _ExtentX        =   2487
               _ExtentY        =   582
               _Version        =   393216
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Format          =   103809026
               CurrentDate     =   40718
            End
            Begin MSComCtl2.DTPicker DtpPreparacion 
               Height          =   330
               Left            =   1575
               TabIndex        =   21
               Top             =   180
               Width           =   1410
               _ExtentX        =   2487
               _ExtentY        =   582
               _Version        =   393216
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Format          =   103809026
               CurrentDate     =   40718
            End
            Begin MSComCtl2.DTPicker DtpllegadaGranja 
               Height          =   330
               Left            =   5085
               TabIndex        =   22
               Top             =   180
               Width           =   1410
               _ExtentX        =   2487
               _ExtentY        =   582
               _Version        =   393216
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Format          =   103809026
               CurrentDate     =   40718
            End
            Begin MSComCtl2.DTPicker DtpPartida 
               Height          =   330
               Left            =   8055
               TabIndex        =   23
               Top             =   180
               Width           =   1410
               _ExtentX        =   2487
               _ExtentY        =   582
               _Version        =   393216
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Format          =   103809026
               CurrentDate     =   40718
            End
            Begin VB.Label lblPreparacion 
               AutoSize        =   -1  'True
               Caption         =   "Preparacion"
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
               Left            =   630
               TabIndex        =   84
               Top             =   225
               Width           =   870
            End
            Begin VB.Label lblllegadaagranja 
               AutoSize        =   -1  'True
               Caption         =   "Llegada a Granja"
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
               Left            =   3780
               TabIndex        =   83
               Top             =   225
               Width           =   1230
            End
            Begin VB.Label lblpartida 
               AutoSize        =   -1  'True
               Caption         =   "Partida Camal"
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
               Left            =   6975
               TabIndex        =   82
               Top             =   225
               Width           =   1515
            End
            Begin VB.Label Label19 
               AutoSize        =   -1  'True
               Caption         =   "Salida Camal"
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
               Left            =   6975
               TabIndex        =   67
               Top             =   630
               Width           =   915
            End
            Begin VB.Label Label18 
               AutoSize        =   -1  'True
               Caption         =   "Atención"
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
               Left            =   3780
               TabIndex        =   66
               Top             =   630
               Width           =   645
            End
            Begin VB.Label Label17 
               AutoSize        =   -1  'True
               Caption         =   "Llegada"
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
               Left            =   630
               TabIndex        =   65
               Top             =   630
               Width           =   570
            End
         End
         Begin VB.CommandButton Cmd_Cliente 
            Height          =   315
            Left            =   9810
            Picture         =   "FrmMantGuiasM.frx":991A
            Style           =   1  'Graphical
            TabIndex        =   64
            Top             =   1755
            Width           =   390
         End
         Begin VB.CommandButton Cmd_Producto 
            Height          =   315
            Left            =   9810
            Picture         =   "FrmMantGuiasM.frx":9CA4
            Style           =   1  'Graphical
            TabIndex        =   56
            Top             =   270
            Width           =   390
         End
         Begin VB.TextBox Txt_GlsProducto 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFC0&
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
            Left            =   2430
            Locked          =   -1  'True
            TabIndex        =   55
            Top             =   270
            Width           =   7350
         End
         Begin VB.TextBox Txt_IdProducto 
            Appearance      =   0  'Flat
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
            Left            =   1575
            TabIndex        =   14
            Top             =   270
            Width           =   825
         End
         Begin VB.TextBox Txt_ValCantidad 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
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
            Left            =   8955
            TabIndex        =   16
            Top             =   675
            Width           =   825
         End
         Begin VB.TextBox Txt_ValFxN 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
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
            Left            =   1575
            TabIndex        =   17
            Top             =   1035
            Width           =   825
         End
         Begin VB.TextBox Txt_ValEX 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
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
            Left            =   8955
            TabIndex        =   18
            Top             =   1035
            Width           =   825
         End
         Begin VB.TextBox Txt_GlsNPrecinto 
            Appearance      =   0  'Flat
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
            Left            =   1575
            TabIndex        =   19
            Top             =   1395
            Width           =   825
         End
         Begin VB.TextBox Txt_GlsCliente 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFC0&
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
            Left            =   2430
            Locked          =   -1  'True
            TabIndex        =   63
            Top             =   1755
            Width           =   7350
         End
         Begin VB.TextBox Txt_IdCliente 
            Appearance      =   0  'Flat
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
            Left            =   1575
            TabIndex        =   20
            Top             =   1755
            Width           =   825
         End
         Begin VB.Label Label22 
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
            Height          =   210
            Left            =   360
            TabIndex        =   57
            Top             =   720
            Width           =   360
         End
         Begin VB.Label Label11 
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
            Height          =   210
            Left            =   360
            TabIndex        =   54
            Top             =   315
            Width           =   495
         End
         Begin VB.Label Label12 
            AutoSize        =   -1  'True
            Caption         =   "Cantidad"
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
            Left            =   7965
            TabIndex        =   58
            Top             =   720
            Width           =   630
         End
         Begin VB.Label Label13 
            AutoSize        =   -1  'True
            Caption         =   "F x N"
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
            Left            =   360
            TabIndex        =   59
            Top             =   1080
            Width           =   375
         End
         Begin VB.Label Label14 
            AutoSize        =   -1  'True
            Caption         =   "EX"
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
            Left            =   7965
            TabIndex        =   60
            Top             =   1080
            Width           =   195
         End
         Begin VB.Label Label20 
            AutoSize        =   -1  'True
            Caption         =   "Nº Precinto"
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
            Left            =   360
            TabIndex        =   61
            Top             =   1440
            Width           =   810
         End
         Begin VB.Label Label21 
            AutoSize        =   -1  'True
            Caption         =   "Cliente"
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
            Left            =   360
            TabIndex        =   62
            Top             =   1845
            Width           =   480
         End
      End
      Begin VB.Frame Frame1 
         Appearance      =   0  'Flat
         ForeColor       =   &H80000008&
         Height          =   4965
         Left            =   450
         TabIndex        =   70
         Top             =   135
         Width           =   10545
         Begin VB.TextBox txtGls_Observacion 
            Appearance      =   0  'Flat
            Height          =   510
            Left            =   1710
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   10
            Top             =   2700
            Width           =   8130
         End
         Begin VB.TextBox txtGls_MotivoTraslado 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFC0&
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
            Left            =   2610
            Locked          =   -1  'True
            TabIndex        =   87
            Top             =   2295
            Width           =   7215
         End
         Begin VB.TextBox txtCod_MotivoTraslado 
            Appearance      =   0  'Flat
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
            Left            =   1710
            TabIndex        =   9
            Top             =   2295
            Width           =   825
         End
         Begin VB.CommandButton cmbAyudaMotivoTraslado 
            Height          =   315
            Left            =   9855
            Picture         =   "FrmMantGuiasM.frx":A02E
            Style           =   1  'Graphical
            TabIndex        =   85
            Top             =   2295
            Width           =   390
         End
         Begin VB.TextBox Txt_RucEmpTrans 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFC0&
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
            Left            =   8010
            Locked          =   -1  'True
            TabIndex        =   40
            Top             =   1890
            Width           =   1815
         End
         Begin VB.TextBox Txt_IdUPP 
            Appearance      =   0  'Flat
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
            Left            =   1710
            TabIndex        =   4
            Top             =   270
            Width           =   825
         End
         Begin VB.CommandButton Cmd_EmpTrans 
            Height          =   315
            Left            =   9855
            Picture         =   "FrmMantGuiasM.frx":A3B8
            Style           =   1  'Graphical
            TabIndex        =   41
            Top             =   1890
            Width           =   390
         End
         Begin VB.CommandButton Cmd_Proveedor 
            Height          =   315
            Left            =   9855
            Picture         =   "FrmMantGuiasM.frx":A742
            Style           =   1  'Graphical
            TabIndex        =   37
            Top             =   1485
            Width           =   390
         End
         Begin VB.CommandButton Cmd_Upp 
            Height          =   315
            Left            =   9855
            Picture         =   "FrmMantGuiasM.frx":AACC
            Style           =   1  'Graphical
            TabIndex        =   28
            Top             =   270
            Width           =   390
         End
         Begin VB.TextBox Txt_GlsUPP 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFC0&
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
            Left            =   2610
            Locked          =   -1  'True
            TabIndex        =   27
            Top             =   270
            Width           =   7215
         End
         Begin VB.Frame Frame2 
            Appearance      =   0  'Flat
            Caption         =   " Unidad de Transporte "
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
            Height          =   1635
            Left            =   270
            TabIndex        =   71
            Top             =   3240
            Width           =   9960
            Begin VB.TextBox Txt_GlsLicConducir 
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFC0&
               BorderStyle     =   0  'None
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
               Left            =   8190
               Locked          =   -1  'True
               TabIndex        =   77
               Top             =   1080
               Width           =   1140
            End
            Begin VB.CommandButton Cmd_Chofer 
               Height          =   315
               Left            =   9380
               Picture         =   "FrmMantGuiasM.frx":AE56
               Style           =   1  'Graphical
               TabIndex        =   53
               Top             =   1080
               Width           =   390
            End
            Begin VB.CommandButton Cmd_Vehiculo 
               Height          =   315
               Left            =   9380
               Picture         =   "FrmMantGuiasM.frx":B1E0
               Style           =   1  'Graphical
               TabIndex        =   44
               Top             =   270
               Width           =   390
            End
            Begin VB.TextBox Txt_IdVehiculo 
               Appearance      =   0  'Flat
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
               Left            =   1035
               TabIndex        =   12
               Top             =   270
               Width           =   1320
            End
            Begin VB.TextBox Txt_GlsVehiculo 
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFC0&
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
               Left            =   2370
               Locked          =   -1  'True
               TabIndex        =   43
               Top             =   270
               Width           =   6945
            End
            Begin VB.TextBox Txt_GlsMarca 
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFC0&
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
               Left            =   1035
               Locked          =   -1  'True
               TabIndex        =   46
               Top             =   675
               Width           =   1320
            End
            Begin VB.TextBox Txt_GlsPlaca 
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFC0&
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
               Left            =   2970
               Locked          =   -1  'True
               TabIndex        =   48
               Top             =   675
               Width           =   1140
            End
            Begin VB.TextBox Txt_GlsInscripcion 
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFC0&
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
               Left            =   8190
               Locked          =   -1  'True
               TabIndex        =   50
               Top             =   675
               Width           =   1140
            End
            Begin VB.TextBox Txt_IdChofer 
               Appearance      =   0  'Flat
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
               Left            =   1035
               TabIndex        =   13
               Top             =   1080
               Width           =   1320
            End
            Begin VB.TextBox Txt_GlsChofer 
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFC0&
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
               Left            =   2385
               Locked          =   -1  'True
               TabIndex        =   52
               Top             =   1080
               Width           =   3705
            End
            Begin CATControls.CATTextBox txt_Flete 
               Height          =   315
               Left            =   4905
               TabIndex        =   90
               Tag             =   "NValPeso"
               Top             =   675
               Width           =   1170
               _ExtentX        =   2064
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
               Container       =   "FrmMantGuiasM.frx":B56A
               Decimales       =   2
               Vacio           =   -1  'True
               EnterTab        =   -1  'True
            End
            Begin VB.Label Label27 
               Caption         =   "Flete"
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
               Left            =   4275
               TabIndex        =   89
               Top             =   720
               Width           =   465
            End
            Begin VB.Label Label23 
               AutoSize        =   -1  'True
               Caption         =   "Licencia Conducir Nº"
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
               Left            =   6525
               TabIndex        =   78
               Top             =   1170
               Width           =   1515
            End
            Begin VB.Label Label6 
               Caption         =   "Vehículo"
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
               Left            =   270
               TabIndex        =   42
               Top             =   315
               Width           =   645
            End
            Begin VB.Label Label7 
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
               Height          =   195
               Left            =   270
               TabIndex        =   45
               Top             =   720
               Width           =   645
            End
            Begin VB.Label Label8 
               Caption         =   "Placa"
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
               Left            =   2475
               TabIndex        =   47
               Top             =   720
               Width           =   510
            End
            Begin VB.Label Label9 
               AutoSize        =   -1  'True
               Caption         =   "Certificado Inscripción"
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
               Left            =   6525
               TabIndex        =   49
               Top             =   765
               Width           =   1605
            End
            Begin VB.Label Label10 
               Caption         =   "Chofer"
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
               Left            =   270
               TabIndex        =   51
               Top             =   1125
               Width           =   645
            End
         End
         Begin VB.TextBox Txt_NumGuia 
            Appearance      =   0  'Flat
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
            Left            =   3690
            TabIndex        =   5
            Top             =   675
            Width           =   960
         End
         Begin VB.TextBox Txt_SerieGuia 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFC0&
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
            Left            =   1710
            Locked          =   -1  'True
            TabIndex        =   30
            Top             =   675
            Width           =   825
         End
         Begin VB.TextBox Txt_GlsPartida 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFC0&
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
            Left            =   1710
            Locked          =   -1  'True
            TabIndex        =   34
            Top             =   1080
            Width           =   8115
         End
         Begin VB.TextBox Txt_IdProveedor 
            Appearance      =   0  'Flat
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
            Left            =   1710
            TabIndex        =   7
            Top             =   1485
            Width           =   825
         End
         Begin VB.TextBox Txt_GlsLlegada 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFC0&
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
            Left            =   2610
            Locked          =   -1  'True
            TabIndex        =   36
            Top             =   1485
            Width           =   7215
         End
         Begin VB.TextBox Txt_IdEmpTrans 
            Appearance      =   0  'Flat
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
            Left            =   1710
            TabIndex        =   8
            Top             =   1890
            Width           =   825
         End
         Begin VB.TextBox Txt_GlsEmpTrans 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFC0&
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
            Left            =   2610
            Locked          =   -1  'True
            TabIndex        =   39
            Top             =   1890
            Width           =   5370
         End
         Begin MSComCtl2.DTPicker Dtp_FechaGuia 
            Height          =   330
            Left            =   8640
            TabIndex        =   6
            Top             =   675
            Width           =   1185
            _ExtentX        =   2090
            _ExtentY        =   582
            _Version        =   393216
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Format          =   103809025
            CurrentDate     =   40718
         End
         Begin VB.Label lblobservaciones 
            Caption         =   "Observaciones"
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
            Left            =   360
            TabIndex        =   88
            Top             =   2790
            Width           =   1185
         End
         Begin VB.Label lblMotivo 
            Caption         =   "Motivo"
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
            Left            =   360
            TabIndex        =   86
            Top             =   2340
            Width           =   960
         End
         Begin VB.Label Label1 
            Caption         =   "Granja"
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
            Left            =   360
            TabIndex        =   11
            Top             =   315
            Width           =   645
         End
         Begin VB.Label Label2 
            Caption         =   "Serie"
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
            Left            =   360
            TabIndex        =   29
            Top             =   720
            Width           =   645
         End
         Begin VB.Label Label3 
            Caption         =   "Punto de Partida"
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
            Left            =   360
            TabIndex        =   33
            Top             =   1125
            Width           =   1185
         End
         Begin VB.Label Label4 
            Caption         =   "Punto de Llegada"
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
            Left            =   360
            TabIndex        =   35
            Top             =   1530
            Width           =   1275
         End
         Begin VB.Label Label5 
            Caption         =   "Transportista"
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
            Left            =   360
            TabIndex        =   38
            Top             =   1935
            Width           =   960
         End
         Begin VB.Label Label15 
            Caption         =   "Número"
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
            Left            =   3060
            TabIndex        =   31
            Top             =   720
            Width           =   645
         End
         Begin VB.Label Label16 
            AutoSize        =   -1  'True
            Caption         =   "Fecha"
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
            Left            =   8010
            TabIndex        =   32
            Top             =   720
            Width           =   450
         End
      End
      Begin VB.Label Lbl_Ayuda 
         Caption         =   "Ayuda F2 ó Dbl. Click"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   6030
         TabIndex        =   81
         Top             =   225
         Visible         =   0   'False
         Width           =   1860
      End
   End
End
Attribute VB_Name = "FrmMantGuiasM"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim IndNuevo                        As Boolean
Dim CIdUPPAux                       As String
Dim CNumGuiaAux                     As String
Dim CIdProveedorAux                 As String

Private Sub Validaciones(StrMsgError As String, PSw1 As Boolean, PSw2 As Boolean)
On Error GoTo Err
    
    If PSw1 Then
        If Val("" & traerCampo("DocVentasGuiasM", "IndImportado", "IdUPP", CIdUPPAux, True, " NumGuia = '" & CNumGuiaAux & "'")) = 1 Then
            StrMsgError = "La Guía Madre se encuentra importada, verifique": GoTo Err
        End If
    End If
    
    If PSw2 Then
        If Txt_IdUPP.Text = "" Then
            StrMsgError = "Ingrese Unidad de Producción": Txt_IdUPP.SetFocus: GoTo Err
        ElseIf Txt_GlsUPP.Text = "" Then
            StrMsgError = "Unidad de Producción ingresada no existe": Txt_IdUPP.SetFocus: GoTo Err
        End If
        
        If Txt_NumGuia.Text = "" Then
            StrMsgError = "Ingrese Número de Guía": Txt_NumGuia.SetFocus: GoTo Err
        End If
        
        If txtCod_MotivoTraslado.Text = "06090013" Then
            If Len(Trim("" & txtGls_Observacion.Text)) = 0 Then
                StrMsgError = "El Motivo Necesita una Observacion": txtGls_Observacion.SetFocus: GoTo Err
            End If
        End If
        
        If Txt_IdProveedor.Text = "" Then
            StrMsgError = "Ingrese Punto de Llegada": Txt_IdProveedor.SetFocus: GoTo Err
        ElseIf Txt_GlsLlegada.Text = "" Then
            StrMsgError = "Punto de Llegada ingresada no existe": Txt_IdProveedor.SetFocus: GoTo Err
        End If
        
        If Txt_IdEmpTrans.Text = "" Then
            StrMsgError = "Ingrese Transportista": Txt_IdEmpTrans.SetFocus: GoTo Err
        ElseIf Txt_GlsEmpTrans.Text = "" Then
            StrMsgError = "Transportista ingresado no existe": Txt_IdEmpTrans.SetFocus: GoTo Err
        End If
        
        If Txt_IdVehiculo.Text = "" Then
            StrMsgError = "Ingrese Vehículo": Txt_IdVehiculo.SetFocus: GoTo Err
        ElseIf Txt_GlsVehiculo.Text = "" Then
            StrMsgError = "Vehículo ingresado no existe": Txt_IdVehiculo.SetFocus: GoTo Err
        End If
        
        If Txt_IdChofer.Text = "" Then
            StrMsgError = "Ingrese Chofer": Txt_IdChofer.SetFocus: GoTo Err
        ElseIf Txt_GlsLicConducir.Text = "" Then
            StrMsgError = "Ingrese Licencia de Conducir del Chofer": Txt_IdChofer.SetFocus: GoTo Err
        End If
        
        If Txt_IdProducto.Text = "" Then
            StrMsgError = "Ingrese Producto": Txt_IdProducto.SetFocus: GoTo Err
        ElseIf Txt_GlsProducto.Text = "" Then
            StrMsgError = "Producto ingresado no existe": Txt_IdProducto.SetFocus: GoTo Err
        End If
        
        If Val("" & Txt_ValPeso.Text) = "0" Then
            StrMsgError = "Ingrese Peso": Txt_ValPeso.SetFocus: GoTo Err
        End If
        
        If Val("" & Txt_ValCantidad.Text) = "0" Then
            StrMsgError = "Ingrese Cantidad": Txt_ValCantidad.SetFocus: GoTo Err
        End If
        
        If Val("" & Txt_ValFxN.Text) = "0" Then
            StrMsgError = "Ingrese FxN": Txt_ValFxN.SetFocus: GoTo Err
        End If
        
        If Val("" & Txt_ValEX.Text) = "0" Then
            StrMsgError = "Ingrese EX": Txt_ValEX.SetFocus: GoTo Err
        End If
        
        If Txt_GlsNPrecinto.Text = "" Then
            StrMsgError = "Ingrese Nº de Precinto": Txt_GlsNPrecinto.SetFocus: GoTo Err
        End If
        
        If Txt_IdCliente.Text = "" Then
            StrMsgError = "Ingrese Cliente": Txt_IdCliente.SetFocus: GoTo Err
        ElseIf Txt_GlsChofer.Text = "" Then
            StrMsgError = "Cliente ingresado no existe": Txt_IdCliente.SetFocus: GoTo Err
        End If
        
    End If
    
    Exit Sub
    
Err:
    If StrMsgError = "" Then StrMsgError = Err.Description
End Sub

Private Sub eliminar(StrMsgError As String)
On Error GoTo Err
Dim cdelete                         As String
Dim indTrans                        As Boolean
Dim CIdValesCabI                    As String
Dim CIdValesCabS                    As String
Dim CArr(2)                         As String
Dim cselect                         As String
Dim stralmacenOri                   As String
Dim stralmacenDes                   As String
Dim StrAlSucursalOri                As String
Dim StrAlSucursalDes                As String
Dim StrperiodoinvOri                As String
Dim StrperiodoinvDes                As String
    
    Cn.BeginTrans
    indTrans = True
        
    stralmacenOri = traerCampo("unidadproduccion", "IdAlmacen", "CodUnidProd", Trim("" & Txt_IdUPP.Text), True)
    StrAlSucursalOri = traerCampo("Almacenes", "idSucursal", "IdAlmacen", Trim("" & stralmacenOri), True)
    
    stralmacenDes = traerCampo("unidadproduccion", "IdAlmacen", "CodUnidProd", Trim("" & Txt_IdProveedor.Text), True)
    StrAlSucursalDes = traerCampo("Almacenes", "idSucursal", "IdAlmacen", Trim("" & stralmacenDes), True)
        
    StrperiodoinvOri = traerCampo("periodosinv", "idPeriodoInv", "idSucursal", Trim("" & StrAlSucursalOri), True, " estPeriodoInv = 'ACT' ")
    StrperiodoinvDes = traerCampo("periodosinv", "idPeriodoInv", "idSucursal", Trim("" & StrAlSucursalDes), True, " estPeriodoInv = 'ACT' ")
        
    traerCampos "DocVentasGuiasM", "IdValesCabI,IdValesCabS", "IdSucursal", glsSucursal, 2, CArr, True, "IdUPP = '" & Txt_IdUPP.Text & "' And NumGuia = '" & Txt_NumGuia.Text & "'"
    
    CIdValesCabI = Trim("" & CArr(0))
    CIdValesCabS = Trim("" & CArr(1))
    
    If Len(Trim(CIdValesCabI)) > 0 Then
        actualizaStock_Liquidaciones CIdValesCabS, 1, StrMsgError, "S", StrAlSucursalOri, False
        If StrMsgError <> "" Then GoTo Err
        
        actualizaStock_Liquidaciones CIdValesCabI, 1, StrMsgError, "I", StrAlSucursalDes, False
        If StrMsgError <> "" Then GoTo Err
        
        cselect = "Delete A,B " & _
                    "From ValesCab A " & _
                    "Inner Join ValesDet B " & _
                    "On A.IdEmpresa = B.IdEmpresa And A.IdSucursal = B.IdSucursal And A.TipoVale = B.TipoVale And A.IdValesCab = B.IdValesCab " & _
                    "Where A.IdEmpresa = '" & glsEmpresa & "' And A.TipoVale = If(A.IdValesCab = '" & CIdValesCabI & "','I','S') " & _
                    "And A.IdValesCab In('" & CIdValesCabI & "','" & CIdValesCabS & "')"
        Cn.Execute cselect
    End If
    
    
    cselect = "Delete from DocReferencia " & _
              "where idempresa = '" & glsEmpresa & "' and TipoDocOrigen = '99' and NumDocOrigen = '" & CIdValesCabS & "' and SerieDocOrigen = '000' " & _
              "and idSucursal = '" & StrAlSucursalOri & "' "
    Cn.Execute cselect
    
    cselect = "Delete from DocReferencia " & _
              "where idempresa = '" & glsEmpresa & "' and TipoDocOrigen = '88' and NumDocOrigen = '" & CIdValesCabI & "' and SerieDocOrigen = '000' " & _
              "and idSucursal = '" & StrAlSucursalDes & "' "
    Cn.Execute cselect
    
    cdelete = "Delete From DocVentasGuiasM " & _
              "Where IdEmpresa = '" & glsEmpresa & "' And IdSucursal = '" & glsSucursal & "' And IdUPP = '" & CIdUPPAux & "' " & _
              "And NumGuia = '" & CNumGuiaAux & "'"
    Cn.Execute cdelete
        
    Cn.CommitTrans
    indTrans = False
    
    MsgBox "Se Eliminó Correctamente", vbInformation, App.Title
    
    Exit Sub
    
Err:
    If indTrans Then Cn.RollbackTrans
    If StrMsgError = "" Then StrMsgError = Err.Description
End Sub

Private Sub Grabar(StrMsgError As String)
On Error GoTo Err
Dim cupdate                         As String
Dim CFecSystem                      As String
Dim indTrans                        As Boolean
    
    Cn.BeginTrans
    indTrans = True
    
    CFecSystem = getFechaHoraSistema
    
    If IndNuevo Then
        
        cupdate = "Insert Into DocVentasGuiasM(IdEmpresa,IdSucursal,IdUPP,SerieGuia,NumGuia,FechaGuia,GlsPartida,IdProveedor,IdEmpTrans,IdVehiculo," & _
                  "IdChofer,IdProducto,ValPeso,ValCantidad,ValFxN,ValEX,HoraLlegada,HoraAtencion,HoraSalida,GlsPrecinto,IdCliente,FecRegistro," & _
                  "HoraRegistro,IdUsuarioRegistro,IndImportado,HoraPreparacion, HoraLlegadaGranja, HoraPartida,GlsObservacion,idMotivoTraslado,Flete,MortCantidad,MortPeso)Values(" & _
                  "'" & glsEmpresa & "','" & glsSucursal & "','" & Txt_IdUPP.Text & "','" & Txt_SerieGuia.Text & "','" & Txt_NumGuia.Text & "'," & _
                  "'" & Format(Dtp_FechaGuia.Value, "yyyy-mm-dd") & "','" & Txt_GlsPartida.Text & "','" & Txt_IdProveedor.Text & "'," & _
                  "'" & Txt_IdEmpTrans.Text & "','" & Txt_IdVehiculo.Text & "','" & Txt_IdChofer.Text & "','" & Txt_IdProducto.Text & "'," & _
                  "" & Txt_ValPeso.Text & "," & Txt_ValCantidad.Text & "," & Txt_ValFxN.Text & "," & Txt_ValEX.Text & ",'" & Format(Dtp_HoraLlegada.Value, "h:mm:ss") & "'," & _
                  "'" & Format(Dtp_HoraAtencion.Value, "h:mm:ss") & "','" & Format(Dtp_HoraSalida.Value, "h:mm:ss") & "','" & Txt_GlsNPrecinto.Text & "','" & Txt_IdCliente.Text & "'," & _
                  "'" & Format(CFecSystem, "yyyy-mm-dd") & "','" & Format(CFecSystem, "h:mm:ss") & "','" & glsUser & "',0, " & _
                  "'" & Format(DtpPreparacion.Value, "h:mm:ss") & "','" & Format(DtpllegadaGranja.Value, "h:mm:ss") & "','" & Format(DtpPartida.Value, "h:mm:ss") & "','" & Trim("" & txtGls_Observacion.Text) & "','" & Trim("" & txtCod_MotivoTraslado.Text) & "'," & Val(Format(txt_Flete.Text, "0.00")) & "," & Val(TxtMortCantidad.Text) & "," & Val(Format(TxtMortPeso.Text, "0.00")) & ")"
        
    Else
    
        cupdate = "Update DocVentasGuiasM " & _
                  "Set IdUPP = '" & Txt_IdUPP.Text & "',SerieGuia = '" & Txt_SerieGuia.Text & "',NumGuia = '" & Txt_NumGuia.Text & "'," & _
                  "FechaGuia = '" & Format(Dtp_FechaGuia.Value, "yyyy-mm-dd") & "',GlsPartida = '" & Txt_GlsPartida.Text & "'," & _
                  "IdProveedor = '" & Txt_IdProveedor.Text & "',IdEmpTrans = '" & Txt_IdEmpTrans.Text & "'," & _
                  "IdVehiculo = '" & Txt_IdVehiculo.Text & "',IdChofer = '" & Txt_IdChofer.Text & "',IdProducto = '" & Txt_IdProducto.Text & "'," & _
                  "ValPeso = " & Txt_ValPeso.Text & ",ValCantidad = " & Txt_ValCantidad.Text & ",ValFxN = " & Txt_ValFxN.Text & "," & _
                  "ValEX = " & Txt_ValEX.Text & ",HoraLlegada = '" & Format(Dtp_HoraLlegada.Value, "h:mm:ss") & "',HoraAtencion = '" & Format(Dtp_HoraAtencion.Value, "h:mm:ss") & "'," & _
                  "HoraSalida = '" & Format(Dtp_HoraSalida.Value, "h:mm:ss") & "',GlsPrecinto = '" & Txt_GlsNPrecinto.Text & "'," & _
                  "IdCliente = '" & Txt_IdCliente.Text & "',FechaModificado = '" & Format(CFecSystem, "yyyy-mm-dd") & "'," & _
                  "HoraModificado = '" & Format(CFecSystem, "h:mm:ss") & "',IdUsuarioModificado = '" & glsUser & "' " & _
                  ",HoraPreparacion = '" & Format(DtpPreparacion.Value, "h:mm:ss") & "' " & _
                  ",HoraLlegadaGranja = '" & Format(DtpllegadaGranja.Value, "h:mm:ss") & "' " & _
                  ",HoraPartida = '" & Format(DtpPartida.Value, "h:mm:ss") & "' " & _
                  ",GlsObservacion = '" & Trim("" & txtGls_Observacion.Text) & "' " & _
                  ",idMotivoTraslado = '" & txtCod_MotivoTraslado.Text & "' " & _
                  ",Flete = " & Val(Format(txt_Flete.Text, "0.00")) & " " & _
                  ",MortCantidad = " & Val(TxtMortCantidad.Text) & " " & _
                  ",MortPeso = " & Val(Format(TxtMortPeso.Text, "0.00")) & " " & _
                  "Where IdEmpresa = '" & glsEmpresa & "' And IdUPP = '" & CIdUPPAux & "' " & _
                  "And NumGuia = '" & CNumGuiaAux & "'"
                  
    End If
    
    Cn.Execute cupdate
    
    Genera_Vale StrMsgError
    If StrMsgError <> "" Then GoTo Err
    
    Cn.CommitTrans
    indTrans = False
    
    CIdUPPAux = Txt_IdUPP.Text
    CNumGuiaAux = Txt_NumGuia.Text
    CIdProveedorAux = Txt_IdProveedor.Text
    
    If IndNuevo Then
        MsgBox "Se Grabó Correctamente", vbInformation, App.Title
    Else
        MsgBox "Se Modificó Correctamente", vbInformation, App.Title
    End If
    
    IndNuevo = False
    
    Exit Sub

Err:
    If indTrans Then Cn.RollbackTrans
    If StrMsgError = "" Then StrMsgError = Err.Description
End Sub

Private Sub Genera_Vale(StrMsgError As String)
On Error GoTo Err
Dim cselect                         As String
Dim CIdValesCabI                    As String
Dim CIdValesCabS                    As String
Dim CIdConceptoI                    As String
Dim CIdConceptoS                    As String
Dim CAbreviaturaDoc                 As String
Dim CArr(2)                         As String
Dim stralmacenOri                   As String
Dim stralmacenDes                   As String
Dim StrAlSucursalOri                As String
Dim StrAlSucursalDes                As String
Dim StrperiodoinvOri                As String
Dim StrperiodoinvDes                As String
Dim CIdUPPAnt                       As String
Dim CIdUPPDesAnt                    As String
    
    
    If Len(Trim("" & CIdUPPAux)) > 0 Then
        
        stralmacenOri = traerCampo("unidadproduccion", "IdAlmacen", "CodUnidProd", CIdUPPAux, True)
        
    Else
    
        stralmacenOri = traerCampo("unidadproduccion", "IdAlmacen", "CodUnidProd", Trim("" & Txt_IdUPP.Text), True)
    
    End If
    
    StrAlSucursalOri = traerCampo("Almacenes", "idSucursal", "IdAlmacen", Trim("" & stralmacenOri), True)
    
    If Len(Trim("" & CIdProveedorAux)) > 0 Then
    
        stralmacenDes = traerCampo("unidadproduccion", "IdAlmacen", "CodUnidProd", CIdProveedorAux, True)
    
    Else
        
        stralmacenDes = traerCampo("unidadproduccion", "IdAlmacen", "CodUnidProd", Trim("" & Txt_IdProveedor.Text), True)
        
    End If
    
    StrAlSucursalDes = traerCampo("Almacenes", "idSucursal", "IdAlmacen", Trim("" & stralmacenDes), True)
        
    StrperiodoinvOri = traerCampo("periodosinv", "idPeriodoInv", "idSucursal", Trim("" & StrAlSucursalOri), True, " estPeriodoInv = 'ACT' ")
    StrperiodoinvDes = traerCampo("periodosinv", "idPeriodoInv", "idSucursal", Trim("" & StrAlSucursalDes), True, " estPeriodoInv = 'ACT' ")
    
    If Not IndNuevo Then
        traerCampos "DocVentasGuiasM", "IdValesCabI,IdValesCabS", "IdSucursal", glsSucursal, 2, CArr, True, "IdUPP = '" & Txt_IdUPP.Text & "' And NumGuia = '" & Txt_NumGuia.Text & "'"
        
        CIdValesCabI = Trim("" & CArr(0))
        CIdValesCabS = Trim("" & CArr(1))
        
        If Len(Trim(CIdValesCabI)) > 0 Then
                        
            actualizaStock_Liquidaciones CIdValesCabS, 1, StrMsgError, "S", StrAlSucursalOri, False
            If StrMsgError <> "" Then GoTo Err
            
            actualizaStock_Liquidaciones CIdValesCabI, 1, StrMsgError, "I", StrAlSucursalDes, False
            If StrMsgError <> "" Then GoTo Err
            
            cselect = "Delete A,B " & _
                        "From ValesCab A " & _
                        "Inner Join ValesDet B " & _
                            "On A.IdEmpresa = B.IdEmpresa And A.IdSucursal = B.IdSucursal And A.TipoVale = B.TipoVale And A.IdValesCab = B.IdValesCab " & _
                        "Where A.IdEmpresa = '" & glsEmpresa & "' And A.TipoVale = A.IdValesCab = 'S' " & _
                        "And A.IdValesCab = '" & CIdValesCabS & "' "
            
            Cn.Execute cselect
            
            cselect = "Delete A,B " & _
                        "From ValesCab A " & _
                        "Inner Join ValesDet B " & _
                        "On A.IdEmpresa = B.IdEmpresa And A.IdSucursal = B.IdSucursal And A.TipoVale = B.TipoVale And A.IdValesCab = B.IdValesCab " & _
                        "Where A.IdEmpresa = '" & glsEmpresa & "' And A.TipoVale = 'I' " & _
                        "And A.IdValesCab = '" & CIdValesCabI & "' "
                        
            Cn.Execute cselect
            
        End If
    End If
    
    stralmacenOri = traerCampo("unidadproduccion", "IdAlmacen", "CodUnidProd", Trim("" & Txt_IdUPP.Text), True)
    StrAlSucursalOri = traerCampo("Almacenes", "idSucursal", "IdAlmacen", Trim("" & stralmacenOri), True)
    
    stralmacenDes = traerCampo("unidadproduccion", "IdAlmacen", "CodUnidProd", Trim("" & Txt_IdProveedor.Text), True)
    StrAlSucursalDes = traerCampo("Almacenes", "idSucursal", "IdAlmacen", Trim("" & stralmacenDes), True)
    
    CAbreviaturaDoc = "" & traerCampo("Documentos", "AbreDocumento", "IdDocumento", "86", False)
    traerCampos "MotivosTraslados", "IdConceptoI,IdConcepto", "IdMotivoTraslado", txtCod_MotivoTraslado.Text, 2, CArr, False
        
    CIdConceptoI = Trim("" & CArr(0))
    CIdConceptoS = Trim("" & CArr(1))
    
    CIdValesCabS = IIf(Len(Trim(CIdValesCabS)) = 0, generaCorrelativoAnoMes_Vale("ValesCab", "idValesCab", "S"), CIdValesCabS)
    
    cselect = "Insert Into ValesCab(IdValesCab,TipoVale,FechaEmision,ValorTotal,IgvTotal,PrecioTotal,IdProvCliente,IdConcepto,IdAlmacen,IdMoneda," & _
              "GlsDocReferencia,TipoCambio,IdEmpresa,IdSucursal,IdPeriodoInv,FechaRegistro,IdUsuarioRegistro) " & _
              "Select '" & CIdValesCabS & "','S',A.FechaGuia,0,0,0,'','" & CIdConceptoS & "',B.IdAlmacen,'PEN'," & _
              "ConCat('" & CAbreviaturaDoc & "',A.SerieGuia,'-',A.NumGuia),0,A.IdEmpresa,'" & StrAlSucursalOri & "','" & StrperiodoinvOri & "'," & _
              "SysDate(),'" & glsUser & "' " & _
              "From DocVentasGuiasM A " & _
              "Inner Join UnidadProduccion B " & _
                 "On A.IdEmpresa = B.IdEmpresa And A.IdUPP = B.CodUnidProd " & _
              "Where A.IdEmpresa = '" & glsEmpresa & "' And A.IdSucursal = '" & glsSucursal & "' And A.IdUPP = '" & Txt_IdUPP.Text & "' " & _
              "And A.NumGuia = '" & Txt_NumGuia.Text & "'"
              
    Cn.Execute cselect

    cselect = "Insert Into ValesDet(TipoVale,IdValesCab,Item,IdProducto,GlsProducto,IdUM,Factor,Afecto,Cantidad,VVUnit,IGVUnit,PVUnit,TotalVVNeto," & _
              "TotalIGVNeto,TotalPVNeto,IdMoneda,IdEmpresa,IdSucursal,Cantidad2) " & _
              "Select 'S','" & CIdValesCabS & "',1,A.IdProducto,P.GlsProducto,P.IdUMCompra,1,0,(A.ValCantidad / 1),0,0,0,0,0,0,'PEN',A.IdEmpresa," & _
              " '" & StrAlSucursalOri & "',a.valpeso " & _
              "From DocVentasGuiasM A " & _
              "Inner Join Productos P " & _
                 "On A.IdEmpresa = P.IdEmpresa And A.IdProducto = P.IdProducto " & _
              "Where A.IdEmpresa = '" & glsEmpresa & "' And A.IdSucursal = '" & glsSucursal & "' And A.IdUPP = '" & Txt_IdUPP.Text & "' " & _
              "And A.NumGuia = '" & Txt_NumGuia.Text & "'"
    Cn.Execute cselect
               
    CIdValesCabI = IIf(Len(Trim(CIdValesCabI)) = 0, generaCorrelativoAnoMes_Vale("ValesCab", "idValesCab", "I"), CIdValesCabI)
    
    cselect = "Insert Into ValesCab(IdValesCab,TipoVale,FechaEmision,ValorTotal,IgvTotal,PrecioTotal,IdProvCliente,IdConcepto,IdAlmacen,IdMoneda," & _
              "GlsDocReferencia,TipoCambio,IdEmpresa,IdSucursal,IdPeriodoInv,FechaRegistro,IdUsuarioRegistro) " & _
              "Select '" & CIdValesCabI & "','I',A.FechaGuia,0,0,0,'','" & CIdConceptoI & "',B.IdAlmacen,'PEN'," & _
              "ConCat('" & CAbreviaturaDoc & "',A.SerieGuia,'-',A.NumGuia),0,A.IdEmpresa,'" & StrAlSucursalDes & "','" & StrperiodoinvDes & "'," & _
              "SysDate(),'" & glsUser & "' " & _
              "From DocVentasGuiasM A " & _
              "Inner Join UnidadProduccion B " & _
                 "On A.IdEmpresa = B.IdEmpresa And A.IdProveedor = B.CodUnidProd " & _
              "Where A.IdEmpresa = '" & glsEmpresa & "' And A.IdSucursal = '" & glsSucursal & "' And A.IdUPP = '" & Txt_IdUPP.Text & "' " & _
              "And A.NumGuia = '" & Txt_NumGuia.Text & "'"
    Cn.Execute cselect
    
    cselect = "Insert Into ValesDet(TipoVale,IdValesCab,Item,IdProducto,GlsProducto,IdUM,Factor,Afecto,Cantidad,VVUnit,IGVUnit,PVUnit,TotalVVNeto," & _
              "TotalIGVNeto,TotalPVNeto,IdMoneda,IdEmpresa,IdSucursal,Cantidad2) " & _
              "Select 'I','" & CIdValesCabI & "',1,A.IdProducto,P.GlsProducto,P.IdUMCompra,1,0,(A.ValCantidad / 1),0,0,0,0,0,0,'PEN',A.IdEmpresa," & _
              "'" & StrAlSucursalDes & "',a.valpeso " & _
              "From DocVentasGuiasM A " & _
              "Inner Join Productos P " & _
                 "On A.IdEmpresa = P.IdEmpresa And A.IdProducto = P.IdProducto " & _
              "Where A.IdEmpresa = '" & glsEmpresa & "' And A.IdSucursal = '" & glsSucursal & "' And A.IdUPP = '" & Txt_IdUPP.Text & "' " & _
              "And A.NumGuia = '" & Txt_NumGuia.Text & "'"
    Cn.Execute cselect
    
    cselect = "Delete from DocReferencia " & _
              "where idempresa = '" & glsEmpresa & "' and TipoDocOrigen = '99' and NumDocOrigen = '" & CIdValesCabS & "' and SerieDocOrigen = '000' " & _
              "and idSucursal = '" & StrAlSucursalOri & "' "
    Cn.Execute cselect
    
    cselect = "Delete from DocReferencia " & _
              "where idempresa = '" & glsEmpresa & "' and TipoDocOrigen = '88' and NumDocOrigen = '" & CIdValesCabI & "' and SerieDocOrigen = '000' " & _
              "and idSucursal = '" & StrAlSucursalDes & "' "
    Cn.Execute cselect
    
    cselect = "Insert Into DocReferencia(TipoDocOrigen,NumDocOrigen,SerieDocOrigen,TipoDocReferencia,NumDocReferencia,SerieDocReferencia,Item," & _
              "IdEmpresa,IdSucursal)Values" & _
              "('99','" & CIdValesCabS & "','000','86','" & Txt_NumGuia.Text & "','" & Txt_SerieGuia.Text & "',1,'" & glsEmpresa & "'," & _
              "'" & StrAlSucursalOri & "')," & _
              "('88','" & CIdValesCabI & "','000','86','" & Txt_NumGuia.Text & "','" & Txt_SerieGuia.Text & "',1,'" & glsEmpresa & "'," & _
              "'" & StrAlSucursalDes & "');"
    Cn.Execute cselect
    
    actualizaStock_Liquidaciones CIdValesCabS, 0, StrMsgError, "S", StrAlSucursalOri, False
    If StrMsgError <> "" Then GoTo Err
    
    actualizaStock_Liquidaciones CIdValesCabI, 0, StrMsgError, "I", StrAlSucursalDes, False
    If StrMsgError <> "" Then GoTo Err
    
    cselect = "Update DocVentasGuiasM " & _
              "Set IdValesCabI = '" & CIdValesCabI & "',IdValesCabS = '" & CIdValesCabS & "' " & _
              "Where IdEmpresa = '" & glsEmpresa & "' And IdSucursal = '" & glsSucursal & "' And IdUPP = '" & Txt_IdUPP.Text & "' " & _
              "And NumGuia = '" & Txt_NumGuia.Text & "'"
    Cn.Execute cselect
    
    Exit Sub
    
Err:
    If StrMsgError = "" Then StrMsgError = Err.Description
End Sub

Private Sub nuevo(StrMsgError As String)
On Error GoTo Err
Dim C   As Object

    For Each C In Me.Controls
        If (TypeOf C Is TextBox Or TypeOf C Is CATTextBox) And C.Name <> "txt_TextoBuscar" And C.Name <> "txt_Ano" Then
            If C.Alignment = 1 Then
                C.Text = 0
            Else
                If C.Name = "txtCod_MotivoTraslado" Then
                    C.Text = ""
                    C.Text = "06090006"
                Else
                    C.Text = ""
                End If
            End If
        End If
        
        If TypeOf C Is DTPicker Then
            If C.Format = 1 Then
                C.Value = getFechaSistema
            Else
                C.Value = Format(getFechaHoraSistema, "h:mm:ss")
            End If
        End If
    Next
    
    CIdUPPAux = ""
    CIdProveedorAux = ""
    CNumGuiaAux = ""
    
    IndNuevo = True
    Fra_Lista.Visible = False
    Fra_Registro.Visible = True
    Fra_Registro.Enabled = True
    
    Exit Sub
    
Err:
    If StrMsgError = "" Then StrMsgError = Err.Description
End Sub

Private Sub cbx_Mes_Click()
On Error GoTo Err
Dim StrMsgError As String
        
    Lista StrMsgError
    If StrMsgError <> "" Then GoTo Err
    
    Exit Sub
    
Err:
    If StrMsgError = "" Then StrMsgError = Err.Description
    MsgBox StrMsgError, vbInformation, App.Title
End Sub

Private Sub cmbAyudaMotivoTraslado_Click()
On Error GoTo Err
Dim StrMsgError As String

    mostrarAyuda "MOTIVOTRASLADO", txtCod_MotivoTraslado, txtGls_MotivoTraslado, "And  idDocumento = '86'"
    
    Exit Sub
    
Err:
    If StrMsgError = "" Then StrMsgError = Err.Description
    MsgBox StrMsgError, vbInformation, App.Title
End Sub

Private Sub Cmd_Chofer_Click()
On Error GoTo Err
Dim StrMsgError As String
    
    mostrarAyuda "CHOFER", Txt_IdChofer, Txt_GlsChofer
    
    Exit Sub
    
Err:
    If StrMsgError = "" Then StrMsgError = Err.Description
    MsgBox StrMsgError, vbInformation, App.Title
End Sub

Private Sub Cmd_Cliente_Click()
On Error GoTo Err
Dim StrMsgError As String
    
    mostrarAyuda "CLIENTE", Txt_IdCliente, Txt_GlsCliente
    
    Exit Sub
    
Err:
    If StrMsgError = "" Then StrMsgError = Err.Description
    MsgBox StrMsgError, vbInformation, App.Title
End Sub

Private Sub Cmd_EmpTrans_Click()
On Error GoTo Err
Dim StrMsgError As String
    
    mostrarAyuda "EMPTRANS", Txt_IdEmpTrans, Txt_GlsEmpTrans
    
    Exit Sub
    
Err:
    If StrMsgError = "" Then StrMsgError = Err.Description
    MsgBox StrMsgError, vbInformation, App.Title
End Sub

Private Sub Cmd_Producto_Click()
On Error GoTo Err
Dim StrMsgError As String
    
    mostrarAyuda "PRODUCTOS", Txt_IdProducto, Txt_GlsProducto
    
    Exit Sub
    
Err:
    If StrMsgError = "" Then StrMsgError = Err.Description
    MsgBox StrMsgError, vbInformation, App.Title
End Sub

Private Sub Cmd_Proveedor_Click()
On Error GoTo Err
Dim StrMsgError As String
    
    mostrarAyuda "UNIDADPRODUC", Txt_IdProveedor, Nothing
    
    Exit Sub
    
Err:
    If StrMsgError = "" Then StrMsgError = Err.Description
    MsgBox StrMsgError, vbInformation, App.Title
End Sub

Private Sub Cmd_Upp_Click()
On Error GoTo Err
Dim StrMsgError As String
    
    mostrarAyuda "UNIDADPRODUC", Txt_IdUPP, Txt_GlsUPP
    
    Exit Sub
    
Err:
    If StrMsgError = "" Then StrMsgError = Err.Description
    MsgBox StrMsgError, vbInformation, App.Title
End Sub

Private Sub Cmd_Vehiculo_Click()
On Error GoTo Err
Dim StrMsgError As String
    
    mostrarAyuda "VEHICULO", Txt_IdVehiculo, Txt_GlsVehiculo
    
    Exit Sub
    
Err:
    If StrMsgError = "" Then StrMsgError = Err.Description
    MsgBox StrMsgError, vbInformation, App.Title
End Sub

Private Sub Form_Load()
On Error GoTo Err
Dim StrMsgError As String

    Me.left = 0
    Me.top = 0
    
    txt_Ano.Text = Year(getFechaSistema)
    cbx_Mes.ListIndex = Month(getFechaSistema) - 1
    ConfGrid gLista, False, False, False, False
    
    Lista StrMsgError
    If StrMsgError <> "" Then GoTo Err
    habilitaBotones StrMsgError, 7
    If StrMsgError <> "" Then GoTo Err
    
    Exit Sub
    
Err:
    If StrMsgError = "" Then StrMsgError = Err.Description
    MsgBox StrMsgError, vbInformation, App.Title
End Sub

Private Sub Mostrar_Datos(StrMsgError As String, PIdUPP As String, PNumero As String)
On Error GoTo Err
Dim csql        As String
Dim RsConsulta  As New ADODB.Recordset
    
    csql = "Select Flete, MortCantidad, MortPeso,IdUPP,NumGuia,FechaGuia,GlsPartida,IdProveedor,IdEmpTrans,IdVehiculo,IdChofer,IdProducto,ValPeso,ValCantidad,ValFxN," & _
           "ValEX,HoraLlegada,HoraAtencion,HoraSalida,GlsPrecinto,IdCliente " & _
           ",HoraPreparacion, HoraLlegadaGranja, HoraPartida, GlsObservacion,idMotivoTraslado " & _
           "From DocVentasGuiasM " & _
           "Where IdEmpresa = '" & glsEmpresa & "' And IdUPP = '" & PIdUPP & "' " & _
           "And NumGuia = '" & PNumero & "'"
    
    With RsConsulta
        .Open csql, Cn, adOpenKeyset, adLockReadOnly
        If Not .EOF Then
            CIdUPPAux = "" & .Fields("IdUPP")
            CNumGuiaAux = "" & .Fields("NumGuia")
            Txt_IdUPP.Text = "" & .Fields("IdUPP")
            Txt_NumGuia.Text = "" & .Fields("NumGuia")
            Dtp_FechaGuia.Value = CVDate("" & .Fields("FechaGuia"))
            Txt_GlsPartida.Text = "" & .Fields("GlsPartida")
            CIdProveedorAux = "" & .Fields("IdProveedor")
            Txt_IdProveedor.Text = "" & .Fields("IdProveedor")
            Txt_IdEmpTrans.Text = "" & .Fields("IdEmpTrans")
            Txt_IdVehiculo.Text = "" & .Fields("IdVehiculo")
            Txt_IdChofer.Text = "" & .Fields("IdChofer")
            Txt_IdProducto.Text = "" & .Fields("IdProducto")
            Txt_ValPeso.Text = Val("" & .Fields("ValPeso"))
            Txt_ValCantidad.Text = Val("" & .Fields("ValCantidad"))
            Txt_ValFxN.Text = Val("" & .Fields("ValFxN"))
            Txt_ValEX.Text = Val("" & .Fields("ValEX"))
            Dtp_HoraLlegada.Value = "" & .Fields("HoraLlegada")
            Dtp_HoraAtencion.Value = "" & .Fields("HoraAtencion")
            Dtp_HoraSalida.Value = "" & .Fields("HoraSalida")
            Txt_GlsNPrecinto.Text = "" & .Fields("GlsPrecinto")
            Txt_IdCliente.Text = "" & .Fields("IdCliente")
            DtpPreparacion.Value = "" & .Fields("HoraPreparacion")
            DtpllegadaGranja.Value = "" & .Fields("HoraLlegadaGranja")
            DtpPartida.Value = "" & .Fields("HoraPartida")
            txtGls_Observacion.Text = "" & .Fields("GlsObservacion")
            txtCod_MotivoTraslado.Text = "" & .Fields("idMotivoTraslado")
            txt_Flete.Text = Val(Format(.Fields("Flete"), "0.00"))
            TxtMortCantidad.Text = Val(.Fields("MortCantidad"))
            TxtMortPeso.Text = Val(Format(.Fields("MortPeso"), "0.00"))
        End If
        .Close: Set RsConsulta = Nothing
    End With
    
    Exit Sub
    
Err:
    If StrMsgError = "" Then StrMsgError = Err.Description
    Exit Sub
    
End Sub

Private Sub gLista_OnDblClick()
On Error GoTo Err
Dim StrMsgError As String
    
    nuevo StrMsgError
    If StrMsgError <> "" Then GoTo Err
    txtCod_MotivoTraslado.Text = ""
    
    Mostrar_Datos StrMsgError, gLista.Columns.ColumnByFieldName("IdUPP").Value, gLista.Columns.ColumnByFieldName("NumGuia").Value
    If StrMsgError <> "" Then GoTo Err
    habilitaBotones StrMsgError, 2
    If StrMsgError <> "" Then GoTo Err
    Fra_Registro.Visible = True
    Fra_Registro.Enabled = False
    
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
            nuevo StrMsgError
            If StrMsgError <> "" Then GoTo Err
        Case 2 'Grabar
            Txt_NumGuia.Text = Format(Txt_NumGuia.Text, "00000000")
            
            Validaciones StrMsgError, False, True
            If StrMsgError <> "" Then GoTo Err
            
            Grabar StrMsgError
            If StrMsgError <> "" Then GoTo Err
            Fra_Registro.Visible = True
            Fra_Registro.Enabled = False
        Case 3 'Modificar
            Validaciones StrMsgError, True, False
            If StrMsgError <> "" Then GoTo Err
            IndNuevo = False
            Fra_Registro.Enabled = True
        Case 4 'Cancelar
            If IndNuevo Then
                Lista StrMsgError
                If StrMsgError <> "" Then GoTo Err
            Else
                Fra_Registro.Visible = True
                Fra_Registro.Enabled = False
            End If
        Case 5 'Eliminar
            Validaciones StrMsgError, True, False
            If StrMsgError <> "" Then GoTo Err
            eliminar StrMsgError
            If StrMsgError <> "" Then GoTo Err
            nuevo StrMsgError
            If StrMsgError <> "" Then GoTo Err
        Case 6 'Imprimir
            FrmrptGuiasMadres.Show
        Case 7 'Imprimir
            FrmrptGuiasMadres_Chofer.Show
        Case 8 'Lista
            Lista StrMsgError
            If StrMsgError <> "" Then GoTo Err
        Case 9 'Salir
            Unload Me
    End Select
    habilitaBotones StrMsgError, Button.Index
    If StrMsgError <> "" Then GoTo Err
    
    Exit Sub
    
Err:
    If StrMsgError = "" Then StrMsgError = Err.Description
    MsgBox StrMsgError, vbInformation, App.Title
End Sub

Private Sub Lista(StrMsgError As String)
On Error GoTo Err
Dim csql    As String
    
    If Len(Trim(txt_TextoBuscar.Text)) > 0 Then
        csql = " And (U.DescUnidad Like'%" & txt_TextoBuscar.Text & "%' Or D.SerieGuia Like'%" & txt_TextoBuscar.Text & "%' Or D.NumGuia Like'%" & txt_TextoBuscar.Text & "%' Or D.FechaGuia Like'%" & txt_TextoBuscar.Text & "%' Or P.F4Direccion Like'%" & txt_TextoBuscar.Text & "%')"
    End If
    
    csql = "Select ConCat(U.DescUnidad,D.NumGuia) As Item,U.DescUnidad,D.SerieGuia,D.NumGuia,D.FechaGuia,P.F4Direccion As Direccion,D.IdUPP " & _
           "From DocVentasGuiasM D " & _
           "Inner Join UnidadProduccion U " & _
               "On D.IdEmpresa = U.IdEmpresa And D.IdUPP = U.CodUnidProd " & _
           "Inner Join UnidadProduccion P " & _
               "On D.IdEmpresa = P.IdEmpresa And D.IdProveedor = P.CodUnidProd " & _
           "Where D.IdEmpresa = '" & glsEmpresa & "' And Year(D.FechaGuia) = '" & txt_Ano.Text & "' " & _
           "And Month(D.FechaGuia) = " & cbx_Mes.ListIndex + 1 & "" & csql
    
    With gLista
        .DefaultFields = False
        .Dataset.ADODataset.ConnectionString = strcn
        .Dataset.ADODataset.CursorLocation = clUseClient
        .Dataset.Active = False
        .Dataset.ADODataset.CommandText = csql
        .Dataset.DisableControls
        .Dataset.Active = True
        .KeyField = "item"
    End With
    Fra_Lista.Visible = True
    Fra_Registro.Visible = False
    
    Exit Sub

Err:
    If StrMsgError = "" Then StrMsgError = Err.Description
End Sub

Private Sub habilitaBotones(StrMsgError As String, indexBoton As Integer)
On Error GoTo Err
Dim indHabilitar As Boolean

    Select Case indexBoton
        Case 1, 2, 3, 5 'Nuevo, Grabar, Modificar,Eliminar
            If indexBoton = 2 Then indHabilitar = True
            Toolbar1.Buttons(1).Visible = indHabilitar 'Nuevo
            Toolbar1.Buttons(2).Visible = Not indHabilitar 'Grabar
            Toolbar1.Buttons(3).Visible = indHabilitar 'Modificar
            Toolbar1.Buttons(4).Visible = Not indHabilitar 'Cancelar
            Toolbar1.Buttons(5).Visible = indHabilitar 'Eliminar
            Toolbar1.Buttons(8).Visible = indHabilitar 'Lista
            Toolbar1.Buttons(6).Visible = False
            Toolbar1.Buttons(7).Visible = False
        Case 4 'Cancelar
            If IndNuevo Then
                Toolbar1.Buttons(1).Visible = True
                Toolbar1.Buttons(2).Visible = False
                Toolbar1.Buttons(3).Visible = False
                Toolbar1.Buttons(4).Visible = False
                Toolbar1.Buttons(5).Visible = False
                Toolbar1.Buttons(6).Visible = True
                Toolbar1.Buttons(7).Visible = True
                Toolbar1.Buttons(8).Visible = False
            Else
                indHabilitar = True
                Toolbar1.Buttons(1).Visible = indHabilitar 'Nuevo
                Toolbar1.Buttons(2).Visible = Not indHabilitar 'Grabar
                Toolbar1.Buttons(3).Visible = indHabilitar 'Modificar
                Toolbar1.Buttons(4).Visible = Not indHabilitar 'Cancelar
                Toolbar1.Buttons(5).Visible = indHabilitar 'Eliminar
                Toolbar1.Buttons(6).Visible = False 'Imprimir
                Toolbar1.Buttons(7).Visible = False 'Imprimir
                Toolbar1.Buttons(8).Visible = indHabilitar 'Lista
            End If
        Case 7 'Lista
            Toolbar1.Buttons(1).Visible = True
            Toolbar1.Buttons(2).Visible = False
            Toolbar1.Buttons(3).Visible = False
            Toolbar1.Buttons(4).Visible = False
            Toolbar1.Buttons(5).Visible = False
            Toolbar1.Buttons(6).Visible = True
            Toolbar1.Buttons(7).Visible = True 'Imprimir
            Toolbar1.Buttons(8).Visible = False
    End Select
    
    Exit Sub
    
Err:
    If StrMsgError = "" Then StrMsgError = Err.Description
End Sub

Private Sub txt_Ano_Change()
On Error GoTo Err
Dim StrMsgError As String
        
    Lista StrMsgError
    If StrMsgError <> "" Then GoTo Err
    
    Exit Sub
    
Err:
    If StrMsgError = "" Then StrMsgError = Err.Description
    MsgBox StrMsgError, vbInformation, App.Title
End Sub

Private Sub Txt_GlsNPrecinto_GotFocus()
On Error GoTo Err
Dim StrMsgError As String
    
    Foco StrMsgError, Txt_GlsNPrecinto
    If StrMsgError <> "" Then GoTo Err
    
    Exit Sub
    
Err:
    If StrMsgError = "" Then StrMsgError = Err.Description
    MsgBox StrMsgError, vbInformation, App.Title
End Sub

Private Sub Txt_GlsNPrecinto_KeyPress(KeyAscii As Integer)
On Error GoTo Err
Dim StrMsgError As String
    
    KeyAscii = ControlaKey(StrMsgError, KeyAscii, "T")
    If StrMsgError <> "" Then GoTo Err
    
    Exit Sub
    
Err:
    If StrMsgError = "" Then StrMsgError = Err.Description
    MsgBox StrMsgError, vbInformation, App.Title
End Sub

Private Sub Txt_IdChofer_Change()
On Error GoTo Err
Dim StrMsgError     As String
Dim CArray(2)       As String
    
    traerCampos "Choferes A Inner Join Personas B On A.IdChofer = B.IdPersona", "B.GlsPersona,A.NroBrevete", "A.IdChofer", Txt_IdChofer.Text, 2, CArray, True
    Txt_GlsChofer.Text = CArray(0)
    Txt_GlsLicConducir.Text = CArray(1)
    
    Exit Sub
    
Err:
    If StrMsgError = "" Then StrMsgError = Err.Description
    MsgBox StrMsgError, vbInformation, App.Title
End Sub

Private Sub Txt_IdChofer_DblClick()
On Error GoTo Err
Dim StrMsgError As String
    
    mostrarAyuda "CHOFER", Txt_IdChofer, Txt_GlsChofer
    
    Exit Sub
    
Err:
    If StrMsgError = "" Then StrMsgError = Err.Description
    MsgBox StrMsgError, vbInformation, App.Title
End Sub

Private Sub Txt_IdChofer_GotFocus()
On Error GoTo Err
Dim StrMsgError As String
    
    Foco StrMsgError, Txt_IdChofer
    If StrMsgError <> "" Then GoTo Err
    Lbl_Ayuda.Visible = True
    
    Exit Sub
    
Err:
    If StrMsgError = "" Then StrMsgError = Err.Description
    MsgBox StrMsgError, vbInformation, App.Title
End Sub

Private Sub Txt_IdChofer_KeyDown(KeyCode As Integer, Shift As Integer)
On Error GoTo Err
Dim StrMsgError As String
    
    If KeyCode = 113 Then
        mostrarAyuda "CHOFER", Txt_IdChofer, Txt_GlsChofer
    End If
    
    Exit Sub
    
Err:
    If StrMsgError = "" Then StrMsgError = Err.Description
    MsgBox StrMsgError, vbInformation, App.Title
End Sub

Private Sub Txt_IdChofer_KeyPress(KeyAscii As Integer)
On Error GoTo Err
Dim StrMsgError As String
    
    KeyAscii = ControlaKey(StrMsgError, KeyAscii, "T")
    If StrMsgError <> "" Then GoTo Err
    
    Exit Sub
    
Err:
    If StrMsgError = "" Then StrMsgError = Err.Description
    MsgBox StrMsgError, vbInformation, App.Title
End Sub

Private Sub Txt_IdChofer_LostFocus()
On Error GoTo Err
Dim StrMsgError As String
    
    Lbl_Ayuda.Visible = False
    
    Exit Sub
    
Err:
    If StrMsgError = "" Then StrMsgError = Err.Description
    MsgBox StrMsgError, vbInformation, App.Title
End Sub

Private Sub Txt_IdCliente_Change()
On Error GoTo Err
Dim StrMsgError As String
    
    Txt_GlsCliente.Text = traerCampo("Personas", "GlsPersona", "IdPersona", Txt_IdCliente.Text, False)
    
    Exit Sub
    
Err:
    If StrMsgError = "" Then StrMsgError = Err.Description
    MsgBox StrMsgError, vbInformation, App.Title
End Sub

Private Sub Txt_IdCliente_DblClick()
On Error GoTo Err
Dim StrMsgError As String
    
    mostrarAyuda "CLIENTE", Txt_IdCliente, Txt_GlsCliente
    
    Exit Sub
    
Err:
    If StrMsgError = "" Then StrMsgError = Err.Description
    MsgBox StrMsgError, vbInformation, App.Title
End Sub

Private Sub Txt_IdCliente_GotFocus()
On Error GoTo Err
Dim StrMsgError As String
    
    Foco StrMsgError, Txt_IdCliente
    If StrMsgError <> "" Then GoTo Err
    Lbl_Ayuda.Visible = True
    
    Exit Sub
    
Err:
    If StrMsgError = "" Then StrMsgError = Err.Description
    MsgBox StrMsgError, vbInformation, App.Title
End Sub

Private Sub Txt_IdCliente_KeyDown(KeyCode As Integer, Shift As Integer)
On Error GoTo Err
Dim StrMsgError As String
    
    If KeyCode = 113 Then
        mostrarAyuda "CLIENTE", Txt_IdCliente, Txt_GlsCliente
    End If
    
    Exit Sub
    
Err:
    If StrMsgError = "" Then StrMsgError = Err.Description
    MsgBox StrMsgError, vbInformation, App.Title
End Sub

Private Sub Txt_IdCliente_LostFocus()
On Error GoTo Err
Dim StrMsgError As String
    
    Lbl_Ayuda.Visible = False
    
    Exit Sub
    
Err:
    If StrMsgError = "" Then StrMsgError = Err.Description
    MsgBox StrMsgError, vbInformation, App.Title
End Sub

Private Sub Txt_IdEmpTrans_Change()
On Error GoTo Err
Dim StrMsgError As String
Dim CArray(2)   As String
    
    traerCampos "Personas", "GlsPersona,Ruc", "IdPersona", Txt_IdEmpTrans.Text, 2, CArray, False
    Txt_GlsEmpTrans.Text = CArray(0)
    Txt_RucEmpTrans.Text = CArray(1)
    
    Exit Sub
    
Err:
    If StrMsgError = "" Then StrMsgError = Err.Description
    MsgBox StrMsgError, vbInformation, App.Title
End Sub

Private Sub Txt_IdEmpTrans_DblClick()
On Error GoTo Err
Dim StrMsgError As String
    
    mostrarAyuda "EMPTRANS", Txt_IdEmpTrans, Txt_GlsEmpTrans
    
    Exit Sub
    
Err:
    If StrMsgError = "" Then StrMsgError = Err.Description
    MsgBox StrMsgError, vbInformation, App.Title
End Sub

Private Sub Txt_IdEmpTrans_GotFocus()
On Error GoTo Err
Dim StrMsgError As String
    
    Foco StrMsgError, Txt_IdEmpTrans
    If StrMsgError <> "" Then GoTo Err
    Lbl_Ayuda.Visible = True
    
    Exit Sub
    
Err:
    If StrMsgError = "" Then StrMsgError = Err.Description
    MsgBox StrMsgError, vbInformation, App.Title
End Sub

Private Sub Txt_IdEmpTrans_KeyDown(KeyCode As Integer, Shift As Integer)
On Error GoTo Err
Dim StrMsgError As String
    
    If KeyCode = 113 Then
        mostrarAyuda "EMPTRANS", Txt_IdEmpTrans, Txt_GlsEmpTrans
    End If
    
    Exit Sub
    
Err:
    If StrMsgError = "" Then StrMsgError = Err.Description
    MsgBox StrMsgError, vbInformation, App.Title
End Sub

Private Sub Txt_IdEmpTrans_KeyPress(KeyAscii As Integer)
On Error GoTo Err
Dim StrMsgError As String
    
    KeyAscii = ControlaKey(StrMsgError, KeyAscii, "T")
    If StrMsgError <> "" Then GoTo Err
    
    Exit Sub
    
Err:
    If StrMsgError = "" Then StrMsgError = Err.Description
    MsgBox StrMsgError, vbInformation, App.Title
End Sub

Private Sub Txt_IdEmpTrans_LostFocus()
On Error GoTo Err
Dim StrMsgError As String
    
    Lbl_Ayuda.Visible = False
    
    Exit Sub
    
Err:
    If StrMsgError = "" Then StrMsgError = Err.Description
    MsgBox StrMsgError, vbInformation, App.Title
End Sub

Private Sub Txt_IdProducto_Change()
On Error GoTo Err
Dim StrMsgError As String
    
    Txt_GlsProducto.Text = traerCampo("Productos", "GlsProducto", "IdProducto", Txt_IdProducto.Text, True)
    
    Exit Sub
    
Err:
    If StrMsgError = "" Then StrMsgError = Err.Description
    MsgBox StrMsgError, vbInformation, App.Title
End Sub

Private Sub Txt_IdProducto_DblClick()
On Error GoTo Err
Dim StrMsgError As String
    
    mostrarAyuda "PRODUCTOS", Txt_IdProducto, Txt_GlsProducto
    
    Exit Sub
    
Err:
    If StrMsgError = "" Then StrMsgError = Err.Description
    MsgBox StrMsgError, vbInformation, App.Title
End Sub

Private Sub Txt_IdProducto_GotFocus()
On Error GoTo Err
Dim StrMsgError As String
    
    Foco StrMsgError, Txt_IdProducto
    If StrMsgError <> "" Then GoTo Err
    Lbl_Ayuda.Visible = True
    
    Exit Sub
    
Err:
    If StrMsgError = "" Then StrMsgError = Err.Description
    MsgBox StrMsgError, vbInformation, App.Title
End Sub

Private Sub Txt_IdProducto_KeyDown(KeyCode As Integer, Shift As Integer)
On Error GoTo Err
Dim StrMsgError As String
    
    If KeyCode = 113 Then
        mostrarAyuda "PRODUCTOS", Txt_IdProducto, Txt_GlsProducto
    End If
    
    Exit Sub
    
Err:
    If StrMsgError = "" Then StrMsgError = Err.Description
    MsgBox StrMsgError, vbInformation, App.Title
End Sub

Private Sub Txt_IdProducto_KeyPress(KeyAscii As Integer)
On Error GoTo Err
Dim StrMsgError As String
    
    KeyAscii = ControlaKey(StrMsgError, KeyAscii, "T")
    If StrMsgError <> "" Then GoTo Err
    
    Exit Sub
    
Err:
    If StrMsgError = "" Then StrMsgError = Err.Description
    MsgBox StrMsgError, vbInformation, App.Title
End Sub

Private Sub Txt_IdProducto_LostFocus()
On Error GoTo Err
Dim StrMsgError As String
    
    Lbl_Ayuda.Visible = False
    
    Exit Sub
    
Err:
    If StrMsgError = "" Then StrMsgError = Err.Description
    MsgBox StrMsgError, vbInformation, App.Title
End Sub

Private Sub Txt_IdProveedor_Change()
On Error GoTo Err
Dim StrMsgError As String
    
    Txt_GlsLlegada.Text = traerCampo("UnidadProduccion", "F4Direccion", "CodUnidProd", Txt_IdProveedor.Text, True)
    
    Exit Sub
    
Err:
    If StrMsgError = "" Then StrMsgError = Err.Description
    MsgBox StrMsgError, vbInformation, App.Title
End Sub

Private Sub Txt_IdProveedor_DblClick()
On Error GoTo Err
Dim StrMsgError As String
    
    mostrarAyuda "UNIDADPRODUC", Txt_IdProveedor, Nothing
    
    Exit Sub
    
Err:
    If StrMsgError = "" Then StrMsgError = Err.Description
    MsgBox StrMsgError, vbInformation, App.Title
End Sub

Private Sub Txt_IdProveedor_GotFocus()
On Error GoTo Err
Dim StrMsgError As String
    
    Foco StrMsgError, Txt_IdProveedor
    If StrMsgError <> "" Then GoTo Err
    Lbl_Ayuda.Visible = True
    
    Exit Sub
    
Err:
    If StrMsgError = "" Then StrMsgError = Err.Description
    MsgBox StrMsgError, vbInformation, App.Title
End Sub

Private Sub Txt_IdProveedor_KeyDown(KeyCode As Integer, Shift As Integer)
On Error GoTo Err
Dim StrMsgError As String
    
    If KeyCode = 113 Then
        mostrarAyuda "UNIDADPRODUC", Txt_IdProveedor, Nothing
    End If
    
    Exit Sub
    
Err:
    If StrMsgError = "" Then StrMsgError = Err.Description
    MsgBox StrMsgError, vbInformation, App.Title
End Sub

Private Sub Txt_IdProveedor_KeyPress(KeyAscii As Integer)
On Error GoTo Err
Dim StrMsgError As String
    
    KeyAscii = ControlaKey(StrMsgError, KeyAscii, "T")
    If StrMsgError <> "" Then GoTo Err
    
    Exit Sub
    
Err:
    If StrMsgError = "" Then StrMsgError = Err.Description
    MsgBox StrMsgError, vbInformation, App.Title
End Sub

Private Sub Txt_IdProveedor_LostFocus()
On Error GoTo Err
Dim StrMsgError As String
    
    Lbl_Ayuda.Visible = False
    
    Exit Sub
    
Err:
    If StrMsgError = "" Then StrMsgError = Err.Description
    MsgBox StrMsgError, vbInformation, App.Title
End Sub

Private Sub Txt_IdUPP_Change()
On Error GoTo Err
Dim StrMsgError As String
Dim CArray(3)   As String
    
    traerCampos "UnidadProduccion", "DescUnidad,SerieGuia,F4Direccion", "CodUnidProd", Txt_IdUPP.Text, 3, CArray, True
    
    Txt_GlsUPP.Text = CArray(0)
    Txt_SerieGuia.Text = CArray(1)
    Txt_GlsPartida.Text = CArray(2)
    
    Exit Sub
    
Err:
    If StrMsgError = "" Then StrMsgError = Err.Description
    MsgBox StrMsgError, vbInformation, App.Title
End Sub

Private Sub Txt_IdUPP_DblClick()
On Error GoTo Err
Dim StrMsgError As String
    
    mostrarAyuda "UNIDADPRODUC", Txt_IdUPP, Txt_GlsUPP
    
    Exit Sub
    
Err:
    If StrMsgError = "" Then StrMsgError = Err.Description
    MsgBox StrMsgError, vbInformation, App.Title
End Sub

Private Sub Txt_IdUPP_GotFocus()
On Error GoTo Err
Dim StrMsgError As String
    
    Foco StrMsgError, Txt_IdUPP
    If StrMsgError <> "" Then GoTo Err
    Lbl_Ayuda.Visible = True
    Lbl_Ayuda.Visible = True
    
    Exit Sub
    
Err:
    If StrMsgError = "" Then StrMsgError = Err.Description
    MsgBox StrMsgError, vbInformation, App.Title
End Sub

Private Sub Foco(StrMsgError As String, PTxt As TextBox)
On Error GoTo Err
    
    PTxt.SelStart = 0: PTxt.SelLength = Len(PTxt.Text)
    
    Exit Sub
    
Err:
    If StrMsgError = "" Then StrMsgError = Err.Description
End Sub

Private Sub Txt_IdUPP_KeyDown(KeyCode As Integer, Shift As Integer)
On Error GoTo Err
Dim StrMsgError As String
    
    If KeyCode = 113 Then
        mostrarAyuda "UNIDADPRODUC", Txt_IdUPP, Txt_GlsUPP
    End If
    
    Exit Sub
    
Err:
    If StrMsgError = "" Then StrMsgError = Err.Description
    MsgBox StrMsgError, vbInformation, App.Title
End Sub

Private Function ControlaKey(StrMsgError As String, ByVal KeyAscii As Integer, PTipo As String) As Integer
On Error GoTo Err
    
    If InStr("0123456789/-", Chr(KeyAscii)) = 0 Then
        ControlaKey = 0
    Else
        ControlaKey = KeyAscii
    End If
    If KeyAscii = 8 Then ControlaKey = KeyAscii ' borrado atras
    If KeyAscii = 13 Then ControlaKey = KeyAscii 'Enter
    If PTipo = "N" Then If KeyAscii = 46 Then ControlaKey = KeyAscii ' PARA EL PUNTO
    
    Exit Function

Err:
    If StrMsgError = "" Then StrMsgError = Err.Description
End Function

Private Sub Txt_IdUPP_KeyPress(KeyAscii As Integer)
On Error GoTo Err
Dim StrMsgError As String
    
    KeyAscii = ControlaKey(StrMsgError, KeyAscii, "T")
    If StrMsgError <> "" Then GoTo Err
    
    Exit Sub
    
Err:
    If StrMsgError = "" Then StrMsgError = Err.Description
    MsgBox StrMsgError, vbInformation, App.Title
End Sub

Private Sub Txt_IdUPP_LostFocus()
Dim StrMsgError                     As String
On Error GoTo Err
    
    Lbl_Ayuda.Visible = False
    
    Exit Sub

Err:
    If StrMsgError = "" Then StrMsgError = Err.Description
    MsgBox StrMsgError, vbInformation, App.Title
End Sub

Private Sub Txt_IdVehiculo_Change()
On Error GoTo Err
Dim StrMsgError As String
Dim CArray(5)   As String
    
    traerCampos "Vehiculos A Left Join Datos B On A.IdMarcaVehi = B.IdDato And '07' = B.IdTipoDatos", "A.GlsVehiculo,B.GlsDato,A.GlsPlaca,A.GlsCodInscripcion,A.IdChofer", "IdVehiculo", Txt_IdVehiculo.Text, 5, CArray, True
    Txt_GlsVehiculo.Text = CArray(0)
    Txt_GlsMarca.Text = CArray(1)
    Txt_GlsPlaca.Text = CArray(2)
    Txt_GlsInscripcion.Text = CArray(3)
    Txt_IdChofer.Text = CArray(4)
    
    Exit Sub

Err:
    If StrMsgError = "" Then StrMsgError = Err.Description
    MsgBox StrMsgError, vbInformation, App.Title
End Sub

Private Sub Txt_IdVehiculo_DblClick()
On Error GoTo Err
Dim StrMsgError As String
    
    mostrarAyuda "VEHICULO", Txt_IdVehiculo, Txt_GlsVehiculo
    
    Exit Sub
    
Err:
    If StrMsgError = "" Then StrMsgError = Err.Description
    MsgBox StrMsgError, vbInformation, App.Title
End Sub

Private Sub Txt_IdVehiculo_GotFocus()
On Error GoTo Err
Dim StrMsgError As String
    
    Foco StrMsgError, Txt_IdVehiculo
    If StrMsgError <> "" Then GoTo Err
    Lbl_Ayuda.Visible = True
    
    Exit Sub

Err:
    If StrMsgError = "" Then StrMsgError = Err.Description
    MsgBox StrMsgError, vbInformation, App.Title
End Sub

Private Sub Txt_IdVehiculo_KeyDown(KeyCode As Integer, Shift As Integer)
On Error GoTo Err
Dim StrMsgError As String
    
    If KeyCode = 113 Then
        mostrarAyuda "VEHICULO", Txt_IdVehiculo, Txt_GlsVehiculo
    End If
    
    Exit Sub
    
Err:
    If StrMsgError = "" Then StrMsgError = Err.Description
    MsgBox StrMsgError, vbInformation, App.Title
End Sub

Private Sub Txt_IdVehiculo_KeyPress(KeyAscii As Integer)
On Error GoTo Err
Dim StrMsgError As String
    
    KeyAscii = ControlaKey(StrMsgError, KeyAscii, "T")
    If StrMsgError <> "" Then GoTo Err
    
    Exit Sub
    
Err:
    If StrMsgError = "" Then StrMsgError = Err.Description
    MsgBox StrMsgError, vbInformation, App.Title
End Sub

Private Sub Txt_IdVehiculo_LostFocus()
On Error GoTo Err
Dim StrMsgError As String
    
    Lbl_Ayuda.Visible = False
    
    Exit Sub
    
Err:
    If StrMsgError = "" Then StrMsgError = Err.Description
    MsgBox StrMsgError, vbInformation, App.Title
End Sub

Private Sub Txt_NumGuia_GotFocus()
On Error GoTo Err
Dim StrMsgError As String
    
    Foco StrMsgError, Txt_NumGuia
    If StrMsgError <> "" Then GoTo Err
    
    Exit Sub
    
Err:
    If StrMsgError = "" Then StrMsgError = Err.Description
    MsgBox StrMsgError, vbInformation, App.Title
End Sub

Private Sub Txt_NumGuia_KeyPress(KeyAscii As Integer)
On Error GoTo Err
Dim StrMsgError As String
    
    If StrMsgError <> "" Then GoTo Err
    
    Exit Sub
    
Err:
    If StrMsgError = "" Then StrMsgError = Err.Description
    MsgBox StrMsgError, vbInformation, App.Title
End Sub

Private Sub Txt_NumGuia_LostFocus()
On Error GoTo Err
Dim StrMsgError As String
    
    Txt_NumGuia.Text = Format(Txt_NumGuia.Text, "00000000")
    
    Exit Sub
    
Err:
    If StrMsgError = "" Then StrMsgError = Err.Description
    MsgBox StrMsgError, vbInformation, App.Title
End Sub

Private Sub txt_TextoBuscar_Change()
On Error GoTo Err
Dim StrMsgError As String
        
    Lista StrMsgError
    If StrMsgError <> "" Then GoTo Err
    
    Exit Sub
    
Err:
    If StrMsgError = "" Then StrMsgError = Err.Description
    MsgBox StrMsgError, vbInformation, App.Title
End Sub

Private Sub Txt_ValCantidad_GotFocus()
On Error GoTo Err
Dim StrMsgError As String
    
    Foco StrMsgError, Txt_ValCantidad
    If StrMsgError <> "" Then GoTo Err
    
    Exit Sub
    
Err:
    If StrMsgError = "" Then StrMsgError = Err.Description
    MsgBox StrMsgError, vbInformation, App.Title
End Sub

Private Sub Txt_ValCantidad_KeyPress(KeyAscii As Integer)
On Error GoTo Err
Dim StrMsgError As String
    
    KeyAscii = ControlaKey(StrMsgError, KeyAscii, "N")
    If StrMsgError <> "" Then GoTo Err
    
    Exit Sub
    
Err:
    If StrMsgError = "" Then StrMsgError = Err.Description
    MsgBox StrMsgError, vbInformation, App.Title
End Sub

Private Sub Txt_ValEX_GotFocus()
On Error GoTo Err
Dim StrMsgError As String
    
    Foco StrMsgError, Txt_ValEX
    If StrMsgError <> "" Then GoTo Err
    
    Exit Sub
    
Err:
    If StrMsgError = "" Then StrMsgError = Err.Description
    MsgBox StrMsgError, vbInformation, App.Title
End Sub

Private Sub Txt_ValEX_KeyPress(KeyAscii As Integer)
On Error GoTo Err
Dim StrMsgError As String
    
    KeyAscii = ControlaKey(StrMsgError, KeyAscii, "N")
    If StrMsgError <> "" Then GoTo Err
    
    Exit Sub
    
Err:
    If StrMsgError = "" Then StrMsgError = Err.Description
    MsgBox StrMsgError, vbInformation, App.Title
End Sub

Private Sub Txt_ValFxN_GotFocus()
On Error GoTo Err
Dim StrMsgError As String
    
    Foco StrMsgError, Txt_ValFxN
    If StrMsgError <> "" Then GoTo Err
    
    Exit Sub
    
Err:
    If StrMsgError = "" Then StrMsgError = Err.Description
    MsgBox StrMsgError, vbInformation, App.Title
End Sub

Private Sub Txt_ValFxN_KeyPress(KeyAscii As Integer)
On Error GoTo Err
Dim StrMsgError As String
    
    KeyAscii = ControlaKey(StrMsgError, KeyAscii, "N")
    If StrMsgError <> "" Then GoTo Err
    
    Exit Sub
    
Err:
    If StrMsgError = "" Then StrMsgError = Err.Description
    MsgBox StrMsgError, vbInformation, App.Title
End Sub

Private Sub Txt_ValPeso_GotFocus()
On Error GoTo Err
Dim StrMsgError As String
    
    Foco StrMsgError, Txt_ValPeso
    If StrMsgError <> "" Then GoTo Err
    
    Exit Sub
    
Err:
    If StrMsgError = "" Then StrMsgError = Err.Description
    MsgBox StrMsgError, vbInformation, App.Title
End Sub

Private Sub Txt_ValPeso_KeyPress(KeyAscii As Integer)
On Error GoTo Err
Dim StrMsgError As String
    
    KeyAscii = ControlaKey(StrMsgError, KeyAscii, "N")
    If StrMsgError <> "" Then GoTo Err
    
    Exit Sub
    
Err:
    If StrMsgError = "" Then StrMsgError = Err.Description
    MsgBox StrMsgError, vbInformation, App.Title
End Sub

Private Sub txtCod_MotivoTraslado_Change()
On Error GoTo Err
Dim StrMsgError As String
    
    txtGls_MotivoTraslado.Text = Trim("" & traerCampo("motivostraslados", "GlsMotivoTraslado", "idMotivoTraslado", txtCod_MotivoTraslado.Text, False))
    
    Exit Sub

Err:
    If StrMsgError = "" Then StrMsgError = Err.Description
    MsgBox StrMsgError, vbInformation, App.Title
End Sub

Private Sub txtCod_MotivoTraslado_KeyDown(KeyCode As Integer, Shift As Integer)
On Error GoTo Err
Dim StrMsgError As String
    
    If KeyCode = 113 Then
        mostrarAyuda "MOTIVOTRASLADO", txtCod_MotivoTraslado, txtGls_MotivoTraslado, "And  idDocumento = '86' "
    End If
    
    Exit Sub
    
Err:
    If StrMsgError = "" Then StrMsgError = Err.Description
    MsgBox StrMsgError, vbInformation, App.Title
End Sub
