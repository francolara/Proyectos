VERSION 5.00
Object = "{6A24B331-7634-11D3-A5B0-0050044A7E1A}#1.5#0"; "DXDBGrid.dll"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{F41D1D30-7878-4923-8CB3-6CCACDC9C9DE}#1.0#0"; "catcontrols.ocx"
Begin VB.Form FrmLiquidacion 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Registro de Liquidaciones"
   ClientHeight    =   9105
   ClientLeft      =   3600
   ClientTop       =   1425
   ClientWidth     =   11850
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   9105
   ScaleWidth      =   11850
   Begin MSComctlLib.ImageList imgDocVentas 
      Left            =   0
      Top             =   0
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
            Picture         =   "FrmLiquidacion.frx":0000
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmLiquidacion.frx":039A
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmLiquidacion.frx":07EC
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmLiquidacion.frx":0B86
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmLiquidacion.frx":0F20
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmLiquidacion.frx":12BA
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmLiquidacion.frx":1654
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmLiquidacion.frx":19EE
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmLiquidacion.frx":1D88
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmLiquidacion.frx":2122
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmLiquidacion.frx":24BC
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmLiquidacion.frx":317E
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmLiquidacion.frx":3518
            Key             =   ""
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmLiquidacion.frx":396A
            Key             =   ""
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmLiquidacion.frx":3D04
            Key             =   ""
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmLiquidacion.frx":4716
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Frame Frame6 
      Appearance      =   0  'Flat
      ForeColor       =   &H80000008&
      Height          =   8340
      Left            =   90
      TabIndex        =   26
      Top             =   675
      Width           =   11670
      Begin VB.Frame Frame9 
         Appearance      =   0  'Flat
         ForeColor       =   &H80000008&
         Height          =   690
         Left            =   135
         TabIndex        =   31
         Top             =   135
         Width           =   11385
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
            ItemData        =   "FrmLiquidacion.frx":4DE8
            Left            =   6885
            List            =   "FrmLiquidacion.frx":4E10
            Style           =   2  'Dropdown List
            TabIndex        =   1
            Top             =   225
            Width           =   1935
         End
         Begin CATControls.CATTextBox txt_TextoBuscar 
            Height          =   315
            Left            =   945
            TabIndex        =   0
            Top             =   225
            Width           =   5025
            _ExtentX        =   8864
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
            Container       =   "FrmLiquidacion.frx":4E79
            Estilo          =   1
            Vacio           =   -1  'True
            EnterTab        =   -1  'True
         End
         Begin CATControls.CATTextBox txt_Ano 
            Height          =   315
            Left            =   10350
            TabIndex        =   2
            Top             =   225
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
            Container       =   "FrmLiquidacion.frx":4E95
            Estilo          =   3
            Vacio           =   -1  'True
            EnterTab        =   -1  'True
         End
         Begin VB.Label Label9 
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
            Left            =   90
            TabIndex        =   34
            Top             =   270
            Width           =   735
         End
         Begin VB.Label Label8 
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
            Left            =   6435
            TabIndex        =   33
            Top             =   270
            Width           =   300
         End
         Begin VB.Label Label7 
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
            Left            =   9900
            TabIndex        =   32
            Top             =   270
            Width           =   300
         End
      End
      Begin VB.Frame Frame8 
         Appearance      =   0  'Flat
         ForeColor       =   &H80000008&
         Height          =   3480
         Left            =   135
         TabIndex        =   28
         Top             =   4680
         Width           =   11400
         Begin DXDBGRIDLibCtl.dxDBGrid gListaDetalle 
            Height          =   3150
            Left            =   135
            OleObjectBlob   =   "FrmLiquidacion.frx":4EB1
            TabIndex        =   4
            Top             =   180
            Width           =   11160
         End
      End
      Begin VB.Frame Frame7 
         Appearance      =   0  'Flat
         ForeColor       =   &H80000008&
         Height          =   3795
         Left            =   135
         TabIndex        =   27
         Top             =   855
         Width           =   11400
         Begin DXDBGRIDLibCtl.dxDBGrid gLista 
            Height          =   3465
            Left            =   135
            OleObjectBlob   =   "FrmLiquidacion.frx":7CFC
            TabIndex        =   3
            Top             =   180
            Width           =   11160
         End
      End
   End
   Begin VB.Frame Frame1 
      Appearance      =   0  'Flat
      ForeColor       =   &H80000008&
      Height          =   8340
      Left            =   90
      TabIndex        =   12
      Top             =   675
      Width           =   11670
      Begin VB.Frame Frame2 
         Appearance      =   0  'Flat
         ForeColor       =   &H80000008&
         Height          =   1590
         Left            =   135
         TabIndex        =   21
         Top             =   180
         Width           =   11400
         Begin VB.CommandButton CmdAyudaCamal 
            Height          =   315
            Left            =   10800
            Picture         =   "FrmLiquidacion.frx":A221
            Style           =   1  'Graphical
            TabIndex        =   29
            Top             =   720
            Width           =   390
         End
         Begin MSComCtl2.DTPicker dtp_Emision 
            Height          =   315
            Left            =   1665
            TabIndex        =   22
            Tag             =   "FFecEmision"
            Top             =   1125
            Width           =   1185
            _ExtentX        =   2090
            _ExtentY        =   556
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
            Format          =   50003969
            CurrentDate     =   38955
         End
         Begin CATControls.CATTextBox txt_NumLiq 
            Height          =   315
            Left            =   10215
            TabIndex        =   23
            Tag             =   "TIdLiquidacion"
            Top             =   225
            Width           =   990
            _ExtentX        =   1746
            _ExtentY        =   556
            BackColor       =   16777215
            Enabled         =   0   'False
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            FontBold        =   -1  'True
            FontName        =   "Arial"
            FontSize        =   9.75
            ForeColor       =   -2147483640
            MaxLength       =   8
            Container       =   "FrmLiquidacion.frx":A5AB
            Vacio           =   -1  'True
            EnterTab        =   -1  'True
         End
         Begin CATControls.CATTextBox txt_CodCamal 
            Height          =   315
            Left            =   1665
            TabIndex        =   5
            Tag             =   "TidCamal"
            Top             =   720
            Width           =   1185
            _ExtentX        =   2090
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
            Container       =   "FrmLiquidacion.frx":A5C7
            Estilo          =   1
            EnterTab        =   -1  'True
         End
         Begin CATControls.CATTextBox txtGls_Camal 
            Height          =   315
            Left            =   2880
            TabIndex        =   30
            Top             =   720
            Width           =   7905
            _ExtentX        =   13944
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
            Container       =   "FrmLiquidacion.frx":A5E3
            Vacio           =   -1  'True
         End
         Begin VB.Label lbl_upp 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            Caption         =   "Unidad Producción"
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
            Left            =   180
            TabIndex        =   35
            Top             =   765
            Width           =   1350
         End
         Begin VB.Label lbl_FechaEmision 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            Caption         =   "Fecha de Emisión"
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
            Left            =   180
            TabIndex        =   25
            Top             =   1170
            Width           =   1260
         End
         Begin VB.Label lbl_NumDoc 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            Caption         =   "Nº"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000007&
            Height          =   195
            Left            =   9855
            TabIndex        =   24
            Top             =   270
            Width           =   210
         End
      End
      Begin VB.Frame Frame3 
         Appearance      =   0  'Flat
         ForeColor       =   &H80000008&
         Height          =   2265
         Left            =   135
         TabIndex        =   14
         Top             =   5895
         Width           =   11355
         Begin VB.TextBox txt_GlsObservacion 
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
            Height          =   735
            Left            =   1260
            ScrollBars      =   2  'Vertical
            TabIndex        =   10
            Top             =   1215
            Width           =   9930
         End
         Begin CATControls.CATTextBox Txt_PesoVivo 
            Height          =   315
            Left            =   1260
            TabIndex        =   7
            Tag             =   "NValPesoVivo"
            Top             =   360
            Width           =   1665
            _ExtentX        =   2937
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
            Locked          =   -1  'True
            Container       =   "FrmLiquidacion.frx":A5FF
            Decimales       =   2
            Vacio           =   -1  'True
            EnterTab        =   -1  'True
         End
         Begin CATControls.CATTextBox txt_Peso 
            Height          =   315
            Left            =   4410
            TabIndex        =   8
            Tag             =   "NValPeso"
            Top             =   360
            Width           =   1665
            _ExtentX        =   2937
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
            Container       =   "FrmLiquidacion.frx":A61B
            Decimales       =   2
            Vacio           =   -1  'True
            EnterTab        =   -1  'True
         End
         Begin CATControls.CATTextBox txt_Rendimiento 
            Height          =   315
            Left            =   7560
            TabIndex        =   15
            Tag             =   "NValRendimiento"
            Top             =   360
            Width           =   1665
            _ExtentX        =   2937
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
            Container       =   "FrmLiquidacion.frx":A637
            Decimales       =   2
            Vacio           =   -1  'True
            EnterTab        =   -1  'True
         End
         Begin CATControls.CATTextBox Txt_Edad 
            Height          =   315
            Left            =   1260
            TabIndex        =   9
            Tag             =   "NValEdad"
            Top             =   765
            Width           =   1665
            _ExtentX        =   2937
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
            Container       =   "FrmLiquidacion.frx":A653
            Decimales       =   2
            Vacio           =   -1  'True
            EnterTab        =   -1  'True
         End
         Begin VB.Label Label1 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            Caption         =   "%"
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
            Left            =   9270
            TabIndex        =   36
            Top             =   405
            Width           =   240
         End
         Begin VB.Label Label2 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            Caption         =   "Peso Vivo"
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
            Left            =   225
            TabIndex        =   20
            Top             =   405
            Width           =   735
         End
         Begin VB.Label Label3 
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
            ForeColor       =   &H80000007&
            Height          =   210
            Left            =   3870
            TabIndex        =   19
            Top             =   405
            Width           =   360
         End
         Begin VB.Label Label4 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            Caption         =   "Rendimiento"
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
            Left            =   6615
            TabIndex        =   18
            Top             =   405
            Width           =   870
         End
         Begin VB.Label Label5 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            Caption         =   "Edad"
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
            Left            =   225
            TabIndex        =   17
            Top             =   810
            Width           =   360
         End
         Begin VB.Label Label6 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            Caption         =   "Observación"
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
            Left            =   225
            TabIndex        =   16
            Top             =   1215
            Width           =   930
         End
      End
      Begin VB.Frame Frame4 
         Appearance      =   0  'Flat
         Caption         =   " Guías "
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
         Height          =   3975
         Left            =   135
         TabIndex        =   13
         Top             =   1845
         Width           =   11400
         Begin DXDBGRIDLibCtl.dxDBGrid gdetalle 
            Height          =   3570
            Left            =   135
            OleObjectBlob   =   "FrmLiquidacion.frx":A66F
            TabIndex        =   6
            Top             =   270
            Width           =   11145
         End
      End
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Height          =   660
      Left            =   0
      TabIndex        =   11
      Top             =   0
      Width           =   11760
      _ExtentX        =   20743
      _ExtentY        =   1164
      ButtonWidth     =   2011
      ButtonHeight    =   1005
      AllowCustomize  =   0   'False
      Appearance      =   1
      ImageList       =   "imgDocVentas"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   11
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "      Nuevo     "
            Object.ToolTipText     =   "Nuevo"
            ImageIndex      =   1
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "   Grabar   "
            Object.ToolTipText     =   "Grabar"
            ImageIndex      =   4
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "   Modificar   "
            Object.ToolTipText     =   "Modificar"
            ImageIndex      =   10
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "   Eliminar   "
            Object.ToolTipText     =   "Eliminar"
            ImageIndex      =   6
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Importar"
            Object.ToolTipText     =   "Lista"
            ImageIndex      =   15
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Generar Vale"
            Object.ToolTipText     =   "Importar"
            ImageIndex      =   3
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Lista"
            ImageIndex      =   12
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Imprimir"
            ImageIndex      =   11
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Reporte 1"
            ImageIndex      =   10
         EndProperty
         BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Reporte 2"
            ImageIndex      =   10
         EndProperty
         BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Salir"
            Object.ToolTipText     =   "Salir"
            ImageIndex      =   2
         EndProperty
      EndProperty
      BorderStyle     =   1
   End
End
Attribute VB_Name = "FrmLiquidacion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim IndNuevo                        As Boolean

Private Sub cbx_Mes_Click()
On Error GoTo Err
Dim StrMsgError As String
    
    listaliquidaciones StrMsgError
    If StrMsgError <> "" Then GoTo Err
    
    Exit Sub
    
Err:
    If StrMsgError = "" Then StrMsgError = Err.Description
    MsgBox StrMsgError, vbInformation, App.Title
End Sub

Private Sub CmdAyudaCamal_Click()
    
    mostrarAyuda "UNIDADPRODUC", txt_CodCamal, txtGls_Camal
    If txt_CodCamal.Text <> "" Then SendKeys "{tab}"

End Sub

Private Sub Form_Load()
Dim StrMsgError As String
    
    Me.top = 0
    Me.left = 0
    txt_Ano.Text = Year(getFechaSistema)
    cbx_Mes.ListIndex = Month(getFechaSistema) - 1
    
    ConfGrid gLista, True, True, False, False
    ConfGrid gListaDetalle, False, False, False, False
    nuevo StrMsgError
    If StrMsgError <> "" Then GoTo Err
    Frame1.Visible = False
    Frame6.Visible = True
    
    Toolbar1.Buttons(1).Visible = True
    Toolbar1.Buttons(2).Visible = False
    Toolbar1.Buttons(3).Visible = False
    Toolbar1.Buttons(4).Visible = False
    Toolbar1.Buttons(7).Visible = False
    Toolbar1.Buttons(5).Visible = False
    Toolbar1.Buttons(6).Visible = False
    
    Exit Sub
    
Err:
    If StrMsgError = "" Then StrMsgError = Err.Description
End Sub

Private Sub gLista_OnChangeNode(ByVal OldNode As DXDBGRIDLibCtl.IdxGridNode, ByVal Node As DXDBGRIDLibCtl.IdxGridNode)
    
    listaDetalle

End Sub

Private Sub gLista_OnDblClick()
On Error GoTo Err
Dim StrMsgError     As String
Dim rscd            As New ADODB.Recordset
Dim csql            As String
    
    mostrarLiquidacion gLista.Columns.ColumnByName("idLiquidacion").Value, StrMsgError
    If StrMsgError <> "" Then GoTo Err
    
    If Len(Trim("" & traerCampo("docventasliqcab", "IdValesCabI", "IdLiquidacion", gLista.Columns.ColumnByName("idLiquidacion").Value, True, " idsucursal = '" & glsSucursal & "' "))) > 0 Then
        Toolbar1.Buttons(6).Visible = False
    Else
        Toolbar1.Buttons(6).Visible = True
    End If
    Frame1.Visible = True
    Frame6.Visible = False
    
    Toolbar1.Buttons(1).Visible = False
    Toolbar1.Buttons(2).Visible = False
    Toolbar1.Buttons(3).Visible = True
    Toolbar1.Buttons(4).Visible = True
    Toolbar1.Buttons(5).Visible = False
    Toolbar1.Buttons(7).Visible = True
    Frame1.Enabled = False
            
    Exit Sub

Err:
    MsgBox StrMsgError, vbInformation, App.Title
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
On Error GoTo Err
Dim rscd As ADODB.Recordset
Dim rsdd As ADODB.Recordset
Dim numGuia As String
Dim serieGuia As String
Dim strEDV As String
Dim StrMsgError As String

    Select Case Button.Index
        Case 1
            IndNuevo = True
            nuevo StrMsgError
            If StrMsgError <> "" Then GoTo Err
            Frame1.Visible = True
            Frame6.Visible = False
            Toolbar1.Buttons(1).Visible = False
            Toolbar1.Buttons(2).Visible = True
            Toolbar1.Buttons(3).Visible = False
            Toolbar1.Buttons(4).Visible = False
            Toolbar1.Buttons(7).Visible = True
            Toolbar1.Buttons(5).Visible = True
            Toolbar1.Buttons(6).Visible = False
            Toolbar1.Buttons(8).Visible = False
            Toolbar1.Buttons(8).Visible = False
            Frame1.Enabled = True
            
        Case 2
            Grabar StrMsgError
            If StrMsgError <> "" Then GoTo Err
            Toolbar1.Buttons(1).Visible = True
            Toolbar1.Buttons(2).Visible = False
            Toolbar1.Buttons(3).Visible = True
            Toolbar1.Buttons(4).Visible = True
            Toolbar1.Buttons(7).Visible = True
            Toolbar1.Buttons(5).Visible = False
            Toolbar1.Buttons(6).Visible = True
            Toolbar1.Buttons(8).Visible = False
            Frame1.Enabled = False
            
        Case 3
        
            If Len(Trim("" & traerCampo("docventasliqcab", "IdValesCabI", "IdLiquidacion", gLista.Columns.ColumnByName("idLiquidacion").Value, True, " idsucursal = '" & glsSucursal & "' "))) > 0 Then
                MsgBox ("EL documento no se puede Modificar porque los Vales ya han sido Generados"), vbInformation, App.Title
            Else
                IndNuevo = False
                Toolbar1.Buttons(1).Visible = False
                Toolbar1.Buttons(2).Visible = True
                Toolbar1.Buttons(3).Visible = False
                Toolbar1.Buttons(4).Visible = False
                Toolbar1.Buttons(7).Visible = True
                Toolbar1.Buttons(5).Visible = True
                Toolbar1.Buttons(6).Visible = False
                Toolbar1.Buttons(8).Visible = False
                Frame1.Enabled = True
            End If
        Case 4
            If MsgBox("¿Seguro de eliminar el registro?" & vbCrLf & "Se eliminaran todas sus dependencias.", vbQuestion + vbYesNo, App.Title) = vbYes Then
                If Len(Trim("" & traerCampo("docventasliqcab", "IdValesCabI", "IdLiquidacion", gLista.Columns.ColumnByName("idLiquidacion").Value, True, " idsucursal = '" & glsSucursal & "' "))) > 0 Then
                    MsgBox ("EL documento no se puede Eliminar porque los Vales ya han sido Generados"), vbInformation, App.Title
                Else
                    eliminar StrMsgError
                    If StrMsgError <> "" Then GoTo Err
                    nuevo StrMsgError
                    
                    If StrMsgError <> "" Then GoTo Err
                    Toolbar1.Buttons(1).Visible = False
                    Toolbar1.Buttons(2).Visible = True
                    Toolbar1.Buttons(3).Visible = False
                    Toolbar1.Buttons(4).Visible = False
                    Toolbar1.Buttons(7).Visible = True
                    Toolbar1.Buttons(5).Visible = True
                    Toolbar1.Buttons(6).Visible = False
                    Toolbar1.Buttons(8).Visible = False
                    MsgBox "Registro eliminado satisfactoriamente", vbInformation, App.Title
                    Frame1.Enabled = True
                    
                End If
            Else
                Exit Sub
            End If
                
        Case 5 'Importar
            If Len(Trim("" & txt_CodCamal.Text)) = 0 Then
                StrMsgError = "Seleccione Unidad de Produccion"
                GoTo Err
            Else
                FrmListaLiquidaciones.MostrarForm txt_CodCamal.Text, rscd, StrMsgError
                If StrMsgError <> "" Then GoTo Err
                
                mostrarDocImportado rscd, StrMsgError
                If StrMsgError <> "" Then GoTo Err
            End If
            Unload frmListaDocExportar
            
        Case 6 'Generar Vales
            If MsgBox("¿Seguro de Generar Vales?" & vbCrLf & "Se eliminaran todas sus dependencias.", vbQuestion + vbYesNo, App.Title) = vbYes Then
                Genera_Vales StrMsgError
                If StrMsgError <> "" Then GoTo Err
                MsgBox "Se han Generado los Vales satisfactoriamente", vbInformation, App.Title
            Else
                Exit Sub
            End If
        Case 7 'Salir
        
            Frame1.Visible = False
            Frame6.Visible = True
        
            Toolbar1.Buttons(1).Visible = True
            Toolbar1.Buttons(2).Visible = False
            Toolbar1.Buttons(3).Visible = False
            Toolbar1.Buttons(4).Visible = False
            Toolbar1.Buttons(7).Visible = False
            Toolbar1.Buttons(5).Visible = False
            Toolbar1.Buttons(6).Visible = False
            Toolbar1.Buttons(8).Visible = True
            listaliquidaciones StrMsgError
            If StrMsgError <> "" Then GoTo Err
            
        Case 8 'Imprimir
            mostrarReporte "RptLiquidacionVentas.rpt", "parEmpresa|parSucursal|parLiquidacion", glsEmpresa & "|" & glsSucursal & "|" & Trim("" & gLista.Columns.ColumnByFieldName("idLiquidacion").Value), "Liquidacion de Ventas", StrMsgError
            If StrMsgError <> "" Then GoTo Err
        
        Case 9 'Reporte 1
            FrmrptLiquidaciones.Show
        
        Case 10 'Reporte 2
            FrmrptCliente_repeso.Show
        Case 11 'Salir
            Unload Me
    End Select
    
    Exit Sub
    
Err:
    If StrMsgError = "" Then StrMsgError = Err.Description
    MsgBox StrMsgError, vbInformation, App.Title
End Sub

Private Sub txt_Ano_Change()
On Error GoTo Err
Dim StrMsgError As String
    
    listaliquidaciones StrMsgError
    If StrMsgError <> "" Then GoTo Err
    
    Exit Sub
    
Err:
    If StrMsgError = "" Then StrMsgError = Err.Description
    MsgBox StrMsgError, vbInformation, App.Title
End Sub

Private Sub txt_CodCamal_Change()

    If txt_CodCamal.Text <> "" Then
        txtGls_Camal.Text = traerCampo("unidadproduccion", "DescUnidad", "CodUnidProd", txt_CodCamal.Text, False)
    Else
        txtGls_Camal.Text = ""
    End If
    
End Sub

Private Sub Txt_Edad_KeyPress(KeyAscii As Integer)
On Error GoTo Err
Dim StrMsgError As String

    KeyAscii = ControlaKey(StrMsgError, KeyAscii, "N")

Err:
    If StrMsgError = "" Then StrMsgError = Err.Description
End Sub

Private Sub txt_Peso_KeyPress(KeyAscii As Integer)
On Error GoTo Err
Dim StrMsgError As String

    KeyAscii = ControlaKey(StrMsgError, KeyAscii, "N")
    txt_Rendimiento.Text = Format((Format(txt_Peso.Text, "0.00") / Format(Txt_PesoVivo.Text, "0.00")), "0.00")
    txt_Rendimiento.Text = Format((txt_Rendimiento.Text * 100), "0.00")
    
    Exit Sub
    
Err:
    If StrMsgError = "" Then StrMsgError = Err.Description
End Sub

Private Sub txt_Peso_LostFocus()
On Error GoTo Err
Dim StrMsgError As String
    
    txt_Rendimiento.Text = Format((Format(txt_Peso.Text, "0.00") / Format(Txt_PesoVivo.Text, "0.00")) * 100, "0.00")
    
    Exit Sub
    
Err:
    If StrMsgError = "" Then StrMsgError = Err.Description
End Sub

Private Sub Txt_PesoVivo_KeyPress(KeyAscii As Integer)
On Error GoTo Err
Dim StrMsgError As String
    
    KeyAscii = ControlaKey(StrMsgError, KeyAscii, "N")

Err:
    If StrMsgError = "" Then StrMsgError = Err.Description
End Sub

Private Sub txt_Rendimiento_KeyPress(KeyAscii As Integer)
On Error GoTo Err
Dim StrMsgError As String
    
    KeyAscii = ControlaKey(StrMsgError, KeyAscii, "N")

Err:
    If StrMsgError = "" Then StrMsgError = Err.Description
End Sub

Private Sub txtCod_UnidProd_Change()
    
    If txtCod_UnidProd.Text = "" Then
        txtGls_UnidProd.Text = ""
    Else
        txtGls_UnidProd.Text = traerCampo("unidadproduccion", "Descunidad", "CodUnidProd", txtCod_UnidProd.Text, True)
    End If

End Sub

Private Sub nuevo(ByRef StrMsgError As String)
Dim rsg As New ADODB.Recordset
Dim rsd As New ADODB.Recordset
Dim strAno As String

    txt_CodCamal.Text = ""
    txtGls_Camal.Text = ""
    txt_GlsObservacion.Text = ""
    txt_Peso.Text = 0#
    txt_Rendimiento.Text = 0#
    Txt_PesoVivo.Text = 0#
    Txt_Edad.Text = 0#
    txt_NumLiq.Text = ""
    
    dtp_Emision.Value = Format(getFechaSistema, "dd/mm/yyyy")
    ConfGrid GDetalle, False, True, False, False

    rsd.Fields.Append "Item", adInteger, , adFldRowID
    rsd.Fields.Append "FechaGuia", adVarChar, 30, adFldIsNullable
    rsd.Fields.Append "idUpp", adVarChar, 30, adFldIsNullable
    rsd.Fields.Append "GlsUpp", adVarChar, 30, adFldIsNullable
    rsd.Fields.Append "SerieGuia", adVarChar, 3, adFldIsNullable
    rsd.Fields.Append "NumGuia", adChar, 8, adFldIsNullable
    rsd.Fields.Append "ValCantidad", adDouble, 14, adFldIsNullable
    rsd.Fields.Append "ValPesoVivo", adDouble, 14, adFldIsNullable
    rsd.Open
    
    rsd.AddNew
    rsd.Fields("Item") = 1
    rsd.Fields("FechaGuia") = ""
    rsd.Fields("idUpp") = ""
    rsd.Fields("GlsUpp") = ""
    rsd.Fields("SerieGuia") = ""
    rsd.Fields("NumGuia") = ""
    rsd.Fields("ValCantidad") = 0#
    rsd.Fields("ValPesoVivo") = 0#
    
    mostrarDatosGridSQL GDetalle, rsd, StrMsgError
    If StrMsgError <> "" Then GoTo Err
        
    Exit Sub
    
Err:
    If StrMsgError = "" Then StrMsgError = Err.Description
End Sub

Private Sub mostrarDocImportado(ByVal rscd As ADODB.Recordset, ByRef StrMsgError As String)
On Error GoTo Err
Dim Total1 As Double
Dim Total2 As Double
Dim rsg As New ADODB.Recordset
Dim i   As Integer

    i = 0
    Total1 = 0
    Total2 = 0
    
    rsg.Fields.Append "Item", adInteger, , adFldRowID
    rsg.Fields.Append "FechaGuia", adVarChar, 30, adFldIsNullable
    rsg.Fields.Append "idUpp", adVarChar, 30, adFldIsNullable
    rsg.Fields.Append "GlsUpp", adVarChar, 30, adFldIsNullable
    rsg.Fields.Append "SerieGuia", adVarChar, 3, adFldIsNullable
    rsg.Fields.Append "NumGuia", adChar, 8, adFldIsNullable
    rsg.Fields.Append "ValCantidad", adDouble, 14, adFldIsNullable
    rsg.Fields.Append "ValPesoVivo", adDouble, 14, adFldIsNullable
    rsg.Fields.Append "ValEX", adDouble, 14, adFldIsNullable
    rsg.Open
    
    If rscd.RecordCount = 0 Then
        rsg.AddNew
        rsg.Fields("Item") = 1
        rsg.Fields("FechaGuia") = ""
        rsg.Fields("idUpp") = ""
        rsg.Fields("GlsUpp") = ""
        rsg.Fields("SerieGuia") = ""
        rsg.Fields("NumGuia") = ""
        rsg.Fields("ValCantidad") = 0#
        rsg.Fields("ValPesoVivo") = 0#
        rsg.Fields("ValEX") = 0#
    Else
        rscd.MoveFirst
        If Not rscd.EOF Then
            Do While Not rscd.EOF
                rsg.AddNew
                i = i + 1
                rsg.Fields("Item") = 1
                rsg.Fields("FechaGuia") = Trim("" & rscd.Fields("FechaGuia"))
                rsg.Fields("idUpp") = Trim("" & rscd.Fields("idUpp"))
                rsg.Fields("GlsUpp") = Trim("" & rscd.Fields("DescUnidada"))
                rsg.Fields("SerieGuia") = Trim("" & rscd.Fields("SerieGuia"))
                rsg.Fields("NumGuia") = Trim("" & rscd.Fields("NumGuia"))
                rsg.Fields("ValCantidad") = Format(rscd.Fields("ValCantidad"), "0.00")
                rsg.Fields("ValPesoVivo") = Format(rscd.Fields("ValPeso"), "0.00")
                Total1 = Total1 + (Format(rscd.Fields("ValCantidad"), "0.00") * Format(rscd.Fields("ValEX"), "0.00"))
                Total2 = Total2 + (Format(rscd.Fields("ValCantidad"), "0.00"))
                rscd.MoveNext
            Loop
        End If
    End If
    
    mostrarDatosGridSQL GDetalle, rsg, StrMsgError
    If StrMsgError <> "" Then GoTo Err
    
    If (Total1 + Total2) = 0 Then
    Else
        Txt_Edad.Text = Val(Format((Total1 / Total2), "0.00"))
    End If
    Me.Refresh
    Txt_PesoVivo.Text = Format(GDetalle.Columns.ColumnByFieldName("ValPesoVivo").SummaryFooterValue, "0.00")
    
    Exit Sub

Err:
    If StrMsgError = "" Then StrMsgError = Err.Description
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

Private Sub Grabar(StrMsgError As String)
On Error GoTo Err
Dim C As Object
Dim csql              As String
Dim strPendiente      As String
Dim strEntregado      As String
Dim indDetaAutomatico As String
Dim campo As String
Dim periodo    As String
Dim rsop       As New ADODB.Recordset
Dim cmysql     As String
    
    eliminaNulosGrilla
    
    Movimiento_Ant = 0
    If IndNuevo = True Then
        EjecutaSQLFrmRegLiquidaciones 0, StrMsgError
        strMsg = "Grabó"
        IndNuevo = False
    Else
        EjecutaSQLFrmRegLiquidaciones 1, StrMsgError
        
        strMsg = "Modificó"
    End If
    MsgBox "Se " & strMsg & " Satisfactoriamente", vbInformation, App.Title
    
    Exit Sub
    
Err:
    If StrMsgError = "" Then StrMsgError = Err.Description
End Sub

Private Sub EjecutaSQLFrmRegLiquidaciones(tipoOperacion As Integer, ByRef StrMsgError As String)
On Error GoTo Err
Dim C                       As Object
Dim csql                    As String
Dim strCampo                As String
Dim strTipoDato             As String
Dim strCampos               As String
Dim strValores              As String
Dim strValCod               As String
Dim FecChqEntregado         As String
Dim StrCod                  As String
Dim i                       As Integer
Dim indTrans                As Boolean
Dim indPrimeroCtas          As Boolean
Dim rst                     As New ADODB.Recordset
Dim cEstado                 As String
Dim X                       As Double
Dim Nro_Comp                As String
Dim Correla_ant             As String
Dim destinoRec              As String
Dim idctaContableRec        As String
Dim idMovBancosRec          As Integer
Dim Correla_pago            As String
Dim Correla_dcto            As String
Dim strNumComprobante       As String
Dim valMontoAplicado        As Double
Dim cli_pro                 As String
Dim nro_comp_ant            As String
Dim Pagos                   As Integer
Dim glsobs                  As String

    csql = ""
    For Each C In Me.Controls
        If TypeOf C Is CATTextBox Or TypeOf C Is DTPicker Or TypeOf C Is CheckBox Or TypeOf C Is TextBox Then
            If C.Tag <> "" Then
                strTipoDato = left(C.Tag, 1)
                strCampo = right(C.Tag, Len(C.Tag) - 1)
                
                If UCase(strCampo) = UCase("FecEmision") Then
                    FecChqEntregado = "'" & Format(C.Value, "yyyy-mm-dd")
                End If
                
                Select Case tipoOperacion
                    Case 0 'inserta
                        strCampos = strCampos & strCampo & ","
                        
                        If UCase(strCampo) = UCase("IdLiquidacion") Then
                            If Trim(C.Value) = "" Then
                                StrCod = generaCorrelativo("docventasliqcab", "IdLiquidacion", 8, , True)
                                C.Text = StrCod
                            End If
                            strValCod = Trim(C.Value)
                        End If
                        
                        Select Case strTipoDato
                            Case "N"
                                strValores = strValores & Val(C.Value) & ","
                            Case "T"
                                strValores = strValores & "'" & Trim(C.Text) & "',"
                            Case "F"
                                strValores = strValores & "'" & Format(C.Value, "yyyy-mm-dd") & "',"
                        End Select
                    Case 1
                        Select Case strTipoDato
                            Case "N"
                                strValores = Val(C.Value)
                            Case "T"
                                strValores = "'" & C.Value & "'"
                            Case "F"
                                strValores = "'" & Format(C.Value, "yyyy-mm-dd") & "'"
                        End Select
                        strCampos = strCampos & strCampo & "=" & strValores & ","
                End Select
            End If
        End If
    Next
    
    If Len(strCampos) > 1 Then strCampos = left(strCampos, Len(strCampos) - 1)
    
    indTrans = True
    Cn.BeginTrans

    Select Case tipoOperacion
        Case 0
            If Len(strValores) > 1 Then strValores = left(strValores, Len(strValores) - 1)
            csql = "INSERT INTO docventasliqcab(" & strCampos & ",idEmpresa,IdSucursal,FecRegistro,IdUsuarioRegistro,HoraRegistro) VALUES(" & strValores & ",'" & glsEmpresa & "','" & glsSucursal & "',sysdate(),'" & glsUser & "', '" & Time & "') "
            Cn.Execute csql
        Case 1
            csql = "UPDATE docventasliqcab SET " & strCampos & ", FecModificado = sysdate(),IdUsuarioModificado = '" & glsUser & "',HoraModificado = '" & Time & "' WHERE idEmpresa = '" & glsEmpresa & "' AND IdLiquidacion = '" & Trim("" & txt_NumLiq.Text) & "' AND idsucursal = '" & glsSucursal & "' "
            Cn.Execute csql
    End Select
    
    glsobs = txt_GlsObservacion.Text
    
    csql = "UPDATE docventasliqcab SET GlsObservacion = '" & glsobs & "' WHERE idEmpresa = '" & glsEmpresa & "' AND IdLiquidacion = '" & Trim("" & txt_NumLiq.Text) & "' AND idsucursal = '" & glsSucursal & "' "
    Cn.Execute csql
    
    glsobservacionventas2 = txt_GlsObservacion.Text
    
    If TypeName(GDetalle) <> "Nothing" Then
        
        Cn.Execute "DELETE FROM docventasliqdet WHERE idempresa = '" & glsEmpresa & "' and idSucursal = '" & glsSucursal & "' and IdLiquidacion = '" & Trim("" & txt_NumLiq.Text) & "' "
        
        GDetalle.Dataset.First
        Do While Not GDetalle.Dataset.EOF
            strCampos = ""
            strValores = ""
            For i = 0 To GDetalle.Columns.Count - 1
                If UCase(left(GDetalle.Columns(i).ObjectName, 1)) = "W" Then
                    If UCase(GDetalle.Columns(i).ObjectName) <> UCase("WTidPag_Dcto_Temp") Then
                        strTipoDato = Mid(GDetalle.Columns(i).ObjectName, 2, 1)
                        strCampo = Mid(GDetalle.Columns(i).ObjectName, 3)
                        strCampos = strCampos & strCampo & ","
                    
                        Select Case strTipoDato
                            Case "N"
                                    strValores = strValores & Val(GDetalle.Columns(i).Value) & ","
                            Case "T"
                                    strValores = strValores & "'" & Trim(GDetalle.Columns(i).Value) & "',"
                            Case "F"
                                    strValores = strValores & "'" & Format(GDetalle.Columns(i).Value, "yyyy-mm-dd") & "',"
                        End Select
                    End If
                End If
            Next
            
            If Len(strCampos) > 1 Then strCampos = left(strCampos, Len(strCampos) - 1)
            If Len(strValores) > 1 Then strValores = left(strValores, Len(strValores) - 1)
            
            csql = "INSERT INTO docventasliqdet(" & strCampos & ", IdLiquidacion, IdSucursal,idEmpresa) VALUES(" & strValores & ",'" & Trim(txt_NumLiq.Text) & "','" & glsSucursal & "','" & glsEmpresa & "' )"
            Cn.Execute csql
            
            GDetalle.Dataset.Next
        Loop
    End If
    
    GDetalle.Dataset.First
    If Not GDetalle.Dataset.EOF Then
        Do While Not GDetalle.Dataset.EOF
            csql = " update docventasguiasm set IndImportado = 1 " & _
                   " Where IdEmpresa = '" & glsEmpresa & "' " & _
                   " and IdUPP = '" & Trim("" & GDetalle.Columns.ColumnByFieldName("idUpp").Value) & "' " & _
                   " and NumGuia = '" & Trim("" & GDetalle.Columns.ColumnByFieldName("NumGuia").Value) & "' "
            Cn.Execute csql
            
            GDetalle.Dataset.Next
        Loop
    End If
    
    Cn.CommitTrans
    If rst.State = 1 Then rst.Close: Set rst = Nothing
    
    Exit Sub
    
Err:
    If rst.State = 1 Then rst.Close: Set rst = Nothing
    If StrMsgError = "" Then StrMsgError = Err.Description
    If indTrans Then Cn.RollbackTrans
End Sub

Private Sub eliminaNulosGrilla()
Dim indWhile As Boolean
Dim indEntro As Boolean
Dim i As Integer

    indWhile = True
    Do While indWhile = True
        If GDetalle.Count >= 1 Then
            GDetalle.Dataset.First
            indEntro = False
            Do While Not GDetalle.Dataset.EOF
                If (Len(Trim(GDetalle.Columns.ColumnByFieldName("NumGuia").Value)) > 0 Or Len(Trim("" & GDetalle.Columns.ColumnByFieldName("SerieGuia").Value)) > 0) Then
                    GDetalle.Dataset.Next
                Else
                    GDetalle.Dataset.Delete
                    indEntro = True
                    Exit Do
                End If
            Loop
            indWhile = indEntro
        Else
            indWhile = False
        End If
    Loop
    
    If GDetalle.Count >= 1 Then
        GDetalle.Dataset.First
        i = 0
        Do While Not GDetalle.Dataset.EOF
            i = i + 1
            GDetalle.Dataset.Edit
            GDetalle.Columns.ColumnByFieldName("Item").Value = i
            If GDetalle.Dataset.State = dsEdit Then GDetalle.Dataset.Post
            GDetalle.Dataset.Next
        Loop
    Else
        indInserta = True
        GDetalle.Dataset.Append
        indInserta = False
    End If
End Sub

Private Sub txt_TextoBuscar_Change()
On Error GoTo Err
Dim StrMsgError As String
    
    listaliquidaciones StrMsgError
    If StrMsgError <> "" Then GoTo Err
    
    Exit Sub
    
Err:
    If StrMsgError = "" Then StrMsgError = Err.Description
    MsgBox StrMsgError, vbInformation, App.Title
End Sub

Private Sub listaliquidaciones(ByRef StrMsgError As String)
On Error GoTo Err
Dim strCond         As String
Dim Vizualiza_Docum As String
Dim Orden           As String
Dim rsl             As New ADODB.Recordset
Dim rdv             As New ADODB.Recordset
    
    strCond = ""
    If Trim(txt_TextoBuscar.Text) <> "" Then
        strCond = Trim(txt_TextoBuscar.Text)
        strCond = " AND (idLiquidacion LIKE '%" & strCond & "%' or GlsObservacion LIKE '%" & strCond & "%' ) "
    End If

    csql = "SELECT (@i:=@i +1) as item,a.idLiquidacion,b.DescUnidad AS glspersona,a.FecEmision,a.GlsObservacion FROM docventasliqcab a " & _
           "inner join unidadproduccion b on a.idcamal = b.CodUnidProd and a.idempresa = b.idempresa inner join (SELECT @i:= 0) foo " & _
           "where a.idempresa = '" & glsEmpresa & "' " & _
           "and a.idsucursal = '" & glsSucursal & "' " & _
           "and year(a.Fecemision) = " & txt_Ano.Text & " " & _
           "and Month(a.Fecemision) = " & Val(cbx_Mes.ListIndex + 1) & " "
    If strCond <> "" Then csql = csql + strCond
    
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
    listaDetalle
    Me.Refresh
    
    Exit Sub
    
Err:
    If StrMsgError = "" Then StrMsgError = Err.Description
End Sub

Private Sub listaDetalle()
    
    csql = " SELECT a.item,b.DescUnidad,a.fechaguia,a.serieguia,a.numguia,a.valcantidad,a.valpesovivo " & _
           " FROM docventasliqdet a inner join unidadproduccion b " & _
           " on a.idupp = b.CodUnidProd " & _
           " and a.idempresa = b.idempresa " & _
           " Where a.idempresa = '" & glsEmpresa & "' " & _
           " and a.idsucursal = '" & glsSucursal & "' " & _
           " and a.idLiquidacion = '" & gLista.Columns.ColumnByFieldName("idLiquidacion").Value & "' "
    
    With gListaDetalle
        .DefaultFields = False
        .Dataset.ADODataset.ConnectionString = strcn
        .Dataset.ADODataset.CursorLocation = clUseClient
        .Dataset.Active = False
        .Dataset.ADODataset.CommandText = csql
        .Dataset.DisableControls
        .Dataset.Active = True
        .KeyField = "item"
    End With

End Sub

Private Sub mostrarLiquidacion(strnum As String, ByRef StrMsgError As String)
On Error GoTo Err
Dim rst                             As New ADODB.Recordset
Dim rsg                             As New ADODB.Recordset
Dim rsd                             As New ADODB.Recordset
Dim RsTabla                         As New ADODB.Recordset

    csql = "SELECT IdEmpresa, IdSucursal, IdLiquidacion, idCamal, GlsObservacion, round(ValPesoVivo,2) as ValPesoVivo, round(ValPeso,2) as ValPeso, round(ValRendimiento,2) as ValRendimiento, round(ValEdad,2) as ValEdad,FecEmision " & _
            "FROM docventasliqcab  " & _
            "WHERE idEmpresa = '" & glsEmpresa & "' AND idSucursal = '" & glsSucursal & "' AND IdLiquidacion = '" & strnum & "' "
    rst.Open csql, Cn, adOpenStatic, adLockReadOnly
    
    If Not rst.EOF Then
        txt_GlsObservacion.Text = Trim("" & rst.Fields("GlsObservacion"))
        txt_CodCamal.Text = Trim("" & rst.Fields("idCamal"))
        dtp_Emision.Value = Format(rst.Fields("FecEmision"), "YYYY-MM-DD")
        txt_NumLiq.Text = strnum
        Txt_PesoVivo.Text = Format(rst.Fields("ValPesoVivo"), "0.00")
        txt_Peso.Text = Format(rst.Fields("ValPeso"), "0.00")
        txt_Rendimiento.Text = Format(rst.Fields("ValRendimiento"), "0.00")
        Txt_Edad.Text = Format(rst.Fields("ValEdad"), "0.00")
    End If
    If rst.State = 1 Then rst.Close: Set rst = Nothing
    
    csql = "SELECT * " & _
           "FROM docventasliqdet " & _
           "WHERE idEmpresa = '" & glsEmpresa & "' AND idSucursal = '" & glsSucursal & "' AND IdLiquidacion = '" & strnum & "' " & _
           "ORDER BY ITEM "
    rst.Open csql, Cn, adOpenStatic, adLockReadOnly
    
    rsg.Fields.Append "Item", adInteger, , adFldRowID
    rsg.Fields.Append "FechaGuia", adVarChar, 30, adFldIsNullable
    rsg.Fields.Append "idUpp", adVarChar, 30, adFldIsNullable
    rsg.Fields.Append "GlsUpp", adVarChar, 150, adFldIsNullable
    rsg.Fields.Append "SerieGuia", adVarChar, 3, adFldIsNullable
    rsg.Fields.Append "NumGuia", adChar, 8, adFldIsNullable
    rsg.Fields.Append "ValCantidad", adDouble, 14, adFldIsNullable
    rsg.Fields.Append "ValPesoVivo", adDouble, 14, adFldIsNullable
    rsg.Open
    
    If rst.RecordCount = 0 Then
        rsg.AddNew
        rsg.Fields("Item") = 1
        rsg.Fields("FechaGuia") = ""
        rsg.Fields("idUpp") = ""
        rsg.Fields("GlsUpp") = ""
        rsg.Fields("SerieGuia") = ""
        rsg.Fields("NumGuia") = ""
        rsg.Fields("ValCantidad") = 0#
        rsg.Fields("ValPesoVivo") = 0#
    Else
        Do While Not rst.EOF
            rsg.AddNew
            rsg.Fields("Item") = Val(rst.Fields("Item"))
            rsg.Fields("FechaGuia") = Trim("" & rst.Fields("FechaGuia"))
            rsg.Fields("idUpp") = Trim("" & rst.Fields("idUpp"))
            rsg.Fields("GlsUpp") = Trim("" & traerCampo("unidadproduccion", "DescUnidad", "CodUnidProd", Trim("" & rst.Fields("idUpp")), True))
            rsg.Fields("SerieGuia") = Trim("" & rst.Fields("SerieGuia"))
            rsg.Fields("NumGuia") = Trim("" & rst.Fields("NumGuia"))
            rsg.Fields("ValCantidad") = Val(Format(rst.Fields("ValCantidad"), "0.00"))
            rsg.Fields("ValPesoVivo") = Val(Format(rst.Fields("ValPesoVivo"), "0.00"))
            rst.MoveNext
        Loop
    End If
    rst.Close: Set rst = Nothing
    mostrarDatosGridSQL GDetalle, rsg, StrMsgError
    If StrMsgError <> "" Then GoTo Err
    
    Exit Sub
    
Err:
    If StrMsgError = "" Then StrMsgError = Err.Description
End Sub

Private Sub eliminar(ByRef StrMsgError As String)
On Error GoTo Err
Dim indTrans        As Boolean
Dim csql            As String
        
    Cn.BeginTrans
    indTrans = True
        
    csql = "delete from docventasliqcab " & _
            " WHERE idEmpresa = '" & glsEmpresa & "' AND idSucursal = '" & glsSucursal & _
            "' AND IdLiquidacion = '" & Trim("" & txt_NumLiq.Text) & "' "
    Cn.Execute csql
        
    csql = "delete from docventasliqdet " & _
            " WHERE idEmpresa = '" & glsEmpresa & "' AND idSucursal = '" & glsSucursal & _
            "' AND IdLiquidacion = '" & Trim("" & txt_NumLiq.Text) & "' "
    Cn.Execute csql
    
    GDetalle.Dataset.First
    If Not GDetalle.Dataset.EOF Then
        Do While Not GDetalle.Dataset.EOF
            csql = " update docventasguiasm set IndImportado = 0 " & _
                  "Where IdEmpresa = '" & glsEmpresa & "' " & _
                  "and IdUPP = '" & Trim("" & GDetalle.Columns.ColumnByFieldName("idUpp").Value) & "' " & _
                  "and NumGuia = '" & Trim("" & GDetalle.Columns.ColumnByFieldName("NumGuia").Value) & "' "
            Cn.Execute csql
            
            GDetalle.Dataset.Next
        Loop
    End If
    Cn.CommitTrans
    
    Exit Sub
    
Err:
    If indTrans Then Cn.RollbackTrans
    If StrMsgError = "" Then StrMsgError = Err.Description
End Sub

Private Sub Genera_Vales(ByRef StrMsgError As String)
On Error GoTo Err
Dim cselect                         As String
Dim CIdValesCabI                    As String
Dim CIdValesCabS                    As String
Dim CIdConceptoI                    As String
Dim CIdConceptoS                    As String
Dim CAbreviaturaDoc                 As String
Dim CArr(2)                         As String
Dim StrAlSucursal                   As String
Dim strAlmacen                      As String
Dim StrperiodoinvOri                As String
    
    strAlmacen = traerCampo("unidadproduccion", "IdAlmacen", "CodUnidProd", Trim("" & txt_CodCamal.Text), True)
    StrAlSucursal = traerCampo("Almacenes", "idSucursal", "IdAlmacen", Trim("" & strAlmacen), True)
    StrperiodoinvOri = traerCampo("periodosinv", "idPeriodoInv", "idSucursal", Trim("" & StrAlSucursal), True, " estPeriodoInv = 'ACT' ")
    
    CAbreviaturaDoc = "" & traerCampo("Documentos", "AbreDocumento", "IdDocumento", "86", False)
        
    CIdConceptoI = Trim("" & CArr(0))
    CIdConceptoS = Trim("" & CArr(1))
    
    CIdValesCabS = IIf(Len(Trim(CIdValesCabS)) = 0, generaCorrelativoAnoMes_Vale("ValesCab", "idValesCab", "S"), CIdValesCabS)
    
    cselect = "Insert Into ValesCab(IdValesCab,TipoVale,FechaEmision,ValorTotal,IgvTotal,PrecioTotal,IdProvCliente,IdConcepto,IdAlmacen,IdMoneda," & _
              "GlsDocReferencia,TipoCambio,IdEmpresa,IdSucursal,IdPeriodoInv,obsValesCab,FechaRegistro,IdUsuarioRegistro) " & _
              "values('" & CIdValesCabS & "','S','" & Format(dtp_Emision.Value, "yyyy-mm-dd") & "',0,0,0,'','20','" & strAlmacen & "','PEN','',0," & _
              "'" & glsEmpresa & "','" & StrAlSucursal & "','" & StrperiodoinvOri & "','" & "Liquidacion de Ventas Nro -" & txt_NumLiq.Text & "'," & _
              "SysDate(),'" & glsUser & "')"
              
    Cn.Execute cselect
    
    cselect = "Insert Into ValesDet(TipoVale,IdValesCab,Item,IdProducto,GlsProducto,IdUM,Factor,Afecto,Cantidad,VVUnit,IGVUnit,PVUnit,TotalVVNeto," & _
              "TotalIGVNeto,TotalPVNeto,IdMoneda,IdEmpresa,IdSucursal,Cantidad2) " & _
              "select 'S','" & CIdValesCabS & "',(@i:=@i +1),IdProducto,GlsProducto,IdUMCompra,1,0,Cantidad,0,0,0,0, " & _
              "0,0,'PEN',IdEmpresa,'" & StrAlSucursal & "',Cantidad2 " & _
              "From (Select A.IdProducto,P.GlsProducto,P.IdUMCompra,0,(A.ValCantidad / 1) as Cantidad,A.IdEmpresa,a.idsucursal,a.valpeso as Cantidad2 " & _
              "From docventasliqdet d " & _
              "inner join DocVentasGuiasm A " & _
              "on d.SerieGuia = a.SerieGuia " & _
              "and d.NumGuia = a.NumGuia " & _
              "and d.idempresa = a.idempresa " & _
              "and d.idsucursal = a.idsucursal " & _
              "Inner Join Productos P On A.IdEmpresa = P.IdEmpresa " & _
              "And A.IdProducto = P.IdProducto " & _
              "Where d.IdEmpresa = '" & glsEmpresa & "' " & _
              "and d.idliquidacion = '" & txt_NumLiq.Text & "' order by A.IdProducto) x ,(SELECT @i:= 0) Correla "
    Cn.Execute cselect
               
    CIdValesCabI = IIf(Len(Trim(CIdValesCabI)) = 0, generaCorrelativoAnoMes_Vale("ValesCab", "idValesCab", "I"), CIdValesCabI)
    
    cselect = "Insert Into ValesCab(IdValesCab,TipoVale,FechaEmision,ValorTotal,IgvTotal,PrecioTotal,IdProvCliente,IdConcepto,IdAlmacen,IdMoneda," & _
              "GlsDocReferencia,TipoCambio,IdEmpresa,IdSucursal,IdPeriodoInv,obsValesCab,FechaRegistro,IdUsuarioRegistro) " & _
              " values('" & CIdValesCabI & "','I','" & Format(dtp_Emision.Value, "yyyy-mm-dd") & "',0,0,0,'','27','" & strAlmacen & "','PEN','',0," & _
              "'" & glsEmpresa & "','" & StrAlSucursal & "','" & StrperiodoinvOri & "','" & "Liquidacion de Ventas Nro -" & txt_NumLiq.Text & "'," & _
              "SysDate(),'" & glsUser & "')"
              
    Cn.Execute csql
    
    Cn.Execute cselect
    
    cselect = "Insert Into ValesDet(TipoVale,IdValesCab,Item,IdProducto,GlsProducto,IdUM,Factor,Afecto,Cantidad,VVUnit,IGVUnit,PVUnit,TotalVVNeto," & _
              "TotalIGVNeto,TotalPVNeto,IdMoneda,IdEmpresa,IdSucursal,Cantidad2) " & _
              "select 'I','" & CIdValesCabI & "',(@i:=@i +1),IdProducto,GlsProducto,IdUM,Factor,0,Cantidad,0,0,0,0, " & _
              "0,0,'PEN',IdEmpresa,'" & StrAlSucursal & "',Cantidad2 " & _
              "From (SELECT " & _
              "a.idproducto,b.glsproducto,b.idUM,b.Factor,sum(a.unidad) as Cantidad,a.idempresa,a.idsucursal, " & _
              "sum(A.kg) As cantidad2 " & _
              "FROM docventasdetliquidacion a " & _
              "inner join docventasdet b " & _
              "on a.iddocumento = b.iddocumento " & _
              "and a.idserie = b.idserie and a.iddocventas = b.iddocventas and a.idempresa = b.idempresa and a.idsucursal = b.idsucursal " & _
              "and a.idproducto = b.idproducto inner join docventas c on b.iddocumento = c.iddocumento and b.idserie = c.idserie " & _
              "and b.iddocventas = c.iddocventas and b.idempresa = c.idempresa and b.idsucursal = c.idsucursal where a.idliquidacion = '" & txt_NumLiq.Text & "' " & _
              "and a.idempresa = '" & glsEmpresa & "' and a.idsucursal = '" & glsSucursal & "' " & _
              "group by a.idproducto order by a.idproducto) x ,(SELECT @i:= 0) Correla "
    Cn.Execute cselect
    
    cselect = "Insert Into DocReferencia(TipoDocOrigen,NumDocOrigen,SerieDocOrigen,TipoDocReferencia,NumDocReferencia,SerieDocReferencia,Item," & _
              "IdEmpresa,IdSucursal)Values" & _
              "('99','" & CIdValesCabS & "','000','00','" & txt_NumLiq.Text & "','000',1,'" & glsEmpresa & "'," & _
              "'" & StrAlSucursal & "')," & _
              "('88','" & CIdValesCabI & "','000','00','" & txt_NumLiq.Text & "','000',1,'" & glsEmpresa & "'," & _
              "'" & StrAlSucursal & "');"
    Cn.Execute cselect
    
    actualizaStock_Liquidaciones CIdValesCabI, 0, StrMsgError, "I", StrAlSucursal, False
    If StrMsgError <> "" Then GoTo Err
    
    actualizaStock_Liquidaciones CIdValesCabS, 0, StrMsgError, "S", StrAlSucursal, False
    If StrMsgError <> "" Then GoTo Err
    
    cselect = "Update docventasliqcab " & _
              "Set IdValesCabI = '" & CIdValesCabI & "',IdValesCabS = '" & CIdValesCabS & "' " & _
              "Where IdEmpresa = '" & glsEmpresa & "' And IdSucursal = '" & glsSucursal & "' And idliquidacion = '" & txt_NumLiq.Text & "' "
    Cn.Execute cselect
    
    Exit Sub
    
Err:
    If StrMsgError = "" Then StrMsgError = Err.Description
End Sub
