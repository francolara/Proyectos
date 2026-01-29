VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{F41D1D30-7878-4923-8CB3-6CCACDC9C9DE}#1.0#0"; "catcontrols.ocx"
Begin VB.Form frmConsCaja 
   Appearance      =   0  'Flat
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Consulta de Caja Activa"
   ClientHeight    =   9165
   ClientLeft      =   5415
   ClientTop       =   1665
   ClientWidth     =   6585
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   9165
   ScaleWidth      =   6585
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame6 
      Appearance      =   0  'Flat
      ForeColor       =   &H80000008&
      Height          =   945
      Left            =   90
      TabIndex        =   55
      Top             =   1440
      Width           =   6360
      Begin CATControls.CATTextBox txtVal_InicialSoles 
         Height          =   315
         Left            =   2070
         TabIndex        =   56
         Tag             =   "NTipoCambio"
         Top             =   495
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   556
         BackColor       =   12640511
         Enabled         =   0   'False
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Alignment       =   1
         FontName        =   "MS Sans Serif"
         FontSize        =   8.25
         ForeColor       =   -2147483640
         Container       =   "frmConsCaja.frx":0000
         Text            =   "0.00"
         Decimales       =   2
         Estilo          =   4
         EnterTab        =   -1  'True
      End
      Begin CATControls.CATTextBox txtVal_InicialDolar 
         Height          =   315
         Left            =   4140
         TabIndex        =   58
         Tag             =   "NTipoCambio"
         Top             =   495
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   556
         BackColor       =   12640511
         Enabled         =   0   'False
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Alignment       =   1
         FontName        =   "MS Sans Serif"
         FontSize        =   8.25
         ForeColor       =   -2147483640
         Container       =   "frmConsCaja.frx":001C
         Text            =   "0.00"
         Decimales       =   2
         Estilo          =   4
         EnterTab        =   -1  'True
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         Caption         =   "S/."
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   15
         Left            =   2115
         TabIndex        =   60
         Top             =   180
         Width           =   1350
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         Caption         =   "US$"
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   13
         Left            =   4230
         TabIndex        =   59
         Top             =   180
         Width           =   1350
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         Caption         =   "Monto Inicial :"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   495
         TabIndex        =   57
         Top             =   540
         Width           =   990
      End
   End
   Begin VB.CommandButton cmbRefrescar 
      Caption         =   "&Refrescar"
      Height          =   360
      Left            =   540
      Style           =   1  'Graphical
      TabIndex        =   54
      Top             =   1035
      Width           =   1750
   End
   Begin VB.CommandButton cmbImprimir 
      Caption         =   "&Imprimir Consolidado"
      Height          =   360
      Left            =   2385
      Style           =   1  'Graphical
      TabIndex        =   53
      Top             =   1035
      Width           =   1750
   End
   Begin VB.CommandButton cmbImprimirDetallado 
      Caption         =   "&Imprimir Detallado"
      Height          =   360
      Left            =   4230
      Style           =   1  'Graphical
      TabIndex        =   52
      Top             =   1035
      Width           =   1750
   End
   Begin VB.Frame Frame5 
      Appearance      =   0  'Flat
      ForeColor       =   &H80000008&
      Height          =   990
      Left            =   105
      TabIndex        =   39
      Top             =   8115
      Width           =   6315
      Begin CATControls.CATTextBox txtVal_SaldoSoles 
         Height          =   315
         Left            =   2040
         TabIndex        =   40
         Tag             =   "NTipoCambio"
         Top             =   495
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   556
         BackColor       =   16777152
         Enabled         =   0   'False
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Alignment       =   1
         FontName        =   "MS Sans Serif"
         FontSize        =   8.25
         ForeColor       =   -2147483640
         Container       =   "frmConsCaja.frx":0038
         Text            =   "0.00"
         Decimales       =   2
         Estilo          =   4
         EnterTab        =   -1  'True
      End
      Begin CATControls.CATTextBox txtVal_SaldoDolares 
         Height          =   315
         Left            =   4140
         TabIndex        =   41
         Tag             =   "NTipoCambio"
         Top             =   465
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   556
         BackColor       =   16777152
         Enabled         =   0   'False
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Alignment       =   1
         FontName        =   "MS Sans Serif"
         FontSize        =   8.25
         ForeColor       =   -2147483640
         Container       =   "frmConsCaja.frx":0054
         Text            =   "0.00"
         Decimales       =   2
         Estilo          =   4
         EnterTab        =   -1  'True
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         Caption         =   "US$"
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   12
         Left            =   4230
         TabIndex        =   51
         Top             =   180
         Width           =   1350
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         Caption         =   "S/."
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   11
         Left            =   2115
         TabIndex        =   50
         Top             =   180
         Width           =   1350
      End
      Begin VB.Label Label3 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         Caption         =   "Saldo:"
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   28
         Left            =   495
         TabIndex        =   42
         Top             =   495
         Width           =   450
      End
   End
   Begin VB.Frame Frame4 
      Appearance      =   0  'Flat
      ForeColor       =   &H80000008&
      Height          =   945
      Left            =   90
      TabIndex        =   32
      Top             =   7140
      Width           =   6315
      Begin CATControls.CATTextBox txtVal_IngresosSoles 
         Height          =   315
         Left            =   2040
         TabIndex        =   33
         Tag             =   "NTipoCambio"
         Top             =   495
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   556
         BackColor       =   16777152
         Enabled         =   0   'False
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Alignment       =   1
         FontName        =   "MS Sans Serif"
         FontSize        =   8.25
         ForeColor       =   -2147483640
         Container       =   "frmConsCaja.frx":0070
         Text            =   "0.00"
         Decimales       =   2
         Estilo          =   4
         EnterTab        =   -1  'True
      End
      Begin CATControls.CATTextBox txtVal_IngresosDolares 
         Height          =   315
         Left            =   4140
         TabIndex        =   34
         Tag             =   "NTipoCambio"
         Top             =   465
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   556
         BackColor       =   16777152
         Enabled         =   0   'False
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Alignment       =   1
         FontName        =   "MS Sans Serif"
         FontSize        =   8.25
         ForeColor       =   -2147483640
         Container       =   "frmConsCaja.frx":008C
         Text            =   "0.00"
         Decimales       =   2
         Estilo          =   4
         EnterTab        =   -1  'True
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         Caption         =   "US$"
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   9
         Left            =   4185
         TabIndex        =   49
         Top             =   180
         Width           =   1350
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         Caption         =   "S/."
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   8
         Left            =   2070
         TabIndex        =   48
         Top             =   180
         Width           =   1350
      End
      Begin VB.Label Label3 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         Caption         =   "Ingresos:"
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   24
         Left            =   495
         TabIndex        =   35
         Top             =   585
         Width           =   645
      End
   End
   Begin VB.Frame Frame3 
      Appearance      =   0  'Flat
      ForeColor       =   &H80000008&
      Height          =   945
      Left            =   90
      TabIndex        =   12
      Top             =   6180
      Width           =   6315
      Begin CATControls.CATTextBox txtVal_EgresosSoles 
         Height          =   315
         Left            =   2040
         TabIndex        =   27
         Tag             =   "NTipoCambio"
         Top             =   495
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   556
         BackColor       =   16777152
         Enabled         =   0   'False
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Alignment       =   1
         FontName        =   "MS Sans Serif"
         FontSize        =   8.25
         ForeColor       =   -2147483640
         Container       =   "frmConsCaja.frx":00A8
         Text            =   "0.00"
         Decimales       =   2
         Estilo          =   4
         EnterTab        =   -1  'True
      End
      Begin CATControls.CATTextBox txtVal_EgresosDolares 
         Height          =   315
         Left            =   4140
         TabIndex        =   28
         Tag             =   "NTipoCambio"
         Top             =   495
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   556
         BackColor       =   16777152
         Enabled         =   0   'False
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Alignment       =   1
         FontName        =   "MS Sans Serif"
         FontSize        =   8.25
         ForeColor       =   -2147483640
         Container       =   "frmConsCaja.frx":00C4
         Text            =   "0.00"
         Decimales       =   2
         Estilo          =   4
         EnterTab        =   -1  'True
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         Caption         =   "US$"
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   5
         Left            =   4230
         TabIndex        =   47
         Top             =   180
         Width           =   1350
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         Caption         =   "S/."
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   4
         Left            =   2115
         TabIndex        =   46
         Top             =   180
         Width           =   1350
      End
      Begin VB.Label Label3 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         Caption         =   "Egresos:"
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   10
         Left            =   495
         TabIndex        =   13
         Top             =   495
         Width           =   615
      End
   End
   Begin VB.Frame Frame2 
      Appearance      =   0  'Flat
      Caption         =   " Documentos Emitidos "
      ForeColor       =   &H80000008&
      Height          =   2100
      Left            =   90
      TabIndex        =   2
      Top             =   4050
      Width           =   6315
      Begin CATControls.CATTextBox txtVal_BoletasDolares 
         Height          =   315
         Left            =   4140
         TabIndex        =   21
         Tag             =   "NTipoCambio"
         Top             =   525
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   556
         BackColor       =   16777152
         Enabled         =   0   'False
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Alignment       =   1
         FontName        =   "MS Sans Serif"
         FontSize        =   8.25
         ForeColor       =   -2147483640
         Container       =   "frmConsCaja.frx":00E0
         Text            =   "0.00"
         Decimales       =   2
         Estilo          =   4
         EnterTab        =   -1  'True
      End
      Begin CATControls.CATTextBox txtVal_BoletasSoles 
         Height          =   315
         Left            =   2040
         TabIndex        =   22
         Tag             =   "NTipoCambio"
         Top             =   525
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   556
         BackColor       =   16777152
         Enabled         =   0   'False
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Alignment       =   1
         FontName        =   "MS Sans Serif"
         FontSize        =   8.25
         ForeColor       =   -2147483640
         Container       =   "frmConsCaja.frx":00FC
         Text            =   "0.00"
         Decimales       =   2
         Estilo          =   4
         EnterTab        =   -1  'True
      End
      Begin CATControls.CATTextBox txtVal_FacturasDolares 
         Height          =   315
         Left            =   4140
         TabIndex        =   23
         Tag             =   "NTipoCambio"
         Top             =   885
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   556
         BackColor       =   16777152
         Enabled         =   0   'False
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Alignment       =   1
         FontName        =   "MS Sans Serif"
         FontSize        =   8.25
         ForeColor       =   -2147483640
         Container       =   "frmConsCaja.frx":0118
         Text            =   "0.00"
         Decimales       =   2
         Estilo          =   4
         EnterTab        =   -1  'True
      End
      Begin CATControls.CATTextBox txtVal_FacturasSoles 
         Height          =   315
         Left            =   2040
         TabIndex        =   24
         Tag             =   "NTipoCambio"
         Top             =   885
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   556
         BackColor       =   16777152
         Enabled         =   0   'False
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Alignment       =   1
         FontName        =   "MS Sans Serif"
         FontSize        =   8.25
         ForeColor       =   -2147483640
         Container       =   "frmConsCaja.frx":0134
         Text            =   "0.00"
         Decimales       =   2
         Estilo          =   4
         EnterTab        =   -1  'True
      End
      Begin CATControls.CATTextBox txtVal_TotalDolares 
         Height          =   315
         Left            =   4140
         TabIndex        =   25
         Tag             =   "NTipoCambio"
         Top             =   1680
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   556
         BackColor       =   16777152
         Enabled         =   0   'False
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Alignment       =   1
         FontName        =   "MS Sans Serif"
         FontSize        =   8.25
         ForeColor       =   -2147483640
         Container       =   "frmConsCaja.frx":0150
         Text            =   "0.00"
         Decimales       =   2
         Estilo          =   4
         EnterTab        =   -1  'True
      End
      Begin CATControls.CATTextBox txtVal_TotalSoles 
         Height          =   315
         Left            =   2040
         TabIndex        =   26
         Tag             =   "NTipoCambio"
         Top             =   1680
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   556
         BackColor       =   16777152
         Enabled         =   0   'False
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Alignment       =   1
         FontName        =   "MS Sans Serif"
         FontSize        =   8.25
         ForeColor       =   -2147483640
         Container       =   "frmConsCaja.frx":016C
         Text            =   "0.00"
         Decimales       =   2
         Estilo          =   4
         EnterTab        =   -1  'True
      End
      Begin CATControls.CATTextBox txtVal_TicketDolares 
         Height          =   315
         Left            =   4140
         TabIndex        =   36
         Tag             =   "NTipoCambio"
         Top             =   1230
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   556
         BackColor       =   16777152
         Enabled         =   0   'False
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Alignment       =   1
         FontName        =   "MS Sans Serif"
         FontSize        =   8.25
         ForeColor       =   -2147483640
         Container       =   "frmConsCaja.frx":0188
         Text            =   "0.00"
         Decimales       =   2
         Estilo          =   4
         EnterTab        =   -1  'True
      End
      Begin CATControls.CATTextBox txtVal_TicketSoles 
         Height          =   315
         Left            =   2040
         TabIndex        =   37
         Tag             =   "NTipoCambio"
         Top             =   1230
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   556
         BackColor       =   16777152
         Enabled         =   0   'False
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Alignment       =   1
         FontName        =   "MS Sans Serif"
         FontSize        =   8.25
         ForeColor       =   -2147483640
         Container       =   "frmConsCaja.frx":01A4
         Text            =   "0.00"
         Decimales       =   2
         Estilo          =   4
         EnterTab        =   -1  'True
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         Caption         =   "US$"
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   3
         Left            =   4230
         TabIndex        =   45
         Top             =   180
         Width           =   1350
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         Caption         =   "S/."
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   1
         Left            =   2115
         TabIndex        =   44
         Top             =   180
         Width           =   1350
      End
      Begin VB.Label Label3 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         Caption         =   "Ticket:"
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   26
         Left            =   495
         TabIndex        =   38
         Top             =   1290
         Width           =   495
      End
      Begin VB.Line Line1 
         X1              =   1560
         X2              =   5880
         Y1              =   1605
         Y2              =   1605
      End
      Begin VB.Label Label3 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         Caption         =   "Boletas:"
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   7
         Left            =   495
         TabIndex        =   11
         Top             =   585
         Width           =   570
      End
      Begin VB.Label Label3 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         Caption         =   "Facturas:"
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   6
         Left            =   495
         TabIndex        =   10
         Top             =   945
         Width           =   660
      End
      Begin VB.Label Label3 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         Caption         =   "Total:"
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   2
         Left            =   495
         TabIndex        =   9
         Top             =   1695
         Width           =   405
      End
   End
   Begin VB.Frame Frame1 
      Appearance      =   0  'Flat
      ForeColor       =   &H80000008&
      Height          =   960
      Left            =   105
      TabIndex        =   1
      Top             =   30
      Width           =   6315
      Begin VB.OptionButton opefectivo 
         Caption         =   "Efectivo"
         Height          =   285
         Left            =   2880
         TabIndex        =   62
         Top             =   180
         Value           =   -1  'True
         Width           =   1275
      End
      Begin VB.OptionButton opcredito 
         Caption         =   "Crédito"
         Height          =   285
         Left            =   4500
         TabIndex        =   61
         Top             =   180
         Width           =   1275
      End
      Begin VB.CommandButton cmbAyudaCaja 
         Height          =   315
         Left            =   5820
         Picture         =   "frmConsCaja.frx":01C0
         Style           =   1  'Graphical
         TabIndex        =   31
         Top             =   555
         Width           =   390
      End
      Begin CATControls.CATTextBox txtCod_Caja 
         Height          =   315
         Left            =   840
         TabIndex        =   3
         Top             =   540
         Width           =   915
         _ExtentX        =   1614
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
         Container       =   "frmConsCaja.frx":054A
         Estilo          =   1
         EnterTab        =   -1  'True
      End
      Begin CATControls.CATTextBox txtGls_Caja 
         Height          =   315
         Left            =   1800
         TabIndex        =   4
         Top             =   540
         Width           =   3990
         _ExtentX        =   7038
         _ExtentY        =   556
         BackColor       =   16777152
         Enabled         =   0   'False
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
         Container       =   "frmConsCaja.frx":0566
      End
      Begin MSComCtl2.DTPicker dtpFecha 
         Height          =   315
         Left            =   840
         TabIndex        =   30
         Top             =   150
         Width           =   1365
         _ExtentX        =   2408
         _ExtentY        =   556
         _Version        =   393216
         Format          =   104005633
         CurrentDate     =   38667
      End
      Begin VB.Label Label3 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         Caption         =   "Fecha:"
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   20
         Left            =   120
         TabIndex        =   29
         Top             =   240
         Width           =   495
      End
      Begin VB.Label Label3 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         Caption         =   "Caja:"
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   0
         Left            =   120
         TabIndex        =   5
         Top             =   600
         Width           =   360
      End
   End
   Begin VB.Frame fraGeneral 
      Appearance      =   0  'Flat
      Caption         =   " Pagos "
      ForeColor       =   &H80000008&
      Height          =   1635
      Left            =   105
      TabIndex        =   0
      Top             =   2400
      Width           =   6315
      Begin CATControls.CATTextBox txtVal_EfectivoSoles 
         Height          =   315
         Left            =   2040
         TabIndex        =   15
         Tag             =   "NTipoCambio"
         Top             =   480
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   556
         BackColor       =   16777152
         Enabled         =   0   'False
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Alignment       =   1
         FontName        =   "MS Sans Serif"
         FontSize        =   8.25
         ForeColor       =   -2147483640
         Container       =   "frmConsCaja.frx":0582
         Text            =   "0.00"
         Decimales       =   2
         Estilo          =   4
         EnterTab        =   -1  'True
      End
      Begin CATControls.CATTextBox txtVal_EfectivoDolares 
         Height          =   315
         Left            =   4140
         TabIndex        =   16
         Tag             =   "NTipoCambio"
         Top             =   495
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   556
         BackColor       =   16777152
         Enabled         =   0   'False
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Alignment       =   1
         FontName        =   "MS Sans Serif"
         FontSize        =   8.25
         ForeColor       =   -2147483640
         Container       =   "frmConsCaja.frx":059E
         Text            =   "0.00"
         Decimales       =   2
         Estilo          =   4
         EnterTab        =   -1  'True
      End
      Begin CATControls.CATTextBox txtVal_TarjetaSoles 
         Height          =   315
         Left            =   2040
         TabIndex        =   17
         Tag             =   "NTipoCambio"
         Top             =   855
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   556
         BackColor       =   16777152
         Enabled         =   0   'False
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Alignment       =   1
         FontName        =   "MS Sans Serif"
         FontSize        =   8.25
         ForeColor       =   -2147483640
         Container       =   "frmConsCaja.frx":05BA
         Text            =   "0.00"
         Decimales       =   2
         Estilo          =   4
         EnterTab        =   -1  'True
      End
      Begin CATControls.CATTextBox txtVal_TarjetaDolares 
         Height          =   315
         Left            =   4140
         TabIndex        =   18
         Tag             =   "NTipoCambio"
         Top             =   855
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   556
         BackColor       =   16777152
         Enabled         =   0   'False
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Alignment       =   1
         FontName        =   "MS Sans Serif"
         FontSize        =   8.25
         ForeColor       =   -2147483640
         Container       =   "frmConsCaja.frx":05D6
         Text            =   "0.00"
         Decimales       =   2
         Estilo          =   4
         EnterTab        =   -1  'True
      End
      Begin CATControls.CATTextBox txtVal_CreditoSoles 
         Height          =   315
         Left            =   2040
         TabIndex        =   19
         Tag             =   "NTipoCambio"
         Top             =   1215
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   556
         BackColor       =   16777152
         Enabled         =   0   'False
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Alignment       =   1
         FontName        =   "MS Sans Serif"
         FontSize        =   8.25
         ForeColor       =   -2147483640
         Container       =   "frmConsCaja.frx":05F2
         Text            =   "0.00"
         Decimales       =   2
         Estilo          =   4
         EnterTab        =   -1  'True
      End
      Begin CATControls.CATTextBox txtVal_CreditoDolares 
         Height          =   315
         Left            =   4140
         TabIndex        =   20
         Tag             =   "NTipoCambio"
         Top             =   1215
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   556
         BackColor       =   16777152
         Enabled         =   0   'False
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Alignment       =   1
         FontName        =   "MS Sans Serif"
         FontSize        =   8.25
         ForeColor       =   -2147483640
         Container       =   "frmConsCaja.frx":060E
         Text            =   "0.00"
         Decimales       =   2
         Estilo          =   4
         EnterTab        =   -1  'True
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         Caption         =   "US$"
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   17
         Left            =   4185
         TabIndex        =   43
         Top             =   195
         Width           =   1350
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         Caption         =   "S/."
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   14
         Left            =   2070
         TabIndex        =   14
         Top             =   195
         Width           =   1350
      End
      Begin VB.Label Label12 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         Caption         =   "Efectivo:"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   495
         TabIndex        =   8
         Top             =   540
         Width           =   630
      End
      Begin VB.Label Label10 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         Caption         =   "Tarjeta de Credito:"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   495
         TabIndex        =   7
         Top             =   900
         Width           =   1305
      End
      Begin VB.Label Label8 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         Caption         =   "Credito:"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   495
         TabIndex        =   6
         Top             =   1260
         Width           =   540
      End
   End
End
Attribute VB_Name = "frmConsCaja"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private strCodMovCaja As String

Private Sub cmbAyudaCaja_Click()
    
    mostrarAyuda "CAJASUSUARIO", txtCod_Caja, txtGls_Caja
    
End Sub

Private Sub cmbImprimir_Click()
On Error GoTo Err
Dim StrMsgError As String
Dim strNomCampos As String
Dim strValCampos As String
Dim ntc         As Double

    strNomCampos = "parEmpresa|parSucursal|parFecha|parCaja|parInicialSoles|parInicialDolar|parEfectivoSoles|parEfectivoDolares|parTarjetaSoles|parTarjetaDolares|parCreditoSoles|parCreditoDolares|parBoletasSoles|parBoletasDolares|parFacturasSoles|parFacturasDolares|parEgresosSoles|parEgresosDolares|parTicketSoles|parTicketDolares|parIngresosSoles|parIngresosDolares"
    
    ntc = 3
    strValCampos = glsEmpresa & "|" & glsSucursal & "|" & Format(DtpFecha.Value, "yyyy-mm-dd") & "|" & txtGls_Caja.Text & "|" & txtVal_InicialSoles.Value & "|" & txtVal_InicialDolar.Value & "|" & txtVal_EfectivoSoles.Value & "|" & txtVal_EfectivoDolares.Value & "|" & txtVal_TarjetaSoles.Value & "|" & txtVal_TarjetaDolares.Value & "|" & txtVal_CreditoSoles.Value & "|" & txtVal_CreditoDolares.Value & "|" & txtVal_BoletasSoles.Value & "|" & txtVal_BoletasDolares.Value & "|" & txtVal_FacturasSoles.Value & "|" & txtVal_FacturasDolares.Value & "|" & txtVal_EgresosSoles.Value & "|" & txtVal_EgresosDolares.Value & "|" & txtVal_TicketSoles.Value & "|" & txtVal_TicketDolares.Value & "|" & txtVal_IngresosSoles.Value & "|" & txtVal_IngresosDolares.Value
    
    Screen.MousePointer = 11
    
    mostrarReporte "rptImprimeCaja.rpt", strNomCampos, strValCampos, "Reporte de Caja Consolidado", StrMsgError
    If StrMsgError <> "" Then GoTo Err
                    
    Exit Sub
    
Err:
    Screen.MousePointer = 0
    If StrMsgError = "" Then StrMsgError = Err.Description
    MsgBox StrMsgError, vbInformation, App.Title

'Dim rsReporte       As New ADODB.Recordset
'Dim fIni    As String
'
'Dim vistaPrevia     As New frmReportePreview
'Dim aplicacion      As New CRAXDRT.Application
'Dim reporte         As CRAXDRT.Report
'
'Dim strMsgError As String
'
'Screen.MousePointer = 11
'On Error GoTo err
'
'    fIni = Format(dtpFecha.Value, "yyyy-mm-dd")
'
'    gStrRutaRpts = App.Path + "\Reportes\"
'
'
'    Set reporte = aplicacion.OpenReport(gStrRutaRpts & "rptImprimeCaja.rpt")
'
'    DoEvents
'
'    Set rsReporte = DataProcedimiento("spu_ImprimeCaja", strMsgError, glsEmpresa, glsSucursal, fIni, txtGls_Caja.Text, txtVal_InicialSoles.Value, txtVal_InicialDolar.Value, txtVal_EfectivoSoles.Value, txtVal_EfectivoDolares.Value, txtVal_TarjetaSoles.Value, txtVal_TarjetaDolares.Value, txtVal_CreditoSoles.Value, txtVal_CreditoDolares.Value, txtVal_BoletasSoles.Value, txtVal_BoletasDolares.Value, txtVal_FacturasSoles.Value, txtVal_FacturasDolares.Value, txtVal_EgresosSoles.Value, txtVal_EgresosDolares.Value, txtVal_TicketSoles.Value, txtVal_TicketDolares.Value, txtVal_IngresosSoles.Value, txtVal_IngresosDolares.Value)
'    If strMsgError <> "" Then GoTo err
'
'
'    If Not rsReporte.EOF And Not rsReporte.BOF Then
'            reporte.Database.SetDataSource rsReporte, 3
'
'            vistaPrevia.CRViewer91.ReportSource = reporte
'            vistaPrevia.Caption = "Reporte de Ventas"
'            vistaPrevia.CRViewer91.ViewReport
'            vistaPrevia.CRViewer91.DisplayGroupTree = False
'            Screen.MousePointer = 0
'            vistaPrevia.WindowState = 2
'
'            vistaPrevia.Show
'
'    Else
'            Screen.MousePointer = 0
'            MsgBox "No existen Registros  Seleccionados", vbInformation, App.Title
'    End If
'Screen.MousePointer = 0
'    Set rsReporte = Nothing
'    Set vistaPrevia = Nothing
'    Set aplicacion = Nothing
'    Set reporte = Nothing
'    Exit Sub
'err:
'Screen.MousePointer = 0
'    If strMsgError = "" Then strMsgError = err.Description
'    Set rsReporte = Nothing
'    Set vistaPrevia = Nothing
'    Set aplicacion = Nothing
'    Set reporte = Nothing
'    MsgBox strMsgError, vbInformation, App.Title
End Sub

Private Sub cmbImprimirDetallado_Click()
On Error GoTo Err
Dim StrMsgError As String

    Screen.MousePointer = 11
    
    If Trim("" & traerCampo("parametros", "ValParametro", "GlsParametro", "FORMATO_LIQUIDACION", True)) = "2" Then
        mostrarReporte "rptLiquidacionCajaDet_formato2.rpt", "parEmpresa|parSucursal|parMovCaja", glsEmpresa & "|" & glsSucursal & "|" & strCodMovCaja, GlsForm, StrMsgError
    Else
        If opefectivo.Value = True Then
            mostrarReporte "rptLiquidacionCajaDetEfec.rpt", "parEmpresa|parSucursal|parMovCaja", glsEmpresa & "|" & glsSucursal & "|" & strCodMovCaja, "Reporte de Caja Detallado", StrMsgError
        Else
            mostrarReporte "rptLiquidacionCajaDetCred.rpt", "parEmpresa|parSucursal|parMovCaja", glsEmpresa & "|" & glsSucursal & "|" & strCodMovCaja, "Reporte de Caja Detallado", StrMsgError
        End If
    End If
    
    If StrMsgError <> "" Then GoTo Err
                    
Exit Sub
Err:
    Screen.MousePointer = 0
    If StrMsgError = "" Then StrMsgError = Err.Description
    MsgBox StrMsgError, vbInformation, App.Title

'Dim rsReporte       As New ADODB.Recordset
'Dim fIni    As String
'
'Dim vistaPrevia     As New frmReportePreview
'Dim aplicacion      As New CRAXDRT.Application
'Dim reporte         As CRAXDRT.Report
'
'Dim strMsgError As String
'
'Screen.MousePointer = 11
'On Error GoTo err
'
''    fIni = Format(dtpFecha.Value, "yyyy-mm-dd")
'
'    gStrRutaRpts = App.Path + "\Reportes\"
'
'
'    Set reporte = aplicacion.OpenReport(gStrRutaRpts & "rptLiquidacionCajaDet.rpt")
'
'    DoEvents
'
'    Set rsReporte = DataProcedimiento("spu_ListaLiquidacionCajaDet", strMsgError, glsEmpresa, glsSucursal, "xxx", strCodMovCaja)
'    If strMsgError <> "" Then GoTo err
'
'
'    If Not rsReporte.EOF And Not rsReporte.BOF Then
'            reporte.Database.SetDataSource rsReporte, 3
'
'            vistaPrevia.CRViewer91.ReportSource = reporte
'            vistaPrevia.Caption = "Reporte de Caja Detallado"
'            vistaPrevia.CRViewer91.ViewReport
'            vistaPrevia.CRViewer91.DisplayGroupTree = False
'            Screen.MousePointer = 0
'            vistaPrevia.WindowState = 2
'
'            vistaPrevia.Show
'
'    Else
'            Screen.MousePointer = 0
'            MsgBox "No existen Registros  Seleccionados", vbInformation, App.Title
'    End If
'Screen.MousePointer = 0
'    Set rsReporte = Nothing
'    Set vistaPrevia = Nothing
'    Set aplicacion = Nothing
'    Set reporte = Nothing
'    Exit Sub
'err:
'Screen.MousePointer = 0
'    If strMsgError = "" Then strMsgError = err.Description
'    Set rsReporte = Nothing
'    Set vistaPrevia = Nothing
'    Set aplicacion = Nothing
'    Set reporte = Nothing
'    MsgBox strMsgError, vbInformation, App.Title

End Sub

Private Sub cmbRefrescar_Click()
On Error GoTo Err
Dim StrMsgError As String

    mostrarValores StrMsgError
    If StrMsgError <> "" Then GoTo Err
    
    Exit Sub
    
Err:
    If StrMsgError = "" Then StrMsgError = Err.Description
    MsgBox StrMsgError, vbInformation, App.Title

End Sub

Private Sub dtpFecha_Change()
On Error GoTo Err
Dim StrMsgError As String

    mostrarCaja StrMsgError
    If StrMsgError <> "" Then GoTo Err

    Exit Sub
    
Err:
    If StrMsgError = "" Then StrMsgError = Err.Description
    MsgBox StrMsgError, vbInformation, App.Title

End Sub

Private Sub Form_Load()
On Error GoTo Err
Dim StrMsgError As String
Dim C As Object

    For Each C In Me.Controls
    
        If TypeOf C Is CATTextBox Then
            If C.Estilo = NumeroContable Then C.Decimales = glsDecimalesCaja
        End If
    
    Next
    
    Me.left = 0
    Me.top = 0
    
    strCodMovCaja = CajaAperturadaUsuario(0, StrMsgError)
    If StrMsgError <> "" Then GoTo Err
    
    DtpFecha.Value = Format(traerCampo("movcajas", "FecCaja", "idMovCaja", strCodMovCaja, True, " idSucursal = '" & glsSucursal & "'"), "dd/mm/yyyy")
    txtCod_Caja.Text = traerCampo("movcajas", "idCaja", "idMovCaja", strCodMovCaja, True, " idSucursal = '" & glsSucursal & "'")
    
    Exit Sub
    
Err:
    If StrMsgError = "" Then StrMsgError = Err.Description
    MsgBox StrMsgError, vbInformation, App.Title

End Sub

Private Sub mostrarValores(ByRef StrMsgError As String)
On Error GoTo Err
Dim C As Object

    For Each C In Me.Controls
    
        If TypeOf C Is CATTextBox Then
            If C.Estilo = NumeroContable Then
                C.Text = 0#
                C.Decimales = glsDecimalesCaja
            End If
        End If
    
    Next
    
    txtVal_InicialSoles.Text = traerCampo("movcajasdet", "ValMonto", "idMovCaja", strCodMovCaja, True, " idTipoMovCaja = '99990001' AND idMoneda = 'PEN'")
    txtVal_InicialDolar.Text = traerCampo("movcajasdet", "ValMonto", "idMovCaja", strCodMovCaja, True, " idTipoMovCaja = '99990001'  AND idMoneda = 'USD'")
    
    mostrarValoresFP StrMsgError
    If StrMsgError <> "" Then GoTo Err
    
    mostrarValoresDocumentos StrMsgError
    If StrMsgError <> "" Then GoTo Err
    
    mostrarValoresEgresos StrMsgError
    If StrMsgError <> "" Then GoTo Err
    
    mostrarValoresIngresos StrMsgError
    If StrMsgError <> "" Then GoTo Err
    
    Me.Refresh
    
    '''txtVal_SaldoSoles.Text = (Val("" & txtVal_InicialSoles.Value) + Val("" & txtVal_TotalSoles.Value) + Val("" & txtVal_IngresosSoles.Value)) - Val("" & txtVal_EgresosSoles.Value)
    '''txtVal_SaldoDolares.Text = (Val("" & txtVal_InicialDolar.Value) + Val("" & txtVal_TotalDolares.Value) + Val("" & txtVal_IngresosDolares.Value)) - Val("" & txtVal_EgresosDolares.Value)
    
    txtVal_SaldoSoles.Text = (Val("" & txtVal_InicialSoles.Value) + Val("" & txtVal_EfectivoSoles.Value) + Val("" & txtVal_IngresosSoles.Value)) - Val("" & txtVal_EgresosSoles.Value)
    txtVal_SaldoDolares.Text = (Val("" & txtVal_InicialDolar.Value) + Val("" & txtVal_EfectivoDolares.Value) + Val("" & txtVal_IngresosDolares.Value)) - Val("" & txtVal_EgresosDolares.Value)
    
    Exit Sub

Err:
    If StrMsgError = "" Then StrMsgError = Err.Description

End Sub

Private Sub mostrarValoresDocumentos(ByRef StrMsgError As String)
On Error GoTo Err
Dim rst As New ADODB.Recordset

    csql = "SELECT SUM(IF(m.idMoneda = 'PEN',IF(m.idDocumento = '01',m.ValMonto,0),0)) AS totalFacturasSOL," & _
                  "SUM(IF(m.idMoneda = 'USD',IF(m.idDocumento = '01',m.ValMonto,0),0)) AS totalFacturasDOL," & _
                  "SUM(IF(m.idMoneda = 'PEN',IF(m.idDocumento = '03',m.ValMonto,0),0)) AS totalBoletasSOL," & _
                  "SUM(IF(m.idMoneda = 'USD',IF(m.idDocumento = '03',m.ValMonto,0),0)) AS totalBoletasDOL, " & _
                  "SUM(IF(m.idMoneda = 'PEN',IF(m.idDocumento = '12',m.ValMonto,0),0)) AS totalTicketSOL," & _
                  "SUM(IF(m.idMoneda = 'USD',IF(m.idDocumento = '12',m.ValMonto,0),0)) AS totalTicketDOL " & _
           "FROM movcajasdet m " & _
           "WHERE m.idEmpresa = '" & glsEmpresa & "' AND m.idSucursal = '" & glsSucursal & "' " & _
             "AND m.idTipoMovCaja = '99990002' AND m.estMovCajaDet <> 'ANU' AND m.idMovCaja = '" & strCodMovCaja & "' AND m.idDocumento IN ('01','03','12')"
    
    rst.Open csql, Cn, adOpenForwardOnly, adLockReadOnly
    
    If Not rst.EOF Then
        txtVal_FacturasSoles.Text = Val("" & rst.Fields("totalFacturasSOL"))
        txtVal_FacturasDolares.Text = Val("" & rst.Fields("totalFacturasDOL"))
        
        txtVal_BoletasSoles.Text = Val("" & rst.Fields("totalBoletasSOL"))
        txtVal_BoletasDolares.Text = Val("" & rst.Fields("totalBoletasDOL"))
        
        txtVal_TicketSoles.Text = Val("" & rst.Fields("totalTicketSOL"))
        txtVal_TicketDolares.Text = Val("" & rst.Fields("totalTicketDOL"))
    End If
    
    txtVal_TotalDolares.Text = Val(txtVal_FacturasDolares.Value) + Val(txtVal_BoletasDolares.Value) + Val(txtVal_TicketDolares.Value)
    
    txtVal_TotalSoles.Text = Val(txtVal_FacturasSoles.Value) + Val(txtVal_BoletasSoles.Value) + Val(txtVal_TicketSoles.Value)
    
    If rst.State = 1 Then rst.Close
    Set rst = Nothing
    Exit Sub
    
Err:
    If rst.State = 1 Then rst.Close
    Set rst = Nothing
    If StrMsgError = "" Then StrMsgError = Err.Description
    
End Sub

Private Sub mostrarValoresFP(ByRef StrMsgError As String)
On Error GoTo Err
Dim rst As New ADODB.Recordset

    csql = "SELECT SUM(IF(m.idMoneda = 'PEN',IF(t.TipoFormaPago = 'C',CASE m.iddocumento when '07' then m.ValMonto*-1 else m.ValMonto end,0),0)) AS totalEfectivoSOL," & _
                  "SUM(IF(m.idMoneda = 'USD',IF(t.TipoFormaPago = 'C',case  m.iddocumento when '07' then  m.ValMonto*-1 else m.ValMonto end,0),0)) AS totalEfectivoDOL," & _
                  "SUM(IF(m.idMoneda = 'PEN',IF(t.TipoFormaPago = 'T',m.ValMonto,0),0)) AS totalTarjetaSOL," & _
                  "SUM(IF(m.idMoneda = 'USD',IF(t.TipoFormaPago = 'T',m.ValMonto,0),0)) AS totalTarjetaDOL, " & _
                  "SUM(IF(m.idMoneda = 'PEN',IF(t.TipoFormaPago = 'R',CASE m.iddocumento when '07' then m.ValMonto*-1 ELSE m.ValMonto END,0),0)) AS totalCreditoSOL," & _
                  "SUM(IF(m.idMoneda = 'USD',IF(t.TipoFormaPago = 'R',CASE m.iddocumento when '07' then m.ValMonto*-1 ELSE m.ValMonto END,0),0)) AS totalCreditoDOL " & _
           "FROM movcajasdet m,formaspagos f,tipoformaspago t " & _
           "WHERE m.idEmpresa = '" & glsEmpresa & "' AND m.idSucursal = '" & glsSucursal & "' " & _
             "AND m.idFormadePago = f.idFormaPago AND f.idEmpresa = '" & glsEmpresa & "' " & _
             "AND f.idTipoFormaPago = t.idTipoFormaPago " & _
             "AND m.idTipoMovCaja = '99990002' AND m.estMovCajaDet <> 'ANU' AND m.idMovCaja = '" & strCodMovCaja & "' "
    'idFormadePago
    'idTipoFormaPago
    rst.Open csql, Cn, adOpenForwardOnly, adLockReadOnly
    
    If Not rst.EOF Then
        txtVal_EfectivoSoles.Text = Val("" & rst.Fields("totalEfectivoSOL"))
        txtVal_EfectivoDolares.Text = Val("" & rst.Fields("totalEfectivoDOL"))
        
        txtVal_TarjetaSoles.Text = Val("" & rst.Fields("totalTarjetaSOL"))
        txtVal_TarjetaDolares.Text = Val("" & rst.Fields("totalTarjetaDOL"))
        
        txtVal_CreditoSoles.Text = Val("" & rst.Fields("totalCreditoSOL"))
        txtVal_CreditoDolares.Text = Val("" & rst.Fields("totalCreditoDOL"))
    End If
    
    If rst.State = 1 Then rst.Close
    Set rst = Nothing
    Exit Sub
    
Err:
    If rst.State = 1 Then rst.Close
    Set rst = Nothing
    If StrMsgError = "" Then StrMsgError = Err.Description
    
End Sub

Private Sub mostrarValoresEgresos(ByRef StrMsgError As String)
On Error GoTo Err
Dim rst As New ADODB.Recordset

    csql = "SELECT SUM(IF(m.idMoneda = 'PEN',m.ValMonto,0)) AS totalEgresosSOL," & _
                  "SUM(IF(m.idMoneda = 'USD',m.ValMonto,0)) AS totalEgresosDOL " & _
           "FROM movcajasdet m,tiposmovcaja t " & _
           "WHERE m.idEmpresa = '" & glsEmpresa & "' AND m.idSucursal = '" & glsSucursal & "' " & _
             "AND m.idTipoMovCaja = t.idTipoMovCaja " & _
             "AND m.estMovCajaDet <> 'ANU' AND m.idMovCaja = '" & strCodMovCaja & "' AND t.indIngresoSalida = 'S' AND left(t.idTipoMovCaja,4) <> '9999'"
    
    rst.Open csql, Cn, adOpenForwardOnly, adLockReadOnly
    
    If Not rst.EOF Then
        txtVal_EgresosSoles.Text = Val("" & rst.Fields("totalEgresosSOL"))
        txtVal_EgresosDolares.Text = Val("" & rst.Fields("totalEgresosDOL"))
    End If
    
    If rst.State = 1 Then rst.Close
    Set rst = Nothing
    Exit Sub
    
Err:
    If rst.State = 1 Then rst.Close
    Set rst = Nothing
    If StrMsgError = "" Then StrMsgError = Err.Description
    
End Sub


Private Sub txtCod_Caja_Change()
On Error GoTo Err
Dim StrMsgError As String

    txtGls_Caja.Text = traerCampo("cajas", "GlsCaja", "idCaja", txtCod_Caja.Text, True)
    
    mostrarCaja StrMsgError
    If StrMsgError <> "" Then GoTo Err
    
    Exit Sub
    
Err:
    If StrMsgError = "" Then StrMsgError = Err.Description
    MsgBox StrMsgError, vbInformation, App.Title
    
End Sub

Private Sub txtCod_Caja_KeyPress(KeyAscii As Integer)

    If KeyAscii <> 13 Then
        mostrarAyudaKeyascii KeyAscii, "CAJASUSUARIO", txtCod_Caja, txtGls_Caja
        KeyAscii = 0
    End If
    
End Sub

Private Sub mostrarCaja(ByRef StrMsgError As String)
On Error GoTo Err
Dim rsTemp As New ADODB.Recordset

    strCodMovCaja = ""
    
    txtVal_InicialSoles.Text = 0
    txtVal_InicialDolar.Text = 0
    
    txtVal_EfectivoSoles.Text = 0
    txtVal_EfectivoDolares.Text = 0
    
    txtVal_TarjetaSoles.Text = 0
    txtVal_TarjetaDolares.Text = 0
    
    txtVal_CreditoSoles.Text = 0
    txtVal_CreditoDolares.Text = 0
    
    txtVal_BoletasSoles.Text = 0
    txtVal_FacturasSoles.Text = 0
    
    txtVal_BoletasDolares.Text = 0
    txtVal_FacturasDolares.Text = 0
    
    txtVal_TicketDolares.Text = 0
    txtVal_TicketDolares.Text = 0
    
    txtVal_TotalSoles.Text = 0
    txtVal_TotalDolares.Text = 0
    
    txtVal_EgresosSoles.Text = 0
    txtVal_EgresosDolares.Text = 0
    
    txtVal_IngresosSoles.Text = 0
    txtVal_IngresosDolares.Text = 0
    
    If txtCod_Caja.Text = "" Then Exit Sub
    
    csql = "SELECT m.idMovCaja " & _
            "FROM movcajas m " & _
            "WHERE m.idUsuario = '" & glsUser & "' " & _
             "AND m.idEmpresa = '" & glsEmpresa & "' " & _
             "AND m.idSucursal = '" & glsSucursal & "' " & _
             "AND m.idCaja = '" & txtCod_Caja.Text & "' " & _
             "AND DATE_FORMAT(m.FecCaja ,'%d/%m/%Y') = DATE_FORMAT('" & Format(DtpFecha.Value, "yyyy-mm-dd") & "','%d/%m/%Y')"
             
    rsTemp.Open csql, Cn, adOpenKeyset, adLockOptimistic
    If Not rsTemp.EOF Then
        strCodMovCaja = "" & rsTemp.Fields("idMovCaja")
        
        mostrarValores StrMsgError
        If StrMsgError <> "" Then GoTo Err
    Else
        StrMsgError = "No hay caja disponible para la fecha indicada"
        GoTo Err
    End If
    
    If rsTemp.State = 1 Then rsTemp.Close
    Set rsTemp = Nothing
    Exit Sub
    
Err:
    If rsTemp.State = 1 Then rsTemp.Close
    Set rsTemp = Nothing
    If StrMsgError = "" Then StrMsgError = Err.Description

End Sub

Private Sub mostrarValoresIngresos(ByRef StrMsgError As String)
On Error GoTo Err
Dim rst As New ADODB.Recordset

    csql = "SELECT SUM(IF(m.idMoneda = 'PEN',m.ValMonto,0)) AS totalIngresosSOL," & _
                  "SUM(IF(m.idMoneda = 'USD',m.ValMonto,0)) AS totalIngresosDOL " & _
           "FROM movcajasdet m,tiposmovcaja t " & _
           "WHERE m.idEmpresa = '" & glsEmpresa & "' AND m.idSucursal = '" & glsSucursal & "' " & _
             "AND m.idTipoMovCaja = t.idTipoMovCaja " & _
             "AND m.estMovCajaDet <> 'ANU' AND m.idMovCaja = '" & strCodMovCaja & "' AND t.indIngresoSalida = 'I' AND left(t.idTipoMovCaja,4) <> '9999'"
    
    rst.Open csql, Cn, adOpenForwardOnly, adLockReadOnly
    
    If Not rst.EOF Then
        txtVal_IngresosSoles.Text = Val("" & rst.Fields("totalIngresosSOL"))
        txtVal_IngresosDolares.Text = Val("" & rst.Fields("totalIngresosDOL"))
    End If
    
    If rst.State = 1 Then rst.Close
    Set rst = Nothing
    Exit Sub
    
Err:
    If rst.State = 1 Then rst.Close
    Set rst = Nothing
    If StrMsgError = "" Then StrMsgError = Err.Description

End Sub

