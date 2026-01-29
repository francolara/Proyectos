VERSION 5.00
Object = "{6A24B331-7634-11D3-A5B0-0050044A7E1A}#1.5#0"; "DXDBGrid.dll"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{F41D1D30-7878-4923-8CB3-6CCACDC9C9DE}#1.0#0"; "CATControls.ocx"
Begin VB.Form FrmDetalladoPorProducto 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Resumen Detallado por Producto"
   ClientHeight    =   8280
   ClientLeft      =   3015
   ClientTop       =   1380
   ClientWidth     =   12225
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   8280
   ScaleWidth      =   12225
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdsalir 
      Caption         =   "&Salir"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   450
      Left            =   7695
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   7710
      Width           =   1450
   End
   Begin VB.CommandButton cmdactualizar 
      Caption         =   "&Actualizar"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   450
      Left            =   2970
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   7710
      Width           =   1450
   End
   Begin VB.CommandButton cmddetallado 
      Caption         =   "&Imprimir Detallado"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   450
      Left            =   4530
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   7710
      Width           =   1450
   End
   Begin VB.CommandButton cmdresumen 
      Caption         =   "&Imprimir Resumen"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   450
      Left            =   6120
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   7710
      Width           =   1450
   End
   Begin VB.Frame fraReportes 
      Appearance      =   0  'Flat
      ForeColor       =   &H00C00000&
      Height          =   4365
      Index           =   0
      Left            =   45
      TabIndex        =   16
      Top             =   3240
      Width           =   12135
      Begin DXDBGRIDLibCtl.dxDBGrid gLista 
         Height          =   4050
         Left            =   90
         OleObjectBlob   =   "FrmDetalladoPorProducto.frx":0000
         TabIndex        =   11
         Top             =   180
         Width           =   11970
      End
   End
   Begin VB.Frame fraReportes 
      Appearance      =   0  'Flat
      ForeColor       =   &H00C00000&
      Height          =   3240
      Index           =   4
      Left            =   45
      TabIndex        =   12
      Top             =   0
      Width           =   12135
      Begin VB.CommandButton cmbAyudaMoneda 
         Height          =   315
         Left            =   10260
         Picture         =   "FrmDetalladoPorProducto.frx":6B4C
         Style           =   1  'Graphical
         TabIndex        =   27
         Top             =   2775
         Width           =   390
      End
      Begin VB.CommandButton cmbAyudaVendedor 
         Height          =   315
         Left            =   10260
         Picture         =   "FrmDetalladoPorProducto.frx":6ED6
         Style           =   1  'Graphical
         TabIndex        =   24
         Top             =   2385
         Width           =   390
      End
      Begin VB.CommandButton cmbAyudaProducto 
         Height          =   315
         Left            =   10260
         Picture         =   "FrmDetalladoPorProducto.frx":7260
         Style           =   1  'Graphical
         TabIndex        =   21
         Top             =   1980
         Width           =   390
      End
      Begin VB.Frame fraReportes 
         Appearance      =   0  'Flat
         Caption         =   " Rango de Fechas "
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
         Height          =   765
         Index           =   1
         Left            =   1260
         TabIndex        =   18
         Top             =   1035
         Width           =   9390
         Begin MSComCtl2.DTPicker dtpfInicio 
            Height          =   315
            Left            =   2820
            TabIndex        =   2
            Top             =   315
            Width           =   1230
            _ExtentX        =   2170
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
            Format          =   143458305
            CurrentDate     =   38667
         End
         Begin MSComCtl2.DTPicker dtpFFinal 
            Height          =   315
            Left            =   6225
            TabIndex        =   3
            Top             =   315
            Width           =   1230
            _ExtentX        =   2170
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
            Format          =   143458305
            CurrentDate     =   38667
         End
         Begin VB.Label Label2 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            Caption         =   "Hasta"
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
            Left            =   5625
            TabIndex        =   20
            Top             =   360
            Width           =   420
         End
         Begin VB.Label Label1 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            Caption         =   "Desde"
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
            Left            =   2205
            TabIndex        =   19
            Top             =   360
            Width           =   465
         End
      End
      Begin VB.CommandButton cmbAyudaSucursal 
         Height          =   315
         Left            =   10260
         Picture         =   "FrmDetalladoPorProducto.frx":75EA
         Style           =   1  'Graphical
         TabIndex        =   13
         Top             =   270
         Width           =   390
      End
      Begin CATControls.CATTextBox txtCod_Sucursal 
         Height          =   315
         Left            =   2265
         TabIndex        =   0
         Tag             =   "TidMoneda"
         Top             =   270
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
         Container       =   "FrmDetalladoPorProducto.frx":7974
         Estilo          =   1
         Vacio           =   -1  'True
         EnterTab        =   -1  'True
      End
      Begin CATControls.CATTextBox txtGls_Sucursal 
         Height          =   315
         Left            =   3240
         TabIndex        =   14
         Top             =   270
         Width           =   6975
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
         Container       =   "FrmDetalladoPorProducto.frx":7990
         Vacio           =   -1  'True
      End
      Begin CATControls.CATTextBox txt_Serie 
         Height          =   315
         Left            =   2265
         TabIndex        =   1
         Top             =   630
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
         MaxLength       =   3
         Container       =   "FrmDetalladoPorProducto.frx":79AC
         Estilo          =   1
         Vacio           =   -1  'True
         EnterTab        =   -1  'True
      End
      Begin CATControls.CATTextBox txtCod_Producto 
         Height          =   315
         Left            =   2265
         TabIndex        =   4
         Tag             =   "TidMoneda"
         Top             =   1995
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
         Container       =   "FrmDetalladoPorProducto.frx":79C8
         Estilo          =   1
         Vacio           =   -1  'True
         EnterTab        =   -1  'True
      End
      Begin CATControls.CATTextBox txtGls_Producto 
         Height          =   315
         Left            =   3240
         TabIndex        =   22
         Top             =   1995
         Width           =   6975
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
         Container       =   "FrmDetalladoPorProducto.frx":79E4
         Vacio           =   -1  'True
      End
      Begin CATControls.CATTextBox txtCod_Vendedor 
         Height          =   315
         Left            =   2265
         TabIndex        =   5
         Tag             =   "TidPerCliente"
         Top             =   2385
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
         Container       =   "FrmDetalladoPorProducto.frx":7A00
         Estilo          =   1
         Vacio           =   -1  'True
         EnterTab        =   -1  'True
      End
      Begin CATControls.CATTextBox txtGls_Vendedor 
         Height          =   315
         Left            =   3240
         TabIndex        =   25
         Tag             =   "TGlsCliente"
         Top             =   2385
         Width           =   6975
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
         Locked          =   -1  'True
         Container       =   "FrmDetalladoPorProducto.frx":7A1C
         Estilo          =   1
         Vacio           =   -1  'True
      End
      Begin CATControls.CATTextBox txtCod_Moneda 
         Height          =   315
         Left            =   2265
         TabIndex        =   6
         Tag             =   "TidMoneda"
         Top             =   2775
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
         Container       =   "FrmDetalladoPorProducto.frx":7A38
         Estilo          =   1
         EnterTab        =   -1  'True
      End
      Begin CATControls.CATTextBox txtGls_Moneda 
         Height          =   315
         Left            =   3240
         TabIndex        =   28
         Top             =   2775
         Width           =   6975
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
         Container       =   "FrmDetalladoPorProducto.frx":7A54
         Vacio           =   -1  'True
      End
      Begin VB.Label lbl_Moneda 
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
         ForeColor       =   &H80000007&
         Height          =   210
         Left            =   1260
         TabIndex        =   29
         Top             =   2850
         Width           =   570
      End
      Begin VB.Label Label12 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         Caption         =   "Vendedor"
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
         Left            =   1260
         TabIndex        =   26
         Top             =   2475
         Width           =   720
      End
      Begin VB.Label Label5 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         Caption         =   "Producto"
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
         Left            =   1260
         TabIndex        =   23
         Top             =   2070
         Width           =   645
      End
      Begin VB.Label Label3 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
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
         ForeColor       =   &H80000007&
         Height          =   210
         Left            =   1260
         TabIndex        =   17
         Top             =   675
         Width           =   375
      End
      Begin VB.Label Label4 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         Caption         =   "Sucursal"
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
         Left            =   1260
         TabIndex        =   15
         Top             =   315
         Width           =   645
      End
   End
End
Attribute VB_Name = "FrmDetalladoPorProducto"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmdactualizar_Click()
On Error GoTo Err
Dim StrMsgError As String
Dim X                   As Integer

    validaFormSQL Me, StrMsgError
    If StrMsgError <> "" Then GoTo Err
    
    csql = "select concat(dc.idDocumento,dc.idSerie,dc.idDocVentas) as Item, dc.idDocumento, dc.idSerie, dc.idDocVentas, DATE_FORMAT(dc.FecEmision,GET_FORMAT(DATE, 'EUR')) as FecEmision, " & _
            "dc.idPerCliente, dc.GlsCliente, dc.idPerVendedor, dc.glsVendedor, dd.idProducto, dd.glsProducto, FORMAT(dd.Cantidad,2) as Kg, FORMAT(dd.Cantidad2,2) as Unidades, d.AbreDocumento, " & _
            "CASE '" & Trim(txtCod_Moneda.Text) & "' WHEN 'PEN' THEN  FORMAT(IF(dc.idMoneda = 'PEN', dd.VVUnit,dd.VVUnit * dc.TipoCambio),2) WHEN 'USD' THEN  FORMAT(IF(dc.idMoneda = 'USD', dd.VVUnit,dd.VVUnit / dc.TipoCambio),2) END AS VVUnit, " & _
            "CASE '" & Trim(txtCod_Moneda.Text) & "' WHEN 'PEN' THEN  FORMAT(IF(dc.idMoneda = 'PEN', dd.IGVUnit,dd.IGVUnit * dc.TipoCambio),2) WHEN 'USD' THEN  FORMAT(IF(dc.idMoneda = 'USD', dd.IGVUnit,dd.IGVUnit / dc.TipoCambio),2) END AS IGVUnit, " & _
            "CASE '" & Trim(txtCod_Moneda.Text) & "' WHEN 'PEN' THEN  FORMAT(IF(dc.idMoneda = 'PEN', dd.PVUnit,dd.PVUnit * dc.TipoCambio),2) WHEN 'USD' THEN  FORMAT(IF(dc.idMoneda = 'USD', dd.PVUnit,dd.PVUnit / dc.TipoCambio),2) END As PVUnit, " & _
            "CASE '" & Trim(txtCod_Moneda.Text) & "' WHEN 'PEN' THEN  FORMAT(IF(dc.idMoneda = 'PEN', dd.TotalVVNeto,dd.TotalVVNeto * dc.TipoCambio),2) WHEN 'USD' THEN  FORMAT(IF(dc.idMoneda = 'USD', dd.TotalVVNeto,dd.TotalVVNeto / dc.TipoCambio),2) END AS TotalVVNeto, " & _
            "CASE '" & Trim(txtCod_Moneda.Text) & "' WHEN 'PEN' THEN  FORMAT(IF(dc.idMoneda = 'PEN', dd.TotalIGVNeto,dd.TotalIGVNeto * dc.TipoCambio),2) WHEN 'USD' THEN  FORMAT(IF(dc.idMoneda = 'USD', dd.TotalIGVNeto,dd.TotalIGVNeto / dc.TipoCambio),2) END AS TotalIGVNeto, " & _
            "CASE '" & Trim(txtCod_Moneda.Text) & "' WHEN 'PEN' THEN  FORMAT(IF(dc.idMoneda = 'PEN', dd.TotalPVNeto,dd.TotalPVNeto * dc.TipoCambio),2) WHEN 'USD' THEN  FORMAT(IF(dc.idMoneda = 'USD', dd.TotalPVNeto,dd.TotalPVNeto / dc.TipoCambio),2) END As TotalPVNeto " & _
            "from docventas dc, docventasdet dd, Documentos d " & _
            "Where dc.idEmpresa = dd.idEmpresa And dc.idSucursal = dd.idSucursal And dc.idDocumento = dd.idDocumento " & _
            "and dc.idSerie = dd.idSerie and dc.idDocVentas = dd.idDocVentas and dc.estDocVentas <> 'ANU' and dc.idDocumento = d.idDocumento " & _
            "and dc.iddocumento in('01','03','08','25') and dc.idEmpresa = '" & glsEmpresa & "' and dc.idSucursal like '%" & Trim(txtCod_Sucursal.Text) & "%' " & _
            "and dc.idPerVendedor like '%" & Trim(txtCod_Vendedor.Text) & "%' and dc.idSerie like '%" & Trim(txt_serie.Text) & "%' and dd.idProducto like '%" & Trim(txtCod_Producto.Text) & "%' " & _
            "and dc.FecEmision between '" & Format(dtpfInicio.Value, "yyyy-mm-dd") & "' and '" & Format(dtpFFinal.Value, "yyyy-mm-dd") & "' " & _
            "order by dd.glsProducto, dc.idDocumento, dc.idSerie, dc.idDocVentas"
    
    With GLista
        .DefaultFields = False
        .Dataset.ADODataset.ConnectionString = strcn
        .Dataset.ADODataset.CursorLocation = clUseClient
        .Dataset.Active = False
        .Dataset.ADODataset.CommandText = csql
        .Dataset.DisableControls
        .Dataset.Active = True
        .KeyField = "item"
    End With
    Me.Refresh
    
    For X = 0 To GLista.Columns.Count - 1
        
        GLista.m.ApplyBestFit GLista.Columns(X)
        
    Next X
    
    Exit Sub
Err:
    If StrMsgError = "" Then StrMsgError = Err.Description
    MsgBox StrMsgError, vbInformation, App.Title
End Sub

Private Sub cmdsalir_Click()
    
    Unload Me

End Sub

Private Sub cmbAyudaMoneda_Click()
    
    mostrarAyuda "MONEDA", txtCod_Moneda, txtGls_Moneda

End Sub

Private Sub cmbAyudaProducto_Click()
    
    mostrarAyuda "PRODUCTOS", txtCod_Producto, txtGls_Producto

End Sub

Private Sub cmbAyudaSucursal_Click()
    
    mostrarAyuda "SUCURSAL", txtCod_Sucursal, txtGls_Sucursal

End Sub

Private Sub cmbAyudaVendedor_Click()
    
    mostrarAyuda "VENDEDOR", txtCod_Vendedor, txtGls_Vendedor

End Sub

Private Sub cmddetallado_Click()
On Error GoTo Err
Dim StrMsgError As String

    mostrarReporte "rptDetalleProducto.rpt", "parEmpresa|parSucursal|parVendedor|parSerie|parProducto|parMoneda|parFecDesde|parFecHasta", glsEmpresa & "|" & Trim(txtCod_Sucursal.Text) & "|" & Trim(txtCod_Vendedor.Text) & "|" & Trim(txt_serie.Text) & "|" & Trim(txtCod_Producto.Text) & "|" & Trim(txtCod_Moneda.Text) & "|" & Format(dtpfInicio.Value, "yyyy-mm-dd") & "|" & Format(dtpFFinal.Value, "yyyy-mm-dd"), "Detalle por Producto", StrMsgError
    If StrMsgError <> "" Then GoTo Err

    Exit Sub
    
Err:
    If StrMsgError = "" Then StrMsgError = Err.Description
    MsgBox StrMsgError, vbInformation, App.Title
End Sub

Private Sub cmdresumen_Click()
On Error GoTo Err
Dim StrMsgError As String

    mostrarReporte "rptResumenProducto.rpt", "parEmpresa|parSucursal|parVendedor|parSerie|parProducto|parMoneda|parFecDesde|parFecHasta", glsEmpresa & "|" & Trim(txtCod_Sucursal.Text) & "|" & Trim(txtCod_Vendedor.Text) & "|" & Trim(txt_serie.Text) & "|" & Trim(txtCod_Producto.Text) & "|" & Trim(txtCod_Moneda.Text) & "|" & Format(dtpfInicio.Value, "yyyy-mm-dd") & "|" & Format(dtpFFinal.Value, "yyyy-mm-dd"), "Detalle por Producto", StrMsgError
    If StrMsgError <> "" Then GoTo Err

    Exit Sub
    
Err:
    If StrMsgError = "" Then StrMsgError = Err.Description
    MsgBox StrMsgError, vbInformation, App.Title
End Sub

Private Sub Form_Load()
On Error GoTo Err
Dim StrMsgError As String

    Me.top = 0
    Me.left = 0

    txtGls_Vendedor.Text = "TODOS LOS VENDEDORES"
    txtGls_Producto.Text = "TODOS LOS PRODUCTOS"
    txtGls_Sucursal.Text = "TODOS LAS SUCURSALES"
    txtCod_Moneda.Text = "PEN"
    dtpfInicio.Value = Format(Date, "dd/mm/yyyy")
    dtpFFinal.Value = Format(Date, "dd/mm/yyyy")
    ConfGrid GLista, False, False, False, False
    
    Exit Sub

Err:
    If StrMsgError = "" Then StrMsgError = Err.Description
    MsgBox StrMsgError, vbInformation, App.Title
End Sub

Private Sub txt_Serie_LostFocus()
    
    txt_serie.Text = Format(txt_serie.Text, "0000")

End Sub

Private Sub txtCod_Moneda_Change()
    
    txtGls_Moneda.Text = traerCampo("monedas", "GlsMoneda", "idMoneda", txtCod_Moneda.Text, False)

End Sub

Private Sub txtCod_Producto_Change()
    
    If txtCod_Producto.Text <> "" Then
        txtGls_Producto.Text = traerCampo("productos", "GlsProducto", "idProducto", txtCod_Producto.Text, True)
    Else
        txtGls_Producto.Text = "TODOS LOS PRODUCTOS"
    End If

End Sub

Private Sub txtCod_Sucursal_Change()
    
    If txtCod_Sucursal.Text <> "" Then
        txtGls_Sucursal.Text = traerCampo("personas", "GlsPersona", "idPersona", txtCod_Sucursal.Text, False)
    Else
        txtGls_Sucursal.Text = "TODOS LAS SUCURSALES"
    End If

End Sub

Private Sub txtCod_Vendedor_Change()
    
    If txtCod_Vendedor.Text = "" Then
        txtGls_Vendedor.Text = "TODOS LOS VENDEDORES"
    Else
        txtGls_Vendedor.Text = traerCampo("personas", "GlsPersona", "idPersona", txtCod_Vendedor.Text, False)
    End If

End Sub
