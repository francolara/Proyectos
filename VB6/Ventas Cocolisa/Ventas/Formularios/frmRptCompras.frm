VERSION 5.00
Object = "{6A24B331-7634-11D3-A5B0-0050044A7E1A}#1.5#0"; "DXDBGrid.dll"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "mscomctl.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{F41D1D30-7878-4923-8CB3-6CCACDC9C9DE}#1.0#0"; "CATControls.ocx"
Begin VB.Form frmRptCompras 
   Appearance      =   0  'Flat
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Registro de Movimientos - Compras"
   ClientHeight    =   9795
   ClientLeft      =   5670
   ClientTop       =   2820
   ClientWidth     =   13125
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   9795
   ScaleWidth      =   13125
   ShowInTaskbar   =   0   'False
   Begin VB.Frame fraListado 
      Appearance      =   0  'Flat
      ForeColor       =   &H00C00000&
      Height          =   9105
      Left            =   60
      TabIndex        =   6
      Top             =   660
      Width           =   13005
      Begin VB.Frame Frame1 
         Appearance      =   0  'Flat
         ForeColor       =   &H80000008&
         Height          =   1380
         Left            =   90
         TabIndex        =   7
         Top             =   120
         Width           =   12840
         Begin VB.CommandButton cmbAyudaAlmacen 
            Height          =   315
            Left            =   6705
            Picture         =   "frmRptCompras.frx":0000
            Style           =   1  'Graphical
            TabIndex        =   18
            Top             =   940
            Width           =   390
         End
         Begin VB.CommandButton cmbAyudaSucursal 
            Height          =   315
            Left            =   6705
            Picture         =   "frmRptCompras.frx":038A
            Style           =   1  'Graphical
            TabIndex        =   9
            Top             =   595
            Width           =   390
         End
         Begin VB.CommandButton cmbAyudaProveedor 
            Height          =   315
            Left            =   6705
            Picture         =   "frmRptCompras.frx":0714
            Style           =   1  'Graphical
            TabIndex        =   8
            Top             =   250
            Width           =   390
         End
         Begin CATControls.CATTextBox txtCod_Sucursal 
            Height          =   315
            Left            =   1035
            TabIndex        =   1
            Tag             =   "TidMoneda"
            Top             =   600
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
            Container       =   "frmRptCompras.frx":0A9E
            Estilo          =   1
            EnterTab        =   -1  'True
         End
         Begin CATControls.CATTextBox txtGls_Sucursal 
            Height          =   315
            Left            =   1980
            TabIndex        =   10
            Top             =   600
            Width           =   4680
            _ExtentX        =   8255
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
            Container       =   "frmRptCompras.frx":0ABA
            Vacio           =   -1  'True
         End
         Begin CATControls.CATTextBox txtCod_Proveedor 
            Height          =   315
            Left            =   1035
            TabIndex        =   0
            Tag             =   "TidMoneda"
            Top             =   255
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
            Container       =   "frmRptCompras.frx":0AD6
            Estilo          =   1
            EnterTab        =   -1  'True
         End
         Begin CATControls.CATTextBox txtGls_Proveedor 
            Height          =   315
            Left            =   1980
            TabIndex        =   11
            Top             =   255
            Width           =   4680
            _ExtentX        =   8255
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
            Container       =   "frmRptCompras.frx":0AF2
            Vacio           =   -1  'True
         End
         Begin MSComCtl2.DTPicker dtpfInicio 
            Height          =   315
            Left            =   8910
            TabIndex        =   3
            Top             =   255
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
            Format          =   132513793
            CurrentDate     =   38667
         End
         Begin MSComCtl2.DTPicker dtpFFinal 
            Height          =   315
            Left            =   11370
            TabIndex        =   4
            Top             =   255
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
            Format          =   132513793
            CurrentDate     =   38667
         End
         Begin CATControls.CATTextBox txtCod_Almacen 
            Height          =   315
            Left            =   1035
            TabIndex        =   2
            Tag             =   "TidAlmacen"
            Top             =   945
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
            Container       =   "frmRptCompras.frx":0B0E
            Estilo          =   1
            EnterTab        =   -1  'True
         End
         Begin CATControls.CATTextBox txtGls_Almacen 
            Height          =   315
            Left            =   1980
            TabIndex        =   19
            Top             =   945
            Width           =   4680
            _ExtentX        =   8255
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
            Container       =   "frmRptCompras.frx":0B2A
            Vacio           =   -1  'True
         End
         Begin VB.Label lbl_Almacen 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            Caption         =   "Almacén"
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
            Left            =   200
            TabIndex        =   20
            Top             =   1020
            Width           =   630
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
            Left            =   200
            TabIndex        =   15
            Top             =   645
            Width           =   645
         End
         Begin VB.Label Label2 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            Caption         =   "Proveedor"
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
            Left            =   200
            TabIndex        =   14
            Top             =   300
            Width           =   750
         End
         Begin VB.Label Label3 
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
            Left            =   10830
            TabIndex        =   13
            Top             =   300
            Width           =   420
         End
         Begin VB.Label Label5 
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
            Left            =   8325
            TabIndex        =   12
            Top             =   300
            Width           =   465
         End
      End
      Begin DXDBGRIDLibCtl.dxDBGrid gCabecera 
         Height          =   4230
         Left            =   120
         OleObjectBlob   =   "frmRptCompras.frx":0B46
         TabIndex        =   16
         Top             =   1620
         Width           =   12795
      End
      Begin DXDBGRIDLibCtl.dxDBGrid gDetalle 
         Height          =   2985
         Left            =   120
         OleObjectBlob   =   "frmRptCompras.frx":4CD1
         TabIndex        =   17
         Top             =   6000
         Width           =   12795
      End
   End
   Begin MSComctlLib.ImageList imgDocVentas 
      Left            =   3540
      Top             =   5100
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
            Picture         =   "frmRptCompras.frx":84E2
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRptCompras.frx":887C
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRptCompras.frx":8CCE
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRptCompras.frx":9068
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRptCompras.frx":9402
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRptCompras.frx":979C
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRptCompras.frx":9B36
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRptCompras.frx":9ED0
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRptCompras.frx":A26A
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRptCompras.frx":A604
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRptCompras.frx":A99E
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRptCompras.frx":B660
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRptCompras.frx":B9FA
            Key             =   ""
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRptCompras.frx":BE4C
            Key             =   ""
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRptCompras.frx":C1E6
            Key             =   ""
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRptCompras.frx":CBF8
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
      Width           =   13125
      _ExtentX        =   23151
      _ExtentY        =   1164
      ButtonWidth     =   2990
      ButtonHeight    =   1005
      Appearance      =   1
      ImageList       =   "imgDocVentas"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   7
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "         Actualizar         "
            Object.ToolTipText     =   "Nuevo"
            ImageIndex      =   14
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
            Caption         =   "Reg. Compras"
            Object.ToolTipText     =   "Registro de Compras"
            ImageIndex      =   10
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Resumen"
            Object.ToolTipText     =   "Imprimir Resumen"
            ImageIndex      =   12
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Detallado"
            Object.ToolTipText     =   "Imprimir Detalle"
            ImageIndex      =   15
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Excel"
            Object.ToolTipText     =   "Cancelar"
            ImageIndex      =   13
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Por Producto"
            ImageIndex      =   1
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Salir"
            Object.ToolTipText     =   "Salir"
            ImageIndex      =   2
         EndProperty
      EndProperty
      BorderStyle     =   1
   End
End
Attribute VB_Name = "frmRptCompras"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmbAyudaAlmacen_Click()
Dim strCondicion As String
    
    If txtCod_Sucursal.Text = "" Then
        MsgBox "Seleccione una sucursal", vbInformation, App.Title
        txtCod_Sucursal.SetFocus
        Exit Sub
    End If

    strCondicion = " AND idSucursal = '" & txtCod_Sucursal.Text & "'"
    mostrarAyuda "ALMACENVTA", txtCod_Almacen, txtGls_Almacen, strCondicion
    
End Sub

Private Sub cmbAyudaSucursal_Click()
    
    mostrarAyuda "SUCURSAL", txtCod_Sucursal, txtGls_Sucursal

End Sub

Private Sub cmbAyudaProveedor_Click()
    
    mostrarAyuda "PROVEEDOR", txtCod_Proveedor, txtGls_Proveedor

End Sub

Private Sub Form_Load()
    
    Me.top = 0
    Me.left = 0

    ConfGrid gCabecera, False, False, True, False
    ConfGrid gDetalle, False, False, True, False
    
    txtGls_Proveedor.Text = "TODOS LOS PROVEEDORES"
    txtGls_Sucursal.Text = "TODAS LAS SUCURSALES"
    txtGls_Almacen.Text = "TODOS LOS ALMACENES"
    dtpfInicio.Value = Format(Date, "dd/mm/yyyy")
    dtpFFinal.Value = Format(Date, "dd/mm/yyyy")
    
End Sub

Private Sub gCabecera_OnChangeNode(ByVal OldNode As DXDBGRIDLibCtl.IdxGridNode, ByVal Node As DXDBGRIDLibCtl.IdxGridNode)
On Error GoTo Err
Dim StrMsgError As String

    listaComprasDetalle StrMsgError
    If StrMsgError <> "" Then GoTo Err
        
    Exit Sub

Err:
    If StrMsgError = "" Then StrMsgError = Err.Description
    MsgBox StrMsgError, vbInformation, App.Title
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
On Error GoTo Err
Dim StrMsgError As String
Dim strSucursal As String
Dim strAlmacen As String
Dim strProveedor As String
Dim strFecIni As String
Dim strFecFin As String

    strSucursal = Trim(txtCod_Sucursal.Text)
    strAlmacen = Trim(txtCod_Almacen.Text)
    strProveedor = Trim(txtCod_Proveedor.Text)
    strFecIni = Format(dtpfInicio.Value, "yyyy-mm-dd")
    strFecFin = Format(dtpFFinal.Value, "yyyy-mm-dd")
        
    Select Case Button.Index
        Case 1 'Actualizar
            listaCompras StrMsgError
            If StrMsgError <> "" Then GoTo Err
        Case 2 'Reg. Compras
        Case 3 'Resumen
            mostrarReporte "rptComprasResumen.rpt", "varEmpresa|varSucursal|varAlmacen|varProveedor|varFechaIni|varFechaFin", glsEmpresa & "|" & strSucursal & "|" & strAlmacen & "|" & strProveedor & "|" & strFecIni & "|" & strFecFin, "Resumen de Compras", StrMsgError
            If StrMsgError <> "" Then GoTo Err
        
        Case 4 'Detallado
            mostrarReporte "rptComprasDetalle.rpt", "varEmpresa|varSucursal|varAlmacen|varProveedor|varFechaIni|varFechaFin", glsEmpresa & "|" & strSucursal & "|" & strAlmacen & "|" & strProveedor & "|" & strFecIni & "|" & strFecFin, "Detalle de Compras", StrMsgError
            If StrMsgError <> "" Then GoTo Err
            
        Case 5 'Excel
            gCabecera.m.ExportToXLS App.Path & "\Temporales\Compras.xls"
            ShellEx App.Path & "\Temporales\Compras.xls", essSW_MAXIMIZE, , , "open", Me.hwnd
        Case 6 'Detallado
            mostrarReporte "rptComprasDetallePorProducto.rpt", "varEmpresa|varSucursal|varAlmacen|varProveedor|varFechaIni|varFechaFin", glsEmpresa & "|" & strSucursal & "|" & strAlmacen & "|" & strProveedor & "|" & strFecIni & "|" & strFecFin, "Detalle de Compras", StrMsgError
            If StrMsgError <> "" Then GoTo Err
        Case 7 'Salir
            Unload Me
    End Select
    
    Exit Sub

Err:
    If StrMsgError = "" Then StrMsgError = Err.Description
    MsgBox StrMsgError, vbInformation, App.Title
End Sub

Private Sub txtCod_Almacen_Change()
    
    If txtCod_Almacen.Text <> "" Then
        txtGls_Almacen.Text = traerCampo("almacenes", "GlsAlmacen", "idAlmacen", txtCod_Almacen.Text, True)
    Else
        txtGls_Almacen.Text = "TODOS LOS ALMACENES"
    End If

End Sub

Private Sub txtCod_Almacen_KeyDown(KeyCode As Integer, Shift As Integer)

    If KeyCode = 8 Then
        txtCod_Almacen.Text = ""
    End If

End Sub

Private Sub txtCod_Almacen_KeyPress(KeyAscii As Integer)
Dim strCondicion As String

    If KeyAscii <> 13 Then
        If txtCod_Sucursal.Text = "" Then
            MsgBox "Seleccione una sucursal", vbInformation, App.Title
            txtCod_Sucursal.SetFocus
            KeyAscii = 0
            Exit Sub
        End If
        strCondicion = " AND idSucursal = '" & txtCod_Sucursal.Text & "'"
        mostrarAyudaKeyascii KeyAscii, "ALMACENVTA", txtCod_Almacen, txtGls_Almacen, strCondicion
        KeyAscii = 0
    End If

End Sub

Private Sub txtCod_Sucursal_Change()
    
    If txtCod_Sucursal.Text <> "" Then
        txtGls_Sucursal.Text = traerCampo("personas", "GlsPersona", "idPersona", txtCod_Sucursal.Text, False)
    Else
        txtGls_Sucursal.Text = "TODOS LAS SUCURSALES"
    End If
    txtCod_Almacen.Text = ""

End Sub

Private Sub txtCod_Sucursal_KeyDown(KeyCode As Integer, Shift As Integer)

    If KeyCode = 8 Then
        txtCod_Sucursal.Text = ""
    End If

End Sub

Private Sub txtCod_Sucursal_KeyPress(KeyAscii As Integer)

    If KeyAscii <> 13 And KeyAscii <> 8 Then
        mostrarAyudaKeyascii KeyAscii, "SUCURSAL", txtCod_Sucursal, txtGls_Sucursal
        KeyAscii = 0
    End If

End Sub

Private Sub txtCod_Proveedor_Change()
    
    If txtCod_Proveedor.Text <> "" Then
        txtGls_Proveedor.Text = traerCampo("personas", "GlsPersona", "idPersona", txtCod_Proveedor.Text, False)
    Else
        txtGls_Proveedor.Text = "TODOS LOS PROVEEDORES"
    End If

End Sub

Private Sub txtCod_Proveedor_KeyDown(KeyCode As Integer, Shift As Integer)

    If KeyCode = 8 Then
        txtCod_Proveedor.Text = ""
    End If

End Sub

Private Sub txtCod_Proveedor_KeyPress(KeyAscii As Integer)

    If KeyAscii <> 13 And KeyAscii <> 8 Then
        mostrarAyudaKeyascii KeyAscii, "PROVEEDOR", txtCod_Proveedor, txtGls_Proveedor
        KeyAscii = 0
    End If

End Sub

Private Sub listaCompras(ByRef StrMsgError As String)
Dim strSucursal As String
Dim strAlmacen As String
Dim strProveedor As String
Dim strFecIni As String
Dim strFecFin As String
Dim rsdatos   As New ADODB.Recordset
On Error GoTo Err

    strSucursal = Trim(txtCod_Sucursal.Text)
    strAlmacen = Trim(txtCod_Almacen.Text)
    strProveedor = Trim(txtCod_Proveedor.Text)
    strFecIni = Format(dtpfInicio.Value, "yyyy-mm-dd")
    strFecFin = Format(dtpFFinal.Value, "yyyy-mm-dd")
       
'    With gCabecera
'        .DefaultFields = False
'        .Dataset.ADODataset.ConnectionString = strcn
'        .Dataset.ADODataset.CursorLocation = clUseClient
'        .Dataset.Active = False
'        .Dataset.ADODataset.CommandText = "CALL spu_ListaCompras ('" & glsEmpresa & "','" & strSucursal & "','" & strAlmacen & "','" & strProveedor & "','" & strFecIni & "','" & strFecFin & "')"
'        .Dataset.DisableControls
'        .Dataset.Active = True
'        .KeyField = "idValesCab"
'    End With
'

    csql = "EXECUTE spu_ListaCompras '" & glsEmpresa & "','" & strSucursal & "','" & strAlmacen & "','" & strProveedor & "','" & strFecIni & "','" & strFecFin & "'"
    If rsdatos.State = 1 Then rsdatos.Close: Set rsdatos = Nothing
    rsdatos.Open csql, Cn, adOpenStatic, adLockOptimistic
        
    Set gCabecera.DataSource = rsdatos

    listaComprasDetalle StrMsgError
    If StrMsgError <> "" Then GoTo Err
            
    Me.Refresh
    Exit Sub

Err:
    If StrMsgError = "" Then StrMsgError = Err.Description
End Sub

Private Sub listaComprasDetalle(ByRef StrMsgError As String)
Dim rsdatos                     As New ADODB.Recordset
On Error GoTo Err
    
    
    csql = "EXECUTE spu_ListaComprasDetGrilla '" & glsEmpresa & "','" & gCabecera.Columns.ColumnByFieldName("idSucursal").Value & "','" & right(gCabecera.Columns.ColumnByFieldName("idValesCab").Value, 8) & "','" & left(gCabecera.Columns.ColumnByFieldName("idValesCab").Value, 1) & "'"
    If rsdatos.State = 1 Then rsdatos.Close: Set rsdatos = Nothing
    rsdatos.Open csql, Cn, adOpenStatic, adLockOptimistic
        
    Set gDetalle.DataSource = rsdatos
    
'    With gDetalle
'        .DefaultFields = False
'        .Dataset.ADODataset.ConnectionString = strcn
'        .Dataset.ADODataset.CursorLocation = clUseClient
'        .Dataset.Active = False
'        .Dataset.ADODataset.CommandText = "CALL spu_ListaComprasDetGrilla ('" & glsEmpresa & "','" & gCabecera.Columns.ColumnByFieldName("idSucursal").Value & "','" & right(gCabecera.Columns.ColumnByFieldName("idValesCab").Value, 8) & "','" & left(gCabecera.Columns.ColumnByFieldName("idValesCab").Value, 1) & "')"
'        .Dataset.DisableControls
'        .Dataset.Active = True
'        .KeyField = "item"
'    End With
    
    
    
    Me.Refresh
    
    Exit Sub

Err:
    If StrMsgError = "" Then StrMsgError = Err.Description
End Sub
