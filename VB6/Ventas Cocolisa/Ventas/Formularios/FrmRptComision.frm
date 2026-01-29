VERSION 5.00
Object = "{6A24B331-7634-11D3-A5B0-0050044A7E1A}#1.5#0"; "DXDBGrid.dll"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{F41D1D30-7878-4923-8CB3-6CCACDC9C9DE}#1.0#0"; "CATControls.ocx"
Begin VB.Form FrmRptComision 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Reporte de Comisiones"
   ClientHeight    =   9105
   ClientLeft      =   1710
   ClientTop       =   705
   ClientWidth     =   13575
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
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   9105
   ScaleWidth      =   13575
   ShowInTaskbar   =   0   'False
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
      Height          =   8340
      Left            =   45
      TabIndex        =   0
      Top             =   675
      Width           =   13470
      Begin VB.CheckBox chkNotaCredito 
         Caption         =   "Considerar Nota Crédito"
         Height          =   240
         Left            =   180
         TabIndex        =   22
         Tag             =   "NindInsertaPrecioLista"
         Top             =   2385
         Width           =   2100
      End
      Begin VB.Frame FraGastosventa 
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
         Height          =   2175
         Left            =   5940
         TabIndex        =   17
         Top             =   120
         Width           =   7380
         Begin DXDBGRIDLibCtl.dxDBGrid gGastosVentas 
            Height          =   1875
            Left            =   45
            OleObjectBlob   =   "FrmRptComision.frx":0000
            TabIndex        =   18
            Top             =   135
            Width           =   7215
         End
      End
      Begin VB.Frame fraReportes 
         Appearance      =   0  'Flat
         Caption         =   " Rango de Fechas "
         ForeColor       =   &H00000000&
         Height          =   675
         Index           =   1
         Left            =   180
         TabIndex        =   8
         Top             =   1635
         Width           =   5685
         Begin MSComCtl2.DTPicker dtpfInicio 
            Height          =   315
            Left            =   1155
            TabIndex        =   9
            Top             =   210
            Width           =   1365
            _ExtentX        =   2408
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
            Format          =   121700353
            CurrentDate     =   38667
         End
         Begin MSComCtl2.DTPicker dtpFFinal 
            Height          =   315
            Left            =   3840
            TabIndex        =   10
            Top             =   210
            Width           =   1365
            _ExtentX        =   2408
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
            Format          =   121700353
            CurrentDate     =   38667
         End
         Begin VB.Label Label1 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            Caption         =   "Desde"
            ForeColor       =   &H80000008&
            Height          =   210
            Left            =   540
            TabIndex        =   12
            Top             =   285
            Width           =   465
         End
         Begin VB.Label Label2 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            Caption         =   "Hasta"
            ForeColor       =   &H80000008&
            Height          =   210
            Left            =   3285
            TabIndex        =   11
            Top             =   285
            Width           =   420
         End
      End
      Begin VB.Frame fraReportes 
         Appearance      =   0  'Flat
         ForeColor       =   &H00C00000&
         Height          =   630
         Index           =   4
         Left            =   180
         TabIndex        =   3
         Top             =   135
         Width           =   5655
         Begin VB.CommandButton cmbAyudaSucursal 
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
            Left            =   5175
            Picture         =   "FrmRptComision.frx":2EBF
            Style           =   1  'Graphical
            TabIndex        =   4
            Top             =   180
            Width           =   390
         End
         Begin CATControls.CATTextBox txtCod_Sucursal 
            Height          =   285
            Left            =   930
            TabIndex        =   5
            Tag             =   "TidMoneda"
            Top             =   180
            Width           =   825
            _ExtentX        =   1455
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
            Container       =   "FrmRptComision.frx":3249
            Estilo          =   1
            EnterTab        =   -1  'True
         End
         Begin CATControls.CATTextBox txtGls_Sucursal 
            Height          =   285
            Left            =   1815
            TabIndex        =   6
            Top             =   180
            Width           =   3330
            _ExtentX        =   5874
            _ExtentY        =   503
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
            Container       =   "FrmRptComision.frx":3265
            Vacio           =   -1  'True
         End
         Begin VB.Label Label4 
            Appearance      =   0  'Flat
            Caption         =   "Sucursal"
            ForeColor       =   &H80000007&
            Height          =   240
            Left            =   105
            TabIndex        =   7
            Top             =   225
            Width           =   765
         End
      End
      Begin VB.Frame fraReportes 
         Appearance      =   0  'Flat
         ForeColor       =   &H00C00000&
         Height          =   675
         Index           =   15
         Left            =   180
         TabIndex        =   2
         Top             =   855
         Width           =   5670
         Begin VB.CommandButton cmbAyudaVendedor 
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
            Left            =   5190
            Picture         =   "FrmRptComision.frx":3281
            Style           =   1  'Graphical
            TabIndex        =   19
            Top             =   225
            Width           =   390
         End
         Begin CATControls.CATTextBox txtCod_Vendedor 
            Height          =   285
            Left            =   945
            TabIndex        =   20
            Tag             =   "TidMoneda"
            Top             =   225
            Width           =   825
            _ExtentX        =   1455
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
            Container       =   "FrmRptComision.frx":360B
            Estilo          =   1
            EnterTab        =   -1  'True
         End
         Begin CATControls.CATTextBox txtGls_Vendedor 
            Height          =   285
            Left            =   1830
            TabIndex        =   21
            Top             =   225
            Width           =   3330
            _ExtentX        =   5874
            _ExtentY        =   503
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
            Container       =   "FrmRptComision.frx":3627
            Vacio           =   -1  'True
         End
         Begin VB.Label Label12 
            Appearance      =   0  'Flat
            Caption         =   "Vendedor"
            ForeColor       =   &H80000007&
            Height          =   240
            Left            =   90
            TabIndex        =   13
            Top             =   270
            Width           =   765
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
         ForeColor       =   &H80000008&
         Height          =   5550
         Left            =   180
         TabIndex        =   1
         Top             =   2595
         Width           =   13200
         Begin DXDBGRIDLibCtl.dxDBGrid g 
            Height          =   5235
            Left            =   90
            OleObjectBlob   =   "FrmRptComision.frx":3643
            TabIndex        =   15
            Top             =   180
            Width           =   13020
         End
      End
   End
   Begin MSComctlLib.ImageList imgDocVentas 
      Left            =   1680
      Top             =   5490
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
            Picture         =   "FrmRptComision.frx":9410
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmRptComision.frx":97AA
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmRptComision.frx":9BFC
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmRptComision.frx":9F96
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmRptComision.frx":A330
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmRptComision.frx":A6CA
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmRptComision.frx":AA64
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmRptComision.frx":ADFE
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmRptComision.frx":B198
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmRptComision.frx":B532
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmRptComision.frx":B8CC
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmRptComision.frx":C58E
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmRptComision.frx":C928
            Key             =   ""
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmRptComision.frx":CD7A
            Key             =   ""
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmRptComision.frx":D114
            Key             =   ""
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmRptComision.frx":DB26
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   660
      Left            =   0
      TabIndex        =   14
      Top             =   0
      Width           =   13575
      _ExtentX        =   23945
      _ExtentY        =   1164
      ButtonWidth     =   1561
      ButtonHeight    =   1005
      Appearance      =   1
      ImageList       =   "imgDocVentas"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   5
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Actualizar"
            Object.ToolTipText     =   "Nuevo"
            ImageIndex      =   14
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Excel"
            Object.ToolTipText     =   "Cancelar"
            ImageIndex      =   13
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Gasto"
            ImageIndex      =   5
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Clientes"
            ImageIndex      =   12
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Salir"
            Object.ToolTipText     =   "Salir"
            ImageIndex      =   2
         EndProperty
      EndProperty
      BorderStyle     =   1
   End
   Begin DXDBGRIDLibCtl.dxDBGrid g___ 
      Height          =   570
      Left            =   0
      OleObjectBlob   =   "FrmRptComision.frx":E1F8
      TabIndex        =   16
      Top             =   0
      Width           =   870
   End
End
Attribute VB_Name = "FrmRptComision"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmbAyudaSucursal_Click()
    
    mostrarAyuda "SUCURSAL", txtCod_Sucursal, txtGls_Sucursal

End Sub

Private Sub cmbAyudaVendedor_Click()
    
    mostrarAyuda "VENDEDOR", txtCod_Vendedor, txtGls_Vendedor

End Sub

Private Sub Form_Load()

    Me.top = 0
    Me.left = 0
    
    ConfGrid1 G, False, True, False, False
    ConfGrid1 gGastosVentas, False, True, False, False
    
    txtGls_Vendedor.Text = "TODOS LOS VENDEDORES"
    txtCod_Sucursal.Text = glsSucursal
    chkNotaCredito.Value = 1
    
    dtpfInicio.Value = Format(Date, "dd/mm/yyyy")
    dtpFFinal.Value = Format(Date, "dd/mm/yyyy")
    
    If traerCampo("PARAMETROS", "VALPARAMETRO", "GLSPARAMETRO", "VISUALIZA_CLIENTESXVENDEDOR", True) = "S" Then
        If traerCampo("usuarios", "indJefe", "idUsuario", glsUser, True) = "0" Then
            txtCod_Vendedor.Text = Trim("" & glsUser)
            fraReportes(15).Enabled = False
        Else
            fraReportes(15).Enabled = True
        End If
    Else
        fraReportes(15).Enabled = True
    End If

End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
On Error GoTo Err
Dim rsReporte As ADODB.Recordset
Dim StrMsgError As String
Dim strSucursal As String
Dim strTipoDoc As String
Dim strSerie As String
Dim strFecIni As String
Dim strFecFin As String

    Select Case Button.Index
        Case 1 'Actualizar
            listaComisiones StrMsgError
            If StrMsgError <> "" Then GoTo Err
        Case 2 'Excel
            G.m.ExportToXLS App.Path & "\Temporales\Comisiones.xls"
            ShellEx App.Path & "\Temporales\Comisiones.xls", essSW_MAXIMIZE, , , "open", Me.hwnd
        Case 3 'Mantenimiento GastoVenta
            FrmGastoVenta.Show 1
        Case 4 'Listado de Clientes que no Comisionan
            FrmListaClientesComisionistas.Show 1
        Case 5 'Salir
            Unload Me
    End Select
    
    Exit Sub

Err:
    If StrMsgError = "" Then StrMsgError = Err.Description
    MsgBox StrMsgError, vbInformation, App.Title
End Sub

Private Sub listaComisiones(ByRef StrMsgError As String)
On Error GoTo Err
Dim strSucursal As String
Dim strTipoDoc As String
Dim strSerie As String
Dim strFecIni As String
Dim strFecFin As String
Dim dblV          As Double
Dim dblVU         As Double
Dim dblTG         As Double
Dim dblAcumulaTG  As Double
Dim strdocg       As String
Dim strdocc       As String
Dim item          As Integer
Dim rst           As New ADODB.Recordset
Dim rsd           As New ADODB.Recordset
Dim o             As Integer
Dim intFor        As Integer
Dim dblMaxMargenN As Double
Dim dblMaxMargenI As Double

    item = 0
    If rsd.State = 1 Then rsd.Close
    rsd.Fields.Append "item", adVarChar, 30, adFldIsNullable
    rsd.Fields.Append "idCodGastoVenta", adVarChar, 10, adFldIsNullable
    rsd.Fields.Append "glsDescGastoVenta", adVarChar, 500, adFldIsNullable
    rsd.Fields.Append "FecRegistro", adVarChar, 50, adFldIsNullable
    rsd.Fields.Append "glsVendedor", adVarChar, 300, adFldIsNullable
    rsd.Fields.Append "TotalGastoVenta", adDouble, 14, adFldIsNullable
    rsd.Fields.Append "TotalGastoxp", adDouble, 14, adFldIsNullable
    rsd.Open
    
    strFecIni = Format(dtpfInicio.Value, "yyyy-mm-dd")
    strFecFin = Format(dtpFFinal.Value, "yyyy-mm-dd")
       
    With G
        .DefaultFields = False
        .Dataset.ADODataset.ConnectionString = strcn '''Cn
        .Dataset.ADODataset.CursorLocation = clUseClient
        .Dataset.Active = False
        If chkNotaCredito.Value = 1 Then
            .Dataset.ADODataset.CommandText = "CALL Spu_ComisionAlsisacNc('" & glsEmpresa & "','" & strFecIni & "','" & strFecFin & "','" & Trim("" & txtCod_Sucursal.Text) & "','" & Trim("" & txtCod_Vendedor.Text) & "')"
        Else
            .Dataset.ADODataset.CommandText = "CALL Spu_ComisionAlsisac('" & glsEmpresa & "','" & strFecIni & "','" & strFecFin & "','" & Trim("" & txtCod_Sucursal.Text) & "','" & Trim("" & txtCod_Vendedor.Text) & "')"
        End If
        .Dataset.DisableControls
        .Dataset.Active = True
        .KeyField = "item"
    End With
 
    For intFor = 0 To G.Columns.Count - 1
        G.m.ApplyBestFit G.Columns(intFor)
    Next intFor
    
'    csql = "Select distinct g.idCodGastoVenta,g.glsDescGastoVenta,  DATE_FORMAT(g.FecRegistro,'%d/%m/%Y') as FecRegistro , g.glsVendedor, g.TotalGastoVenta ,d.NroDocumento " & _
'           "From GastosVenta g  Inner Join GastosVentadet d On g.idCodGastoVenta = d.idCodGastoVenta And g.idEmpresa = d.idEmpresa  INNER Join temcomision2 t  On t.Nro_Comp = d.NroDocumento   And t.idEmpresa = d.idEmpresa  Inner Join Clientes c On c.idCliente = d.idCliente And c.idEmpresa = d.idEmpresa  " & _
'           "Where g.idEmpresa = '" & glsEmpresa & "' And  c.indComision Not In (1) " & _
'           "And g.idvendedor in (Select  T.idvendedor From temcomision1 T  Where t.idVendedor='" & txtCod_Vendedor.Text & "') "
'
'    If rst.State = 1 Then rst.Close
'    rst.Open csql, strcn, adOpenStatic, adLockOptimistic
'    If Not rst.EOF Then
'        rst.MoveFirst
'        Do While Not rst.EOF
'            dblV = 0
'            dblVU = 0
'            dblAcumulaTG = 0
'            strdocg = "" & rst.Fields("NroDocumento").Value
'            If G.Count > 0 Then
'                G.Dataset.First
'                dblMaxMargenN = traerCampo("margenescomisiones", "Max(porComision)", "idSucursal", glsSucursal, True, "idtipo = 'N'")
'                dblMaxMargenI = traerCampo("margenescomisiones", "Max(porComision)", "idSucursal", glsSucursal, True, "idtipo = 'I'")
'
'                Do While Not G.Dataset.EOF
'                    strdocc = "" & G.Columns.ColumnByFieldName("Nro_Comp").Value
'                    If strdocg = strdocc Then
'                        If G.Columns.ColumnByFieldName("idMarca").Value = "N" Then
'                            dblV = Format(G.Columns.ColumnByFieldName("VVUnitCompra").Value * (dblMaxMargenN), "0.00")
'                        Else
'                            dblV = Format(G.Columns.ColumnByFieldName("VVUnitCompra").Value * (dblMaxMargenI), "0.00")
'                        End If
'                        dblVU = Format(G.Columns.ColumnByFieldName("TotalVVNeto").Value / G.Columns.ColumnByFieldName("VVUnit").Value, "0.00")
'                        dblTG = Format(((Format(G.Columns.ColumnByFieldName("VVUnit").Value, "0.00") - Format(G.Columns.ColumnByFieldName("VVUnitCompra").Value, "0.00")) * dblVU) - ((dblV - Format(G.Columns.ColumnByFieldName("VVUnitCompra").Value, "0.00")) * dblVU), "0.00")
'                        dblTG = IIf(left(dblTG, 1) = "-", 0, dblTG)
'                    End If
'                    If strdocg = strdocc Then
'                        dblAcumulaTG = dblAcumulaTG + dblTG
'                        dblV = 0
'                        dblVU = 0
'                        dblTG = 0
'                    End If
'                    G.Dataset.Next
'                Loop
'            End If
'
'            rsd.AddNew
'            item = item + 1
'            rsd.Fields("item") = item
'            rsd.Fields("idCodGastoVenta") = "" & rst.Fields("idCodGastoVenta")
'            rsd.Fields("glsDescGastoVenta") = "" & rst.Fields("glsDescGastoVenta")
'            rsd.Fields("FecRegistro") = "" & rst.Fields("FecRegistro")
'            rsd.Fields("glsVendedor") = "" & rst.Fields("glsVendedor")
'            rsd.Fields("TotalGastoVenta") = rst.Fields("TotalGastoVenta")
'            rsd.Fields("TotalGastoxp") = dblAcumulaTG
'            rst.MoveNext
'
'            dblV = 0
'            dblVU = 0
'            dblAcumulaTG = 0
'        Loop
'    End If
'
'    Set gGastosVentas.DataSource = Nothing
'    mostrarDatosGridSQL gGastosVentas, rsd, StrMsgError
    
    For intFor = 0 To gGastosVentas.Columns.Count - 1
       gGastosVentas.m.ApplyBestFit gGastosVentas.Columns(intFor)
    Next intFor
    Me.Refresh

    Exit Sub
Err:
    If StrMsgError = "" Then StrMsgError = Err.Description
    Exit Sub
    Resume
End Sub

Private Sub txtCod_Sucursal_Change()
    
    If txtCod_Sucursal.Text <> "" Then
        txtGls_Sucursal.Text = traerCampo("personas", "GlsPersona", "idPersona", txtCod_Sucursal.Text, False)
    Else
        txtGls_Sucursal.Text = "TODOS LAS SUCURSALES"
    End If

End Sub

Private Sub txtCod_Vendedor_Change()
    
    If txtCod_Vendedor.Text <> "" Then
        txtGls_Vendedor.Text = traerCampo("personas", "GlsPersona", "idPersona", txtCod_Vendedor.Text, False)
    Else
        txtGls_Vendedor.Text = "TODOS LOS VENDEDORES"
    End If

End Sub
