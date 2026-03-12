VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{F41D1D30-7878-4923-8CB3-6CCACDC9C9DE}#1.0#0"; "catcontrols.ocx"
Begin VB.Form FrmRptRentabilidad 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Reporte de Rentabilidad"
   ClientHeight    =   2805
   ClientLeft      =   8700
   ClientTop       =   2400
   ClientWidth     =   7665
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
   ScaleHeight     =   2805
   ScaleWidth      =   7665
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2295
      Left            =   45
      TabIndex        =   2
      Top             =   0
      Width           =   7530
      Begin VB.CheckBox ChkProcesaCostos 
         Caption         =   "Procesar Costos"
         Height          =   240
         Left            =   5895
         TabIndex        =   16
         Top             =   1845
         Value           =   1  'Checked
         Width           =   1545
      End
      Begin VB.Frame FraOrden 
         Appearance      =   0  'Flat
         Caption         =   " Orden "
         ForeColor       =   &H00000000&
         Height          =   765
         Left            =   135
         TabIndex        =   13
         Top             =   990
         Width           =   7275
         Begin VB.OptionButton OptOrden 
            Caption         =   "Valor Venta"
            Height          =   240
            Index           =   0
            Left            =   1350
            TabIndex        =   15
            Top             =   360
            Value           =   -1  'True
            Width           =   2025
         End
         Begin VB.OptionButton OptOrden 
            Caption         =   "Porcentaje"
            Height          =   240
            Index           =   1
            Left            =   4995
            TabIndex        =   14
            Top             =   360
            Width           =   2025
         End
      End
      Begin VB.CommandButton cmbAyudaAlmacen 
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
         Left            =   6960
         Picture         =   "FrmRptRentabilidad.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   270
         Width           =   390
      End
      Begin VB.CommandButton cmbAyudaProducto 
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
         Left            =   6960
         Picture         =   "FrmRptRentabilidad.frx":038A
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   660
         Width           =   390
      End
      Begin CATControls.CATTextBox txtCod_Almacen 
         Height          =   315
         Left            =   885
         TabIndex        =   5
         Tag             =   "TidAlmacen"
         Top             =   285
         Width           =   1005
         _ExtentX        =   1773
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
         Container       =   "FrmRptRentabilidad.frx":0714
         Estilo          =   1
         EnterTab        =   -1  'True
      End
      Begin CATControls.CATTextBox txtGls_Almacen 
         Height          =   315
         Left            =   1935
         TabIndex        =   6
         Top             =   285
         Width           =   4995
         _ExtentX        =   8811
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
         Container       =   "FrmRptRentabilidad.frx":0730
         Vacio           =   -1  'True
      End
      Begin CATControls.CATTextBox txtCod_Producto 
         Height          =   315
         Left            =   885
         TabIndex        =   7
         Tag             =   "TidMoneda"
         Top             =   660
         Width           =   1005
         _ExtentX        =   1773
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
         Container       =   "FrmRptRentabilidad.frx":074C
         Estilo          =   1
         EnterTab        =   -1  'True
      End
      Begin CATControls.CATTextBox txtGls_Producto 
         Height          =   315
         Left            =   1935
         TabIndex        =   8
         Top             =   660
         Width           =   4995
         _ExtentX        =   8811
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
         Container       =   "FrmRptRentabilidad.frx":0768
         Vacio           =   -1  'True
      End
      Begin MSComCtl2.DTPicker dtpFFinal 
         Height          =   315
         Left            =   3495
         TabIndex        =   9
         Top             =   1845
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
         Format          =   134021121
         CurrentDate     =   38667
      End
      Begin VB.Label lbl_Almacen 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         Caption         =   "Almacén"
         ForeColor       =   &H80000007&
         Height          =   210
         Left            =   165
         TabIndex        =   12
         Top             =   315
         Width           =   630
      End
      Begin VB.Label Label5 
         Appearance      =   0  'Flat
         Caption         =   "Producto"
         ForeColor       =   &H80000007&
         Height          =   240
         Left            =   165
         TabIndex        =   11
         Top             =   690
         Width           =   765
      End
      Begin VB.Label Label2 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         Caption         =   "Hasta"
         ForeColor       =   &H80000008&
         Height          =   210
         Left            =   2775
         TabIndex        =   10
         Top             =   1875
         Width           =   420
      End
   End
   Begin VB.CommandButton cmdsalir 
      Caption         =   "&Salir"
      Height          =   435
      Left            =   3885
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   2385
      Width           =   1230
   End
   Begin VB.CommandButton cmdaceptar 
      Caption         =   "&Aceptar"
      Height          =   435
      Left            =   2610
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   2385
      Width           =   1230
   End
End
Attribute VB_Name = "FrmRptRentabilidad"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmbAyudaAlmacen_Click()
Dim strCondicion As String
    
    mostrarAyuda "ALMACENVTA", txtCod_Almacen, txtGls_Almacen, strCondicion

End Sub

Private Sub cmbAyudaProducto_Click()
    
    mostrarAyuda "PRODUCTOS", txtCod_Producto, txtGls_Producto

End Sub

Private Sub cmdaceptar_Click()
On Error GoTo Err
Dim StrMsgError As String
Dim Ffin        As String
 
    Ffin = Format(dtpFFinal.Value, "yyyy-mm-dd")
    
    If ChkProcesaCostos.Value Then
        ProcesarCostos StrMsgError
        If StrMsgError <> "" Then GoTo Err
    End If
    
    If OptOrden.item(0).Value Then
        mostrarReporte "rptRentabilidadOrdXVVTot.rpt", "parEmpresa|parSucursal|parAlmacen|parProducto|parFecHasta|parMov", glsEmpresa & "|" & glsSucursal & "|" & Trim(txtCod_Almacen.Text) & "|" & Trim(txtCod_Producto.Text) & "|" & Ffin & "|" & 0, Me.Caption, StrMsgError
        If StrMsgError <> "" Then GoTo Err
    Else
        mostrarReporte "rptRentabilidadOrdXPorcen.rpt", "parEmpresa|parSucursal|parAlmacen|parProducto|parFecHasta|parMov", glsEmpresa & "|" & glsSucursal & "|" & Trim(txtCod_Almacen.Text) & "|" & Trim(txtCod_Producto.Text) & "|" & Ffin & "|" & 0, Me.Caption, StrMsgError
        If StrMsgError <> "" Then GoTo Err
    End If
    
    Exit Sub

Err:
    If StrMsgError = "" Then StrMsgError = Err.Description
    MsgBox StrMsgError, vbInformation, App.Title
End Sub

Private Sub cmdsalir_Click()
    
    Unload Me

End Sub

Private Sub Form_Load()
        
    Me.top = 0
    Me.left = 0
    txtGls_Producto.Text = "TODOS LOS PRODUCTOS"
    txtGls_Almacen.Text = "TODOS LOS ALMACENES"
    dtpFFinal.Value = Format(Date, "dd/mm/yyyy")

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
        mostrarAyudaKeyascii KeyAscii, "ALMACENVTA", txtCod_Almacen, txtGls_Almacen, strCondicion
        KeyAscii = 0
    End If
    
End Sub

Private Sub txtCod_Producto_Change()

    If txtCod_Producto.Text <> "" Then
        txtGls_Producto.Text = traerCampo("productos", "GlsProducto", "idProducto", txtCod_Producto.Text, True)
    Else
        txtGls_Producto.Text = "TODOS LOS PRODUCTOS"
    End If
    
End Sub

Private Sub ProcesarCostos(ByRef StrMsgError As String)
On Error GoTo Err
Dim rsvales                    As New ADODB.Recordset
Dim strCodAlmacen              As String
Dim strCodProducto             As String
Dim strTransferencia           As String
Dim indInicio                  As Boolean

Dim dblCostoUnit               As Double
Dim dblPromedio                As Double
Dim dblStockAct                As Double
Dim dblRptVVUnit               As Double
Dim dblRptIGVUnit              As Double
Dim dblRptPVUnit               As Double
Dim dblRptTotVVUnit            As Double
Dim dblRptTotIGVUnit           As Double
Dim dblRptTotPVUnit            As Double
Dim dblCostoTransferencia      As Double
Dim dblTotVVUnitCosto          As Double

Dim dblCostoUnitDol            As Double
Dim dblPromedioDol             As Double
Dim dblStockActDol             As Double
Dim dblRptVVUnitDol            As Double
Dim dblRptIGVUnitDol           As Double
Dim dblRptPVUnitDol            As Double
Dim dblRptTotVVUnitDol         As Double
Dim dblRptTotIGVUnitDol        As Double
Dim dblRptTotPVUnitDol         As Double
Dim dblCostoTransferenciaDol   As Double
Dim dblTotVVUnitCostoDol       As Double

Dim rsvaleSal                  As New ADODB.Recordset
Dim ValeSalida                 As String
Dim StrValeIngreso             As String
Dim strCadCodPro               As String
Dim CIdProducto                As String
Dim strCadAnnio                As String

    If CCodProducto = "CodigoRapido" Then
        If Len(Trim("" & txtCod_Producto.Text)) > 0 Then
            CIdProducto = traerCampo("Productos", "IdProducto", "CodigoRapido", txtCod_Producto.Text, True)
        Else
            CIdProducto = txtCod_Producto.Text
        End If
    Else
        CIdProducto = txtCod_Producto.Text
    End If
             
   If Len(Trim(CIdProducto)) > 0 Then
      strCadCodPro = "And d.idProducto = '" & CIdProducto & "'  "
   Else
      strCadCodPro = ""
   End If
       
   strCadAnnio = "AND (year(v.fechaEmision) = " & Year(dtpFFinal.Value) & ")  "
   
    
   Me.MousePointer = 11

   csql = "SELECT v.idAlmacen, v.tipoVale, ifnull(t.tcventa,v.TipoCambio) as TipoCambio, v.fechaEmision, d.Cantidad, c.indCosto, v.idSucursal, " & _
             "v.idValesCab, d.idProducto, d.item, v.idMoneda, v.idConcepto,ifnull(t.tcventa,0) as tcventa, " & _
             "if(v.idMoneda = 'PEN', d.TotalPVNeto, d.TotalPVNeto * ifnull(t.tcventa,v.TipoCambio)) as TotalPVNeto, " & _
             "if(v.idMoneda = 'PEN', d.TotalIGVNeto, d.TotalIGVNeto * ifnull(t.tcventa,v.TipoCambio)) as TotalIGVNeto, " & _
             "if(v.idMoneda = 'PEN', d.VVUnit, d.VVUnit * ifnull(t.tcventa,v.TipoCambio)) as VVUnit, " & _
             "if(v.idMoneda = 'PEN', d.TotalVVNeto, d.TotalVVNeto * ifnull(t.tcventa,v.TipoCambio)) as TotalVVNeto " & _
             "FROM valescab v inner join valesdet d " & _
             "ON v.idEmpresa = d.idEmpresa " & _
             "AND v.IdSucursal = d.idSucursal " & _
             "AND v.idValesCab = d.idValesCab " & _
             "AND v.tipoVale = d.tipoVale " & _
             "INNER JOIN conceptos c " & _
             "ON v.idConcepto = c.idConcepto " & _
             "INNER JOIN periodosInv p " & _
             "ON v.idPeriodoInv = p.idPeriodoInv " & _
             "AND v.idEmpresa = p.idEmpresa " & _
             "left join  tiposdecambio t " & _
                        "On t.fecha = v.fechaEmision " & _
             "WHERE estValeCab <> 'ANU' " & _
              strCadAnnio & _
             "AND v.idEmpresa = '" & glsEmpresa & "' AND estPeriodoInv = 'ACT' " & strCadCodPro & _
             "ORDER BY v.idAlmacen, d.idProducto,v.fechaEmision,v.tipovale,v.idValesCab"
    
    If rsvales.State = 1 Then rsvales.Close
    rsvales.Open csql, Cn, adOpenStatic, adLockReadOnly
    If Not rsvales.EOF Then
        rsvales.MoveFirst
        Do While Not rsvales.EOF
                strCodAlmacen = Trim(rsvales.Fields("idAlmacen") & "")
            Do While strCodAlmacen = Trim(rsvales.Fields("idAlmacen") & "") And Not rsvales.EOF
                indInicio = True
                
                dblPromedio = 0#
                dblCostoUnit = 0#
                dblStockAct = 0#
                dblTotVVUnitCosto = 0#
                
                dblPromedioDol = 0#
                dblCostoUnitDol = 0#
                dblStockActDol = 0#
                dblTotVVUnitCostoDol = 0#
                
                strCodProducto = Trim(rsvales.Fields("idProducto") & "")
                
                Do While strCodAlmacen = Trim(rsvales.Fields("idAlmacen") & "") And strCodProducto = Trim(rsvales.Fields("idProducto") & "") And Not rsvales.EOF
                'Do While strCodProducto = Trim(rsvales.Fields("idProducto") & "") And Not rsvales.EOF
                    
'                    If rsvales.Fields("idvalescab") = "10030183" And strCodProducto = "11080038" Then
'                        strCodProducto = strCodProducto
'                    End If

                    'If rsvales.Fields("idvalescab") = "12010019" And strCodProducto = "10010541" Then
                    '    strCodProducto = strCodProducto
                    'End If
                    
                    'If rsvales.Fields("idvalescab") = "12010030" And strCodProducto = "10010541" Then
                    '    strCodProducto = strCodProducto
                    'End If
                    
                    If indInicio = True Then
                        If Val(rsvales.Fields("VVUnit") & "") = 0 Then
                            strTransferencia = ""
                            strTransferencia = traerCampo("valestrans", "idAlmacenOrigen", "idAlmacenDestino", rsvales.Fields("idAlmacen"), True, "idValeIngreso = '" & rsvales.Fields("idValesCab") & "' and estValeTrans <> 'ANU' ")
                            If strTransferencia = "" Then
                                dblCostoUnit = 0
                                dblCostoUnitDol = 0
                            End If
                            
                            ValeSalida = traerCampo("valestrans", "idValeSalida", "idValeIngreso", rsvales.Fields("idValesCab"), True, "estValeTrans <> 'ANU' ")
                            csql = "Select VVUnit,IGVUnit,PVUnit,TotalVVNeto,TotalIGVNeto,TotalPVNeto " & _
                                       "From valesdet  WHERE idEmpresa = '" & glsEmpresa & "' " & _
                                       "AND idValesCab = '" & ValeSalida & "' " & _
                                       "AND tipoVale = 'S' " & _
                                       "AND idProducto = '" & rsvales.Fields("idProducto") & "' "
                            If rsvaleSal.State = 1 Then rsvaleSal.Close
                            rsvaleSal.Open csql, Cn, adOpenStatic, adLockReadOnly
                            If Not rsvaleSal.EOF Then
                                dblCostoUnit = rsvaleSal.Fields("VVUnit").Value
                                dblRptTotVVUnit = rsvaleSal.Fields("TotalVVNeto").Value
                                dblRptTotIGVUnit = rsvaleSal.Fields("TotalIGVNeto").Value
                                dblRptTotPVUnit = rsvaleSal.Fields("TotalPVNeto").Value
                            End If
                                
                            If rsvales.Fields("idMoneda") & "" = "PEN" Then
                                csql = "UPDATE valesdet SET " & _
                                        "VVUnit = " & dblCostoUnit & "," & _
                                        "IGVUnit = " & dblCostoUnit * glsIGV & "," & _
                                        "PVUnit = " & dblCostoUnit * (glsIGV + 1) & "," & _
                                        "TotalVVNeto = " & dblRptTotVVUnit & "," & _
                                        "TotalIGVNeto = " & dblRptTotIGVUnit & "," & _
                                        "TotalPVNeto = " & dblRptTotPVUnit & " " & _
                                        "WHERE idEmpresa = '" & glsEmpresa & "' " & _
                                        "AND idSucursal = '" & rsvales.Fields("idSucursal") & "' " & _
                                        "AND idValesCab = '" & rsvales.Fields("idValesCab") & "' " & _
                                        "AND tipoVale = '" & rsvales.Fields("tipoVale") & "' " & _
                                        "AND idProducto = '" & rsvales.Fields("idProducto") & "' " & _
                                        "AND item = " & rsvales.Fields("item") & ""
                                Cn.Execute csql
                                
                                
                                csql = "UPDATE valescab a inner join " & _
                                       "(select idEmpresa,idSucursal,tipovale,idvalescab,sum(TotalVVNeto) as TotalVVNeto,sum(TotalIGVNeto) as TotalIGVNeto,sum(TotalPVNeto) as TotalPVNeto from valesdet " & _
                                       "WHERE idEmpresa = '" & glsEmpresa & "' " & _
                                       "AND idSucursal = '" & rsvales.Fields("idSucursal") & "' " & _
                                       "AND idValesCab = '" & rsvales.Fields("idValesCab") & "' " & _
                                       "AND tipoVale = '" & rsvales.Fields("tipoVale") & "' " & _
                                       "group by idvalescab,tipovale) x " & _
                                       "on a.idvalescab = x.idvalescab " & _
                                       "and a.tipovale = x.tipovale " & _
                                       "and a.idempresa = x.idempresa " & _
                                       "and a.idsucursal = x.idsucursal " & _
                                       "Set a.valorTotal = x.TotalVVNeto, a.igvTotal = x.TotalIGVNeto, a.precioTotal = x.TotalPVNeto " & _
                                       "WHERE a.idEmpresa = '" & glsEmpresa & "' " & _
                                       "AND a.idSucursal = '" & rsvales.Fields("idSucursal") & "' " & _
                                       "AND a.idValesCab = '" & rsvales.Fields("idValesCab") & "' " & _
                                       "AND a.tipoVale = '" & rsvales.Fields("tipoVale") & "' "
                                       
                                Cn.Execute csql
                            Else
                                If Val(rsvales.Fields("TipoCambio")) = 0 Then
                                    csql = "UPDATE valesdet SET " & _
                                            "VVUnit = 0," & _
                                            "IGVUnit = 0," & _
                                            "PVUnit = 0," & _
                                            "TotalVVNeto = 0," & _
                                            "TotalIGVNeto = 0," & _
                                            "TotalPVNeto = 0 " & _
                                            "WHERE idEmpresa = '" & glsEmpresa & "' " & _
                                            "AND idSucursal = '" & rsvales.Fields("idSucursal") & "' " & _
                                            "AND idValesCab = '" & rsvales.Fields("idValesCab") & "' " & _
                                            "AND tipoVale = '" & rsvales.Fields("tipoVale") & "' " & _
                                            "AND idProducto = '" & rsvales.Fields("idProducto") & "' " & _
                                            "AND item = " & rsvales.Fields("item") & ""
                                    Cn.Execute csql
                                    
                                    
                                csql = "UPDATE valescab a inner join " & _
                                       "(select idEmpresa,idSucursal,tipovale,idvalescab,sum(TotalVVNeto) as TotalVVNeto,sum(TotalIGVNeto) as TotalIGVNeto,sum(TotalPVNeto) as TotalPVNeto from valesdet " & _
                                       "WHERE idEmpresa = '" & glsEmpresa & "' " & _
                                       "AND idSucursal = '" & rsvales.Fields("idSucursal") & "' " & _
                                       "AND idValesCab = '" & rsvales.Fields("idValesCab") & "' " & _
                                       "AND tipoVale = '" & rsvales.Fields("tipoVale") & "' " & _
                                       "group by idvalescab,tipovale) x " & _
                                       "on a.idvalescab = x.idvalescab " & _
                                       "and a.tipovale = x.tipovale " & _
                                       "and a.idempresa = x.idempresa " & _
                                       "and a.idsucursal = x.idsucursal " & _
                                       "Set a.valorTotal = x.TotalVVNeto, a.igvTotal = x.TotalIGVNeto, a.precioTotal = x.TotalPVNeto " & _
                                       "WHERE a.idEmpresa = '" & glsEmpresa & "' " & _
                                       "AND a.idSucursal = '" & rsvales.Fields("idSucursal") & "' " & _
                                       "AND a.idValesCab = '" & rsvales.Fields("idValesCab") & "' " & _
                                       "AND a.tipoVale = '" & rsvales.Fields("tipoVale") & "' "
                                       
                                       Cn.Execute csql
                                
                                Else
                                    csql = "UPDATE valesdet SET " & _
                                            "VVUnit = " & Format(dblCostoUnit / Val(rsvales.Fields("TipoCambio")), "0.00") & "," & _
                                            "IGVUnit = " & Format((dblCostoUnit * glsIGV) / Val(rsvales.Fields("TipoCambio")), "0.00") & "," & _
                                            "PVUnit = " & Format((dblCostoUnit * (glsIGV + 1)) / Val(rsvales.Fields("TipoCambio")), "0.00") & "," & _
                                            "TotalVVNeto = " & Format(dblRptTotVVUnit / Val(rsvales.Fields("TipoCambio")), "0.00") & "," & _
                                            "TotalIGVNeto = " & Format(dblRptTotIGVUnit / Val(rsvales.Fields("TipoCambio")), "0.00") & "," & _
                                            "TotalPVNeto = " & Format(dblRptTotPVUnit / Val(rsvales.Fields("TipoCambio")), "0.00") & " " & _
                                            "WHERE idEmpresa = '" & glsEmpresa & "' " & _
                                            "AND idSucursal = '" & rsvales.Fields("idSucursal") & "' " & _
                                            "AND idValesCab = '" & rsvales.Fields("idValesCab") & "' " & _
                                            "AND tipoVale = '" & rsvales.Fields("tipoVale") & "' " & _
                                            "AND idProducto = '" & rsvales.Fields("idProducto") & "' " & _
                                            "AND item = " & rsvales.Fields("item") & ""
                                    Cn.Execute csql
                                    
                                    
                                    csql = "UPDATE valescab a inner join " & _
                                           "(select idEmpresa,idSucursal,tipovale,idvalescab,sum(TotalVVNeto) as TotalVVNeto,sum(TotalIGVNeto) as TotalIGVNeto,sum(TotalPVNeto) as TotalPVNeto from valesdet " & _
                                           "WHERE idEmpresa = '" & glsEmpresa & "' " & _
                                           "AND idSucursal = '" & rsvales.Fields("idSucursal") & "' " & _
                                           "AND idValesCab = '" & rsvales.Fields("idValesCab") & "' " & _
                                           "AND tipoVale = '" & rsvales.Fields("tipoVale") & "' " & _
                                           "group by idvalescab,tipovale) x " & _
                                           "on a.idvalescab = x.idvalescab " & _
                                           "and a.tipovale = x.tipovale " & _
                                           "and a.idempresa = x.idempresa " & _
                                           "and a.idsucursal = x.idsucursal " & _
                                           "Set a.valorTotal = x.TotalVVNeto, a.igvTotal = x.TotalIGVNeto, a.precioTotal = x.TotalPVNeto " & _
                                           "WHERE a.idEmpresa = '" & glsEmpresa & "' " & _
                                           "AND a.idSucursal = '" & rsvales.Fields("idSucursal") & "' " & _
                                           "AND a.idValesCab = '" & rsvales.Fields("idValesCab") & "' " & _
                                           "AND a.tipoVale = '" & rsvales.Fields("tipoVale") & "' "
                                           
                                     Cn.Execute csql
                                       
                                End If
                            End If
                            
                        Else
                            'Antes
                            dblCostoUnit = Val(rsvales.Fields("VVUnit") & "")
                            dblRptTotVVUnit = dblCostoUnit * Val(rsvales.Fields("Cantidad") & "")
                            dblRptTotIGVUnit = Format(dblRptTotVVUnit * glsIGV, "0.00")
                            dblRptTotPVUnit = dblRptTotVVUnit + dblRptTotIGVUnit
                            
                            '11-06-12
'                            If rsvales.Fields("idMoneda") = "PEN" Then
'                                dblCostoUnit = Val(rsvales.Fields("VVUnit") & "")
'                                dblRptTotVVUnit = dblCostoUnit * Val(rsvales.Fields("Cantidad") & "")
'                                dblRptTotIGVUnit = Format(dblRptTotVVUnit * glsIGV, "0.00")
'                                dblRptTotPVUnit = dblRptTotVVUnit + dblRptTotIGVUnit
'                            Else
'                                dblCostoUnit = Val(rsvales.Fields("VVUnit") & "") / Val(rsvales.Fields("TipoCambio"))
'                                dblRptTotVVUnit = dblCostoUnit * Val(rsvales.Fields("Cantidad") & "")
'                                dblRptTotIGVUnit = Format(dblRptTotVVUnit * glsIGV, "0.00")
'                                dblRptTotPVUnit = dblRptTotVVUnit + dblRptTotIGVUnit
'                            End If
                            
                        End If
                        
                        dblPromedio = dblCostoUnit
                        dblStockAct = Val(rsvales.Fields("Cantidad") & "")
  
                        If dblRptTotVVUnit = 0 Then
                            dblRptTotVVUnit = Val(rsvales.Fields("Cantidad") & "") * dblCostoUnit
                        End If
                        
                        If dblRptTotIGVUnit = 0 Then
                            dblRptTotIGVUnit = dblRptTotVVUnit * glsIGV
                        End If

                        If dblRptTotPVUnit = 0 Then
                            dblRptTotPVUnit = dblRptTotVVUnit + dblRptTotIGVUnit
                        End If
                        
                        dblTotVVUnitCosto = Val(dblCostoUnit) * Val(rsvales.Fields("Cantidad") & "")
                        
                        indInicio = False
                        
                    Else
                        If Trim(rsvales.Fields("indCosto") & "") = "S" Then
                            If IIf(dblStockAct < 0, 0, dblStockAct) + Val(rsvales.Fields("Cantidad") & "") > 0# Then
                                strTransferencia = ""
                                strTransferencia = traerCampo("valestrans", "idAlmacenOrigen", "idAlmacenDestino", rsvales.Fields("idAlmacen"), True, "idValeIngreso = '" & rsvales.Fields("idValesCab") & "' and estValeTrans <> 'ANU' ")
                                
                                If strTransferencia <> "" Then
                                    ValeSalida = traerCampo("valestrans", "idValeSalida", "idValeIngreso", rsvales.Fields("idValesCab"), True, "estValeTrans <> 'ANU' ")
                                    csql = "Select VVUnit,IGVUnit,PVUnit,TotalVVNeto,TotalIGVNeto,TotalPVNeto " & _
                                            "From valesdet  WHERE idEmpresa = '" & glsEmpresa & "' " & _
                                            "AND idValesCab = '" & ValeSalida & "' " & _
                                            "AND tipoVale = 'S' " & _
                                            "AND idProducto = '" & rsvales.Fields("idProducto") & "' "
                                    If rsvaleSal.State = 1 Then rsvaleSal.Close
                                    rsvaleSal.Open csql, Cn, adOpenStatic, adLockReadOnly
                                    If Not rsvaleSal.EOF Then
                                        dblRptVVUnit = rsvaleSal.Fields("VVUnit").Value
                                        dblRptIGVUnit = rsvaleSal.Fields("IGVUnit").Value
                                        dblRptPVUnit = rsvaleSal.Fields("PVUnit").Value
                                        dblRptTotVVUnit = rsvaleSal.Fields("TotalVVNeto").Value
                                        dblRptTotIGVUnit = rsvaleSal.Fields("TotalIGVNeto").Value
                                        dblRptTotPVUnit = rsvaleSal.Fields("TotalPVNeto").Value
                                    End If
                                    
                                    If rsvales.Fields("idMoneda") = "PEN" Then
                                        csql = "UPDATE valesdet SET " & _
                                                "VVUnit = " & dblRptVVUnit & "," & _
                                                "IGVUnit = " & dblRptIGVUnit & "," & _
                                                "PVUnit = " & dblRptPVUnit & "," & _
                                                "TotalVVNeto = " & dblRptTotVVUnit & "," & _
                                                "TotalIGVNeto = " & dblRptTotIGVUnit & "," & _
                                                "TotalPVNeto = " & dblRptTotPVUnit & " " & _
                                                "WHERE idEmpresa = '" & glsEmpresa & "' " & _
                                                "AND idValesCab = '" & rsvales.Fields("idValesCab") & "' " & _
                                                "AND tipoVale = '" & rsvales.Fields("tipoVale") & "' " & _
                                                "AND idProducto = '" & rsvales.Fields("idProducto") & "' " & _
                                                "AND item = " & rsvales.Fields("item") & ""
                                                
                                    Cn.Execute csql
                                                
                                    csql = "UPDATE valescab a inner join " & _
                                           "(select idEmpresa,idSucursal,tipovale,idvalescab,sum(TotalVVNeto) as TotalVVNeto,sum(TotalIGVNeto) as TotalIGVNeto,sum(TotalPVNeto) as TotalPVNeto from valesdet " & _
                                           "WHERE idEmpresa = '" & glsEmpresa & "' " & _
                                           "AND idValesCab = '" & rsvales.Fields("idValesCab") & "' " & _
                                           "AND tipoVale = '" & rsvales.Fields("tipoVale") & "' " & _
                                           "group by idvalescab,tipovale) x " & _
                                           "on a.idvalescab = x.idvalescab " & _
                                           "and a.tipovale = x.tipovale " & _
                                           "and a.idempresa = x.idempresa " & _
                                           "Set a.valorTotal = x.TotalVVNeto, a.igvTotal = x.TotalIGVNeto, a.precioTotal = x.TotalPVNeto " & _
                                           "WHERE a.idEmpresa = '" & glsEmpresa & "' " & _
                                           "AND a.idValesCab = '" & rsvales.Fields("idValesCab") & "' " & _
                                           "AND a.tipoVale = '" & rsvales.Fields("tipoVale") & "' "
                                           
                                     Cn.Execute csql
                                     
                                    Else
                                        If Val(rsvales.Fields("TipoCambio")) = 0 Then
                                            csql = "UPDATE valesdet SET " & _
                                                    "VVUnit = 0," & _
                                                    "IGVUnit = 0," & _
                                                    "PVUnit = 0," & _
                                                    "TotalVVNeto = 0," & _
                                                    "TotalIGVNeto = 0," & _
                                                    "TotalPVNeto = 0 " & _
                                                    "WHERE idEmpresa = '" & glsEmpresa & "' " & _
                                                    "AND idValesCab = '" & rsvales.Fields("idValesCab") & "' " & _
                                                    "AND tipoVale = '" & rsvales.Fields("tipoVale") & "' " & _
                                                    "AND idProducto = '" & rsvales.Fields("idProducto") & "' " & _
                                                    "AND item = " & rsvales.Fields("item") & ""
                                                    
                                            Cn.Execute csql
                                                        
                                            csql = "UPDATE valescab a inner join " & _
                                                   "(select idEmpresa,idSucursal,tipovale,idvalescab,sum(TotalVVNeto) as TotalVVNeto,sum(TotalIGVNeto) as TotalIGVNeto,sum(TotalPVNeto) as TotalPVNeto from valesdet " & _
                                                   "WHERE idEmpresa = '" & glsEmpresa & "' " & _
                                                   "AND idValesCab = '" & rsvales.Fields("idValesCab") & "' " & _
                                                   "AND tipoVale = '" & rsvales.Fields("tipoVale") & "' " & _
                                                   "group by idvalescab,tipovale) x " & _
                                                   "on a.idvalescab = x.idvalescab " & _
                                                   "and a.tipovale = x.tipovale " & _
                                                   "and a.idempresa = x.idempresa " & _
                                                   "Set a.valorTotal = x.TotalVVNeto, a.igvTotal = x.TotalIGVNeto, a.precioTotal = x.TotalPVNeto " & _
                                                   "WHERE a.idEmpresa = '" & glsEmpresa & "' " & _
                                                   "AND a.idValesCab = '" & rsvales.Fields("idValesCab") & "' " & _
                                                   "AND a.tipoVale = '" & rsvales.Fields("tipoVale") & "' "
                                                   
                                             Cn.Execute csql
                                        
                                        Else
                                            csql = "UPDATE valesdet SET " & _
                                                    "VVUnit = " & Format(dblRptVVUnit / Val(rsvales.Fields("TipoCambio")), "0.00") & "," & _
                                                    "IGVUnit = " & Format(dblRptIGVUnit / Val(rsvales.Fields("TipoCambio")), "0.00") & "," & _
                                                    "PVUnit = " & Format(dblRptPVUnit / Val(rsvales.Fields("TipoCambio")), "0.00") & "," & _
                                                    "TotalVVNeto = " & Format(dblRptTotVVUnit / Val(rsvales.Fields("TipoCambio")), "0.00") & "," & _
                                                    "TotalIGVNeto = " & Format(dblRptTotIGVUnit / Val(rsvales.Fields("TipoCambio")), "0.00") & "," & _
                                                    "TotalPVNeto = " & Format(dblRptTotPVUnit / Val(rsvales.Fields("TipoCambio")), "0.00") & " " & _
                                                    "WHERE idEmpresa = '" & glsEmpresa & "' " & _
                                                    "AND idValesCab = '" & rsvales.Fields("idValesCab") & "' " & _
                                                    "AND tipoVale = '" & rsvales.Fields("tipoVale") & "' " & _
                                                    "AND idProducto = '" & rsvales.Fields("idProducto") & "' " & _
                                                    "AND item = " & rsvales.Fields("item") & ""
                                                            
                                            Cn.Execute csql
                                                        
                                            csql = "UPDATE valescab a inner join " & _
                                                   "(select idEmpresa,idSucursal,tipovale,idvalescab,sum(TotalVVNeto) as TotalVVNeto,sum(TotalIGVNeto) as TotalIGVNeto,sum(TotalPVNeto) as TotalPVNeto from valesdet " & _
                                                   "WHERE idEmpresa = '" & glsEmpresa & "' " & _
                                                   "AND idValesCab = '" & rsvales.Fields("idValesCab") & "' " & _
                                                   "AND tipoVale = '" & rsvales.Fields("tipoVale") & "' " & _
                                                   "group by idvalescab,tipovale) x " & _
                                                   "on a.idvalescab = x.idvalescab " & _
                                                   "and a.tipovale = x.tipovale " & _
                                                   "and a.idempresa = x.idempresa " & _
                                                   "Set a.valorTotal = x.TotalVVNeto, a.igvTotal = x.TotalIGVNeto, a.precioTotal = x.TotalPVNeto " & _
                                                   "WHERE a.idEmpresa = '" & glsEmpresa & "' " & _
                                                   "AND a.idValesCab = '" & rsvales.Fields("idValesCab") & "' " & _
                                                   "AND a.tipoVale = '" & rsvales.Fields("tipoVale") & "' "
                                                   
                                             Cn.Execute csql
                                        End If
                                    
                                    End If
                                    
                                    dblPromedio = (dblPromedio * dblStockAct + dblRptTotVVUnit) / (dblStockAct + Val(rsvales.Fields("Cantidad") & ""))
                                
                                Else
                                    'Antes
                                    dblRptTotVVUnit = Val(rsvales.Fields("Cantidad") & "") * Val(rsvales.Fields("VVUnit") & "")
                                    If dblRptTotVVUnit = 0 Then
                                        dblRptTotVVUnit = Val(rsvales.Fields("TotalVVNeto") & "")
                                    End If
                                    
                                    If (dblStockAct + Val(rsvales.Fields("Cantidad") & "")) <> 0 Then
                                        'dblPromedio = (dblPromedio * dblStockAct + dblRptTotVVUnit) / (dblStockAct + Val(rsvales.Fields("Cantidad") & ""))
                                        dblPromedio = (dblPromedio * dblStockAct + IIf(rsvales.Fields("TipoVale") = "I", dblRptTotVVUnit, dblRptTotVVUnit * -1)) / (dblStockAct + IIf(rsvales.Fields("TipoVale") = "I", Val(rsvales.Fields("Cantidad") & ""), Val(rsvales.Fields("Cantidad") & "") * -1))
                                    End If
                                    
                                End If
                                
                            End If
                        
                        Else
                             If Val(rsvales.Fields("Cantidad") & "") > 0# Then
                                                                                       
                                strTransferencia = ""
                                strTransferencia = traerCampo("valestrans", "idAlmacenOrigen", "idAlmacenDestino", rsvales.Fields("idAlmacen"), True, "idValeIngreso = '" & rsvales.Fields("idValesCab") & "' and estValeTrans <> 'ANU' ")
                                                                 
                                If Len(Trim(traerCampo("ValesCab", "idValesCab", "idValesCab", rsvales.Fields("idValesCab"), True, "TipoVale = 'S' And idConcepto NOT IN('26') "))) > 0 Then
                                   strTransferencia = ""
                                End If
                                  
                                If strTransferencia <> "" Then
                                    If rsvales.Fields("tipoVale") = "I" Then
                                        ' NO HACE NADA LA SALIDA VA ACTUALIZAR EL INGRESO
                                    End If
                                Else
                                                                                       
                                    dblRptVVUnit = dblPromedio
                                    dblRptIGVUnit = Format(dblRptVVUnit * glsIGV, "0.00")
                                    dblRptPVUnit = dblRptVVUnit + dblRptIGVUnit
                                    dblRptTotVVUnit = Val(Format(dblPromedio * Val(rsvales.Fields("Cantidad") & ""), "0.000"))
                                    dblRptTotIGVUnit = Format(dblRptTotVVUnit * glsIGV, "0.00")
                                    dblRptTotPVUnit = dblRptTotVVUnit + dblRptTotIGVUnit
                                    
                                    If rsvales.Fields("idMoneda") & "" = "PEN" Then
                                        csql = "UPDATE valesdet SET " & _
                                                "VVUnit = " & dblRptVVUnit & "," & _
                                                "IGVUnit = " & dblRptIGVUnit & "," & _
                                                "PVUnit = " & dblRptPVUnit & "," & _
                                                "TotalVVNeto = " & dblRptTotVVUnit & "," & _
                                                "TotalIGVNeto = " & dblRptTotIGVUnit & "," & _
                                                "TotalPVNeto = " & dblRptTotPVUnit & " " & _
                                                "WHERE idEmpresa = '" & glsEmpresa & "' " & _
                                                "AND idSucursal = '" & rsvales.Fields("idSucursal") & "' " & _
                                                "AND idValesCab = '" & rsvales.Fields("idValesCab") & "' " & _
                                                "AND tipoVale = '" & rsvales.Fields("tipoVale") & "' " & _
                                                "AND idProducto = '" & rsvales.Fields("idProducto") & "' " & _
                                                "AND item = " & rsvales.Fields("item") & ""
                                                
                                        Cn.Execute csql
                                                
                                        csql = "UPDATE valescab a inner join " & _
                                               "(select idEmpresa,idSucursal,tipovale,idvalescab,sum(TotalVVNeto) as TotalVVNeto,sum(TotalIGVNeto) as TotalIGVNeto,sum(TotalPVNeto) as TotalPVNeto from valesdet " & _
                                               "WHERE idEmpresa = '" & glsEmpresa & "' " & _
                                               "AND idSucursal = '" & rsvales.Fields("idSucursal") & "' " & _
                                               "AND idValesCab = '" & rsvales.Fields("idValesCab") & "' " & _
                                               "AND tipoVale = '" & rsvales.Fields("tipoVale") & "' " & _
                                               "group by idvalescab,tipovale) x " & _
                                               "on a.idvalescab = x.idvalescab " & _
                                               "and a.tipovale = x.tipovale " & _
                                               "and a.idempresa = x.idempresa " & _
                                               "and a.idsucursal = x.idsucursal " & _
                                               "Set a.valorTotal = x.TotalVVNeto, a.igvTotal = x.TotalIGVNeto, a.precioTotal = x.TotalPVNeto " & _
                                               "WHERE a.idEmpresa = '" & glsEmpresa & "' " & _
                                               "AND a.idSucursal = '" & rsvales.Fields("idSucursal") & "' " & _
                                               "AND a.idValesCab = '" & rsvales.Fields("idValesCab") & "' " & _
                                               "AND a.tipoVale = '" & rsvales.Fields("tipoVale") & "' "
                                               
                                        Cn.Execute csql
                                        
                                    Else
                                        If Val(rsvales.Fields("TipoCambio")) = 0 Then
                                            csql = "UPDATE valesdet SET " & _
                                                    "VVUnit = 0," & _
                                                    "IGVUnit = 0," & _
                                                    "PVUnit = 0," & _
                                                    "TotalVVNeto = 0," & _
                                                    "TotalIGVNeto = 0," & _
                                                    "TotalPVNeto = 0 " & _
                                                    "WHERE idEmpresa = '" & glsEmpresa & "' " & _
                                                    "AND idSucursal = '" & rsvales.Fields("idSucursal") & "' " & _
                                                    "AND idValesCab = '" & rsvales.Fields("idValesCab") & "' " & _
                                                    "AND tipoVale = '" & rsvales.Fields("tipoVale") & "' " & _
                                                    "AND idProducto = '" & rsvales.Fields("idProducto") & "' " & _
                                                    "AND item = " & rsvales.Fields("item") & ""
                                            Cn.Execute csql
                                                    
                                            csql = "UPDATE valescab a inner join " & _
                                                   "(select idEmpresa,idSucursal,tipovale,idvalescab,sum(TotalVVNeto) as TotalVVNeto,sum(TotalIGVNeto) as TotalIGVNeto,sum(TotalPVNeto) as TotalPVNeto from valesdet " & _
                                                   "WHERE idEmpresa = '" & glsEmpresa & "' " & _
                                                   "AND idSucursal = '" & rsvales.Fields("idSucursal") & "' " & _
                                                   "AND idValesCab = '" & rsvales.Fields("idValesCab") & "' " & _
                                                   "AND tipoVale = '" & rsvales.Fields("tipoVale") & "' " & _
                                                   "group by idvalescab,tipovale) x " & _
                                                   "on a.idvalescab = x.idvalescab " & _
                                                   "and a.tipovale = x.tipovale " & _
                                                   "and a.idempresa = x.idempresa " & _
                                                   "and a.idsucursal = x.idsucursal " & _
                                                   "Set a.valorTotal = x.TotalVVNeto, a.igvTotal = x.TotalIGVNeto, a.precioTotal = x.TotalPVNeto " & _
                                                   "WHERE a.idEmpresa = '" & glsEmpresa & "' " & _
                                                   "AND a.idSucursal = '" & rsvales.Fields("idSucursal") & "' " & _
                                                   "AND a.idValesCab = '" & rsvales.Fields("idValesCab") & "' " & _
                                                   "AND a.tipoVale = '" & rsvales.Fields("tipoVale") & "' "
                                                   
                                            Cn.Execute csql
                                        
                                        Else
                                            csql = "UPDATE valesdet SET " & _
                                                    "VVUnit = " & dblRptVVUnit / Val(rsvales.Fields("TipoCambio")) & "," & _
                                                    "IGVUnit = " & dblRptIGVUnit / Val(rsvales.Fields("TipoCambio")) & "," & _
                                                    "PVUnit = " & dblRptPVUnit / Val(rsvales.Fields("TipoCambio")) & "," & _
                                                    "TotalVVNeto = " & dblRptTotVVUnit / Val(rsvales.Fields("TipoCambio")) & "," & _
                                                    "TotalIGVNeto = " & dblRptTotIGVUnit / Val(rsvales.Fields("TipoCambio")) & "," & _
                                                    "TotalPVNeto = " & dblRptTotPVUnit / Val(rsvales.Fields("TipoCambio")) & " " & _
                                                    "WHERE idEmpresa = '" & glsEmpresa & "' " & _
                                                    "AND idSucursal = '" & rsvales.Fields("idSucursal") & "' " & _
                                                    "AND idValesCab = '" & rsvales.Fields("idValesCab") & "' " & _
                                                    "AND tipoVale = '" & rsvales.Fields("tipoVale") & "' " & _
                                                    "AND idProducto = '" & rsvales.Fields("idProducto") & "' " & _
                                                    "AND item = " & rsvales.Fields("item") & ""
                                            Cn.Execute csql
                                                    
                                            csql = "UPDATE valescab a inner join " & _
                                                   "(select idEmpresa,idSucursal,tipovale,idvalescab,sum(TotalVVNeto) as TotalVVNeto,sum(TotalIGVNeto) as TotalIGVNeto,sum(TotalPVNeto) as TotalPVNeto from valesdet " & _
                                                   "WHERE idEmpresa = '" & glsEmpresa & "' " & _
                                                   "AND idSucursal = '" & rsvales.Fields("idSucursal") & "' " & _
                                                   "AND idValesCab = '" & rsvales.Fields("idValesCab") & "' " & _
                                                   "AND tipoVale = '" & rsvales.Fields("tipoVale") & "' " & _
                                                   "group by idvalescab,tipovale) x " & _
                                                   "on a.idvalescab = x.idvalescab " & _
                                                   "and a.tipovale = x.tipovale " & _
                                                   "and a.idempresa = x.idempresa " & _
                                                   "and a.idsucursal = x.idsucursal " & _
                                                   "Set a.valorTotal = x.TotalVVNeto, a.igvTotal = x.TotalIGVNeto, a.precioTotal = x.TotalPVNeto " & _
                                                   "WHERE a.idEmpresa = '" & glsEmpresa & "' " & _
                                                   "AND a.idSucursal = '" & rsvales.Fields("idSucursal") & "' " & _
                                                   "AND a.idValesCab = '" & rsvales.Fields("idValesCab") & "' " & _
                                                   "AND a.tipoVale = '" & rsvales.Fields("tipoVale") & "' "
                                                   
                                            Cn.Execute csql
                                            
                                        End If
                                    End If
                                End If
                                                                
                                If rsvales.Fields("tipoVale") = "S" Then
                                    strTransferencia = ""
                                    StrValeIngreso = ""
                                    strTransferencia = traerCampo("valestrans", "idAlmacenOrigen", "idAlmacenOrigen", rsvales.Fields("idAlmacen"), True, "idValeSalida = '" & rsvales.Fields("idValesCab") & "' and estValeTrans <> 'ANU' ")
                            
                                    If strTransferencia <> "" Then
                                        If rsvales.Fields("tipoVale") = "I" Then
                                
                                        Else
                                            ValeSalida = traerCampo("valestrans", "idValeSalida", "idValeSalida", rsvales.Fields("idValesCab"), True, "estValeTrans <> 'ANU' ")
                                            StrValeIngreso = traerCampo("valestrans", "idValeIngreso", "idValeSalida", rsvales.Fields("idValesCab"), True, "estValeTrans <> 'ANU' ")
                                            
                                            csql = "Select VVUnit,IGVUnit,PVUnit,TotalVVNeto,TotalIGVNeto,TotalPVNeto " & _
                                                    "From valesdet  WHERE idEmpresa = '" & glsEmpresa & "' " & _
                                                    "AND idValesCab = '" & ValeSalida & "' " & _
                                                    "AND tipoVale = 'S' " & _
                                                    "AND idProducto = '" & rsvales.Fields("idProducto") & "' "
                                            If rsvaleSal.State = 1 Then rsvaleSal.Close
                                            rsvaleSal.Open csql, Cn, adOpenStatic, adLockReadOnly
                                            If Not rsvaleSal.EOF Then
                                                dblRptVVUnit = rsvaleSal.Fields("VVUnit").Value
                                                dblRptIGVUnit = rsvaleSal.Fields("IGVUnit").Value
                                                dblRptPVUnit = rsvaleSal.Fields("PVUnit").Value
                                                dblRptTotVVUnit = rsvaleSal.Fields("TotalVVNeto").Value
                                                dblRptTotIGVUnit = rsvaleSal.Fields("TotalIGVNeto").Value
                                                dblRptTotPVUnit = rsvaleSal.Fields("TotalPVNeto").Value
                                            End If
                                            
                                            If rsvales.Fields("idMoneda") = "PEN" Then
                                                csql = "UPDATE valesdet SET " & _
                                                        "VVUnit = " & dblRptVVUnit & "," & _
                                                        "IGVUnit = " & dblRptIGVUnit & "," & _
                                                        "PVUnit = " & dblRptPVUnit & "," & _
                                                        "TotalVVNeto = " & dblRptTotVVUnit & "," & _
                                                        "TotalIGVNeto = " & dblRptTotIGVUnit & "," & _
                                                        "TotalPVNeto = " & dblRptTotPVUnit & " " & _
                                                        "WHERE idEmpresa = '" & glsEmpresa & "' " & _
                                                        "AND idValesCab = '" & StrValeIngreso & "' " & _
                                                        "AND tipoVale = 'I' " & _
                                                        "AND idProducto = '" & rsvales.Fields("idProducto") & "' " & _
                                                        "AND item = " & rsvales.Fields("item") & ""
                                                        
                                                Cn.Execute csql
                                                        
                                                csql = "UPDATE valescab a inner join " & _
                                                       "(select idEmpresa,idSucursal,tipovale,idvalescab,sum(TotalVVNeto) as TotalVVNeto,sum(TotalIGVNeto) as TotalIGVNeto,sum(TotalPVNeto) as TotalPVNeto from valesdet " & _
                                                       "WHERE idEmpresa = '" & glsEmpresa & "' " & _
                                                       "AND idValesCab = '" & StrValeIngreso & "' " & _
                                                       "AND tipoVale = 'I' " & _
                                                       "group by idvalescab,tipovale) x " & _
                                                       "on a.idvalescab = x.idvalescab " & _
                                                       "and a.tipovale = x.tipovale " & _
                                                       "and a.idempresa = x.idempresa " & _
                                                       "Set a.valorTotal = x.TotalVVNeto, a.igvTotal = x.TotalIGVNeto, a.precioTotal = x.TotalPVNeto " & _
                                                       "WHERE a.idEmpresa = '" & glsEmpresa & "' " & _
                                                       "AND a.idValesCab = '" & StrValeIngreso & "' " & _
                                                       "AND a.tipoVale = 'I' "
                                                       
                                                 Cn.Execute csql
                                             
                                            Else
                                                If Val(rsvales.Fields("TipoCambio")) = 0 Then
                                                    csql = "UPDATE valesdet SET " & _
                                                            "VVUnit = 0," & _
                                                            "IGVUnit = 0," & _
                                                            "PVUnit = 0," & _
                                                            "TotalVVNeto = 0," & _
                                                            "TotalIGVNeto = 0," & _
                                                            "TotalPVNeto = 0 " & _
                                                            "WHERE idEmpresa = '" & glsEmpresa & "' " & _
                                                            "AND idValesCab = '" & StrValeIngreso & "' " & _
                                                            "AND tipoVale = 'I' " & _
                                                            "AND idProducto = '" & rsvales.Fields("idProducto") & "' " & _
                                                            "AND item = " & rsvales.Fields("item") & ""
                                                            
                                                    Cn.Execute csql
                                                                
                                                    csql = "UPDATE valescab a inner join " & _
                                                           "(select idEmpresa,idSucursal,tipovale,idvalescab,sum(TotalVVNeto) as TotalVVNeto,sum(TotalIGVNeto) as TotalIGVNeto,sum(TotalPVNeto) as TotalPVNeto from valesdet " & _
                                                           "WHERE idEmpresa = '" & glsEmpresa & "' " & _
                                                           "AND idValesCab = '" & StrValeIngreso & "' " & _
                                                           "AND tipoVale = 'I' " & _
                                                           "group by idvalescab,tipovale) x " & _
                                                           "on a.idvalescab = x.idvalescab " & _
                                                           "and a.tipovale = x.tipovale " & _
                                                           "and a.idempresa = x.idempresa " & _
                                                           "Set a.valorTotal = x.TotalVVNeto, a.igvTotal = x.TotalIGVNeto, a.precioTotal = x.TotalPVNeto " & _
                                                           "WHERE a.idEmpresa = '" & glsEmpresa & "' " & _
                                                           "AND a.idValesCab = '" & StrValeIngreso & "' " & _
                                                           "AND a.tipoVale = 'I' "
                                                           
                                                     Cn.Execute csql
                                                
                                                Else
'                                                    csql = "UPDATE valesdet SET " & _
'                                                            "VVUnit = " & Format(dblRptVVUnit / Val(rsvales.Fields("TipoCambio")), "0.00") & "," & _
'                                                            "IGVUnit = " & Format(dblRptIGVUnit / Val(rsvales.Fields("TipoCambio")), "0.00") & "," & _
'                                                            "PVUnit = " & Format(dblRptPVUnit / Val(rsvales.Fields("TipoCambio")), "0.00") & "," & _
'                                                            "TotalVVNeto = " & Format(dblRptTotVVUnit / Val(rsvales.Fields("TipoCambio")), "0.00") & "," & _
'                                                            "TotalIGVNeto = " & Format(dblRptTotIGVUnit / Val(rsvales.Fields("TipoCambio")), "0.00") & "," & _
'                                                            "TotalPVNeto = " & Format(dblRptTotPVUnit / Val(rsvales.Fields("TipoCambio")), "0.00") & " " & _
'                                                            "WHERE idEmpresa = '" & glsEmpresa & "' " & _
'                                                            "AND idValesCab = '" & StrValeIngreso & "' " & _
'                                                            "AND tipoVale = 'I' " & _
'                                                            "AND idProducto = '" & rsvales.Fields("idProducto") & "' " & _
'                                                            "AND item = " & rsvales.Fields("item") & ""
                                                            
                                                            
                                                            csql = "UPDATE valesdet SET " & _
                                                            "VVUnit = " & dblRptVVUnit & "," & _
                                                            "IGVUnit = " & dblRptIGVUnit & "," & _
                                                            "PVUnit = " & dblRptPVUnit & "," & _
                                                            "TotalVVNeto = " & dblRptTotVVUnit & "," & _
                                                            "TotalIGVNeto = " & dblRptTotIGVUnit & "," & _
                                                            "TotalPVNeto = " & dblRptTotPVUnit & " " & _
                                                            "WHERE idEmpresa = '" & glsEmpresa & "' " & _
                                                            "AND idValesCab = '" & StrValeIngreso & "' " & _
                                                            "AND tipoVale = 'I' " & _
                                                            "AND idProducto = '" & rsvales.Fields("idProducto") & "' " & _
                                                            "AND item = " & rsvales.Fields("item") & ""
                                                                    
                                                    Cn.Execute csql
                                                                
                                                    csql = "UPDATE valescab a inner join " & _
                                                           "(select idEmpresa,idSucursal,tipovale,idvalescab,sum(TotalVVNeto) as TotalVVNeto,sum(TotalIGVNeto) as TotalIGVNeto,sum(TotalPVNeto) as TotalPVNeto from valesdet " & _
                                                           "WHERE idEmpresa = '" & glsEmpresa & "' " & _
                                                           "AND idValesCab = '" & StrValeIngreso & "' " & _
                                                           "AND tipoVale = 'I' " & _
                                                           "group by idvalescab,tipovale) x " & _
                                                           "on a.idvalescab = x.idvalescab " & _
                                                           "and a.tipovale = x.tipovale " & _
                                                           "and a.idempresa = x.idempresa " & _
                                                           "Set a.valorTotal = x.TotalVVNeto, a.igvTotal = x.TotalIGVNeto, a.precioTotal = x.TotalPVNeto " & _
                                                           "WHERE a.idEmpresa = '" & glsEmpresa & "' " & _
                                                           "AND a.idValesCab = '" & StrValeIngreso & "' " & _
                                                           "AND a.tipoVale = 'I' "
                                                           
                                                     Cn.Execute csql
                                                End If
                                            End If
                                        End If
                                    End If
                                 End If
                            Else
                                csql = "UPDATE valesdet SET " & _
                                        "VVUnit = 0," & _
                                        "IGVUnit = 0," & _
                                        "PVUnit = 0," & _
                                        "TotalVVNeto = 0," & _
                                        "TotalIGVNeto = 0," & _
                                        "TotalPVNeto = 0 " & _
                                        "WHERE idEmpresa = '" & glsEmpresa & "' " & _
                                        "AND idSucursal = '" & rsvales.Fields("idSucursal") & "' " & _
                                        "AND idValesCab = '" & rsvales.Fields("idValesCab") & "' " & _
                                        "AND tipoVale = '" & rsvales.Fields("tipoVale") & "' " & _
                                        "AND idProducto = '" & rsvales.Fields("idProducto") & "' " & _
                                        "AND item = " & rsvales.Fields("item") & ""
                                Cn.Execute csql
                                        
                                csql = "UPDATE valescab a inner join " & _
                                       "(select idEmpresa,idSucursal,tipovale,idvalescab,sum(TotalVVNeto) as TotalVVNeto,sum(TotalIGVNeto) as TotalIGVNeto,sum(TotalPVNeto) as TotalPVNeto from valesdet " & _
                                       "WHERE idEmpresa = '" & glsEmpresa & "' " & _
                                       "AND idSucursal = '" & rsvales.Fields("idSucursal") & "' " & _
                                       "AND idValesCab = '" & rsvales.Fields("idValesCab") & "' " & _
                                       "AND tipoVale = '" & rsvales.Fields("tipoVale") & "' " & _
                                       "group by idvalescab,tipovale) x " & _
                                       "on a.idvalescab = x.idvalescab " & _
                                       "and a.tipovale = x.tipovale " & _
                                       "and a.idempresa = x.idempresa " & _
                                       "and a.idsucursal = x.idsucursal " & _
                                       "Set a.valorTotal = x.TotalVVNeto, a.igvTotal = x.TotalIGVNeto, a.precioTotal = x.TotalPVNeto " & _
                                       "WHERE a.idEmpresa = '" & glsEmpresa & "' " & _
                                       "AND a.idSucursal = '" & rsvales.Fields("idSucursal") & "' " & _
                                       "AND a.idValesCab = '" & rsvales.Fields("idValesCab") & "' " & _
                                       "AND a.tipoVale = '" & rsvales.Fields("tipoVale") & "' "
                                       
                                Cn.Execute csql
                            End If
                        End If

                        If Trim(rsvales.Fields("tipoVale") & "") = "I" Then
                            dblStockAct = dblStockAct + Val(rsvales.Fields("Cantidad") & "")
                        Else
                            dblStockAct = dblStockAct - Val(rsvales.Fields("Cantidad") & "")
                        End If
                        
                    End If
                    
                    rsvales.MoveNext
                    
                    If rsvales.EOF Then Exit Do
                    If strCodProducto <> Trim(rsvales.Fields("idProducto") & "") And strCodAlmacen = Trim(rsvales.Fields("idAlmacen") & "") Then Exit Do
                Loop
                
                If rsvales.EOF Then Exit Do
                If strCodAlmacen <> Trim(rsvales.Fields("idAlmacen") & "") Then Exit Do

            Loop
        Loop
    End If
    
    If rsvales.State = 1 Then rsvales.Close: Set rsvales = Nothing
    
    '---------------------------------------------------------
    csql = "UPDATE productosalmacen " & _
           "Set CantidadStock = 0"
    Cn.Execute csql
    
    csql = "UPDATE productosalmacen a, vw_temp_stock v " & _
           "Set A.CantidadStock = v.STOCK " & _
           "Where A.idEmpresa = v.idEmpresa " & _
           "AND a.idSucursal = v.idSucursal " & _
           "AND a.idAlmacen = v.idAlmacen " & _
           "AND a.idProducto = v.idProducto " & _
           "AND a.idUMCompra = v.idUM"
    Cn.Execute csql
    '---------------------------------------------------------

    Me.MousePointer = 1
    
Exit Sub
Err:
    If rsvales.State = 1 Then rsvales.Close: Set rsvales = Nothing
    
    Me.MousePointer = 1
    If StrMsgError = "" Then StrMsgError = Err.Description
End Sub
