VERSION 5.00
Object = "{F41D1D30-7878-4923-8CB3-6CCACDC9C9DE}#1.0#0"; "CATControls.ocx"
Begin VB.Form frmProcesaCostosVales 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Procesar Costos"
   ClientHeight    =   4710
   ClientLeft      =   7155
   ClientTop       =   3570
   ClientWidth     =   6315
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4710
   ScaleWidth      =   6315
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame3 
      Appearance      =   0  'Flat
      ForeColor       =   &H80000008&
      Height          =   780
      Left            =   45
      TabIndex        =   13
      Top             =   2970
      Width           =   6180
      Begin VB.CommandButton cmbAyudaProducto 
         Height          =   315
         Left            =   5700
         Picture         =   "frmProcesaCostosVales.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   14
         Top             =   225
         Width           =   390
      End
      Begin CATControls.CATTextBox txtCod_Producto 
         Height          =   315
         Left            =   735
         TabIndex        =   15
         Tag             =   "TidMoneda"
         Top             =   225
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
         Container       =   "frmProcesaCostosVales.frx":038A
         Estilo          =   1
         EnterTab        =   -1  'True
      End
      Begin CATControls.CATTextBox txtGls_Producto 
         Height          =   315
         Left            =   1665
         TabIndex        =   16
         Top             =   225
         Width           =   4005
         _ExtentX        =   7064
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
         Container       =   "frmProcesaCostosVales.frx":03A6
         Vacio           =   -1  'True
      End
      Begin VB.Label Label2 
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
         Left            =   45
         TabIndex        =   17
         Top             =   255
         Width           =   645
      End
   End
   Begin VB.Frame Frame2 
      Appearance      =   0  'Flat
      ForeColor       =   &H80000008&
      Height          =   780
      Left            =   45
      TabIndex        =   6
      Top             =   3780
      Width           =   6180
      Begin VB.CommandButton cmdSalir 
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
         Height          =   465
         Left            =   3105
         TabIndex        =   8
         Top             =   180
         Width           =   1635
      End
      Begin VB.CommandButton cmdprocesar 
         Caption         =   "&Procesar"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   465
         Left            =   1350
         TabIndex        =   7
         Top             =   180
         Width           =   1635
      End
   End
   Begin VB.Frame Frame1 
      Appearance      =   0  'Flat
      ForeColor       =   &H80000008&
      Height          =   825
      Left            =   45
      TabIndex        =   0
      Top             =   2115
      Width           =   6180
      Begin VB.ComboBox cbxAno 
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
         ItemData        =   "frmProcesaCostosVales.frx":03C2
         Left            =   2025
         List            =   "frmProcesaCostosVales.frx":03E1
         Style           =   2  'Dropdown List
         TabIndex        =   11
         Top             =   270
         Width           =   2145
      End
      Begin VB.ComboBox cbx_Mes 
         Height          =   315
         ItemData        =   "frmProcesaCostosVales.frx":041C
         Left            =   1485
         List            =   "frmProcesaCostosVales.frx":041E
         Style           =   2  'Dropdown List
         TabIndex        =   10
         Top             =   810
         Width           =   2115
      End
      Begin VB.Label Label1 
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
         Height          =   210
         Index           =   0
         Left            =   1395
         TabIndex        =   12
         Top             =   315
         Width           =   300
      End
      Begin VB.Label Label1 
         Caption         =   "Mes"
         Height          =   285
         Index           =   1
         Left            =   585
         TabIndex        =   9
         Top             =   855
         Width           =   555
      End
   End
   Begin VB.Label Label7 
      Alignment       =   2  'Center
      BackColor       =   &H000000FF&
      Caption         =   "* No Olvidar correr el proceso hasta el mes Actual*"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   900
      TabIndex        =   5
      Top             =   1485
      Width           =   4605
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Center
      BackColor       =   &H000000FF&
      Caption         =   "TODOS los usuarios estén fuera del sistema"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   900
      TabIndex        =   4
      Top             =   945
      Width           =   4605
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      BackColor       =   &H000000FF&
      Caption         =   "No ejecutar este proceso, si no está seguro que"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   945
      TabIndex        =   3
      Top             =   675
      Width           =   4605
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      BackColor       =   &H000000FF&
      Caption         =   "ADVERTENCIA"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   90
      TabIndex        =   2
      Top             =   225
      Width           =   6135
   End
   Begin VB.Label Label3 
      BackColor       =   &H000000FF&
      Height          =   1995
      Left            =   45
      TabIndex        =   1
      Top             =   45
      Width           =   6180
   End
End
Attribute VB_Name = "frmProcesaCostosVales"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim CCodProducto                    As String
Option Explicit

Private Sub cmbAyudaProducto_Click()
    mostrarAyuda "PRODUCTOS", txtCod_Producto, txtGls_Producto
End Sub

Private Sub cmdprocesar_Click()
On Error GoTo Err
Dim StrMsgError As String

    If MsgBox("Está seguro(a) de realizar el proceso ? ", vbQuestion + vbYesNo, App.Title) = vbYes Then
        ProcesarCostos StrMsgError
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
Dim strAnnio                   As String
Dim CArray()                   As String
Dim intAño                     As Integer
Dim RsProductosError            As New ADODB.Recordset
Dim IndSalir                    As Boolean
    
    IndSalir = False
    
    RsProductosError.Fields.Append "IdProducto", adVarChar, 8, adFldRowID
    RsProductosError.Fields.Append "GlsProducto", adVarChar, 800, adFldIsNullable
    RsProductosError.Open
    
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
      
   'Todos
   If Trim(cbxAno.Text) = "Periodo Actual" Then
        'TOMASINI 06/06/2013
        'strCadAnnio = ""
        'ReDim cArray(2)
        'traerCampos "PeriodosInv", "Year(FecInicio) As FecIni,IfNull(Year(FecFin),'') As FecFin", "estPeriodoInv", "ACT", 2, cArray(), True, "idSucursal = '" & glsSucursal & "'"
        'strCadAnnio = "AND (year(v.fechaEmision) In('" & cArray(0) & "','" & IIf(cArray(1) = "", Year(getFechaSistema), cArray(1)) & "'))"
   
        'jach 07/06/2013
        ReDim CArray(1)
        strCadAnnio = ""
        traerCampos "PeriodosInv", "Year(FecInicio) As FecIni,IfNull(Year(FecFin),'') As FecFin", "estPeriodoInv", "ACT", 2, CArray(), True, "idSucursal = '" & glsSucursal & "'"
        For intAño = Val(CArray(0)) To Year(getFechaSistema)
            strCadAnnio = strCadAnnio & "'" & intAño & "', "
        Next intAño
        If Len(strCadAnnio) > 0 Then
            strCadAnnio = left(strCadAnnio, Len(strCadAnnio) - 2)
        End If
        
        strCadAnnio = "AND (year(v.fechaEmision) In(" & strCadAnnio & "))"
   Else
        strCadAnnio = "AND (year(v.fechaEmision) = " & Val(cbxAno.Text) & ") "
   End If
    
   Me.MousePointer = 11

   csql = "SELECT v.idAlmacen, v.tipoVale, isnull(v.TipoCambio,t.tcventa) as TipoCambio, v.fechaEmision, Sum(d.Cantidad) Cantidad, c.indCosto, v.idSucursal, " & _
             "v.idValesCab, d.idProducto, d.item, v.idMoneda, v.idConcepto,isnull(V.TipoCambio,0) as tcventa, " & _
             "Sum(CASE WHEN v.idMoneda = 'PEN' THEN d.TotalPVNeto ELSE d.TotalPVNeto * isnull(v.TipoCambio,t.tcventa) END) as TotalPVNeto, " & _
             "Sum(CASE WHEN v.idMoneda = 'PEN' THEN d.TotalIGVNeto ELSE d.TotalIGVNeto * isnull(v.TipoCambio,t.tcventa) END) as TotalIGVNeto, " & _
             "CASE WHEN v.idMoneda = 'PEN' THEN d.VVUnit ELSE d.VVUnit * isnull(v.TipoCambio,t.tcventa) END as VVUnit, " & _
             "Sum(CASE WHEN v.idMoneda = 'PEN' THEN d.TotalVVNeto ELSE d.TotalVVNeto * isnull(v.TipoCambio,t.tcventa) END) as TotalVVNeto " & _
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
             "AND v.idEmpresa = '" & glsEmpresa & "' " & strCadCodPro & _
             "Group By v.idAlmacen, v.tipoVale,v.idValesCab, d.idProducto, d.item, v.idMoneda, v.idConcepto,v.fechaEmision ,c.indCosto, v.idSucursal,v.TipoCambio,t.tcventa,d.VVUnit ORDER BY v.idAlmacen, d.idProducto,v.fechaEmision,v.tipovale,v.idValesCab"
    
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
                    
                    If rsvales.Fields("idvalescab") = "16050119" Then
                        strCodProducto = strCodProducto
                    End If
                    
                    'JACH 06/06/20134
                    'solo lo activo cuando ratreo el costo de un producto, coloco el numero de vale
                    'If rsvales("idValesCab") = "15070245" Then MsgBox ""
                    
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
                                       "AND idProducto = '" & rsvales.Fields("idProducto") & "' And Item = " & rsvales.Fields("Item") & ""
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
                                        ""
                                        'AND item = " & rsvales.Fields("item") & "
                                Cn.Execute csql
                                
                                
                                csql = "UPDATE a Set a.valorTotal = x.TotalVVNeto, a.igvTotal = x.TotalIGVNeto, a.precioTotal = x.TotalPVNeto from valescab a inner join " & _
                                       "(select idEmpresa,idSucursal,tipovale,idvalescab,sum(TotalVVNeto) as TotalVVNeto,sum(TotalIGVNeto) as TotalIGVNeto,sum(TotalPVNeto) as TotalPVNeto from valesdet " & _
                                       "WHERE idEmpresa = '" & glsEmpresa & "' " & _
                                       "AND idSucursal = '" & rsvales.Fields("idSucursal") & "' " & _
                                       "AND idValesCab = '" & rsvales.Fields("idValesCab") & "' " & _
                                       "AND tipoVale = '" & rsvales.Fields("tipoVale") & "' " & _
                                       "group by idEmpresa,idSucursal,idvalescab,tipovale) x " & _
                                       "on a.idvalescab = x.idvalescab " & _
                                       "and a.tipovale = x.tipovale " & _
                                       "and a.idempresa = x.idempresa " & _
                                       "and a.idsucursal = x.idsucursal " & _
                                       " " & _
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
                                            ""
                                            'AND item = " & rsvales.Fields("item") & "
                                    Cn.Execute csql
                                    
                                    
                                csql = "UPDATE a Set a.valorTotal = x.TotalVVNeto, a.igvTotal = x.TotalIGVNeto, a.precioTotal = x.TotalPVNeto from valescab a inner join " & _
                                       "(select idEmpresa,idSucursal,tipovale,idvalescab,sum(TotalVVNeto) as TotalVVNeto,sum(TotalIGVNeto) as TotalIGVNeto,sum(TotalPVNeto) as TotalPVNeto from valesdet " & _
                                       "WHERE idEmpresa = '" & glsEmpresa & "' " & _
                                       "AND idSucursal = '" & rsvales.Fields("idSucursal") & "' " & _
                                       "AND idValesCab = '" & rsvales.Fields("idValesCab") & "' " & _
                                       "AND tipoVale = '" & rsvales.Fields("tipoVale") & "' " & _
                                       "group by idEmpresa,idSucursal,idvalescab,tipovale) x " & _
                                       "on a.idvalescab = x.idvalescab " & _
                                       "and a.tipovale = x.tipovale " & _
                                       "and a.idempresa = x.idempresa " & _
                                       "and a.idsucursal = x.idsucursal " & _
                                       " " & _
                                       "WHERE a.idEmpresa = '" & glsEmpresa & "' " & _
                                       "AND a.idSucursal = '" & rsvales.Fields("idSucursal") & "' " & _
                                       "AND a.idValesCab = '" & rsvales.Fields("idValesCab") & "' " & _
                                       "AND a.tipoVale = '" & rsvales.Fields("tipoVale") & "' "
                                       
                                       Cn.Execute csql
                                
                                Else
                                    'JACH
                                    '09/02/2016
                                    csql = "UPDATE valesdet SET " & _
                                            "VVUnit = " & dblCostoUnit / Val(rsvales.Fields("TipoCambio")) & "," & _
                                            "IGVUnit = " & (dblCostoUnit * glsIGV) / Val(rsvales.Fields("TipoCambio")) & "," & _
                                            "PVUnit = " & (dblCostoUnit * (glsIGV + 1)) / Val(rsvales.Fields("TipoCambio")) & "," & _
                                            "TotalVVNeto = " & dblRptTotVVUnit / Val(rsvales.Fields("TipoCambio")) & "," & _
                                            "TotalIGVNeto = " & dblRptTotIGVUnit / Val(rsvales.Fields("TipoCambio")) & "," & _
                                            "TotalPVNeto = " & dblRptTotPVUnit / Val(rsvales.Fields("TipoCambio")) & " " & _
                                            "WHERE idEmpresa = '" & glsEmpresa & "' " & _
                                            "AND idSucursal = '" & rsvales.Fields("idSucursal") & "' " & _
                                            "AND idValesCab = '" & rsvales.Fields("idValesCab") & "' " & _
                                            "AND tipoVale = '" & rsvales.Fields("tipoVale") & "' " & _
                                            "AND idProducto = '" & rsvales.Fields("idProducto") & "' " & _
                                            ""
                                            'AND item = " & rsvales.Fields("item") & "
                                    'csql = "UPDATE valesdet SET " & _
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
                                            ""
                                            'AND item = " & rsvales.Fields("item") & "
                                    Cn.Execute csql
                                    
                                    
                                    csql = "UPDATE a Set a.valorTotal = x.TotalVVNeto, a.igvTotal = x.TotalIGVNeto, a.precioTotal = x.TotalPVNeto from valescab a inner join " & _
                                           "(select idEmpresa,idSucursal,tipovale,idvalescab,sum(TotalVVNeto) as TotalVVNeto,sum(TotalIGVNeto) as TotalIGVNeto,sum(TotalPVNeto) as TotalPVNeto from valesdet " & _
                                           "WHERE idEmpresa = '" & glsEmpresa & "' " & _
                                           "AND idSucursal = '" & rsvales.Fields("idSucursal") & "' " & _
                                           "AND idValesCab = '" & rsvales.Fields("idValesCab") & "' " & _
                                           "AND tipoVale = '" & rsvales.Fields("tipoVale") & "' " & _
                                           "group by idEmpresa,idSucursal,idvalescab,tipovale) x " & _
                                           "on a.idvalescab = x.idvalescab " & _
                                           "and a.tipovale = x.tipovale " & _
                                           "and a.idempresa = x.idempresa " & _
                                           "and a.idsucursal = x.idsucursal " & _
                                           " " & _
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
                            dblRptTotIGVUnit = dblRptTotVVUnit * glsIGV
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
                                            "AND idProducto = '" & rsvales.Fields("idProducto") & "' And Item = " & rsvales.Fields("Item") & ""
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
                                                ""
                                                'AND item = " & rsvales.Fields("item") & "
                                    Cn.Execute csql
                                                
                                    csql = "UPDATE a Set a.valorTotal = x.TotalVVNeto, a.igvTotal = x.TotalIGVNeto, a.precioTotal = x.TotalPVNeto from valescab a inner join " & _
                                           "(select idEmpresa,idSucursal,tipovale,idvalescab,sum(TotalVVNeto) as TotalVVNeto,sum(TotalIGVNeto) as TotalIGVNeto,sum(TotalPVNeto) as TotalPVNeto from valesdet " & _
                                           "WHERE idEmpresa = '" & glsEmpresa & "' " & _
                                           "AND idValesCab = '" & rsvales.Fields("idValesCab") & "' " & _
                                           "AND tipoVale = '" & rsvales.Fields("tipoVale") & "' " & _
                                           "group by idEmpresa,idSucursal,idvalescab,tipovale) x " & _
                                           "on a.idvalescab = x.idvalescab " & _
                                           "and a.tipovale = x.tipovale " & _
                                           "and a.idempresa = x.idempresa " & _
                                           " " & _
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
                                                    ""
                                                    'AND item = " & rsvales.Fields("item") & "
                                            Cn.Execute csql
                                                        
                                            csql = "UPDATE a Set a.valorTotal = x.TotalVVNeto, a.igvTotal = x.TotalIGVNeto, a.precioTotal = x.TotalPVNeto from valescab a inner join " & _
                                                   "(select idEmpresa,idSucursal,tipovale,idvalescab,sum(TotalVVNeto) as TotalVVNeto,sum(TotalIGVNeto) as TotalIGVNeto,sum(TotalPVNeto) as TotalPVNeto from valesdet " & _
                                                   "WHERE idEmpresa = '" & glsEmpresa & "' " & _
                                                   "AND idValesCab = '" & rsvales.Fields("idValesCab") & "' " & _
                                                   "AND tipoVale = '" & rsvales.Fields("tipoVale") & "' " & _
                                                   "group by idEmpresa,idSucursal,idvalescab,tipovale) x " & _
                                                   "on a.idvalescab = x.idvalescab " & _
                                                   "and a.tipovale = x.tipovale " & _
                                                   "and a.idempresa = x.idempresa " & _
                                                   " " & _
                                                   "WHERE a.idEmpresa = '" & glsEmpresa & "' " & _
                                                   "AND a.idValesCab = '" & rsvales.Fields("idValesCab") & "' " & _
                                                   "AND a.tipoVale = '" & rsvales.Fields("tipoVale") & "' "
                                                   
                                             Cn.Execute csql
                                        
                                        Else
                                            'JACH 06/06/2013
                                            'lo coloque en comentario por que no deberia dividir entre el tc
                                            'csql = "UPDATE valesdet SET " & _
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
                                                    ""
                                            'AND item = " & rsvales.Fields("item") & "
                                            Cn.Execute csql
                                                        
                                            csql = "UPDATE a Set a.valorTotal = x.TotalVVNeto, a.igvTotal = x.TotalIGVNeto, a.precioTotal = x.TotalPVNeto from valescab a inner join " & _
                                                   "(select idEmpresa,idSucursal,tipovale,idvalescab,sum(TotalVVNeto) as TotalVVNeto,sum(TotalIGVNeto) as TotalIGVNeto,sum(TotalPVNeto) as TotalPVNeto from valesdet " & _
                                                   "WHERE idEmpresa = '" & glsEmpresa & "' " & _
                                                   "AND idValesCab = '" & rsvales.Fields("idValesCab") & "' " & _
                                                   "AND tipoVale = '" & rsvales.Fields("tipoVale") & "' " & _
                                                   "group by idEmpresa,idSucursal,idvalescab,tipovale) x " & _
                                                   "on a.idvalescab = x.idvalescab " & _
                                                   "and a.tipovale = x.tipovale " & _
                                                   "and a.idempresa = x.idempresa " & _
                                                   " " & _
                                                   "WHERE a.idEmpresa = '" & glsEmpresa & "' " & _
                                                   "AND a.idValesCab = '" & rsvales.Fields("idValesCab") & "' " & _
                                                   "AND a.tipoVale = '" & rsvales.Fields("tipoVale") & "' "
                                                   
                                             Cn.Execute csql
                                            
                                            'JACH 06/06/2013
                                            'si no convierte los montos a dolares el promedio se distorciona
                                            dblRptVVUnit = dblRptVVUnit * Val(rsvales.Fields("TipoCambio"))
                                            dblRptIGVUnit = dblRptIGVUnit * Val(rsvales.Fields("TipoCambio"))
                                            dblRptPVUnit = dblRptPVUnit * Val(rsvales.Fields("TipoCambio"))
                                            dblRptTotVVUnit = dblRptTotVVUnit * Val(rsvales.Fields("TipoCambio"))
                                            dblRptTotIGVUnit = dblRptTotIGVUnit * Val(rsvales.Fields("TipoCambio"))
                                            dblRptTotPVUnit = dblRptTotPVUnit * Val(rsvales.Fields("TipoCambio"))
                                        End If
                                    
                                    End If
                                    
                                    If dblStockAct + Val(rsvales.Fields("Cantidad")) <> 0 Then
                                        dblPromedio = (dblPromedio * dblStockAct + dblRptTotVVUnit) / (dblStockAct + Val(rsvales.Fields("Cantidad") & ""))
                                    End If
                                
                                Else
                                    'Antes
                                    dblRptTotVVUnit = Val(rsvales.Fields("Cantidad") & "") * Val(rsvales.Fields("VVUnit") & "")
                                    If dblRptTotVVUnit = 0 Then
                                        dblRptTotVVUnit = Val(rsvales.Fields("TotalVVNeto") & "")
                                    End If
                                    
                                    'If (dblStockAct + Val(rsvales.Fields("Cantidad") & "")) <> 0 Then
                                    If (dblStockAct + IIf(rsvales.Fields("TipoVale") = "I", Val(rsvales.Fields("Cantidad") & ""), Val(rsvales.Fields("Cantidad") & "") * -1)) > 0 Then
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
                                    'dblRptIGVUnit = Format(dblRptVVUnit * glsIGV, "0.00")
                                    dblRptIGVUnit = dblRptVVUnit * glsIGV
                                    dblRptPVUnit = dblRptVVUnit + dblRptIGVUnit
                                    dblRptTotVVUnit = Val(dblPromedio * Val(rsvales.Fields("Cantidad") & ""))
                                    'dblRptTotIGVUnit = Format(dblRptTotVVUnit * glsIGV, "0.00")
                                    dblRptTotIGVUnit = dblRptTotVVUnit * glsIGV
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
                                                
                                        csql = "UPDATE a Set a.valorTotal = x.TotalVVNeto, a.igvTotal = x.TotalIGVNeto, a.precioTotal = x.TotalPVNeto from valescab a inner join " & _
                                               "(select idEmpresa,idSucursal,tipovale,idvalescab,sum(TotalVVNeto) as TotalVVNeto,sum(TotalIGVNeto) as TotalIGVNeto,sum(TotalPVNeto) as TotalPVNeto from valesdet " & _
                                               "WHERE idEmpresa = '" & glsEmpresa & "' " & _
                                               "AND idSucursal = '" & rsvales.Fields("idSucursal") & "' " & _
                                               "AND idValesCab = '" & rsvales.Fields("idValesCab") & "' " & _
                                               "AND tipoVale = '" & rsvales.Fields("tipoVale") & "' " & _
                                               "group by idEmpresa,idSucursal,idvalescab,tipovale) x " & _
                                               "on a.idvalescab = x.idvalescab " & _
                                               "and a.tipovale = x.tipovale " & _
                                               "and a.idempresa = x.idempresa " & _
                                               "and a.idsucursal = x.idsucursal " & _
                                               " " & _
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
                                                    ""
                                                    'AND item = " & rsvales.Fields("item") & "
                                            Cn.Execute csql
                                                    
                                            csql = "UPDATE a Set a.valorTotal = x.TotalVVNeto, a.igvTotal = x.TotalIGVNeto, a.precioTotal = x.TotalPVNeto from valescab a inner join " & _
                                                   "(select idEmpresa,idSucursal,tipovale,idvalescab,sum(TotalVVNeto) as TotalVVNeto,sum(TotalIGVNeto) as TotalIGVNeto,sum(TotalPVNeto) as TotalPVNeto from valesdet " & _
                                                   "WHERE idEmpresa = '" & glsEmpresa & "' " & _
                                                   "AND idSucursal = '" & rsvales.Fields("idSucursal") & "' " & _
                                                   "AND idValesCab = '" & rsvales.Fields("idValesCab") & "' " & _
                                                   "AND tipoVale = '" & rsvales.Fields("tipoVale") & "' " & _
                                                   "group by idEmpresa,idSucursal,idvalescab,tipovale) x " & _
                                                   "on a.idvalescab = x.idvalescab " & _
                                                   "and a.tipovale = x.tipovale " & _
                                                   "and a.idempresa = x.idempresa " & _
                                                   "and a.idsucursal = x.idsucursal " & _
                                                   " " & _
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
                                                    ""
                                                    'AND item = " & rsvales.Fields("item") & "
                                            Cn.Execute csql
                                                    
                                            csql = "UPDATE a Set a.valorTotal = x.TotalVVNeto, a.igvTotal = x.TotalIGVNeto, a.precioTotal = x.TotalPVNeto from valescab a inner join " & _
                                                   "(select idEmpresa,idSucursal,tipovale,idvalescab,sum(TotalVVNeto) as TotalVVNeto,sum(TotalIGVNeto) as TotalIGVNeto,sum(TotalPVNeto) as TotalPVNeto from valesdet " & _
                                                   "WHERE idEmpresa = '" & glsEmpresa & "' " & _
                                                   "AND idSucursal = '" & rsvales.Fields("idSucursal") & "' " & _
                                                   "AND idValesCab = '" & rsvales.Fields("idValesCab") & "' " & _
                                                   "AND tipoVale = '" & rsvales.Fields("tipoVale") & "' " & _
                                                   "group by idEmpresa,idSucursal,idvalescab,tipovale) x " & _
                                                   "on a.idvalescab = x.idvalescab " & _
                                                   "and a.tipovale = x.tipovale " & _
                                                   "and a.idempresa = x.idempresa " & _
                                                   "and a.idsucursal = x.idsucursal " & _
                                                   " " & _
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
                                                        ""
                                                        'AND item = " & rsvales.Fields("item") & "
                                                Cn.Execute csql
                                                        
                                                csql = "UPDATE a Set a.valorTotal = x.TotalVVNeto, a.igvTotal = x.TotalIGVNeto, a.precioTotal = x.TotalPVNeto from valescab a inner join " & _
                                                       "(select idEmpresa,idSucursal,tipovale,idvalescab,sum(TotalVVNeto) as TotalVVNeto,sum(TotalIGVNeto) as TotalIGVNeto,sum(TotalPVNeto) as TotalPVNeto from valesdet " & _
                                                       "WHERE idEmpresa = '" & glsEmpresa & "' " & _
                                                       "AND idValesCab = '" & StrValeIngreso & "' " & _
                                                       "AND tipoVale = 'I' " & _
                                                       "group by idEmpresa,idSucursal,idvalescab,tipovale) x " & _
                                                       "on a.idvalescab = x.idvalescab " & _
                                                       "and a.tipovale = x.tipovale " & _
                                                       "and a.idempresa = x.idempresa " & _
                                                       " " & _
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
                                                            ""
                                                            'AND item = " & rsvales.Fields("item") & "
                                                    Cn.Execute csql
                                                                
                                                    csql = "UPDATE a Set a.valorTotal = x.TotalVVNeto, a.igvTotal = x.TotalIGVNeto, a.precioTotal = x.TotalPVNeto from valescab a inner join " & _
                                                           "(select idEmpresa,idSucursal,tipovale,idvalescab,sum(TotalVVNeto) as TotalVVNeto,sum(TotalIGVNeto) as TotalIGVNeto,sum(TotalPVNeto) as TotalPVNeto from valesdet " & _
                                                           "WHERE idEmpresa = '" & glsEmpresa & "' " & _
                                                           "AND idValesCab = '" & StrValeIngreso & "' " & _
                                                           "AND tipoVale = 'I' " & _
                                                           "group by idEmpresa,idSucursal,idvalescab,tipovale) x " & _
                                                           "on a.idvalescab = x.idvalescab " & _
                                                           "and a.tipovale = x.tipovale " & _
                                                           "and a.idempresa = x.idempresa " & _
                                                           " " & _
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
                                                    ""
                                                       'AND item = " & rsvales.Fields("item") & "
                                                    Cn.Execute csql
                                                                
                                                    csql = "UPDATE a Set a.valorTotal = x.TotalVVNeto, a.igvTotal = x.TotalIGVNeto, a.precioTotal = x.TotalPVNeto from valescab a inner join " & _
                                                           "(select idEmpresa,idSucursal,tipovale,idvalescab,sum(TotalVVNeto) as TotalVVNeto,sum(TotalIGVNeto) as TotalIGVNeto,sum(TotalPVNeto) as TotalPVNeto from valesdet " & _
                                                           "WHERE idEmpresa = '" & glsEmpresa & "' " & _
                                                           "AND idValesCab = '" & StrValeIngreso & "' " & _
                                                           "AND tipoVale = 'I' " & _
                                                           "group by idEmpresa,idSucursal,idvalescab,tipovale) x " & _
                                                           "on a.idvalescab = x.idvalescab " & _
                                                           "and a.tipovale = x.tipovale " & _
                                                           "and a.idempresa = x.idempresa " & _
                                                           " " & _
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
                                        ""
                                        'AND item = " & rsvales.Fields("item") & "
                                Cn.Execute csql
                                        
                                csql = "UPDATE a Set a.valorTotal = x.TotalVVNeto, a.igvTotal = x.TotalIGVNeto, a.precioTotal = x.TotalPVNeto from valescab a inner join " & _
                                       "(select idEmpresa,idSucursal,tipovale,idvalescab,sum(TotalVVNeto) as TotalVVNeto,sum(TotalIGVNeto) as TotalIGVNeto,sum(TotalPVNeto) as TotalPVNeto from valesdet " & _
                                       "WHERE idEmpresa = '" & glsEmpresa & "' " & _
                                       "AND idSucursal = '" & rsvales.Fields("idSucursal") & "' " & _
                                       "AND idValesCab = '" & rsvales.Fields("idValesCab") & "' " & _
                                       "AND tipoVale = '" & rsvales.Fields("tipoVale") & "' " & _
                                       "group by idEmpresa,idSucursal,idvalescab,tipovale) x " & _
                                       "on a.idvalescab = x.idvalescab " & _
                                       "and a.tipovale = x.tipovale " & _
                                       "and a.idempresa = x.idempresa " & _
                                       "and a.idsucursal = x.idsucursal " & _
                                       " " & _
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
                    
                    If RsProductosError.RecordCount > 0 Then
                        
                        Do While Not IndSalir
                        
                            RsProductosError.Filter = "IdProducto = '" & Trim(rsvales.Fields("IdProducto") & "") & "'"
                            
                            If Not RsProductosError.EOF Then
                                
                                rsvales.MoveNext
                                RsProductosError.Filter = ""
                                
                                If rsvales.EOF Then
                                    IndSalir = True
                                End If
                                
                            Else
                            
                                IndSalir = True
                                RsProductosError.Filter = ""
                                
                            End If
                        
                        Loop
                        
                        IndSalir = False
                        
                    End If
                    
                    If rsvales.EOF Then Exit Do
                    
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
    
    csql = "UPDATE a Set A.CantidadStock = v.STOCK from productosalmacen a, vw_temp_stock v " & _
           " " & _
           "Where A.idEmpresa = v.idEmpresa " & _
           "AND a.idSucursal = v.idSucursal " & _
           "AND a.idAlmacen = v.idAlmacen " & _
           "AND a.idProducto = v.idProducto " & _
           "AND a.idUMCompra = v.idUM"
    Cn.Execute csql
    '---------------------------------------------------------

    Me.MousePointer = 1
    
    MsgBox "Fin del proceso", vbInformation, App.Title
    
    If RsProductosError.RecordCount > 0 Then
        
        MsgBox "Algunos Productos no se procesaron correctamente, revisar el Kardex del listado de Productos que se mostrará a continuación.", vbInformation, App.Title
        
        Open App.Path & "\Temporales\ResumenProductosError.TXT" For Output As #1
    
        RsProductosError.MoveFirst
        Do While Not RsProductosError.EOF
            
            Print #1, "Producto: " & Trim("" & RsProductosError.Fields("IdProducto")) & " - " & Trim("" & RsProductosError.Fields("GlsProducto"))
            
            RsProductosError.MoveNext
            
        Loop
        
        Close #1
        
        ShellEx App.Path & "\Temporales\ResumenProductosError.TXT", essSW_MAXIMIZE, , , "open", Me.hwnd
        
    End If
    
    Exit Sub
Err:
    Close #1
    If Err.Number = -2147217871 Or InStr(Err.Description, "Out of range value adjusted for column") > 0 Then
        
        RsProductosError.Filter = "IdProducto = '" & rsvales.Fields("idProducto") & "'"
        If RsProductosError.EOF Then
            RsProductosError.AddNew
            RsProductosError.Fields("IdProducto") = rsvales.Fields("idProducto")
            RsProductosError.Fields("GlsProducto") = traerCampo("Productos", "GlsProducto", "IdProducto", rsvales.Fields("idProducto"), True)
        End If
        
        RsProductosError.Filter = ""
        
        Resume Next
        
    End If
    If rsvales.State = 1 Then rsvales.Close: Set rsvales = Nothing
    Me.MousePointer = 1
    If StrMsgError = "" Then StrMsgError = Err.Description
    Exit Sub
    Resume
End Sub

Public Function costoPromedioTransferencia(idAlmacenOrigen As String, ByRef fecEmision As String, ByRef idProducto As String, StrMsgError As String) As Double
On Error GoTo Err
Dim rs As New ADODB.Recordset
Dim precioPromedio As Double
Dim precio As Double
Dim stock As Double
Dim primero As Boolean

    csql = "SELECT v.idAlmacen, v.tipoVale, v.TipoCambio, v.fechaEmision, d.Cantidad, c.indCosto, v.idSucursal, " & _
             "v.idValesCab, d.idProducto, d.item, v.idMoneda, v.idConcepto, " & _
             "if(v.idMoneda = 'PEN', d.VVUnit, d.VVUnit * v.TipoCambio) as VVUnit " & _
             "FROM valescab v, valesdet d, conceptos c " & _
             "WHERE v.idEmpresa = d.idEmpresa " & _
             "AND v.IdSucursal = d.idSucursal " & _
             "AND v.idValesCab = d.idValesCab " & _
             "AND v.tipoVale = d.tipoVale " & _
             "AND v.idConcepto = c.idConcepto " & _
             "AND estValeCab <> 'ANU' " & _
             "AND v.idEmpresa = '" & glsEmpresa & "' " & _
             "AND v.idAlmacen = '" & idAlmacenOrigen & "' " & _
             "AND v.FechaEmision <= CAST('" & Format(fecEmision, "yyyy-mm-dd") & "' AS DATE) " & _
             "AND d.idProducto = '" & idProducto & "' " & _
             "ORDER BY v.idSucursal,v.idAlmacen, d.idProducto, v.fechaEmision, v.idValesCab"
    If rs.State = 1 Then rs.Close
    rs.Open csql, Cn, adOpenStatic, adLockReadOnly
    
    If rs.RecordCount <> 0 Then
        precioPromedio = 0#
        stock = 0#
        primero = True
        rs.MoveFirst
        Do While Not rs.EOF
            If primero = True Then
                precioPromedio = Val(rs.Fields("VVUnit") & "")
                stock = Val(rs.Fields("Cantidad") & "")
                primero = False
            Else
                If rs.Fields("tipoVale") = "I" Then
                    precio = Val(rs.Fields("VVUnit") & "") * Val(rs.Fields("Cantidad") & "")
                    stock = stock + Val(rs.Fields("Cantidad") & "")
                    precioPromedio = (precioPromedio * stock + precio) / (stock + Val(rs.Fields("Cantidad") & ""))
                Else
                    stock = stock - Val(rs.Fields("Cantidad") & "")
                End If
            End If
            rs.MoveNext
        Loop
        
        costoPromedioTransferencia = precioPromedio
    Else
        costoPromedioTransferencia = 0
    End If

    Exit Function
    
Err:
    If StrMsgError = "" Then StrMsgError = Err.Description

End Function

Private Sub Form_Load()
    Dim xMes      As Integer
    Dim strAno    As Integer
    Dim i         As Integer
    
    Me.top = 0
    Me.left = 0
    
    xMes = Format(getFechaSistema, "mm")
    strAno = Format(Year(getFechaSistema), "0000")
    
    cbxAno.Clear
    For i = 2008 To Val(strAno)
        cbxAno.AddItem i
    Next
    cbxAno.AddItem "Periodo Actual"
    
    cbx_Mes.AddItem "ENERO" & Space(80) & "01"
    cbx_Mes.AddItem "FEBRERO" & Space(80) & "02"
    cbx_Mes.AddItem "MARZO" & Space(80) & "03"
    cbx_Mes.AddItem "ABRIL" & Space(80) & "04"
    cbx_Mes.AddItem "MAYO" & Space(80) & "05"
    cbx_Mes.AddItem "JUNIO" & Space(80) & "06"
    cbx_Mes.AddItem "JULIO" & Space(80) & "07"
    cbx_Mes.AddItem "AGOSTO" & Space(80) & "08"
    cbx_Mes.AddItem "SETIEMBRE" & Space(80) & "09"
    cbx_Mes.AddItem "OCTUBRE" & Space(80) & "10"
    cbx_Mes.AddItem "NOVIEMBRE" & Space(80) & "11"
    cbx_Mes.AddItem "DICIEMBRE" & Space(80) & "12"
    
    cbx_Mes.ListIndex = xMes - 1
    
    'strAno = Format(Year(getFechaSistema), "0000")
    
    For i = 0 To cbxAno.ListCount - 1
        cbxAno.ListIndex = i
        If cbxAno.Text = strAno Then Exit For
    Next
    
    If leeParametro("VIZUALIZA_CODIGO_RAPIDO") = "S" Then
        txtCod_Producto.MaxLength = 20
        CCodProducto = "CodigoRapido"
    Else
        txtCod_Producto.MaxLength = 8
        CCodProducto = "IdProducto"
    End If
    
    txtGls_Producto.Text = "TODOS LOS PRODUCTOS"
End Sub

Private Sub txtCod_Producto_Change()
    If txtCod_Producto.Text <> "" Then
        txtGls_Producto.Text = traerCampo("productos", "GlsProducto", "idProducto", txtCod_Producto.Text, True)
    Else
        txtGls_Producto.Text = "TODOS LOS PRODUCTOS"
    End If
End Sub

 
