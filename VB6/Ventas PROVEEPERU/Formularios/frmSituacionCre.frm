VERSION 5.00
Object = "{6A24B331-7634-11D3-A5B0-0050044A7E1A}#1.5#0"; "DXDBGrid.dll"
Object = "{F41D1D30-7878-4923-8CB3-6CCACDC9C9DE}#1.0#0"; "CATControls.ocx"
Begin VB.Form frmSituacionCre 
   Appearance      =   0  'Flat
   Caption         =   "Situación Crediticia"
   ClientHeight    =   9105
   ClientLeft      =   2385
   ClientTop       =   1335
   ClientWidth     =   10065
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
   ScaleHeight     =   9105
   ScaleWidth      =   10065
   Begin VB.CommandButton CmdSalir 
      Appearance      =   0  'Flat
      Caption         =   "Salir"
      Height          =   420
      Left            =   4365
      TabIndex        =   22
      Top             =   8595
      Width           =   1770
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
      Height          =   8430
      Left            =   45
      TabIndex        =   0
      Top             =   45
      Width           =   9915
      Begin VB.Frame Frame4 
         Appearance      =   0  'Flat
         Caption         =   " Letras en Responsabilidad "
         ForeColor       =   &H80000008&
         Height          =   2400
         Left            =   135
         TabIndex        =   20
         Top             =   4410
         Width           =   9645
         Begin DXDBGRIDLibCtl.dxDBGrid GRID1 
            Height          =   1965
            Left            =   135
            OleObjectBlob   =   "frmSituacionCre.frx":0000
            TabIndex        =   21
            Top             =   315
            Width           =   9390
         End
      End
      Begin VB.Frame Frame3 
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
         Height          =   1500
         Left            =   2700
         TabIndex        =   10
         Top             =   6840
         Width           =   4605
         Begin CATControls.CATTextBox TxtLinea_Cre 
            Height          =   315
            Left            =   2070
            TabIndex        =   11
            Top             =   225
            Width           =   1695
            _ExtentX        =   2990
            _ExtentY        =   556
            BackColor       =   12640511
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
            Container       =   "frmSituacionCre.frx":4332
            Vacio           =   -1  'True
         End
         Begin CATControls.CATTextBox txtCtas_Cobrar 
            Height          =   315
            Left            =   2070
            TabIndex        =   12
            Top             =   630
            Width           =   1695
            _ExtentX        =   2990
            _ExtentY        =   556
            BackColor       =   12640511
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
            Container       =   "frmSituacionCre.frx":434E
            Vacio           =   -1  'True
         End
         Begin CATControls.CATTextBox TxtDisponible 
            Height          =   315
            Left            =   2070
            TabIndex        =   13
            Top             =   1080
            Width           =   1695
            _ExtentX        =   2990
            _ExtentY        =   556
            BackColor       =   12640511
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
            Container       =   "frmSituacionCre.frx":436A
            Vacio           =   -1  'True
         End
         Begin VB.Label Label12 
            Appearance      =   0  'Flat
            Caption         =   "__________________________"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000007&
            Height          =   240
            Left            =   1890
            TabIndex        =   18
            Top             =   810
            Width           =   2145
         End
         Begin VB.Label Label11 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            Caption         =   "Disponible"
            ForeColor       =   &H80000007&
            Height          =   210
            Left            =   135
            TabIndex        =   17
            Top             =   1125
            Width           =   735
         End
         Begin VB.Label Label10 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            Caption         =   "Cuentas por Cobrar"
            ForeColor       =   &H80000007&
            Height          =   210
            Left            =   135
            TabIndex        =   16
            Top             =   630
            Width           =   1425
         End
         Begin VB.Label lblSinboloLinea 
            Appearance      =   0  'Flat
            Caption         =   "x"
            ForeColor       =   &H80000007&
            Height          =   195
            Left            =   1710
            TabIndex        =   15
            Top             =   270
            Width           =   390
         End
         Begin VB.Label Label8 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            Caption         =   "Línea de Crédito"
            ForeColor       =   &H80000007&
            Height          =   210
            Left            =   135
            TabIndex        =   14
            Top             =   270
            Width           =   1170
         End
      End
      Begin VB.Frame Frame1 
         Appearance      =   0  'Flat
         Caption         =   " Documentos por Cobrar "
         ForeColor       =   &H80000008&
         Height          =   2580
         Left            =   135
         TabIndex        =   9
         Top             =   1710
         Width           =   9645
         Begin DXDBGRIDLibCtl.dxDBGrid GRID2 
            Height          =   2100
            Left            =   135
            OleObjectBlob   =   "frmSituacionCre.frx":4386
            TabIndex        =   19
            Top             =   315
            Width           =   9390
         End
      End
      Begin VB.Frame FraLineaCredito 
         Appearance      =   0  'Flat
         Caption         =   " Situación Crediticia "
         ForeColor       =   &H80000008&
         Height          =   1455
         Left            =   135
         TabIndex        =   1
         Top             =   180
         Width           =   9645
         Begin CATControls.CATTextBox txt_Cod_Cli_Linea 
            Height          =   315
            Left            =   1665
            TabIndex        =   2
            Top             =   315
            Width           =   1020
            _ExtentX        =   1799
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
            Container       =   "frmSituacionCre.frx":86B8
            Vacio           =   -1  'True
         End
         Begin CATControls.CATTextBox tztGlsLinea 
            Height          =   315
            Left            =   2745
            TabIndex        =   3
            Top             =   315
            Width           =   6735
            _ExtentX        =   11880
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
            Container       =   "frmSituacionCre.frx":86D4
            Vacio           =   -1  'True
         End
         Begin CATControls.CATTextBox txtRuc_Linea 
            Height          =   315
            Left            =   1665
            TabIndex        =   4
            Top             =   675
            Width           =   1605
            _ExtentX        =   2831
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
            Container       =   "frmSituacionCre.frx":86F0
            Vacio           =   -1  'True
         End
         Begin CATControls.CATTextBox txtfomaPagoLinea 
            Height          =   315
            Left            =   1665
            TabIndex        =   5
            Top             =   1035
            Width           =   7815
            _ExtentX        =   13785
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
            Container       =   "frmSituacionCre.frx":870C
            Vacio           =   -1  'True
         End
         Begin VB.Label Label7 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            Caption         =   "Forma de Pago"
            ForeColor       =   &H80000007&
            Height          =   210
            Left            =   315
            TabIndex        =   8
            Top             =   1080
            Width           =   1080
         End
         Begin VB.Label Label6 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            Caption         =   "R.U.C."
            ForeColor       =   &H80000007&
            Height          =   210
            Left            =   315
            TabIndex        =   7
            Top             =   720
            Width           =   450
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            Caption         =   "Cliente"
            Height          =   210
            Left            =   315
            TabIndex        =   6
            Top             =   360
            Width           =   480
         End
      End
   End
End
Attribute VB_Name = "frmSituacionCre"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim idformaPagoLinea      As String
Dim GlsformaPagoLinea     As String
Dim LineaCre              As Double
Dim RsCtadcto             As New ADODB.Recordset
Dim MonedaLinea           As String
Dim StrClienteFm          As String
Dim GlsClienteFm          As String
Dim RucClienteFm          As String

Private Sub cmdsalir_Click()
    
    Unload Me

End Sub

Public Sub MostrarFrom(StrCliente As String, GlsCliente As String, RucCliente As String)
Dim CSqlC                   As String

    If Len(Trim(StrCliente)) > 0 Then
    
        TxtLinea_Cre.Text = Val(Format(TxtLinea_Cre.Text, "###,###,##0.00"))
        txtCtas_Cobrar.Text = Val(Format(txtCtas_Cobrar.Text, "###,###,##0.00"))
        
        txt_Cod_Cli_Linea.Text = Trim("" & StrCliente)
        tztGlsLinea.Text = Trim("" & GlsCliente)
        txtRuc_Linea.Text = Trim("" & RucCliente)
        
        idformaPagoLinea = Trim("" & traerCampo("clientes", "idFormaPago", "idCliente", Trim("" & StrCliente), True))
        GlsformaPagoLinea = Trim("" & traerCampo("formaspagos", "GlsFormaPago", "idformapago", Trim("" & idformaPagoLinea), True))
        txtfomaPagoLinea.Text = Trim("" & GlsformaPagoLinea)
        
        MonedaLinea = Trim("" & traerCampo("clientes", "idMonedaLineaCredito", "idCliente", Trim("" & StrCliente), True))
        LineaCre = Val(Format((Trim("" & traerCampo("clientes", "Val_LineaCredito", "idCliente", Trim("" & StrCliente), True))), "0.00"))
        lblSinboloLinea = traerCampo("Monedas", "simbolo", "idmoneda", MonedaLinea, False)
        
        TxtLinea_Cre.Text = Val(Format(LineaCre, "0.00"))
        lblSinboloLinea.Caption = Trim("" & lblSinboloLinea)
        
        CSqlC = "Select Sum(A.Total) Total " & _
                "From (" & _
                    "Select A.IdCliente,Case '" & MonedaLinea & "' " & _
                    "When 'PEN' Then If(A.IdMoneda = 'PEN',A.Saldo,If(A.ValTipoCambio Is Null Or A.ValTipoCambio = 0, 0, A.Saldo * A.ValTipoCambio)) " & _
                    "When 'USD' Then If(A.IdMoneda = 'USD',A.Saldo,If(A.ValTipoCambio Is Null Or A.ValTipoCambio = 0, 0, A.Saldo / A.ValTipoCambio)) " & _
                    "End Total " & _
                    "From(" & _
                        "Select A.IdCliente,A.ValTipoCambio,A.IdMoneda,A.IndDeb_Hab,A.ValTotal - IfNull(Sum(" & _
                        "If(C.IdCta_Dcto Is Not Null,If(C.IndDeb_Hab = 'H',If(A.IdMoneda = 'PEN',If(C.IdMoneda = 'PEN',B.ValImputaSo " & _
                        ",Round(B.ValImputaDo * B.ValTipoCambio,2)),If(C.IdMoneda = 'PEN',Round(B.ValImputaSo / B.ValTipoCambio,2), " & _
                        "B.ValImputaDo)),If(A.IdMoneda = 'PEN',If(C.IdMoneda = 'PEN',B.ValImputaSo * -1,Round((B.ValImputaDo * B.ValTipoCambio),2) * -1) " & _
                        ",If(C.IdMoneda = 'PEN',Round((B.ValImputaSo / B.ValTipoCambio),2) * -1,B.ValImputaDo * -1))),0)),0) Saldo " & _
                        "From Cta_Dcto A " & _
                        "Left Join Cta_Mvto B " & _
                            "On A.IdEmpresa = B.IdEmpresa And A.IdCta_Dcto = B.IdCta_Dcto " & _
                        "Left Join Cta_Dcto C " & _
                            "On B.IdEmpresa = C.IdEmpresa And B.IdCta_Comp = C.IdCta_Dcto " & _
                        "Where A.IdEmpresa = '" & glsEmpresa & "' And (A.IdCliente = '" & Trim("" & StrCliente) & "') " & _
                        "And A.IndDeb_Hab = 'D' And Mid(A.Nro_Comp,1,3) <> 'Aju' " & _
                        "Group By A.IdCta_Dcto" & _
                    ") A " & _
                    "Where A.Saldo <> Round(0,2) Union All "
        
        CSqlC = CSqlC & "Select A.IdCliente,Case '" & MonedaLinea & "' " & _
                    "When 'PEN' Then If(A.IdMoneda = 'PEN',A.Saldo * -1,If(A.ValTipoCambio Is Null Or A.ValTipoCambio = 0, 0,(A.Saldo * A.ValTipoCambio) * -1)) " & _
                    "When 'USD' Then If(A.IdMoneda = 'USD',A.Saldo * -1,If(A.ValTipoCambio Is Null Or A.ValTipoCambio = 0, 0,(A.Saldo / A.ValTipoCambio) * -1)) End Total " & _
                    "From(" & _
                        "Select A.IdCliente,A.ValTipoCambio,A.IdMoneda,A.IndDeb_Hab,A.ValTotal -IfNull(Sum(" & _
                        "If(C.IdCta_Dcto Is Not Null,If(C.IndDeb_Hab = 'H',If(A.IdMoneda = 'PEN',B.ValImputaSo * -1,B.ValImputaDo * -1)," & _
                        "If(A.IdMoneda = 'PEN',B.ValImputaSo,B.ValImputaDo)),0)),0) Saldo " & _
                        "From Cta_Dcto A " & _
                        "Left Join Cta_Mvto B " & _
                            "On A.IdEmpresa = B.IdEmpresa And A.IdCta_Dcto = B.IdCta_Comp " & _
                        "Left Join Cta_Dcto C " & _
                            "On B.IdEmpresa = C.IdEmpresa And B.IdCta_Dcto = C.IdCta_Dcto " & _
                        "Where A.IdEmpresa = '" & glsEmpresa & "' And (A.IdCliente = '" & Trim("" & StrCliente) & "') " & _
                        "And A.IndDeb_Hab = 'H' And UCase(Left(A.Nro_Comp,3)) In('CRE','ANT','NDE') " & _
                        "Group By A.IdCta_Dcto" & _
                    ") A " & _
                    "Where A.Saldo <> Round(0,2)" & _
                ") A " & _
                "Group By A.IdCliente"
        
        If RsCtadcto.State = adStateOpen Then RsCtadcto.Close
        RsCtadcto.Open CSqlC, Cn, adOpenKeyset, adLockOptimistic
                    
        If Not RsCtadcto.EOF Then
            txtCtas_Cobrar.Text = Val(Format(RsCtadcto.Fields("total"), "0.00"))
        End If
        
        TxtDisponible.Text = Val(Format((TxtLinea_Cre.Text - txtCtas_Cobrar.Text), "0.00"))
        TxtLinea_Cre.Text = Format(TxtLinea_Cre.Text, "###,###,##0.00")
        txtCtas_Cobrar.Text = Format(txtCtas_Cobrar.Text, "###,###,##0.00")
        TxtDisponible.Text = Format(TxtDisponible.Text, "###,###,##0.00")
        
        ConfGrid GRID2, False, True, False, False
        ConfGrid GRID1, False, True, False, False
        
        CSqlC = "Select A.*,If(A.IdMoneda = 'PEN','S./','US$') Moneda,IfNull((Case '" & MonedaLinea & "' " & _
                "When 'PEN' Then If(A.IdMoneda = 'PEN',A.Saldo,If(A.ValTipoCambio Is Null Or A.ValTipoCambio = 0, 0,A.Saldo * A.ValTipoCambio)) " & _
                "When 'USD' Then If(A.IdMoneda = 'USD',A.Saldo,If(A.ValTipoCambio Is Null Or A.ValTipoCambio = 0, 0,A.Saldo / A.ValTipoCambio)) " & _
                "End),0) ValSaldo,IfNull((Case '" & MonedaLinea & "' " & _
                "When 'PEN' Then If(A.IdMoneda = 'PEN',A.ValTotal,If(A.ValTipoCambio Is Null Or A.ValTipoCambio = 0, 0,A.ValTotal * A.ValTipoCambio)) " & _
                "When 'USD' Then If(A.IdMoneda = 'USD',A.ValTotal,If(A.ValTipoCambio Is Null Or A.ValTipoCambio = 0, 0,A.ValTotal / A.ValTipoCambio)) " & _
                "End),0) ValTotal " & _
                "From(" & _
                    "Select A.IdCta_Dcto,A.Fec_Comp,A.Nro_Comp,A.Fec_Vcto,A.ValTipoCambio,A.IdMoneda,Concat(A.IdCta_Dcto,A.Nro_Comp) Item," & _
                    "A.IndDeb_Hab,A.ValTotal,A.ValTotal - IfNull(Sum(" & _
                    "If(C.IdCta_Dcto Is Not Null,If(C.IndDeb_Hab = 'H',If(A.IdMoneda = 'PEN',If(C.IdMoneda = 'PEN',B.ValImputaSo " & _
                    ",Round(B.ValImputaDo * B.ValTipoCambio,2)),If(C.IdMoneda = 'PEN',Round(B.ValImputaSo / B.ValTipoCambio,2), " & _
                    "B.ValImputaDo)),If(A.IdMoneda = 'PEN',If(C.IdMoneda = 'PEN',B.ValImputaSo * -1,Round((B.ValImputaDo * B.ValTipoCambio),2) * -1) " & _
                    ",If(C.IdMoneda = 'PEN',Round((B.ValImputaSo / B.ValTipoCambio),2) * -1,B.ValImputaDo * -1))),0)),0) Saldo " & _
                    "From Cta_Dcto A " & _
                    "Left Join Cta_Mvto B " & _
                        "On A.IdEmpresa = B.IdEmpresa And A.IdCta_Dcto = B.IdCta_Dcto " & _
                    "Left Join Cta_Dcto C " & _
                        "On B.IdEmpresa = C.IdEmpresa And B.IdCta_Comp = C.IdCta_Dcto " & _
                    "Where A.IdEmpresa = '" & glsEmpresa & "' And (A.IdCliente = '" & Trim("" & StrCliente) & "') " & _
                    "And A.IndDeb_Hab = 'D' And Mid(A.Nro_Comp,1,3) Not In('Aju','Ler') " & _
                    "Group By A.IdCta_Dcto" & _
                ") A " & _
                "Where A.Saldo <> Round(0,2) Union All "
        
        CSqlC = CSqlC & "Select A.*,If(A.IdMoneda = 'PEN','S./','US$') Moneda,IfNull((Case '" & MonedaLinea & "' " & _
                "When 'PEN' Then If(A.IdMoneda = 'PEN',A.Saldo,If(A.ValTipoCambio Is Null Or A.ValTipoCambio = 0, 0,A.Saldo * A.ValTipoCambio)) " & _
                "When 'USD' Then If(A.IdMoneda = 'USD',A.Saldo,If(A.ValTipoCambio Is Null Or A.ValTipoCambio = 0, 0,A.Saldo / A.ValTipoCambio)) " & _
                "End),0) ValSaldo,IfNull((Case '" & MonedaLinea & "' " & _
                "When 'PEN' Then If(A.IdMoneda = 'PEN',A.ValTotal,If(A.ValTipoCambio Is Null Or A.ValTipoCambio = 0, 0,A.ValTotal * A.ValTipoCambio)) " & _
                "When 'USD' Then If(A.IdMoneda = 'USD',A.ValTotal,If(A.ValTipoCambio Is Null Or A.ValTipoCambio = 0, 0,A.ValTotal / A.ValTipoCambio)) " & _
                "End),0) ValTotal " & _
                "From(" & _
                    "Select A.IdCta_Dcto,A.Fec_Comp,A.Nro_Comp,A.Fec_Vcto,A.ValTipoCambio,A.IdMoneda,Concat(A.IdCta_Dcto,A.Nro_Comp) Item," & _
                    "A.IndDeb_Hab,A.ValTotal,(A.ValTotal - IfNull(Sum(" & _
                    "If(C.IdCta_Dcto Is Not Null,If(C.IndDeb_Hab = 'H',If(A.IdMoneda = 'PEN',B.ValImputaSo * -1,B.ValImputaDo * -1)," & _
                    "If(A.IdMoneda = 'PEN',B.ValImputaSo,B.ValImputaDo)),0)),0)) * -1 Saldo " & _
                    "From Cta_Dcto A " & _
                    "Left Join Cta_Mvto B " & _
                        "On A.IdEmpresa = B.IdEmpresa And A.IdCta_Dcto = B.IdCta_Comp " & _
                    "Left Join Cta_Dcto C " & _
                        "On B.IdEmpresa = C.IdEmpresa And B.IdCta_Dcto = C.IdCta_Dcto " & _
                    "Where A.IdEmpresa = '" & glsEmpresa & "' And (A.IdCliente = '" & Trim("" & StrCliente) & "') " & _
                    "And A.IndDeb_Hab = 'H' And UCase(Left(A.Nro_Comp,3)) In('CRE','ANT','NDE') " & _
                    "Group By A.IdCta_Dcto" & _
                ") A " & _
                "Where A.Saldo <> Round(0,2) Order By 2"
        
        With GRID2
             .DefaultFields = False
             .Dataset.ADODataset.ConnectionString = strcn
             .Dataset.ADODataset.CursorLocation = clUseClient
             .Dataset.Active = False
             .Dataset.ADODataset.CommandText = CSqlC
             .Dataset.DisableControls
             .Dataset.Active = True
             .KeyField = "idcta_dcto"
             .Dataset.Refresh
        End With
        
        CSqlC = "Select A.*,If(A.IdMoneda = 'PEN','S./','US$') Moneda,IfNull((Case '" & MonedaLinea & "' " & _
                "When 'PEN' Then If(A.IdMoneda = 'PEN',A.Saldo,If(A.ValTipoCambio Is Null Or A.ValTipoCambio = 0, 0,A.Saldo * A.ValTipoCambio)) " & _
                "When 'USD' Then If(A.IdMoneda = 'USD',A.Saldo,If(A.ValTipoCambio Is Null Or A.ValTipoCambio = 0, 0,A.Saldo / A.ValTipoCambio)) " & _
                "End),0) ValSaldo,IfNull((Case '" & MonedaLinea & "' " & _
                "When 'PEN' Then If(A.IdMoneda = 'PEN',A.ValTotal,If(A.ValTipoCambio Is Null Or A.ValTipoCambio = 0, 0,A.ValTotal * A.ValTipoCambio)) " & _
                "When 'USD' Then If(A.IdMoneda = 'USD',A.ValTotal,If(A.ValTipoCambio Is Null Or A.ValTipoCambio = 0, 0,A.ValTotal / A.ValTipoCambio)) " & _
                "End),0) ValTotal " & _
                "From(" & _
                    "Select A.IdCta_Dcto,A.Fec_Comp,A.Nro_Comp,A.Fec_Vcto,A.ValTipoCambio,A.IdMoneda,Concat(A.IdCta_Dcto,A.Nro_Comp) Item," & _
                    "A.IndDeb_Hab,A.ValTotal,A.ValTotal - IfNull(Sum(" & _
                    "If(C.IdCta_Dcto Is Not Null,If(C.IndDeb_Hab = 'H',If(A.IdMoneda = 'PEN',If(C.IdMoneda = 'PEN',B.ValImputaSo " & _
                    ",Round(B.ValImputaDo * B.ValTipoCambio,2)),If(C.IdMoneda = 'PEN',Round(B.ValImputaSo / B.ValTipoCambio,2), " & _
                    "B.ValImputaDo)),If(A.IdMoneda = 'PEN',If(C.IdMoneda = 'PEN',B.ValImputaSo * -1,Round((B.ValImputaDo * B.ValTipoCambio),2) * -1) " & _
                    ",If(C.IdMoneda = 'PEN',Round((B.ValImputaSo / B.ValTipoCambio),2) * -1,B.ValImputaDo * -1))),0)),0) Saldo " & _
                    "From Cta_Dcto A " & _
                    "Left Join Cta_Mvto B " & _
                        "On A.IdEmpresa = B.IdEmpresa And A.IdCta_Dcto = B.IdCta_Dcto " & _
                    "Left Join Cta_Dcto C " & _
                        "On B.IdEmpresa = C.IdEmpresa And B.IdCta_Comp = C.IdCta_Dcto " & _
                    "Where A.IdEmpresa = '" & glsEmpresa & "' And (A.IdCliente = '" & Trim("" & StrCliente) & "') " & _
                    "And A.IndDeb_Hab = 'D' And Mid(A.Nro_Comp,1,3) In('Ler') " & _
                    "Group By A.IdCta_Dcto" & _
                ") A " & _
                "Where A.Saldo <> Round(0,2) " & _
                "Order By 2"
        
        With GRID1
             .DefaultFields = False
             .Dataset.ADODataset.ConnectionString = strcn
             .Dataset.ADODataset.CursorLocation = clUseClient
             .Dataset.Active = False
             .Dataset.ADODataset.CommandText = CSqlC
             .Dataset.DisableControls
             .Dataset.Active = True
             .KeyField = "idcta_dcto"
             .Dataset.Refresh
        End With
        frmSituacionCre.Show 1
    End If

End Sub
