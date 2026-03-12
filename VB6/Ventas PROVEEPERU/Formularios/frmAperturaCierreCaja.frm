VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{F41D1D30-7878-4923-8CB3-6CCACDC9C9DE}#1.0#0"; "catcontrols.ocx"
Begin VB.Form frmAperturaCierreCaja 
   Appearance      =   0  'Flat
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Apertura y Cierre de Caja"
   ClientHeight    =   3975
   ClientLeft      =   3420
   ClientTop       =   1890
   ClientWidth     =   8325
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3975
   ScaleWidth      =   8325
   ShowInTaskbar   =   0   'False
   Begin MSComCtl2.DTPicker dtpFecCaja 
      Height          =   315
      Left            =   6405
      TabIndex        =   17
      Top             =   930
      Width           =   1335
      _ExtentX        =   2355
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
      Format          =   103809025
      CurrentDate     =   39392
   End
   Begin VB.CommandButton cmbAyudaCaja 
      Height          =   315
      Left            =   7725
      Picture         =   "frmAperturaCierreCaja.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   15
      Top             =   1350
      Width           =   390
   End
   Begin CATControls.CATTextBox txtCod_Caja 
      Height          =   315
      Left            =   2250
      TabIndex        =   16
      Tag             =   "TidCaja"
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
      Container       =   "frmAperturaCierreCaja.frx":038A
      Estilo          =   1
      EnterTab        =   -1  'True
   End
   Begin MSComctlLib.ImageList imgDocVentas 
      Left            =   480
      Top             =   5040
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
            Picture         =   "frmAperturaCierreCaja.frx":03A6
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAperturaCierreCaja.frx":0740
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAperturaCierreCaja.frx":0B92
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAperturaCierreCaja.frx":0F2C
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAperturaCierreCaja.frx":12C6
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAperturaCierreCaja.frx":1660
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAperturaCierreCaja.frx":19FA
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAperturaCierreCaja.frx":1D94
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAperturaCierreCaja.frx":212E
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAperturaCierreCaja.frx":24C8
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAperturaCierreCaja.frx":2862
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAperturaCierreCaja.frx":3524
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   660
      Left            =   0
      TabIndex        =   3
      Top             =   0
      Width           =   8325
      _ExtentX        =   14684
      _ExtentY        =   1164
      ButtonWidth     =   1032
      ButtonHeight    =   1005
      Appearance      =   1
      ImageList       =   "imgDocVentas"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   3
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Abrir"
            Object.ToolTipText     =   "Grabar"
            ImageIndex      =   4
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Cerrar"
            Object.ToolTipText     =   "Cancelar"
            ImageIndex      =   9
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Salir"
            Object.ToolTipText     =   "Salir"
            ImageIndex      =   2
         EndProperty
      EndProperty
      BorderStyle     =   1
   End
   Begin VB.Frame fraCaja 
      Appearance      =   0  'Flat
      ForeColor       =   &H00C00000&
      Height          =   3210
      Left            =   30
      TabIndex        =   0
      Top             =   720
      Width           =   8235
      Begin CATControls.CATTextBox txtGls_Caja 
         Height          =   315
         Left            =   3195
         TabIndex        =   1
         Top             =   615
         Width           =   4500
         _ExtentX        =   7938
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
         Container       =   "frmAperturaCierreCaja.frx":38BE
      End
      Begin CATControls.CATTextBox txtVal_InicialSoles 
         Height          =   315
         Left            =   2220
         TabIndex        =   9
         Tag             =   "TidPersona"
         Top             =   1110
         Width           =   1875
         _ExtentX        =   3307
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
         Container       =   "frmAperturaCierreCaja.frx":38DA
         Text            =   "0.00"
         Decimales       =   2
         Estilo          =   4
         EnterTab        =   -1  'True
      End
      Begin CATControls.CATTextBox txtVal_InicialDolar 
         Height          =   315
         Left            =   5760
         TabIndex        =   10
         Tag             =   "TidPersona"
         Top             =   1110
         Width           =   1875
         _ExtentX        =   3307
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
         Container       =   "frmAperturaCierreCaja.frx":38F6
         Text            =   "0.00"
         Decimales       =   2
         Estilo          =   4
         EnterTab        =   -1  'True
      End
      Begin CATControls.CATTextBox txtCod_MovCaja 
         Height          =   285
         Left            =   7170
         TabIndex        =   13
         Tag             =   "TidMovCaja"
         Top             =   2760
         Visible         =   0   'False
         Width           =   915
         _ExtentX        =   1614
         _ExtentY        =   503
         BackColor       =   12640511
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
         Container       =   "frmAperturaCierreCaja.frx":3912
         Estilo          =   1
         EnterTab        =   -1  'True
      End
      Begin VB.Label Label2 
         Appearance      =   0  'Flat
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
         ForeColor       =   &H80000008&
         Height          =   210
         Left            =   5595
         TabIndex        =   14
         Top             =   240
         Width           =   450
      End
      Begin VB.Label lblFechaCierre 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         Caption         =   "---"
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
         Left            =   2220
         TabIndex        =   12
         Top             =   2190
         Width           =   180
      End
      Begin VB.Label lblFechaApertura 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         Caption         =   "---"
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
         Left            =   2220
         TabIndex        =   11
         Top             =   1710
         Width           =   180
      End
      Begin VB.Label Label6 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         Caption         =   "Monto Inicial US$"
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
         Left            =   4320
         TabIndex        =   8
         Top             =   1170
         Width           =   1215
      End
      Begin VB.Label Label5 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         Caption         =   "Monto Inicial S/."
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
         Left            =   240
         TabIndex        =   7
         Top             =   1170
         Width           =   1110
      End
      Begin VB.Label Label4 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         Caption         =   "Fecha de Cierre"
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
         Left            =   240
         TabIndex        =   6
         Top             =   2190
         Width           =   1155
      End
      Begin VB.Label Label3 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         Caption         =   "Fecha de Apertura"
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
         Left            =   240
         TabIndex        =   5
         Top             =   1710
         Width           =   1365
      End
      Begin VB.Label lblEstado 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         Caption         =   "---"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H008080FF&
         Height          =   360
         Left            =   3690
         TabIndex        =   4
         Top             =   2520
         Width           =   345
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         Caption         =   "Caja"
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
         Left            =   270
         TabIndex        =   2
         Top             =   690
         Width           =   315
      End
   End
End
Attribute VB_Name = "frmAperturaCierreCaja"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim indInserta As Boolean

Private Sub cmbAyudaCaja_Click()
    
    mostrarAyuda "CAJASUSUARIO", txtCod_Caja, txtGls_Caja

End Sub

Private Sub dtpFecCaja_Change()
On Error GoTo Err
Dim StrMsgError As String

    If Trim(txtCod_Caja.Text) <> "" Then
        mostrarAperturaCierre Trim(txtCod_Caja.Text), StrMsgError
        If StrMsgError <> "" Then GoTo Err
    End If
    
    Exit Sub

Err:
    If StrMsgError = "" Then StrMsgError = Err.Description
    MsgBox StrMsgError, vbInformation, App.Title
End Sub

Private Sub Form_Load()
    
    nuevo

End Sub

Private Sub txtCod_Caja_Change()
On Error GoTo Err
Dim StrMsgError As String

    If Trim(txtCod_Caja.Text) <> "" Then
        mostrarAperturaCierre Trim(txtCod_Caja.Text), StrMsgError
        If StrMsgError <> "" Then GoTo Err
    End If
    
    Exit Sub

Err:
    If StrMsgError = "" Then StrMsgError = Err.Description
    MsgBox StrMsgError, vbInformation, App.Title
End Sub

Private Sub txtCod_Caja_KeyPress(KeyAscii As Integer)
    
    If KeyAscii <> 13 Then
        mostrarAyudaKeyascii KeyAscii, "CAJASUSUARIO", txtCod_Caja, txtGls_Caja
        KeyAscii = 0
        If txtCod_Caja.Text <> "" Then SendKeys "{tab}"
    End If

End Sub

Private Sub nuevo()
        
    limpiaForm Me
    If Trim(txtCod_MovCaja.Text) = "" Then
        Toolbar1.Buttons(2).Enabled = False
    Else
        Toolbar1.Buttons(2).Enabled = True
    End If

End Sub

Private Sub mostrarAperturaCierre(StrCod As String, ByRef StrMsgError As String)
On Error GoTo Err
Dim rst As New ADODB.Recordset

    fraCaja.Enabled = True
    Toolbar1.Buttons(1).Enabled = True
    Toolbar1.Buttons(2).Enabled = True
    
    csql = "SELECT m.idMovCaja,m.idCaja,m.indEstado,c.GlsCaja,m.FecApertura,m.FecCierre " & _
            "FROM movcajas m,cajas c " & _
            "WHERE m.idCaja = c.idCaja " & _
             "AND m.idCaja = '" & StrCod & "' " & _
             "AND m.idUsuario = '" & glsUser & "' " & _
             "AND m.idEmpresa = '" & glsEmpresa & "' " & _
             "AND c.idEmpresa = '" & glsEmpresa & "' " & _
             "AND m.idSucursal = '" & glsSucursal & "' " & _
             "AND m.FecCaja = '" & Format(dtpFecCaja.Value, "yyyy-mm-dd") & "' " & _
            " ORDER BY m.indEstado ASC"
            
    If rst.State = 1 Then rst.Close
    rst.Open csql, Cn, adOpenKeyset, adLockOptimistic
    If Not rst.EOF Then
        txtCod_MovCaja.Text = "" & rst.Fields("idMovCaja")
        lblFechaApertura.Caption = "" & rst.Fields("FecApertura")
        lblFechaCierre.Caption = "" & rst.Fields("FecCierre")
                       
        If ("" & rst.Fields("indEstado")) = "C" Then
            lblEstado.Caption = "Cerrado"
            lblEstado.ForeColor = &HFF&
            Toolbar1.Buttons(1).Enabled = True
            Toolbar1.Buttons(2).Enabled = False
        Else
            lblEstado.Caption = "Abierto"
            lblEstado.ForeColor = &HC000&
            Toolbar1.Buttons(1).Enabled = False
            Toolbar1.Buttons(2).Enabled = True
        
            lblFechaCierre.Caption = "---"
        End If
        txtVal_InicialSoles.Text = traerCampo("movcajasdet", "ValMonto", "idMovCaja", txtCod_MovCaja.Text, True, " idTipoMovCaja = '99990001' AND idMoneda = 'PEN'")
        txtVal_InicialDolar.Text = traerCampo("movcajasdet", "ValMonto", "idMovCaja", txtCod_MovCaja.Text, True, " idTipoMovCaja = '99990001'  AND idMoneda = 'USD'")
    
    Else
        txtCod_MovCaja.Text = ""
        lblFechaApertura.Caption = "---"
        lblFechaCierre.Caption = "---"
        lblEstado.Caption = "Dia no Abierto"
        lblEstado.ForeColor = &HFF&
    End If
    
    If rst.State = 1 Then rst.Close: Set rst = Nothing
    Me.Refresh
    
    Exit Sub
    
Err:
    If rst.State = 1 Then rst.Close: Set rst = Nothing
    If StrMsgError = "" Then StrMsgError = Err.Description
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
On Error GoTo Err
Dim StrMsgError As String

    Select Case Button.Index
        Case 1 'Abrir
            Grabar "A", StrMsgError
            If StrMsgError <> "" Then GoTo Err
        Case 2 'Cerrar
            Grabar "C", StrMsgError
            If StrMsgError <> "" Then GoTo Err
        Case 3 'Salir
            Unload Me
    End Select
    
    Exit Sub

Err:
    MsgBox StrMsgError, vbInformation, App.Title
End Sub

Private Sub Grabar(StrTipo As String, ByRef StrMsgError As String)
On Error GoTo Err
Dim rst As New ADODB.Recordset
Dim indIniTrans As Boolean
Dim strCodMovDet As String
Dim indEvaluacion As Integer
Dim strCodUsuarioAutorizacion As String

    indIniTrans = False
    If txtCod_Caja.Text = "" Then
        StrMsgError = "Seleccione una Caja"
        GoTo Err
    End If

    Cn.BeginTrans
    indIniTrans = True

    If Trim(txtCod_MovCaja.Text) = "" Then
        If rst.State = 1 Then rst.Close
        csql = "SELECT m.FecCaja " & _
                "FROM movcajas m " & _
                "WHERE m.idCaja = '" & txtCod_Caja.Text & "' " & _
                 "AND m.idUsuario = '" & glsUser & "' " & _
                 "AND m.idEmpresa = '" & glsEmpresa & "' " & _
                 "AND m.idSucursal = '" & glsSucursal & "' " & _
                 "AND m.FecCaja < '" & Format(dtpFecCaja.Value, "yyyy-mm-dd") & "' " & _
                 "AND m.indEstado = 'A' " & _
                " ORDER BY m.indEstado ASC"
        rst.Open csql, Cn, adOpenKeyset, adLockOptimistic
        If Not rst.EOF Then
            StrMsgError = "El dia " & rst.Fields("FecCaja") & ", se encuentra todavia abierto"
            GoTo Err
        End If
        
        txtCod_MovCaja.Text = GeneraCorrelativoAnoMes("movcajas", "idMovCaja")
    
        csql = "INSERT INTO movcajas (idMovCaja,idCaja,idUsuario,indEstado,FecApertura,FecCierre,idEmpresa,idSucursal,FecCaja) VALUES(" & _
               "'" & Trim(txtCod_MovCaja.Text) & "','" & txtCod_Caja.Text & "','" & glsUser & "','" & StrTipo & "',SYSDATE(),SYSDATE(),'" & glsEmpresa & "','" & glsSucursal & "','" & Format(dtpFecCaja.Value, "yyyy-mm-dd") & "')"
        Cn.Execute csql
        
        strCodMovDet = GeneraCorrelativoAnoMes("movcajasdet", "idMovCajaDet")
            
        csql = "INSERT INTO movcajasdet (idMovCajaDet,idMovCaja,idTipoMovCaja,idMoneda,ValMonto,FecRegistro,idEmpresa,idSucursal,ValTipoCambio) VALUES(" & _
                "'" & strCodMovDet & "','" & txtCod_MovCaja.Text & "','99990001','PEN'," & txtVal_InicialSoles.Value & ",sysdate(),'" & glsEmpresa & "','" & glsSucursal & "'," & glsTC & ")"
        Cn.Execute csql
            
        strCodMovDet = GeneraCorrelativoAnoMes("movcajasdet", "idMovCajaDet")
            
        csql = "INSERT INTO movcajasdet (idMovCajaDet,idMovCaja,idTipoMovCaja,idMoneda,ValMonto,FecRegistro,idEmpresa,idSucursal,ValTipoCambio) VALUES(" & _
                "'" & strCodMovDet & "','" & txtCod_MovCaja.Text & "','99990001','USD'," & txtVal_InicialDolar.Value & ",sysdate(),'" & glsEmpresa & "','" & glsSucursal & "'," & glsTC & ")"
        Cn.Execute csql
    
    Else
        If StrTipo = "A" Then
            indEvaluacion = 0
    
            frmAprobacion.MostrarForm "06", indEvaluacion, strCodUsuarioAutorizacion, StrMsgError
            If StrMsgError <> "" Then GoTo Err
            
            If indEvaluacion = 0 Then
                If rst.State = 1 Then rst.Close
                Set rst = Nothing
                If indIniTrans = True Then Cn.RollbackTrans
                Exit Sub
            End If
        
            csql = "UPDATE movcajas SET  indEstado = 'A' WHERE idMovCaja = '" & Trim(txtCod_MovCaja.Text) & "' AND idSucursal = '" & glsSucursal & "' AND idEmpresa = '" & glsEmpresa & "'"
        Else
            csql = "UPDATE movcajas SET FecCierre = SYSDATE(), indEstado = 'C' WHERE idMovCaja = '" & Trim(txtCod_MovCaja.Text) & "' AND idSucursal = '" & glsSucursal & "' AND idEmpresa = '" & glsEmpresa & "'"
        End If
        
        Cn.Execute csql
    End If
    
    Cn.CommitTrans

    mostrarAperturaCierre txtCod_Caja.Text, StrMsgError
    If StrMsgError <> "" Then GoTo Err
    MsgBox "Se grabo Satisfactoriamente", vbInformation, App.Title
    If rst.State = 1 Then rst.Close: Set rst = Nothing

    Exit Sub
    
Err:
    If rst.State = 1 Then rst.Close: Set rst = Nothing
    If indIniTrans = True Then Cn.RollbackTrans
    If StrMsgError = "" Then StrMsgError = Err.Description
End Sub

Private Sub txtCod_MovCaja_Change()
    
    If Trim(txtCod_MovCaja.Text) = "" Then
        Toolbar1.Buttons(2).Enabled = False
    Else
        Toolbar1.Buttons(2).Enabled = True
    End If
    
End Sub
