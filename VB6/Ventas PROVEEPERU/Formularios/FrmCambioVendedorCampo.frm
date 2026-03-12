VERSION 5.00
Object = "{F41D1D30-7878-4923-8CB3-6CCACDC9C9DE}#1.0#0"; "catcontrols.ocx"
Begin VB.Form FrmCambioVendedorCampo 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Cambio de Vendedor de Campo"
   ClientHeight    =   4350
   ClientLeft      =   2475
   ClientTop       =   2625
   ClientWidth     =   7170
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
   ScaleHeight     =   4350
   ScaleWidth      =   7170
   ShowInTaskbar   =   0   'False
   Begin VB.Frame fraReportes 
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
      ForeColor       =   &H00C00000&
      Height          =   800
      Index           =   0
      Left            =   90
      TabIndex        =   19
      Top             =   2880
      Width           =   6960
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
         Left            =   6420
         Picture         =   "FrmCambioVendedorCampo.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   9
         Top             =   220
         Width           =   390
      End
      Begin CATControls.CATTextBox txtCod_Vendedor 
         Height          =   315
         Left            =   1545
         TabIndex        =   5
         Tag             =   "TidPerCliente"
         Top             =   240
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
         Container       =   "FrmCambioVendedorCampo.frx":038A
         Estilo          =   1
         EnterTab        =   -1  'True
      End
      Begin CATControls.CATTextBox txtGls_Vendedor 
         Height          =   315
         Left            =   2580
         TabIndex        =   20
         Tag             =   "TGlsCliente"
         Top             =   225
         Width           =   3825
         _ExtentX        =   6747
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
         Container       =   "FrmCambioVendedorCampo.frx":03A6
         Estilo          =   1
         Vacio           =   -1  'True
      End
      Begin VB.Label Label12 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         Caption         =   "Nuevo Vendedor"
         ForeColor       =   &H80000007&
         Height          =   210
         Left            =   210
         TabIndex        =   21
         Top             =   285
         Width           =   1230
      End
   End
   Begin VB.CommandButton btnCancelar 
      Caption         =   "&Salir"
      Height          =   435
      Left            =   3645
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   3780
      Width           =   1230
   End
   Begin VB.CommandButton btnAceptar 
      Caption         =   "&Aceptar"
      Height          =   435
      Left            =   2340
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   3780
      Width           =   1230
   End
   Begin VB.Frame fraReportes 
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   800
      Index           =   15
      Left            =   90
      TabIndex        =   16
      Top             =   2070
      Width           =   6960
      Begin CATControls.CATTextBox txtCod_Vendedor_Ori 
         Height          =   315
         Left            =   1545
         TabIndex        =   4
         Tag             =   "TidPerCliente"
         Top             =   225
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
         Container       =   "FrmCambioVendedorCampo.frx":03C2
         Estilo          =   1
         EnterTab        =   -1  'True
      End
      Begin CATControls.CATTextBox txtGls_Vendedor_Ori 
         Height          =   315
         Left            =   2580
         TabIndex        =   17
         Tag             =   "TGlsCliente"
         Top             =   225
         Width           =   3825
         _ExtentX        =   6747
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
         Container       =   "FrmCambioVendedorCampo.frx":03DE
         Estilo          =   1
         Vacio           =   -1  'True
      End
      Begin VB.Label VendedorOri 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         Caption         =   "Vendedor"
         ForeColor       =   &H80000007&
         Height          =   210
         Left            =   210
         TabIndex        =   18
         Top             =   285
         Width           =   720
      End
   End
   Begin VB.Frame fraReportes 
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   800
      Index           =   13
      Left            =   90
      TabIndex        =   13
      Top             =   1260
      Width           =   6960
      Begin CATControls.CATTextBox txtCod_Cliente_Ori 
         Height          =   315
         Left            =   1545
         TabIndex        =   3
         Tag             =   "TidPerCliente"
         Top             =   270
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
         Container       =   "FrmCambioVendedorCampo.frx":03FA
         Estilo          =   1
         EnterTab        =   -1  'True
      End
      Begin CATControls.CATTextBox txtGls_Cliente_Ori 
         Height          =   315
         Left            =   2580
         TabIndex        =   14
         Tag             =   "TGlsCliente"
         Top             =   270
         Width           =   3825
         _ExtentX        =   6747
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
         Container       =   "FrmCambioVendedorCampo.frx":0416
         Estilo          =   1
         Vacio           =   -1  'True
      End
      Begin VB.Label lbl_Cliente 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         Caption         =   "Cliente"
         ForeColor       =   &H80000007&
         Height          =   210
         Left            =   210
         TabIndex        =   15
         Top             =   330
         Width           =   480
      End
   End
   Begin VB.Frame fraReportes 
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
      ForeColor       =   &H00C00000&
      Height          =   1215
      Index           =   11
      Left            =   90
      TabIndex        =   10
      Top             =   0
      Width           =   6960
      Begin VB.CommandButton cmbAyudaTipoDoc 
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
         Left            =   6420
         Picture         =   "FrmCambioVendedorCampo.frx":0432
         Style           =   1  'Graphical
         TabIndex        =   11
         Top             =   300
         Width           =   390
      End
      Begin CATControls.CATTextBox txtCod_Documento 
         Height          =   315
         Left            =   1545
         TabIndex        =   0
         Tag             =   "TidMoneda"
         Top             =   315
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
         Container       =   "FrmCambioVendedorCampo.frx":07BC
         Estilo          =   1
         EnterTab        =   -1  'True
      End
      Begin CATControls.CATTextBox txtGls_Documento 
         Height          =   315
         Left            =   2580
         TabIndex        =   8
         Top             =   315
         Width           =   3825
         _ExtentX        =   6747
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
         Container       =   "FrmCambioVendedorCampo.frx":07D8
         Vacio           =   -1  'True
      End
      Begin CATControls.CATTextBox txt_numero 
         Height          =   315
         Left            =   4455
         TabIndex        =   2
         Top             =   720
         Width           =   1935
         _ExtentX        =   3413
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
         Alignment       =   2
         FontName        =   "Arial"
         FontSize        =   8.25
         ForeColor       =   -2147483640
         MaxLength       =   8
         Container       =   "FrmCambioVendedorCampo.frx":07F4
         Estilo          =   3
      End
      Begin CATControls.CATTextBox txt_serie 
         Height          =   315
         Left            =   1545
         TabIndex        =   1
         Top             =   720
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
         Alignment       =   2
         FontName        =   "Arial"
         FontSize        =   8.25
         ForeColor       =   -2147483640
         MaxLength       =   3
         Container       =   "FrmCambioVendedorCampo.frx":0810
         Estilo          =   3
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Serie"
         Height          =   210
         Left            =   210
         TabIndex        =   23
         Top             =   765
         Width           =   375
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Número"
         Height          =   210
         Left            =   3750
         TabIndex        =   22
         Top             =   765
         Width           =   555
      End
      Begin VB.Label Label11 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         Caption         =   "Tipo Documento"
         ForeColor       =   &H80000007&
         Height          =   210
         Left            =   210
         TabIndex        =   12
         Top             =   360
         Width           =   1155
      End
   End
End
Attribute VB_Name = "FrmCambioVendedorCampo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Btnaceptar_Click()
On Error GoTo Err
Dim rsd     As New ADODB.Recordset
Dim rsv     As New ADODB.Recordset
Dim Documento  As String
Dim A As String
 
    If txtCod_Documento.Text = "" Then
        MsgBox "Ingrese Documento", vbInformation, App.Title
        Exit Sub
    End If
 
    If txt_serie.Text = "" Then
        MsgBox "Ingrese Serie del Documento", vbInformation, App.Title
        Exit Sub
    End If
 
    If txt_serie.Text = "" Then
        MsgBox "Ingrese Nùmero del Documento", vbInformation, App.Title
        Exit Sub
    End If
 
    If MsgBox("¿Està Seguro(a) de Modificar el Vendedor?", vbInformation + vbYesNo, App.Title) = vbYes Then
        csql = "select IdDocVentas from docventas " & _
                "Where IdDocumento = '" & txtCod_Documento.Text & "' and IdSerie = '" & txt_serie.Text & "'" & _
                "and IdDocVentas = '" & txt_numero.Text & "' "
        rsv.Open csql, Cn, adOpenForwardOnly, adLockReadOnly
       
        If rsv.RecordCount = 0 Then
            MsgBox "El Documento No Existe.Verifique ", vbInformation, App.Title
            Exit Sub
        End If
 
        csql = "select AbreDocumento from Documentos " & _
                "where IdDocumento = '" & txtCod_Documento.Text & "'"
        rsd.Open csql, Cn, adOpenForwardOnly, adLockReadOnly
   
        Documento = "" & rsd.Fields("Abredocumento") + txt_serie.Text + "/" + txt_numero.Text
                                         
        csql = "Update docventas " & _
                "Set IdPerVendedorCampo = '" & txtCod_Vendedor.Text & "',GlsVendedorCampo = '" & txtGls_Vendedor.Text & "',IdPerVendedorCampo_Ant ='" & txtCod_Vendedor_Ori.Text & "' " & _
                "Where IdDocumento = '" & txtCod_Documento.Text & "' and IdSerie = '" & txt_serie.Text & "'" & _
                "and IdDocVentas = '" & txt_numero.Text & "' "
        Cn.Execute csql

        csql = "Update Cta_Dcto " & _
               "Set IdVendedor = '" & txtCod_Vendedor.Text & "'" & _
               "Where Nro_Comp = '" & Documento & "'"
        Cn.Execute csql
   
    Else
        Exit Sub
    End If
    MsgBox "Se Mofifico Satisfactoriamente ", vbInformation, App.Title
     
    If rsd.State = 1 Then rsd.Close: Set rst = Nothing
    If rsv.State = 1 Then rsv.Close: Set rst = Nothing
    
    Exit Sub
    
Err:
    If rsd.State = 1 Then rsd.Close: Set rsd = Nothing
    If StrMsgError = "" Then StrMsgError = Err.Description
    If rsv.State = 1 Then rsv.Close: Set rsv = Nothing
    If StrMsgError = "" Then StrMsgError = Err.Description
End Sub

Private Sub btnCancelar_Click()
    
    Unload Me

End Sub

Private Sub cmbAyudaTipoDoc_Click()
    
    mostrarAyuda "DOCUMENTOS", txtCod_Documento, txtGls_Documento

End Sub

Public Sub Vendedor()
On Error GoTo Err
Dim rst     As New ADODB.Recordset
  
    txt_numero.Text = Format("" & txt_numero.Text, "00000000")
    csql = "Select IdPerCliente,GlsCliente,IdPerVendedorCampo,GlsVendedorCampo from Docventas " & _
            "Where IdDocumento = '" & txtCod_Documento.Text & "' and IdSerie = '" & txt_serie.Text & "'" & _
            "and IdDocVentas = '" & txt_numero.Text & "' "
    rst.Open csql, Cn, adOpenForwardOnly, adLockReadOnly

    If rst.RecordCount <> 0 Then
        txtCod_Cliente_Ori.Text = "" & rst.Fields("IdPerCliente")
        txtGls_Cliente_Ori.Text = "" & rst.Fields("GlsCliente")
        txtCod_Vendedor_Ori.Text = "" & rst.Fields("IdPerVendedorCampo")
        txtGls_Vendedor_Ori.Text = "" & rst.Fields("GlsVendedorCampo")
    Else
        MsgBox "El Documento No Existe.Verifique ", vbInformation, App.Title
        Exit Sub
    End If
    If rst.State = 1 Then rst.Close: Set rst = Nothing
    
    Exit Sub
    
Err:
    If rst.State = 1 Then rst.Close: Set rst = Nothing
    If StrMsgError = "" Then StrMsgError = Err.Description
End Sub

Private Sub cmbAyudaVendedor_Click()
    
    mostrarAyuda "VENDEDOR", txtCod_Vendedor, txtGls_Vendedor

End Sub

Private Sub Form_Load()

    Me.top = 0
    Me.left = 0
    
End Sub

Private Sub txt_numero_KeyPress(KeyAscii As Integer)
 
    If txtCod_Documento.Text = "" Then
        MsgBox "Ingrese Documento", vbInformation, App.Title
        Exit Sub
    Else
        If KeyAscii = 13 Then
            Vendedor
            If txt_numero.Text <> "" Then SendKeys "{tab}"
        Else
            txtCod_Vendedor_Ori.Text = ""
            txtGls_Vendedor_Ori.Text = ""
            txtCod_Cliente_Ori.Text = ""
            txtGls_Cliente_Ori.Text = ""
            txtCod_Vendedor.Text = ""
            txtGls_Vendedor.Text = ""
        End If
    End If

End Sub

Public Sub txt_numero_LostFocus()
    
    txt_numero.Text = Format("" & txt_numero.Text, "00000000")

End Sub

Private Sub txt_Serie_KeyPress(KeyAscii As Integer)

    If KeyAscii = 13 Then
        If txt_serie.Text <> "" Then SendKeys "{tab}"
    End If

End Sub

Private Sub txt_Serie_LostFocus()
   
    txt_serie.Text = Format("" & txt_serie.Text, "000")

End Sub

Private Sub txtCod_Documento_KeyPress(KeyAscii As Integer)

    If KeyAscii <> 13 Then
        mostrarAyuda "DOCUMENTOS", txtCod_Documento, txtGls_Documento
    End If
    
End Sub
