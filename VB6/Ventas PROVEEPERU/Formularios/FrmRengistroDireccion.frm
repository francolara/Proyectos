VERSION 5.00
Object = "{6A24B331-7634-11D3-A5B0-0050044A7E1A}#1.5#0"; "DXDBGrid.dll"
Object = "{F41D1D30-7878-4923-8CB3-6CCACDC9C9DE}#1.0#0"; "CATControls.ocx"
Begin VB.Form FrmRengistroDireccion 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Ingrese Dirección"
   ClientHeight    =   5385
   ClientLeft      =   6525
   ClientTop       =   4500
   ClientWidth     =   7260
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5385
   ScaleWidth      =   7260
   Begin VB.Frame Frame2 
      Caption         =   "Direcciones de entrega"
      Height          =   2565
      Left            =   90
      TabIndex        =   20
      Top             =   90
      Width           =   7125
      Begin DXDBGRIDLibCtl.dxDBGrid gLista 
         Height          =   2265
         Left            =   60
         OleObjectBlob   =   "FrmRengistroDireccion.frx":0000
         TabIndex        =   21
         Top             =   210
         Width           =   6975
      End
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Aceptar"
      Height          =   390
      Index           =   0
      Left            =   2730
      Style           =   1  'Graphical
      TabIndex        =   19
      Top             =   4920
      Width           =   1890
   End
   Begin VB.Frame Frame1 
      Height          =   2175
      Left            =   90
      TabIndex        =   0
      Top             =   2670
      Width           =   7095
      Begin VB.CommandButton cmbAyudaPais 
         Height          =   315
         Left            =   6630
         Picture         =   "FrmRengistroDireccion.frx":1BAD
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   225
         Width           =   390
      End
      Begin VB.CommandButton cmbAyudaDepa 
         Height          =   315
         Left            =   6630
         Picture         =   "FrmRengistroDireccion.frx":1F37
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   600
         Width           =   390
      End
      Begin VB.CommandButton cmbAyudaProv 
         Height          =   315
         Left            =   6630
         Picture         =   "FrmRengistroDireccion.frx":22C1
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   975
         Width           =   390
      End
      Begin VB.CommandButton cmbAyudaDistrito 
         Height          =   315
         Left            =   6630
         Picture         =   "FrmRengistroDireccion.frx":264B
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   1350
         Width           =   390
      End
      Begin CATControls.CATTextBox txtCod_Pais 
         Height          =   315
         Left            =   1290
         TabIndex        =   5
         Tag             =   "TidPais"
         Top             =   240
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
         Container       =   "FrmRengistroDireccion.frx":29D5
         Estilo          =   1
         EnterTab        =   -1  'True
      End
      Begin CATControls.CATTextBox txtGls_Pais 
         Height          =   315
         Left            =   2265
         TabIndex        =   6
         Top             =   240
         Width           =   4335
         _ExtentX        =   7646
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
         Container       =   "FrmRengistroDireccion.frx":29F1
      End
      Begin CATControls.CATTextBox txtCod_Depa 
         Height          =   315
         Left            =   1290
         TabIndex        =   7
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
         Container       =   "FrmRengistroDireccion.frx":2A0D
         Estilo          =   1
         EnterTab        =   -1  'True
      End
      Begin CATControls.CATTextBox txtGls_Depa 
         Height          =   315
         Left            =   2265
         TabIndex        =   8
         Top             =   600
         Width           =   4335
         _ExtentX        =   7646
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
         Container       =   "FrmRengistroDireccion.frx":2A29
      End
      Begin CATControls.CATTextBox txtCod_Prov 
         Height          =   315
         Left            =   1290
         TabIndex        =   9
         Top             =   990
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
         Container       =   "FrmRengistroDireccion.frx":2A45
         Estilo          =   1
         EnterTab        =   -1  'True
      End
      Begin CATControls.CATTextBox txtGls_Prov 
         Height          =   315
         Left            =   2265
         TabIndex        =   10
         Top             =   990
         Width           =   4335
         _ExtentX        =   7646
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
         Container       =   "FrmRengistroDireccion.frx":2A61
      End
      Begin CATControls.CATTextBox txtCod_Distrito 
         Height          =   315
         Left            =   1290
         TabIndex        =   11
         Tag             =   "TidDistrito"
         Top             =   1365
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
         Container       =   "FrmRengistroDireccion.frx":2A7D
         Estilo          =   1
         EnterTab        =   -1  'True
      End
      Begin CATControls.CATTextBox txtGls_Distrito 
         Height          =   315
         Left            =   2265
         TabIndex        =   12
         Top             =   1365
         Width           =   4335
         _ExtentX        =   7646
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
         Container       =   "FrmRengistroDireccion.frx":2A99
      End
      Begin CATControls.CATTextBox txtGls_Direccion 
         Height          =   315
         Left            =   1290
         TabIndex        =   17
         Tag             =   "Tdireccion"
         Top             =   1740
         Width           =   5715
         _ExtentX        =   10081
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
         MaxLength       =   255
         Container       =   "FrmRengistroDireccion.frx":2AB5
         Estilo          =   1
         EnterTab        =   -1  'True
      End
      Begin VB.Label Label10 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         Caption         =   "Dirección"
         ForeColor       =   &H80000008&
         Height          =   210
         Left            =   120
         TabIndex        =   18
         Top             =   1770
         Width           =   675
      End
      Begin VB.Label Label8 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         Caption         =   "Provincia"
         ForeColor       =   &H80000008&
         Height          =   210
         Left            =   120
         TabIndex        =   16
         Top             =   960
         Width           =   660
      End
      Begin VB.Label Label7 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         Caption         =   "Departamento"
         ForeColor       =   &H80000008&
         Height          =   210
         Left            =   120
         TabIndex        =   15
         Top             =   585
         Width           =   1005
      End
      Begin VB.Label Label5 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         Caption         =   "País"
         ForeColor       =   &H80000008&
         Height          =   210
         Left            =   120
         TabIndex        =   14
         Top             =   210
         Width           =   300
      End
      Begin VB.Label Label2 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         Caption         =   "Distrito"
         ForeColor       =   &H80000008&
         Height          =   210
         Left            =   120
         TabIndex        =   13
         Top             =   1410
         Width           =   495
      End
   End
End
Attribute VB_Name = "FrmRengistroDireccion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim xStrdireccion As String, xStrUbigeo As String
Dim zDireccion    As String
Dim zUbigeo       As String
Dim zStrCliente   As String

Private Sub cmbAyudaDepa_Click()
    
    mostrarAyuda "DEPARTAMENTO", txtCod_Depa, txtGls_Depa, " AND idPais = '" & txtCod_Pais.Text & "'"
    'If txtCod_Depa.Text <> "" Then SendKeys "{tab}"

End Sub

Private Sub cmbAyudaDistrito_Click()
    mostrarAyuda "DISTRITO", txtCod_Distrito, txtGls_Distrito, "AND idPais = '" & txtCod_Pais.Text & "' AND idDpto = '" & txtCod_Depa.Text & "' and idProv = '" + txtCod_Prov.Text + "'"
    'If txtCod_Distrito.Text <> "" Then SendKeys "{tab}"

End Sub

Private Sub cmbAyudaPais_Click()
    mostrarAyuda "PAIS", txtCod_Pais, txtGls_Pais
    'If txtCod_Pais.Text <> "" Then SendKeys "{tab}"

End Sub

Private Sub cmbAyudaProv_Click()
    mostrarAyuda "PROVINCIA", txtCod_Prov, txtGls_Prov, "AND idPais = '" & txtCod_Pais.Text & "' AND idDpto = '" & txtCod_Depa.Text + "'"
    'If txtCod_Prov.Text <> "" Then SendKeys "{tab}"

End Sub

Private Sub Command1_Click(Index As Integer)
Dim StrMsgError As String
On Error GoTo Err
    
    xStrUbigeo = Trim("" & txtCod_Distrito.Text)
    xStrdireccion = Trim("" & txtGls_Direccion.Text)
    If Len(xStrUbigeo) <> 6 Then
        StrMsgError = "Ubigeo invalido, Verifique": GoTo Err
    End If
    
    If Trim("" & txtGls_Direccion.Text) = "" Then
        StrMsgError = "Dirección invalido, Verifique": GoTo Err
    End If
    
    Me.Hide
    
Exit Sub
Err:
    If StrMsgError = "" Then StrMsgError = Err.Description
    MsgBox StrMsgError, vbInformation, App.Title
End Sub

Private Sub Form_Load()
Dim StrMsgError As String
On Error GoTo Err

   ConfGrid gLista, False, False, False, False
   txtCod_Pais.Text = "02001"
   
   If Trim("" & zUbigeo) <> "" Then
        txtCod_Depa.Text = left(Trim("" & zUbigeo), 2)
        txtCod_Prov.Text = Mid(Trim("" & zUbigeo), 3, 2)
        txtCod_Distrito.Text = Trim("" & zUbigeo)
        txtGls_Direccion.Text = Trim("" & zDireccion)
   Else
        txtGls_Direccion.Text = ""
        txtCod_Depa.Text = ""
        txtCod_Prov.Text = ""
        txtCod_Distrito.Text = ""
   End If
   
   
    listaDireccion StrMsgError
    If StrMsgError <> "" Then GoTo Err
    
Exit Sub
Err:
    If StrMsgError = "" Then StrMsgError = Err.Description
    MsgBox StrMsgError, vbInformation, App.Title
End Sub

Private Sub listaDireccion(ByRef StrMsgError As String)
Dim rsdatos As New ADODB.Recordset
On Error GoTo Err

csql = "EXEC Spu_TraeDirecciones_GR '" & zStrCliente & "','" & glsEmpresa & "'"

If rsdatos.State = 1 Then rsdatos.Close: Set rsdatos = Nothing
rsdatos.Open csql, Cn, adOpenStatic, adLockOptimistic
Set gLista.DataSource = rsdatos

Me.Refresh
Exit Sub
Err:
    If StrMsgError = "" Then StrMsgError = Err.Description

End Sub

Public Sub MostrarForm(Strdireccion As String, StrUbigeo As String, StrCliente As String, StrMsgError As String)
On Error GoTo Err
    
    zUbigeo = StrUbigeo
    zDireccion = Strdireccion
    zStrCliente = StrCliente
    
    Me.Show 1
    
    Strdireccion = xStrdireccion
    StrUbigeo = xStrUbigeo
    
Exit Sub
Err:
    If StrMsgError = "" Then StrMsgError = Err.Description
End Sub

Private Sub gLista_OnDblClick()
On Error GoTo Err
Dim StrMsgError As String

   If Trim("" & gLista.Columns.ColumnByName("idUbigeo").Value) <> "" Then
        txtCod_Depa.Text = left(Trim("" & Trim("" & gLista.Columns.ColumnByName("idUbigeo").Value)), 2)
        txtCod_Prov.Text = Mid(Trim("" & Trim("" & gLista.Columns.ColumnByName("idUbigeo").Value)), 3, 2)
        txtCod_Distrito.Text = Trim("" & Trim("" & gLista.Columns.ColumnByName("idUbigeo").Value))
        txtGls_Direccion.Text = Trim("" & Trim("" & gLista.Columns.ColumnByName("glsdireccion").Value))
   Else
        txtGls_Direccion.Text = ""
        txtCod_Depa.Text = ""
        txtCod_Prov.Text = ""
        txtCod_Distrito.Text = ""
   End If
    
Exit Sub
Err:
    MsgBox StrMsgError, vbInformation, App.Title
End Sub


Private Sub txtCod_Depa_Change()
    txtGls_Depa.Text = traerCampo("Ubigeo", "GlsUbigeo", "idDpto", txtCod_Depa.Text, False, " idProv = '00' And idPais = '" & txtCod_Pais.Text & "' ")
    txtCod_Prov.Text = ""
    txtGls_Prov.Text = ""
    txtCod_Distrito.Text = ""
    txtGls_Distrito.Text = ""
End Sub

Private Sub txtCod_Depa_KeyPress(KeyAscii As Integer)
    If KeyAscii <> 13 Then
        mostrarAyudaKeyascii KeyAscii, "DEPARTAMENTO", txtCod_Depa, txtGls_Depa
        KeyAscii = 0
        If txtCod_Depa.Text <> "" Then SendKeys "{tab}"
    End If
End Sub

Private Sub txtCod_Distrito_Change()
    txtGls_Distrito.Text = traerCampo("Ubigeo", "GlsUbigeo", "idDistrito", txtCod_Distrito.Text, False, "idPais = '" & txtCod_Pais.Text & "'")
    If indCopiaDireccion Then
        indEditando = True
        txtGls_DireccionFiscal.Text = txtGls_Direccion.Text & "/" & txtGls_Distrito.Text & "/" & txtGls_Prov.Text & "/" & txtGls_Depa.Text
        indEditando = False
    End If

End Sub

Private Sub txtCod_Distrito_KeyPress(KeyAscii As Integer)
    If KeyAscii <> 13 Then
        mostrarAyudaKeyascii KeyAscii, "DISTRITO", txtCod_Distrito, txtGls_Distrito, "AND idDpto = '" & txtCod_Depa.Text & "' and idProv = '" + txtCod_Prov.Text + "'"
        KeyAscii = 0
        If txtCod_Distrito.Text <> "" Then SendKeys "{tab}"
    End If
End Sub

Private Sub txtCod_Pais_Change()
On Error GoTo Err
Dim StrMsgError                 As String

    txtGls_Pais.Text = traerCampo("Datos", "GlsDato", "idDato", txtCod_Pais.Text, False)
    txtCod_Depa.Text = ""
    txtGls_Depa.Text = ""
    txtCod_Prov.Text = ""
    txtGls_Prov.Text = ""
    txtCod_Distrito.Text = ""
    txtGls_Distrito.Text = ""
    
    If txtGls_Pais.Text = "" Or traerCampo("Datos", "IdSunat", "IdDato", txtCod_Pais.Text, False) = "9589" Then
        txtCod_Depa.Vacio = False
        txtGls_Depa.Vacio = False
        txtCod_Prov.Vacio = False
        txtGls_Prov.Vacio = False
        txtCod_Distrito.Vacio = False
        txtGls_Distrito.Vacio = False
    Else
        txtCod_Depa.Vacio = True
        txtGls_Depa.Vacio = True
        txtCod_Prov.Vacio = True
        txtGls_Prov.Vacio = True
        txtCod_Distrito.Vacio = True
        txtGls_Distrito.Vacio = True
    End If
    
    Exit Sub
Err:
    If StrMsgError = "" Then StrMsgError = Err.Description
    MsgBox StrMsgError, vbInformation, App.Title
End Sub

Private Sub txtCod_Pais_KeyPress(KeyAscii As Integer)
    If KeyAscii <> 13 Then
        mostrarAyudaKeyascii KeyAscii, "PAIS", txtCod_Pais, txtGls_Pais
        KeyAscii = 0
        If txtCod_Pais.Text <> "" Then SendKeys "{tab}"
    End If
End Sub

Private Sub txtCod_Prov_Change()
    txtGls_Prov.Text = traerCampo("Ubigeo", "GlsUbigeo", "idProv", txtCod_Prov.Text, False, " idDpto = '" & txtCod_Depa.Text & "' and idProv <> '00' and idDist = '00' And idPais = '" & txtCod_Pais.Text & "' ")
    txtCod_Distrito.Text = ""
    txtGls_Distrito.Text = ""
    If indCopiaDireccion Then
        indEditando = True
        txtGls_DireccionFiscal.Text = txtGls_Direccion.Text & "/" & txtGls_Distrito.Text & "/" & txtGls_Prov.Text & "/" & txtGls_Depa.Text
        indEditando = False
    End If
End Sub

Private Sub txtCod_Prov_KeyPress(KeyAscii As Integer)
    If KeyAscii <> 13 Then
        mostrarAyudaKeyascii KeyAscii, "PROVINCIA", txtCod_Prov, txtGls_Prov, "AND idDpto = '" & txtCod_Depa.Text + "'"
        KeyAscii = 0
        If txtCod_Prov.Text <> "" Then SendKeys "{tab}"
    End If
End Sub
