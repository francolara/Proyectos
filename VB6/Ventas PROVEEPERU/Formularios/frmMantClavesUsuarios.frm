VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{F41D1D30-7878-4923-8CB3-6CCACDC9C9DE}#1.0#0"; "catcontrols.ocx"
Begin VB.Form frmMantClavesUsuarios 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Su contraseña ha caducado. Actualice su nueva contraseña."
   ClientHeight    =   3825
   ClientLeft      =   5730
   ClientTop       =   3090
   ClientWidth     =   6225
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3825
   ScaleWidth      =   6225
   Begin VB.Frame Frm 
      Appearance      =   0  'Flat
      ForeColor       =   &H80000008&
      Height          =   3075
      Left            =   45
      TabIndex        =   0
      Top             =   675
      Width           =   6135
      Begin VB.Frame Frame1 
         Appearance      =   0  'Flat
         ForeColor       =   &H80000008&
         Height          =   2220
         Left            =   360
         TabIndex        =   5
         Top             =   630
         Width           =   5415
         Begin CATControls.CATTextBox txt_Usu 
            Height          =   315
            Left            =   2800
            TabIndex        =   6
            Tag             =   "TvarUsuario"
            Top             =   270
            Width           =   2010
            _ExtentX        =   3545
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
            MaxLength       =   20
            Container       =   "frmMantClavesUsuarios.frx":0000
            EnterTab        =   -1  'True
         End
         Begin CATControls.CATTextBox txt_PassAnt 
            Height          =   315
            Left            =   2800
            TabIndex        =   1
            Tag             =   "TvarPass"
            Top             =   675
            Width           =   2010
            _ExtentX        =   3545
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
            MaxLength       =   20
            PasswordChar    =   "X"
            Container       =   "frmMantClavesUsuarios.frx":001C
            EnterTab        =   -1  'True
         End
         Begin CATControls.CATTextBox txt_PassNuevo 
            Height          =   315
            Left            =   2800
            TabIndex        =   2
            Tag             =   "TvarPass"
            Top             =   1125
            Width           =   2010
            _ExtentX        =   3545
            _ExtentY        =   556
            BackColor       =   16777215
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            FontBold        =   -1  'True
            FontName        =   "Arial"
            FontSize        =   8.25
            ForeColor       =   -2147483640
            MaxLength       =   20
            PasswordChar    =   "X"
            Container       =   "frmMantClavesUsuarios.frx":0038
            EnterTab        =   -1  'True
         End
         Begin CATControls.CATTextBox txt_PassAntNuevoN 
            Height          =   315
            Left            =   2800
            TabIndex        =   3
            Tag             =   "TvarPass"
            Top             =   1530
            Width           =   2010
            _ExtentX        =   3545
            _ExtentY        =   556
            BackColor       =   16777215
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            FontBold        =   -1  'True
            FontName        =   "Arial"
            FontSize        =   8.25
            ForeColor       =   -2147483640
            MaxLength       =   20
            PasswordChar    =   "X"
            Container       =   "frmMantClavesUsuarios.frx":0054
            EnterTab        =   -1  'True
         End
         Begin VB.Label Label11 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            Caption         =   "Contraseña Actual"
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
            Left            =   585
            TabIndex        =   10
            Top             =   765
            Width           =   1350
         End
         Begin VB.Label Label9 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            Caption         =   "Usuario"
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
            Left            =   585
            TabIndex        =   9
            Top             =   360
            Width           =   555
         End
         Begin VB.Label Label1 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            Caption         =   "Nueva contraseña"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   210
            Left            =   585
            TabIndex        =   8
            Top             =   1185
            Width           =   1470
         End
         Begin VB.Label Label2 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            Caption         =   "Repetir Contraseña"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   210
            Left            =   585
            TabIndex        =   7
            Top             =   1590
            Width           =   1605
         End
      End
      Begin CATControls.CATTextBox txtGls_Persona 
         Height          =   315
         Left            =   360
         TabIndex        =   11
         Top             =   270
         Width           =   5430
         _ExtentX        =   9578
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
         Alignment       =   2
         FontName        =   "Arial"
         FontSize        =   8.25
         ForeColor       =   -2147483640
         Container       =   "frmMantClavesUsuarios.frx":0070
      End
      Begin CATControls.CATTextBox txtCod_Persona 
         Height          =   315
         Left            =   90
         TabIndex        =   12
         Tag             =   "TidUsuario"
         Top             =   1800
         Visible         =   0   'False
         Width           =   330
         _ExtentX        =   582
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
         Container       =   "frmMantClavesUsuarios.frx":008C
         Estilo          =   1
         EnterTab        =   -1  'True
      End
   End
   Begin MSComctlLib.ImageList imgDocVentas 
      Left            =   180
      Top             =   7260
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
            Picture         =   "frmMantClavesUsuarios.frx":00A8
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMantClavesUsuarios.frx":0442
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMantClavesUsuarios.frx":0894
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMantClavesUsuarios.frx":0C2E
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMantClavesUsuarios.frx":0FC8
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMantClavesUsuarios.frx":1362
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMantClavesUsuarios.frx":16FC
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMantClavesUsuarios.frx":1A96
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMantClavesUsuarios.frx":1E30
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMantClavesUsuarios.frx":21CA
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMantClavesUsuarios.frx":2564
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMantClavesUsuarios.frx":3226
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   660
      Left            =   0
      TabIndex        =   4
      Top             =   0
      Width           =   6225
      _ExtentX        =   10980
      _ExtentY        =   1164
      ButtonWidth     =   1905
      ButtonHeight    =   1005
      Appearance      =   1
      ImageList       =   "imgDocVentas"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   1
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "     Grabar     "
            Object.ToolTipText     =   "Grabar"
            ImageIndex      =   4
         EndProperty
      EndProperty
      BorderStyle     =   1
   End
End
Attribute VB_Name = "frmMantClavesUsuarios"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Form_Load()

    txtCod_Persona.Text = glsUser
    txt_PassAnt.Text = ""
    txt_PassNuevo.Text = ""
    txt_PassAntNuevoN.Text = ""

End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
On Error GoTo Err
Dim StrMsgError As String

    Select Case Button.Index
        Case 1: 'Grabar
            If traerCampo("usuarios", "varPass", "idUsuario", txtCod_Persona.Text, True) <> txt_PassAnt.Text Then
                txt_PassAnt.Text = ""
                txt_PassAnt.OnError = True
                StrMsgError = "Los Datos son Incorrectos. Verifique."
                GoTo Err
            End If
            
            If txt_PassNuevo.Text <> txt_PassAntNuevoN.Text Then
                txt_PassAntNuevoN.Text = ""
                StrMsgError = "Los Datos son Incorrectos. Verifique."
                txt_PassAntNuevoN.OnError = True
                GoTo Err
            End If
            
            Grabar StrMsgError
            If StrMsgError <> "" Then GoTo Err
            
            MsgBox "Su contraseña ha sido Modificada Satisfactoriamente.", vbInformation, App.Title
                    
            Unload Me
            frmPrincipal.Show
    End Select

    Exit Sub
    
Err:
    If StrMsgError = "" Then StrMsgError = Err.Description
    MsgBox StrMsgError, vbInformation, App.Title

End Sub

Private Sub txt_PassAnt_LostFocus()
On Error GoTo Err
Dim StrMsgError As String

    If traerCampo("usuarios", "varPass", "idUsuario", txtCod_Persona.Text, True) <> txt_PassAnt.Text Then
        txt_PassAnt.Text = ""
        txt_PassAnt.OnError = True
    End If
    
    Exit Sub

Err:
    If StrMsgError = "" Then StrMsgError = Err.Description
    MsgBox StrMsgError, vbInformation, App.Title

End Sub

Private Sub txt_PassAntNuevoN_LostFocus()
On Error GoTo Err
Dim StrMsgError As String

    If txt_PassNuevo.Text <> txt_PassAntNuevoN.Text Then
        txt_PassAntNuevoN.Text = ""
        txt_PassAntNuevoN.OnError = True
    End If
    
    Exit Sub

Err:
    If StrMsgError = "" Then StrMsgError = Err.Description
    MsgBox StrMsgError, vbInformation, App.Title

End Sub

Private Sub txtCod_Persona_Change()
    
    txtGls_Persona.Text = Trim("" & traerCampo("personas", "GlsPersona", "idPersona", txtCod_Persona.Text, False))
    txt_Usu.Text = Trim("" & traerCampo("usuarios", "varUsuario", "idUsuario", txtCod_Persona.Text, True))

End Sub

Private Sub Grabar(StrMsgError As String)
On Error GoTo Err
Dim Cadmysql    As String

    Cadmysql = "Update Usuarios set varPass = '" & txt_PassAntNuevoN.Text & "' , FecModClave = (CAST(CONCAT(CURDATE(),' ',CURTIME()) AS DATETIME)) " & _
               "where idempresa = '" & glsEmpresa & "' and idUsuario = '" & glsUser & "' "
    Cn.Execute Cadmysql
    
    Exit Sub

Err:
    If StrMsgError = "" Then StrMsgError = Err.Description

End Sub
