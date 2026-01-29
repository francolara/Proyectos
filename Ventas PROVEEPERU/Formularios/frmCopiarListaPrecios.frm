VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{F41D1D30-7878-4923-8CB3-6CCACDC9C9DE}#1.0#0"; "catcontrols.ocx"
Begin VB.Form frmCopiarListaPrecios 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Copiar Lista de Precios"
   ClientHeight    =   2880
   ClientLeft      =   3930
   ClientTop       =   1890
   ClientWidth     =   8535
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2880
   ScaleWidth      =   8535
   ShowInTaskbar   =   0   'False
   Begin VB.Frame fraCabecera 
      Appearance      =   0  'Flat
      ForeColor       =   &H00C00000&
      Height          =   2100
      Left            =   60
      TabIndex        =   1
      Top             =   720
      Width           =   8385
      Begin VB.Frame Frame1 
         Appearance      =   0  'Flat
         Caption         =   "Datos de la nueva lista"
         ForeColor       =   &H00C00000&
         Height          =   1140
         Left            =   60
         TabIndex        =   6
         Top             =   840
         Width           =   8265
         Begin CATControls.CATTextBox txtGls_Lista 
            Height          =   315
            Left            =   1335
            TabIndex        =   7
            Tag             =   "TglsLista"
            Top             =   240
            Width           =   6840
            _ExtentX        =   12065
            _ExtentY        =   556
            BackColor       =   16777215
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
            MaxLength       =   255
            Container       =   "frmCopiarListaPrecios.frx":0000
            Estilo          =   1
            EnterTab        =   -1  'True
         End
         Begin MSComCtl2.DTPicker dtp_Vcto 
            Height          =   315
            Left            =   1320
            TabIndex        =   8
            Tag             =   "FfecVcto"
            Top             =   675
            Width           =   1515
            _ExtentX        =   2672
            _ExtentY        =   556
            _Version        =   393216
            Format          =   50069505
            CurrentDate     =   38955
         End
         Begin VB.Label Label2 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            Caption         =   "Fecha Vcto:"
            ForeColor       =   &H80000008&
            Height          =   195
            Left            =   120
            TabIndex        =   10
            Top             =   705
            Width           =   870
         End
         Begin VB.Label Label1 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            Caption         =   "Descripcion:"
            ForeColor       =   &H80000008&
            Height          =   195
            Left            =   120
            TabIndex        =   9
            Top             =   285
            Width           =   885
         End
      End
      Begin VB.CommandButton cmbAyudaListaCopiar 
         Height          =   315
         Left            =   7845
         Picture         =   "frmCopiarListaPrecios.frx":001C
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   360
         Width           =   390
      End
      Begin CATControls.CATTextBox txtCod_ListaCopiar 
         Height          =   315
         Left            =   1350
         TabIndex        =   3
         Tag             =   "TidLista"
         Top             =   360
         Width           =   915
         _ExtentX        =   1614
         _ExtentY        =   556
         BackColor       =   16777215
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
         Container       =   "frmCopiarListaPrecios.frx":03A6
         Estilo          =   1
         EnterTab        =   -1  'True
      End
      Begin CATControls.CATTextBox txtGls_ListaCopiar 
         Height          =   315
         Left            =   2325
         TabIndex        =   4
         Top             =   360
         Width           =   5460
         _ExtentX        =   9631
         _ExtentY        =   556
         BackColor       =   16777152
         Enabled         =   0   'False
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
         Container       =   "frmCopiarListaPrecios.frx":03C2
         Vacio           =   -1  'True
      End
      Begin VB.Label lbl_Lista 
         Appearance      =   0  'Flat
         Caption         =   "Lista a copiar:"
         ForeColor       =   &H80000007&
         Height          =   240
         Left            =   135
         TabIndex        =   5
         Top             =   420
         Width           =   1125
      End
   End
   Begin MSComctlLib.ImageList imgDocVentas 
      Left            =   60
      Top             =   1080
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
            Picture         =   "frmCopiarListaPrecios.frx":03DE
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCopiarListaPrecios.frx":0778
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCopiarListaPrecios.frx":0BCA
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCopiarListaPrecios.frx":0F64
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCopiarListaPrecios.frx":12FE
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCopiarListaPrecios.frx":1698
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCopiarListaPrecios.frx":1A32
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCopiarListaPrecios.frx":1DCC
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCopiarListaPrecios.frx":2166
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCopiarListaPrecios.frx":2500
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCopiarListaPrecios.frx":289A
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCopiarListaPrecios.frx":355C
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   660
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   8535
      _ExtentX        =   15055
      _ExtentY        =   1164
      ButtonWidth     =   1111
      ButtonHeight    =   1005
      Appearance      =   1
      ImageList       =   "imgDocVentas"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   3
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Grabar"
            Object.ToolTipText     =   "Grabar"
            ImageIndex      =   4
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
            Caption         =   "Cancelar"
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
End
Attribute VB_Name = "frmCopiarListaPrecios"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmbAyudaListaCopiar_Click()
    mostrarAyuda "LISTAPRECIOS", txtCod_ListaCopiar, txtGls_ListaCopiar
    If txtCod_ListaCopiar.Text <> "" Then SendKeys "{tab}"
End Sub

Private Sub Form_Load()
dtp_Vcto.Value = Format(getFechaSistema, "dd/mm/yyyy")
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
Dim StrMsgError As String
On Error GoTo Err
Select Case Button.Index
    Case 1 'Grabar
    
        Grabar StrMsgError
        If StrMsgError <> "" Then GoTo Err
        
        Unload Me
        
    Case 3 'Salir
        Unload Me
End Select

Exit Sub
Err:
If StrMsgError = "" Then StrMsgError = Err.Description
MsgBox StrMsgError, vbInformation, App.Title
End Sub

Private Sub txtCod_ListaCopiar_Change()
    txtGls_ListaCopiar.Text = traerCampo("listaprecios", "GlsLista", "idLista", txtCod_ListaCopiar.Text, True)
End Sub

Private Sub txtCod_ListaCopiar_KeyPress(KeyAscii As Integer)
If KeyAscii <> 13 Then
    mostrarAyudaKeyascii KeyAscii, "LISTAPRECIOS", txtCod_ListaCopiar, txtGls_ListaCopiar
    KeyAscii = 0
    If txtCod_ListaCopiar.Text <> "" Then SendKeys "{tab}"
End If
End Sub

Private Sub Grabar(ByRef StrMsgError As String)
Dim strCodLista As String
Dim indTrans As Boolean

On Error GoTo Err

If txtCod_ListaCopiar.Text = "" Then
    StrMsgError = "Seleccione una lista a copiar"
    txtCod_ListaCopiar.OnError = True
    GoTo Err
End If

If txtGls_Lista.Text = "" Then
    StrMsgError = "Ingresa la descripcio a la lista a generar"
    txtCod_ListaCopiar.OnError = True
    GoTo Err
End If

strCodLista = GeneraCorrelativoAnoMes("listaprecios", "idLista")

indTrans = True
Cn.BeginTrans

'Insertamos la lista
csql = "INSERT INTO listaprecios (idEmpresa,idLista,glsLista,estLista,FecVcto) VALUES('" & _
                                  glsEmpresa & "','" & strCodLista & "','" & txtGls_Lista.Text & "',1,'" & Format(dtp_Vcto.Value, "yyyy-mm-dd") & "')"

Cn.Execute csql


'Insertamos los precios de la lista
csql = "INSERT INTO preciosventa (idEmpresa,idLista,idProducto,idUM,VVUnit,IGVUnit,PVUnit) " & _
       "SELECT idEmpresa,'" & strCodLista & "',idProducto,idUM,VVUnit,IGVUnit,PVUnit FROM preciosventa " & _
       "WHERE idEmpresa = '" & glsEmpresa & "' AND idLista = '" & txtCod_ListaCopiar.Text & "'"
       
Cn.Execute csql

Cn.CommitTrans
Exit Sub
Err:
If StrMsgError = "" Then StrMsgError = Err.Description
If indTrans Then Cn.RollbackTrans
End Sub
