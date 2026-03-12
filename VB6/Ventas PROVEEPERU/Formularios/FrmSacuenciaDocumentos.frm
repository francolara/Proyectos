VERSION 5.00
Object = "{6A24B331-7634-11D3-A5B0-0050044A7E1A}#1.5#0"; "DXDBGrid.dll"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{F41D1D30-7878-4923-8CB3-6CCACDC9C9DE}#1.0#0"; "catcontrols.ocx"
Begin VB.Form FrmSecuenciaDocumentos 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Consulta de Documentos"
   ClientHeight    =   9090
   ClientLeft      =   2550
   ClientTop       =   945
   ClientWidth     =   12750
   LinkTopic       =   "Form3"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   9090
   ScaleWidth      =   12750
   ShowInTaskbar   =   0   'False
   Begin VB.Frame FrmSecuencia 
      Height          =   8430
      Left            =   90
      TabIndex        =   0
      Top             =   585
      Width           =   12600
      Begin VB.Frame FraLista1 
         Height          =   3435
         Left            =   1980
         TabIndex        =   12
         Top             =   1080
         Visible         =   0   'False
         Width           =   8205
         Begin DXDBGRIDLibCtl.dxDBGrid gdetalle1 
            Height          =   3120
            Left            =   135
            OleObjectBlob   =   "FrmSacuenciaDocumentos.frx":0000
            TabIndex        =   16
            Top             =   180
            Width           =   7950
         End
      End
      Begin VB.Frame FraLista2 
         Height          =   3300
         Left            =   1980
         TabIndex        =   13
         Top             =   1980
         Visible         =   0   'False
         Width           =   8205
         Begin DXDBGRIDLibCtl.dxDBGrid gdetalle2 
            Height          =   2895
            Left            =   90
            OleObjectBlob   =   "FrmSacuenciaDocumentos.frx":3745
            TabIndex        =   17
            Top             =   315
            Width           =   7995
         End
      End
      Begin VB.Frame FraLista3 
         Height          =   3165
         Left            =   1980
         TabIndex        =   15
         Top             =   3195
         Visible         =   0   'False
         Width           =   8205
         Begin DXDBGRIDLibCtl.dxDBGrid gdetalle3 
            Height          =   2850
            Left            =   90
            OleObjectBlob   =   "FrmSacuenciaDocumentos.frx":6E8A
            TabIndex        =   18
            Top             =   180
            Width           =   8040
         End
      End
      Begin VB.Frame FraLista4 
         Height          =   3255
         Left            =   1980
         TabIndex        =   14
         Top             =   4770
         Visible         =   0   'False
         Width           =   8205
         Begin DXDBGRIDLibCtl.dxDBGrid gDetalle4 
            Height          =   2895
            Left            =   135
            OleObjectBlob   =   "FrmSacuenciaDocumentos.frx":A5CF
            TabIndex        =   19
            Top             =   270
            Width           =   7950
         End
      End
      Begin VB.Frame Frame1 
         Height          =   870
         Left            =   90
         TabIndex        =   1
         Top             =   135
         Width           =   12390
         Begin VB.CommandButton cmbAyudaTipoDoc 
            Height          =   315
            Left            =   6495
            Picture         =   "FrmSacuenciaDocumentos.frx":DD14
            Style           =   1  'Graphical
            TabIndex        =   2
            Top             =   345
            Width           =   390
         End
         Begin CATControls.CATTextBox txtCod_TipoDoc 
            Height          =   315
            Left            =   1395
            TabIndex        =   3
            Tag             =   "TidMoneda"
            Top             =   345
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
            Alignment       =   2
            FontName        =   "Arial"
            FontSize        =   8.25
            ForeColor       =   -2147483640
            MaxLength       =   2
            Container       =   "FrmSacuenciaDocumentos.frx":E09E
            Estilo          =   3
            EnterTab        =   -1  'True
         End
         Begin CATControls.CATTextBox txtGls_TipoDoc 
            Height          =   315
            Left            =   2340
            TabIndex        =   4
            Top             =   345
            Width           =   4140
            _ExtentX        =   7303
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
            Container       =   "FrmSacuenciaDocumentos.frx":E0BA
            Vacio           =   -1  'True
         End
         Begin CATControls.CATTextBox txt_Serie 
            Height          =   315
            Left            =   8550
            TabIndex        =   20
            Top             =   345
            Width           =   960
            _ExtentX        =   1693
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
            Container       =   "FrmSacuenciaDocumentos.frx":E0D6
            Estilo          =   1
            EnterTab        =   -1  'True
         End
         Begin CATControls.CATTextBox txtNum_Documento 
            Height          =   315
            Left            =   10275
            TabIndex        =   21
            Top             =   345
            Width           =   1860
            _ExtentX        =   3281
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
            Container       =   "FrmSacuenciaDocumentos.frx":E0F2
            Estilo          =   1
            EnterTab        =   -1  'True
         End
         Begin VB.Label Label3 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            Caption         =   "Número"
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
            Left            =   9630
            TabIndex        =   7
            Top             =   405
            Width           =   555
         End
         Begin VB.Label Label1 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            Caption         =   "Serie"
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
            Left            =   8100
            TabIndex        =   6
            Top             =   405
            Width           =   375
         End
         Begin VB.Label Label2 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            Caption         =   "Tipo Documento"
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
            Left            =   165
            TabIndex        =   5
            Top             =   400
            Width           =   1155
         End
      End
      Begin DXDBGRIDLibCtl.dxDBGrid gLista1 
         Height          =   1725
         Left            =   90
         OleObjectBlob   =   "FrmSacuenciaDocumentos.frx":E10E
         TabIndex        =   8
         Top             =   1080
         Width           =   12360
      End
      Begin DXDBGRIDLibCtl.dxDBGrid gLista2 
         Height          =   1725
         Left            =   90
         OleObjectBlob   =   "FrmSacuenciaDocumentos.frx":1351B
         TabIndex        =   9
         Top             =   2925
         Width           =   12360
      End
      Begin DXDBGRIDLibCtl.dxDBGrid gLista3 
         Height          =   1725
         Left            =   90
         OleObjectBlob   =   "FrmSacuenciaDocumentos.frx":18928
         TabIndex        =   10
         Top             =   4770
         Width           =   12360
      End
      Begin DXDBGRIDLibCtl.dxDBGrid gLista4 
         Height          =   1725
         Left            =   90
         OleObjectBlob   =   "FrmSacuenciaDocumentos.frx":1DD35
         TabIndex        =   11
         Top             =   6570
         Width           =   12360
      End
      Begin MSComctlLib.ImageList imgDocVentas 
         Left            =   0
         Top             =   0
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
               Picture         =   "FrmSacuenciaDocumentos.frx":23142
               Key             =   ""
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmSacuenciaDocumentos.frx":234DC
               Key             =   ""
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmSacuenciaDocumentos.frx":2392E
               Key             =   ""
            EndProperty
            BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmSacuenciaDocumentos.frx":23CC8
               Key             =   ""
            EndProperty
            BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmSacuenciaDocumentos.frx":24062
               Key             =   ""
            EndProperty
            BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmSacuenciaDocumentos.frx":243FC
               Key             =   ""
            EndProperty
            BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmSacuenciaDocumentos.frx":24796
               Key             =   ""
            EndProperty
            BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmSacuenciaDocumentos.frx":24B30
               Key             =   ""
            EndProperty
            BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmSacuenciaDocumentos.frx":24ECA
               Key             =   ""
            EndProperty
            BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmSacuenciaDocumentos.frx":25264
               Key             =   ""
            EndProperty
            BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmSacuenciaDocumentos.frx":255FE
               Key             =   ""
            EndProperty
            BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmSacuenciaDocumentos.frx":262C0
               Key             =   ""
            EndProperty
            BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmSacuenciaDocumentos.frx":2665A
               Key             =   ""
            EndProperty
            BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmSacuenciaDocumentos.frx":26AAC
               Key             =   ""
            EndProperty
            BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmSacuenciaDocumentos.frx":26E46
               Key             =   ""
            EndProperty
            BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmSacuenciaDocumentos.frx":27858
               Key             =   ""
            EndProperty
         EndProperty
      End
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Height          =   660
      Left            =   45
      TabIndex        =   22
      Top             =   0
      Width           =   12645
      _ExtentX        =   22304
      _ExtentY        =   1164
      ButtonWidth     =   1402
      ButtonHeight    =   1005
      Appearance      =   1
      ImageList       =   "imgDocVentas"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   3
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Imprimir"
            Object.ToolTipText     =   "Imprimir"
            ImageIndex      =   11
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Object.Visible         =   0   'False
            Caption         =   "Excel"
            ImageIndex      =   13
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
Attribute VB_Name = "FrmSecuenciaDocumentos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub SecuenciaCotizacion()
Dim csql                         As String
Dim rst                          As New ADODB.Recordset
Dim rsg                          As New ADODB.Recordset
Dim rsDetalle                    As New ADODB.Recordset
Dim rsdetalleCot                 As New ADODB.Recordset
Dim rsdetallePed                 As New ADODB.Recordset
Dim StrMsgError                  As String
Dim item                         As Integer
Dim rsdetalleFacGuia             As New ADODB.Recordset
Dim rsdetFacGuia                 As New ADODB.Recordset
Dim rstOriFacGuia                As New ADODB.Recordset


    csql = " Select (@i:=@i +1) as item,d.idempresa,d.idsucursal,d.iddocumento,d.glsdocreferencia,d.glsdocreferencia,dc.abredocumento,m.simbolo,d.idserie,d.iddocventas,d.fecemision,d.estdocventas,d.idpercliente," & _
            " d.glscliente,d.idmoneda,d.totalprecioventa from docventas d inner join documentos dc on d.iddocumento = dc.iddocumento " & _
            " inner join monedas m on d.idmoneda = m.idmoneda,(select @i:=0) foo " & _
            " where d.idempresa = '" & glsEmpresa & "' and d.iddocumento = '" & txtCod_TipoDoc.Text & "' " & _
            " and d.idserie = '" & txt_serie.Text & "' and d.iddocventas = '" & txtNum_Documento.Text & "'"
              
    If rst.State = 1 Then rst.Close
    rst.Open csql, Cn, adOpenKeyset, adLockOptimistic
     
    rsg.Fields.Append "item", adInteger, , adFldRowID
    rsg.Fields.Append "idempresa", adChar, 8, adFldIsNullable
    rsg.Fields.Append "idsucursal", adChar, 8, adFldIsNullable
    rsg.Fields.Append "iddocumento", adChar, 2, adFldIsNullable
    rsg.Fields.Append "glsdocreferencia", adChar, 500, adFldIsNullable
    rsg.Fields.Append "abredocumento", adChar, 5, adFldIsNullable
    rsg.Fields.Append "simbolo", adChar, 5, adFldIsNullable
    rsg.Fields.Append "idserie", adChar, 4, adFldIsNullable
    rsg.Fields.Append "iddocventas", adChar, 8, adFldIsNullable
    rsg.Fields.Append "fecemision", adChar, 15, adFldIsNullable
    rsg.Fields.Append "estdocventas", adChar, 3, adFldIsNullable
    rsg.Fields.Append "idpercliente", adChar, 8, adFldIsNullable
    rsg.Fields.Append "glscliente", adVarChar, 100, adFldIsNullable
    rsg.Fields.Append "idmoneda", adChar, 5, adFldIsNullable
    rsg.Fields.Append "totalprecioventa", adDouble, adFldRowID
    rsg.Open
    
    If rst.RecordCount = 0 Then
        rsg.Fields("item") = 1
        rsg.Fields("idempresa") = ""
        rsg.Fields("idsucursal") = ""
        rsg.Fields("iddocumento") = ""
        rsg.Fields("glsdocreferencia") = ""
        rsg.Fields("abredocumento") = ""
        rsg.Fields("simbolo") = ""
        rsg.Fields("idserie") = ""
        rsg.Fields("iddocventas") = ""
        rsg.Fields("fecemision") = ""
        rsg.Fields("estdocventas") = ""
        rsg.Fields("idpercliente") = ""
        rsg.Fields("glscliente") = ""
        rsg.Fields("idmoneda") = ""
        rsg.Fields("totalprecioventa") = ""
                
    Else
        If Not rst.EOF Then
            rst.MoveFirst
            Do While Not rst.EOF
                rsg.AddNew
                rsg.Fields("item") = Val("" & rst.Fields("item"))
                rsg.Fields("idempresa") = Trim("" & rst.Fields("idempresa"))
                rsg.Fields("idsucursal") = Trim("" & rst.Fields("idsucursal"))
                rsg.Fields("iddocumento") = Trim("" & rst.Fields("iddocumento"))
                rsg.Fields("glsdocreferencia") = Trim("" & rst.Fields("glsdocreferencia"))
                rsg.Fields("abredocumento") = Trim("" & rst.Fields("abredocumento"))
                rsg.Fields("simbolo") = Trim("" & rst.Fields("simbolo"))
                rsg.Fields("idserie") = Trim("" & rst.Fields("idserie"))
                rsg.Fields("iddocventas") = Trim("" & rst.Fields("iddocventas"))
                rsg.Fields("fecemision") = "" & rst.Fields("fecemision")
                rsg.Fields("estdocventas") = "" & rst.Fields("estdocventas")
                rsg.Fields("idpercliente") = Trim("" & rst.Fields("idpercliente"))
                rsg.Fields("glscliente") = Trim("" & rst.Fields("glscliente"))
                rsg.Fields("idmoneda") = Trim("" & rst.Fields("idmoneda"))
                rsg.Fields("totalprecioventa") = Trim("" & rst.Fields("totalprecioventa"))
                rst.MoveNext
            Loop
        End If
        mostrarDatosGridSQL gLista1, rsg, StrMsgError
    End If
     
    csql = "select idempresa,tipodocorigen, seriedocorigen,numdocorigen from docreferencia " & _
            "where idempresa = '" & Trim("" & rsg.Fields("idempresa")) & "' and tipodocreferencia = '" & Trim("" & rsg.Fields("iddocumento")) & "' and seriedocreferencia = '" & Trim("" & rsg.Fields("idserie")) & "' and numdocreferencia = '" & Trim("" & rsg.Fields("iddocventas")) & "' "
    If rst.State = 1 Then rst.Close
    rst.Open csql, Cn, adOpenKeyset, adLockOptimistic
    If rst.RecordCount > 0 Then
        rsdetallePed.Fields.Append "Item", adInteger, , adFldRowID
        rsdetallePed.Fields.Append "idempresa", adChar, 8, adFldRowID
        rsdetallePed.Fields.Append "idsucursal", adChar, 8, adFldRowID
        rsdetallePed.Fields.Append "iddocumento", adChar, 2, adFldRowID
        rsdetallePed.Fields.Append "glsdocreferencia", adChar, 500, adFldRowID
        rsdetallePed.Fields.Append "abredocumento", adChar, 5, adFldRowID
        rsdetallePed.Fields.Append "simbolo", adChar, 5, adFldRowID
        rsdetallePed.Fields.Append "idserie", adVarChar, 4, adFldRowID
        rsdetallePed.Fields.Append "iddocventas", adChar, 8, adFldRowID
        rsdetallePed.Fields.Append "fecEmision", adVarChar, 15, adFldRowID
        rsdetallePed.Fields.Append "estdocventas", adChar, 3, adFldRowID
        rsdetallePed.Fields.Append "idpercliente", adVarChar, 8, adFldRowID
        rsdetallePed.Fields.Append "glscliente", adVarChar, 100, adFldRowID
        rsdetallePed.Fields.Append "idmoneda", adVarChar, 3, adFldRowID
        rsdetallePed.Fields.Append "totalprecioventa", adDouble, adFldRowID
        rsdetallePed.Open
                
        rst.MoveFirst
        Do While Not rst.EOF
            csql = " Select d.idempresa,d.idsucursal,d.iddocumento,d.glsdocreferencia,d.glsdocreferencia,dc.abredocumento,m.simbolo,d.idserie,d.iddocventas,d.fecemision,d.estdocventas,d.idpercliente," & _
                    " d.glscliente,d.idmoneda,d.totalprecioventa from docventas d inner join documentos dc on d.iddocumento = dc.iddocumento " & _
                    " inner join monedas m  on d.idmoneda = m.idmoneda " & _
                    " where d.idempresa = '" & Trim("" & rst.Fields("idempresa")) & "' and d.iddocumento = '" & Trim("" & rst.Fields("tipodocorigen")) & "' " & _
                    " and d.idserie = '" & Trim("" & rst.Fields("seriedocorigen")) & "' and d.iddocventas = '" & Trim("" & rst.Fields("numdocorigen")) & "'"
            If rsDetalle.State = 1 Then rsDetalle.Close
            rsDetalle.Open csql, Cn, adOpenKeyset, adLockOptimistic
            
            item = 0
            rsDetalle.MoveFirst
            Do While Not rsDetalle.EOF
                item = item + 1
                rsdetallePed.AddNew
                rsdetallePed!item = item
                rsdetallePed!idempresa = "" & rsDetalle.Fields("idempresa")
                rsdetallePed!idsucursal = "" & rsDetalle.Fields("idsucursal")
                rsdetallePed!idDocumento = "" & rsDetalle.Fields("iddocumento")
                rsdetallePed!glsdocreferencia = "" & rsDetalle.Fields("glsdocreferencia")
                rsdetallePed!abredocumento = "" & rsDetalle.Fields("abredocumento")
                rsdetallePed!simbolo = "" & rsDetalle.Fields("simbolo")
                rsdetallePed!idserie = "" & rsDetalle.Fields("idserie")
                rsdetallePed!idDocVentas = "" & rsDetalle.Fields("iddocventas")
                rsdetallePed!fecEmision = "" & rsDetalle.Fields("fecemision")
                rsdetallePed!estdocventas = "" & rsDetalle.Fields("estdocventas")
                rsdetallePed!idpercliente = "" & rsDetalle.Fields("idpercliente")
                rsdetallePed!GlsCliente = "" & rsDetalle.Fields("glscliente")
                rsdetallePed!idMoneda = "" & rsDetalle.Fields("idmoneda")
                rsdetallePed!totalprecioventa = "" & rsDetalle.Fields("totalprecioventa")
                rsdetallePed.Update
                rsDetalle.MoveNext
            Loop
            
            rst.MoveNext
        Loop
        mostrarDatosGridSQL gLista2, rsdetallePed, StrMsgError
    End If
     
    If gLista2.Count <> 0 Then
        If rsdetallePed.RecordCount > 0 Then
            item = 0
            rsdetalleFacGuia.Fields.Append "Item", adInteger, , adFldRowID
            rsdetalleFacGuia.Fields.Append "idempresa", adChar, 8, adFldRowID
            rsdetalleFacGuia.Fields.Append "idsucursal", adChar, 8, adFldRowID
            rsdetalleFacGuia.Fields.Append "iddocumento", adChar, 2, adFldRowID
            rsdetalleFacGuia.Fields.Append "glsdocreferencia", adChar, 500, adFldRowID
            rsdetalleFacGuia.Fields.Append "abredocumento", adChar, 5, adFldRowID
            rsdetalleFacGuia.Fields.Append "simbolo", adChar, 5, adFldRowID
            rsdetalleFacGuia.Fields.Append "idserie", adVarChar, 4, adFldRowID
            rsdetalleFacGuia.Fields.Append "iddocventas", adChar, 8, adFldRowID
            rsdetalleFacGuia.Fields.Append "fecEmision", adVarChar, 15, adFldRowID
            rsdetalleFacGuia.Fields.Append "estdocventas", adChar, 3, adFldRowID
            rsdetalleFacGuia.Fields.Append "idpercliente", adVarChar, 8, adFldRowID
            rsdetalleFacGuia.Fields.Append "glscliente", adVarChar, 100, adFldRowID
            rsdetalleFacGuia.Fields.Append "idmoneda", adVarChar, 3, adFldRowID
            rsdetalleFacGuia.Fields.Append "totalprecioventa", adDouble, adFldRowID
            rsdetalleFacGuia.Open
                
            rsdetallePed.MoveFirst
            Do While Not rsdetallePed.EOF
                csql = " select idempresa,tipodocorigen, seriedocorigen,numdocorigen from docreferencia " & _
                       " where idempresa = '" & Trim("" & rsdetallePed.Fields("idempresa")) & "' and tipodocreferencia = '" & Trim("" & rsdetallePed.Fields("iddocumento")) & "' " & _
                       " and seriedocreferencia = '" & Trim("" & rsdetallePed.Fields("idserie")) & "' and numdocreferencia = '" & Trim("" & rsdetallePed.Fields("iddocventas")) & "' "
                If rst.State = 1 Then rst.Close
                rst.Open csql, Cn, adOpenKeyset, adLockOptimistic
                   
                rst.MoveFirst
                Do While Not rst.EOF
                    If "" & rst.Fields("tipodocorigen") = "99" Then
                    Else
                        csql = " Select d.idempresa,d.idsucursal,d.iddocumento,d.glsdocreferencia,dc.abredocumento,m.simbolo,d.idserie,d.iddocventas,d.fecemision,d.estdocventas,d.idpercliente," & _
                            " d.glscliente,d.idmoneda,d.totalprecioventa from docventas d inner join documentos dc on d.iddocumento = dc.iddocumento" & _
                            " inner join monedas m on d.idmoneda = m.idmoneda " & _
                            " where d.idempresa = '" & Trim("" & rst.Fields("idempresa")) & "' and d.iddocumento = '" & Trim("" & rst.Fields("tipodocorigen")) & "' " & _
                            " and d.idserie = '" & Trim("" & rst.Fields("seriedocorigen")) & "' and d.iddocventas = '" & Trim("" & rst.Fields("numdocorigen")) & "' "
                        
                        If rsdetFacGuia.State = 1 Then rsdetFacGuia.Close
                        rsdetFacGuia.Open csql, Cn, adOpenKeyset, adLockOptimistic
                        rsdetFacGuia.MoveFirst
                        Do While Not rsdetFacGuia.EOF
                            item = item + 1
                            rsdetalleFacGuia.AddNew
                            rsdetalleFacGuia!item = item
                            rsdetalleFacGuia!idempresa = "" & rsdetFacGuia.Fields("idempresa")
                            rsdetalleFacGuia!idsucursal = "" & rsdetFacGuia.Fields("idsucursal")
                            rsdetalleFacGuia!idDocumento = "" & rsdetFacGuia.Fields("iddocumento")
                            rsdetalleFacGuia!glsdocreferencia = "" & rsdetFacGuia.Fields("glsdocreferencia")
                            rsdetalleFacGuia!abredocumento = "" & rsdetFacGuia.Fields("abredocumento")
                            rsdetalleFacGuia!simbolo = "" & rsdetFacGuia.Fields("simbolo")
                            rsdetalleFacGuia!idserie = "" & rsdetFacGuia.Fields("idserie")
                            rsdetalleFacGuia!idDocVentas = "" & rsdetFacGuia.Fields("iddocventas")
                            rsdetalleFacGuia!fecEmision = "" & rsdetFacGuia.Fields("fecemision")
                            rsdetalleFacGuia!estdocventas = "" & rsdetFacGuia.Fields("estdocventas")
                            rsdetalleFacGuia!idpercliente = "" & rsdetFacGuia.Fields("idpercliente")
                            rsdetalleFacGuia!GlsCliente = "" & rsdetFacGuia.Fields("glscliente")
                            rsdetalleFacGuia!idMoneda = "" & rsdetFacGuia.Fields("idmoneda")
                            rsdetalleFacGuia!totalprecioventa = "" & rsdetFacGuia.Fields("totalprecioventa")
                            rsdetalleFacGuia.Update
                            rsdetFacGuia.MoveNext
                        Loop
                    End If
                    rst.MoveNext
                Loop
                rsdetallePed.MoveNext
            Loop
            mostrarDatosGridSQL gLista3, rsdetalleFacGuia, StrMsgError
        End If
    End If
     
    If gLista3.Count <> 0 Then
        If rsdetalleFacGuia.RecordCount > 0 Then
            item = 0
            rstOriFacGuia.Fields.Append "Item", adInteger, , adFldRowID
            rstOriFacGuia.Fields.Append "idempresa", adChar, 8, adFldRowID
            rstOriFacGuia.Fields.Append "idsucursal", adChar, 8, adFldRowID
            rstOriFacGuia.Fields.Append "iddocumento", adChar, 2, adFldRowID
            rstOriFacGuia.Fields.Append "glsdocreferencia", adChar, 500, adFldRowID
            rstOriFacGuia.Fields.Append "abredocumento", adChar, 5, adFldRowID
            rstOriFacGuia.Fields.Append "simbolo", adChar, 5, adFldRowID
            rstOriFacGuia.Fields.Append "idserie", adVarChar, 4, adFldRowID
            rstOriFacGuia.Fields.Append "iddocventas", adChar, 8, adFldRowID
            rstOriFacGuia.Fields.Append "fecEmision", adVarChar, 15, adFldRowID
            rstOriFacGuia.Fields.Append "estdocventas", adChar, 3, adFldRowID
            rstOriFacGuia.Fields.Append "idpercliente", adVarChar, 8, adFldRowID
            rstOriFacGuia.Fields.Append "glscliente", adVarChar, 100, adFldRowID
            rstOriFacGuia.Fields.Append "idmoneda", adVarChar, 3, adFldRowID
            rstOriFacGuia.Fields.Append "totalprecioventa", adDouble, adFldRowID
            rstOriFacGuia.Open
                
            rsdetalleFacGuia.MoveFirst
            Do While Not rsdetalleFacGuia.EOF
                csql = " select idempresa,tipodocorigen, seriedocorigen,numdocorigen from docreferencia " & _
                       " where idempresa = '" & Trim("" & rsdetalleFacGuia.Fields("idempresa")) & "' and tipodocreferencia = '" & Trim("" & rsdetalleFacGuia.Fields("iddocumento")) & "' " & _
                       " and seriedocreferencia = '" & Trim("" & rsdetalleFacGuia.Fields("idserie")) & "' and numdocreferencia = '" & Trim("" & rsdetalleFacGuia.Fields("iddocventas")) & "' "
                If rst.State = 1 Then rst.Close
                rst.Open csql, Cn, adOpenKeyset, adLockOptimistic
                   
                rst.MoveFirst
                Do While Not rst.EOF
                    If "" & rst.Fields("tipodocorigen") = "99" Then
                    Else
                        csql = " Select d.idempresa,d.idsucursal,d.iddocumento,d.glsdocreferencia,dc.abredocumento,m.simbolo,d.idserie,d.iddocventas,d.fecemision,d.estdocventas,d.idpercliente," & _
                            " d.glscliente,d.idmoneda,d.totalprecioventa from docventas d inner join documentos dc on d.iddocumento = dc.iddocumento" & _
                            " inner join monedas m on d.idmoneda = m.idmoneda " & _
                            " where d.idempresa = '" & Trim("" & rst.Fields("idempresa")) & "' and d.iddocumento = '" & Trim("" & rst.Fields("tipodocorigen")) & "' " & _
                            " and d.idserie = '" & Trim("" & rst.Fields("seriedocorigen")) & "' and d.iddocventas = '" & Trim("" & rst.Fields("numdocorigen")) & "' "
                    
                        If rsdetFacGuia.State = 1 Then rsdetFacGuia.Close
                        rsdetFacGuia.Open csql, Cn, adOpenKeyset, adLockOptimistic
                         
                        rsdetFacGuia.MoveFirst
                        Do While Not rsdetFacGuia.EOF
                            item = item + 1
                            rstOriFacGuia.AddNew
                            rstOriFacGuia!item = item
                            rstOriFacGuia!idempresa = "" & rsdetFacGuia.Fields("idempresa")
                            rstOriFacGuia!idsucursal = "" & rsdetFacGuia.Fields("idsucursal")
                            rstOriFacGuia!idDocumento = "" & rsdetFacGuia.Fields("iddocumento")
                            rstOriFacGuia!glsdocreferencia = "" & rsdetFacGuia.Fields("glsdocreferencia")
                            rstOriFacGuia!abredocumento = "" & rsdetFacGuia.Fields("abredocumento")
                            rstOriFacGuia!simbolo = "" & rsdetFacGuia.Fields("simbolo")
                            rstOriFacGuia!idserie = "" & rsdetFacGuia.Fields("idserie")
                            rstOriFacGuia!idDocVentas = "" & rsdetFacGuia.Fields("iddocventas")
                            rstOriFacGuia!fecEmision = "" & rsdetFacGuia.Fields("fecemision")
                            rstOriFacGuia!estdocventas = "" & rsdetFacGuia.Fields("estdocventas")
                            rstOriFacGuia!idpercliente = "" & rsdetFacGuia.Fields("idpercliente")
                            rstOriFacGuia!GlsCliente = "" & rsdetFacGuia.Fields("glscliente")
                            rstOriFacGuia!idMoneda = "" & rsdetFacGuia.Fields("idmoneda")
                            rstOriFacGuia!totalprecioventa = "" & rsdetFacGuia.Fields("totalprecioventa")
                            rstOriFacGuia.Update
                            rsdetFacGuia.MoveNext
                        Loop
                    End If
                    rst.MoveNext
                Loop
                rsdetalleFacGuia.MoveNext
            Loop
        End If
        mostrarDatosGridSQL gLista4, rstOriFacGuia, StrMsgError
    End If
    
End Sub

Private Sub SecuenciaPedido()
Dim csql                 As String
Dim rst                  As New ADODB.Recordset
Dim rsg                  As New ADODB.Recordset
Dim rsDetalle            As New ADODB.Recordset
Dim rsdetalleCot         As New ADODB.Recordset
Dim rsdetallePed         As New ADODB.Recordset
Dim StrMsgError          As String
Dim item                 As Integer
Dim rsdetalleFacGuia     As New ADODB.Recordset
Dim rsdetFacGuia         As New ADODB.Recordset

    csql = " select idempresa,tipodocreferencia,seriedocreferencia,numdocreferencia from docreferencia" & _
            " where idempresa = '" & glsEmpresa & "' and tipodocorigen = '" & txtCod_TipoDoc.Text & "' and seriedocorigen = '" & txt_serie.Text & "' and numdocorigen = '" & txtNum_Documento.Text & "' "
    If rst.State = 1 Then rst.Close
    rst.Open csql, Cn, adOpenKeyset, adLockOptimistic
     
    If rst.RecordCount > 0 Then
        item = 0
        rsdetalleCot.Fields.Append "Item", adInteger, , adFldRowID
        rsdetalleCot.Fields.Append "idempresa", adChar, 8, adFldRowID
        rsdetalleCot.Fields.Append "idsucursal", adChar, 8, adFldRowID
        rsdetalleCot.Fields.Append "iddocumento", adChar, 2, adFldRowID
        rsdetalleCot.Fields.Append "glsdocreferencia", adChar, 500, adFldRowID
        rsdetalleCot.Fields.Append "abredocumento", adChar, 5, adFldRowID
        rsdetalleCot.Fields.Append "simbolo", adChar, 5, adFldRowID
        rsdetalleCot.Fields.Append "idserie", adVarChar, 4, adFldRowID
        rsdetalleCot.Fields.Append "iddocventas", adChar, 8, adFldRowID
        rsdetalleCot.Fields.Append "fecEmision", adVarChar, 15, adFldRowID
        rsdetalleCot.Fields.Append "estdocventas", adChar, 3, adFldRowID
        rsdetalleCot.Fields.Append "idpercliente", adVarChar, 8, adFldRowID
        rsdetalleCot.Fields.Append "glscliente", adVarChar, 100, adFldRowID
        rsdetalleCot.Fields.Append "idmoneda", adVarChar, 3, adFldRowID
        rsdetalleCot.Fields.Append "totalprecioventa", adDouble, adFldRowID
        rsdetalleCot.Open
            
        rst.MoveFirst
        Do While Not rst.EOF
            csql = " Select d.idempresa,d.idsucursal,d.iddocumento,d.glsdocreferencia,dc.abredocumento,m.simbolo,d.idserie,d.iddocventas,d.fecemision,d.estdocventas,d.idpercliente," & _
                   " d.glscliente,d.idmoneda,d.totalprecioventa from docventas d inner join documentos dc on d.iddocumento = dc.iddocumento " & _
                   " inner join monedas m on d.idmoneda = m.idmoneda " & _
                   " where d.idempresa = '" & Trim("" & rst.Fields("idempresa")) & "' and d.iddocumento = '" & Trim("" & rst.Fields("tipodocreferencia")) & "' " & _
                   " and d.idserie = '" & Trim("" & rst.Fields("seriedocreferencia")) & "' and d.iddocventas = '" & Trim("" & rst.Fields("numdocreferencia")) & "'"
            If rsDetalle.State = 1 Then rsDetalle.Close
            rsDetalle.Open csql, Cn, adOpenKeyset, adLockOptimistic

            rsDetalle.MoveFirst
            Do While Not rsDetalle.EOF
                item = item + 1
                rsdetalleCot.AddNew
                rsdetalleCot!item = item
                rsdetalleCot!idempresa = "" & rsDetalle.Fields("idempresa")
                rsdetalleCot!idsucursal = "" & rsDetalle.Fields("idsucursal")
                rsdetalleCot!idDocumento = "" & rsDetalle.Fields("iddocumento")
                rsdetalleCot!glsdocreferencia = "" & rsDetalle.Fields("glsdocreferencia")
                rsdetalleCot!abredocumento = "" & rsDetalle.Fields("abredocumento")
                rsdetalleCot!simbolo = "" & rsDetalle.Fields("simbolo")
                rsdetalleCot!idserie = "" & rsDetalle.Fields("idserie")
                rsdetalleCot!idDocVentas = "" & rsDetalle.Fields("iddocventas")
                rsdetalleCot!fecEmision = "" & rsDetalle.Fields("fecemision")
                rsdetalleCot!estdocventas = "" & rsDetalle.Fields("estdocventas")
                rsdetalleCot!idpercliente = "" & rsDetalle.Fields("idpercliente")
                rsdetalleCot!GlsCliente = "" & rsDetalle.Fields("glscliente")
                rsdetalleCot!idMoneda = "" & rsDetalle.Fields("idmoneda")
                rsdetalleCot!totalprecioventa = "" & rsDetalle.Fields("totalprecioventa")
                rsdetalleCot.Update
                
                rsDetalle.MoveNext
            Loop
            rst.MoveNext
        Loop
        mostrarDatosGridSQL gLista1, rsdetalleCot, StrMsgError
    End If
            
    csql = " Select d.idempresa,d.idsucursal,d.iddocumento,d.glsdocreferencia,dc.abredocumento,m.simbolo,d.idserie,d.iddocventas,d.fecemision,d.estdocventas,d.idpercliente," & _
            " d.glscliente,d.idmoneda,d.totalprecioventa from docventas d inner join documentos dc on d.iddocumento = dc.iddocumento " & _
            " inner join monedas m on d.idmoneda = m.idmoneda " & _
            " where d.idempresa =  '" & glsEmpresa & "' and d.iddocumento = '" & txtCod_TipoDoc.Text & "' " & _
            " and d.idserie = '" & txt_serie.Text & "' and d.iddocventas = '" & txtNum_Documento.Text & "'"
    If rst.State = 1 Then rst.Close
    rst.Open csql, Cn, adOpenKeyset, adLockOptimistic
            
    rsg.Fields.Append "Item", adInteger, , adFldRowID
    rsg.Fields.Append "idempresa", adChar, 8, adFldRowID
    rsg.Fields.Append "idsucursal", adChar, 8, adFldRowID
    rsg.Fields.Append "iddocumento", adChar, 2, adFldRowID
    rsg.Fields.Append "glsdocreferencia", adChar, 500, adFldRowID
    rsg.Fields.Append "abredocumento", adChar, 5, adFldRowID
    rsg.Fields.Append "simbolo", adChar, 5, adFldRowID
    rsg.Fields.Append "idserie", adVarChar, 4, adFldRowID
    rsg.Fields.Append "iddocventas", adChar, 8, adFldRowID
    rsg.Fields.Append "fecEmision", adVarChar, 15, adFldRowID
    rsg.Fields.Append "estdocventas", adChar, 3, adFldRowID
    rsg.Fields.Append "idpercliente", adVarChar, 8, adFldRowID
    rsg.Fields.Append "glscliente", adVarChar, 100, adFldRowID
    rsg.Fields.Append "idmoneda", adVarChar, 3, adFldRowID
    rsg.Fields.Append "totalprecioventa", adDouble, adFldRowID
    rsg.Open
            
    If rst.RecordCount = 0 Then
        rsg.Fields("item") = 1
        rsg.Fields("idempresa") = ""
        rsg.Fields("idsucursal") = ""
        rsg.Fields("iddocumento") = ""
        rsg.Fields("glsdocreferencia") = ""
        rsg.Fields("abredocumento") = ""
        rsg.Fields("simbolo") = ""
        rsg.Fields("idserie") = ""
        rsg.Fields("iddocventas") = ""
        rsg.Fields("fecemision") = ""
        rsg.Fields("estdocventas") = ""
        rsg.Fields("idpercliente") = ""
        rsg.Fields("glscliente") = ""
        rsg.Fields("idmoneda") = ""
        rsg.Fields("totalprecioventa") = ""
                        
    Else
        If Not rst.EOF Then
            item = 0
            rst.MoveFirst
            Do While Not rst.EOF
                rsg.AddNew
                item = item + 1
                rsg.Fields("item") = item
                rsg.Fields("idempresa") = Trim("" & rst.Fields("idempresa"))
                rsg.Fields("idsucursal") = Trim("" & rst.Fields("idsucursal"))
                rsg.Fields("iddocumento") = Trim("" & rst.Fields("iddocumento"))
                rsg.Fields("glsdocreferencia") = Trim("" & rst.Fields("glsdocreferencia"))
                rsg.Fields("abredocumento") = Trim("" & rst.Fields("abredocumento"))
                rsg.Fields("simbolo") = Trim("" & rst.Fields("simbolo"))
                rsg.Fields("idserie") = Trim("" & rst.Fields("idserie"))
                rsg.Fields("iddocventas") = Trim("" & rst.Fields("iddocventas"))
                rsg.Fields("fecemision") = "" & rst.Fields("fecemision")
                rsg.Fields("estdocventas") = Trim("" & rst.Fields("estdocventas"))
                rsg.Fields("idpercliente") = Trim("" & rst.Fields("idpercliente"))
                rsg.Fields("glscliente") = Trim("" & rst.Fields("glscliente"))
                rsg.Fields("idmoneda") = Trim("" & rst.Fields("idmoneda"))
                rsg.Fields("totalprecioventa") = Trim("" & rst.Fields("totalprecioventa"))
                rst.MoveNext
            Loop
        End If
        mostrarDatosGridSQL gLista2, rsg, StrMsgError
    End If
     
    csql = " select idempresa,tipodocorigen, seriedocorigen,numdocorigen from docreferencia" & _
            " where idempresa = '" & Trim("" & rsg.Fields("idempresa")) & "' and tipodocreferencia = '" & Trim("" & rsg.Fields("iddocumento")) & "' and seriedocreferencia = '" & Trim("" & rsg.Fields("idserie")) & "' and numdocreferencia = '" & Trim("" & rsg.Fields("iddocventas")) & "' "
    If rst.State = 1 Then rst.Close
    rst.Open csql, Cn, adOpenKeyset, adLockOptimistic
       
    If rst.RecordCount > 0 Then
        item = 0
        rsdetallePed.Fields.Append "Item", adInteger, , adFldRowID
        rsdetallePed.Fields.Append "idempresa", adChar, 8, adFldRowID
        rsdetallePed.Fields.Append "idsucursal", adChar, 8, adFldRowID
        rsdetallePed.Fields.Append "iddocumento", adChar, 2, adFldRowID
        rsdetallePed.Fields.Append "glsdocreferencia", adChar, 500, adFldRowID
        rsdetallePed.Fields.Append "abredocumento", adChar, 5, adFldRowID
        rsdetallePed.Fields.Append "simbolo", adChar, 5, adFldRowID
        rsdetallePed.Fields.Append "idserie", adVarChar, 4, adFldRowID
        rsdetallePed.Fields.Append "iddocventas", adChar, 8, adFldRowID
        rsdetallePed.Fields.Append "fecEmision", adVarChar, 15, adFldRowID
        rsdetallePed.Fields.Append "estdocventas", adChar, 3, adFldRowID
        rsdetallePed.Fields.Append "idpercliente", adVarChar, 8, adFldRowID
        rsdetallePed.Fields.Append "glscliente", adVarChar, 100, adFldRowID
        rsdetallePed.Fields.Append "idmoneda", adVarChar, 3, adFldRowID
        rsdetallePed.Fields.Append "totalprecioventa", adDouble, adFldRowID
        rsdetallePed.Open
                
        rst.MoveFirst
        Do While Not rst.EOF
            If "" & rst.Fields("tipodocorigen") = "99" Then
            Else
                csql = " Select d.idempresa,d.idsucursal,d.iddocumento,d.glsdocreferencia,dc.abredocumento,m.simbolo,d.idserie,d.iddocventas,d.fecemision,d.estdocventas,d.idpercliente," & _
                        " d.glscliente,d.idmoneda,d.totalprecioventa from docventas d inner join documentos dc on d.iddocumento = dc.iddocumento" & _
                        " inner join monedas m on d.idmoneda = m.idmoneda " & _
                        " where d.idempresa = '" & Trim("" & rst.Fields("idempresa")) & "' and d.iddocumento = '" & Trim("" & rst.Fields("tipodocorigen")) & "' " & _
                        " and d.idserie = '" & Trim("" & rst.Fields("seriedocorigen")) & "' and d.iddocventas = '" & Trim("" & rst.Fields("numdocorigen")) & "'"
                If rsDetalle.State = 1 Then rsDetalle.Close
                rsDetalle.Open csql, Cn, adOpenKeyset, adLockOptimistic
                       
                rsDetalle.MoveFirst
                Do While Not rsDetalle.EOF
                    item = item + 1
                    rsdetallePed.AddNew
                    rsdetallePed!item = item
                    rsdetallePed!idempresa = "" & rsDetalle.Fields("idempresa")
                    rsdetallePed!idsucursal = "" & rsDetalle.Fields("idsucursal")
                    rsdetallePed!idDocumento = "" & rsDetalle.Fields("iddocumento")
                    rsdetallePed!glsdocreferencia = "" & rsDetalle.Fields("glsdocreferencia")
                    rsdetallePed!abredocumento = "" & rsDetalle.Fields("abredocumento")
                    rsdetallePed!simbolo = "" & rsDetalle.Fields("simbolo")
                    rsdetallePed!idserie = "" & rsDetalle.Fields("idserie")
                    rsdetallePed!idDocVentas = "" & rsDetalle.Fields("iddocventas")
                    rsdetallePed!fecEmision = "" & rsDetalle.Fields("fecemision")
                    rsdetallePed!estdocventas = "" & rsDetalle.Fields("estdocventas")
                    rsdetallePed!idpercliente = "" & rsDetalle.Fields("idpercliente")
                    rsdetallePed!GlsCliente = "" & rsDetalle.Fields("glscliente")
                    rsdetallePed!idMoneda = "" & rsDetalle.Fields("idmoneda")
                    rsdetallePed!totalprecioventa = "" & rsDetalle.Fields("totalprecioventa")
                    rsdetallePed.Update
                    
                    rsDetalle.MoveNext
                Loop
            End If
            rst.MoveNext
        Loop
        mostrarDatosGridSQL gLista3, rsdetallePed, StrMsgError
    End If
    
    If rsdetallePed.State <> 0 Then
        If rsdetallePed.RecordCount > 0 Then
            item = 0
            rsdetalleFacGuia.Fields.Append "Item", adInteger, , adFldRowID
            rsdetalleFacGuia.Fields.Append "idempresa", adChar, 8, adFldRowID
            rsdetalleFacGuia.Fields.Append "idsucursal", adChar, 8, adFldRowID
            rsdetalleFacGuia.Fields.Append "iddocumento", adChar, 2, adFldRowID
            rsdetalleFacGuia.Fields.Append "glsdocreferencia", adChar, 500, adFldRowID
            rsdetalleFacGuia.Fields.Append "abredocumento", adChar, 5, adFldRowID
            rsdetalleFacGuia.Fields.Append "simbolo", adChar, 5, adFldRowID
            rsdetalleFacGuia.Fields.Append "idserie", adVarChar, 4, adFldRowID
            rsdetalleFacGuia.Fields.Append "iddocventas", adChar, 8, adFldRowID
            rsdetalleFacGuia.Fields.Append "fecEmision", adVarChar, 15, adFldRowID
            rsdetalleFacGuia.Fields.Append "estdocventas", adChar, 3, adFldRowID
            rsdetalleFacGuia.Fields.Append "idpercliente", adVarChar, 8, adFldRowID
            rsdetalleFacGuia.Fields.Append "glscliente", adVarChar, 100, adFldRowID
            rsdetalleFacGuia.Fields.Append "idmoneda", adVarChar, 3, adFldRowID
            rsdetalleFacGuia.Fields.Append "totalprecioventa", adDouble, adFldRowID
            rsdetalleFacGuia.Open
                
            rsdetallePed.MoveFirst
            Do While Not rsdetallePed.EOF
                csql = " select idempresa,tipodocorigen, seriedocorigen,numdocorigen from docreferencia " & _
                       " where idempresa = '" & Trim("" & rsdetallePed.Fields("idempresa")) & "' and tipodocreferencia = '" & Trim("" & rsdetallePed.Fields("iddocumento")) & "' " & _
                       " and seriedocreferencia = '" & Trim("" & rsdetallePed.Fields("idserie")) & "' and numdocreferencia = '" & Trim("" & rsdetallePed.Fields("iddocventas")) & "' "
                If rst.State = 1 Then rst.Close
                rst.Open csql, Cn, adOpenKeyset, adLockOptimistic
                   
                If rst.RecordCount > 0 Then
                    rst.MoveFirst
                    Do While Not rst.EOF
                        If "" & rst.Fields("tipodocorigen") = "99" Then
                        Else
                            csql = " Select d.idempresa,d.idsucursal,d.iddocumento,d.glsdocreferencia,dc.abredocumento,m.simbolo,d.idserie,d.iddocventas,d.fecemision,d.estdocventas,d.idpercliente," & _
                                   " d.glscliente,d.idmoneda,d.totalprecioventa from docventas d inner join documentos dc on d.iddocumento = dc.iddocumento" & _
                                   " inner join monedas m on d.idmoneda = m.idmoneda " & _
                                   " where d.idempresa = '" & Trim("" & rst.Fields("idempresa")) & "' and d.iddocumento = '" & Trim("" & rst.Fields("tipodocorigen")) & "' " & _
                                   " and d.idserie = '" & Trim("" & rst.Fields("seriedocorigen")) & "' and d.iddocventas = '" & Trim("" & rst.Fields("numdocorigen")) & "'"
                            
                            If rsdetFacGuia.State = 1 Then rsdetFacGuia.Close
                            rsdetFacGuia.Open csql, Cn, adOpenKeyset, adLockOptimistic
                
                            rsdetFacGuia.MoveFirst
                            Do While Not rsdetFacGuia.EOF
                                item = item + 1
                                rsdetalleFacGuia.AddNew
                                rsdetalleFacGuia!item = item
                                rsdetalleFacGuia!idempresa = "" & rsdetFacGuia.Fields("idempresa")
                                rsdetalleFacGuia!idsucursal = "" & rsdetFacGuia.Fields("idsucursal")
                                rsdetalleFacGuia!idDocumento = "" & rsdetFacGuia.Fields("iddocumento")
                                rsdetalleFacGuia!glsdocreferencia = "" & rsdetFacGuia.Fields("glsdocreferencia")
                                rsdetalleFacGuia!abredocumento = "" & rsdetFacGuia.Fields("abredocumento")
                                rsdetalleFacGuia!simbolo = "" & rsdetFacGuia.Fields("simbolo")
                                rsdetalleFacGuia!idserie = "" & rsdetFacGuia.Fields("idserie")
                                rsdetalleFacGuia!idDocVentas = "" & rsdetFacGuia.Fields("iddocventas")
                                rsdetalleFacGuia!fecEmision = "" & rsdetFacGuia.Fields("fecemision")
                                rsdetalleFacGuia!estdocventas = "" & rsdetFacGuia.Fields("estdocventas")
                                rsdetalleFacGuia!idpercliente = "" & rsdetFacGuia.Fields("idpercliente")
                                rsdetalleFacGuia!GlsCliente = "" & rsdetFacGuia.Fields("glscliente")
                                rsdetalleFacGuia!idMoneda = "" & rsdetFacGuia.Fields("idmoneda")
                                rsdetalleFacGuia!totalprecioventa = "" & rsdetFacGuia.Fields("totalprecioventa")
                                rsdetalleFacGuia.Update
                                rsdetFacGuia.MoveNext
                            Loop
                        End If
                        rst.MoveNext
                    Loop
                End If
                rsdetallePed.MoveNext
            Loop
            mostrarDatosGridSQL gLista4, rsdetalleFacGuia, StrMsgError
        End If
    End If
    If StrMsgError <> "" Then GoTo Err
     
    Exit Sub

Err:
    If StrMsgError = "" Then StrMsgError = Err.Description
    MsgBox StrMsgError, vbInformation, App.Title
End Sub

Private Sub SecuenciaFactura1()
Dim csql                 As String
Dim rst                  As New ADODB.Recordset
Dim rsg                  As New ADODB.Recordset
Dim rsDetalle            As New ADODB.Recordset
Dim rsdetalleCot         As New ADODB.Recordset
Dim rsdetallePed         As New ADODB.Recordset
Dim rsdetalleFacGuia     As New ADODB.Recordset
Dim rsdetFacGuia         As New ADODB.Recordset
Dim rstdetCot            As New ADODB.Recordset
Dim StrMsgError          As String
Dim item                 As Integer


    csql = " Select (@i:=@i +1) as item,d.idempresa,d.idsucursal,d.iddocumento,d.glsdocreferencia,dc.abredocumento,m.simbolo,d.idserie,d.iddocventas,d.fecemision,d.estdocventas,d.idpercliente," & _
            " d.glscliente,d.idmoneda,d.totalprecioventa from docventas d inner join documentos dc on d.iddocumento = dc.iddocumento " & _
            " inner join monedas m on d.idmoneda = m.idmoneda,(select @i:=0) foo " & _
            " where d.idempresa = '" & glsEmpresa & "' and d.iddocumento = '" & txtCod_TipoDoc.Text & "' " & _
            " and d.idserie = '" & txt_serie.Text & "' and d.iddocventas = '" & txtNum_Documento.Text & "'"
               
    If rst.State = 1 Then rst.Close
    rst.Open csql, Cn, adOpenKeyset, adLockOptimistic
     
    rsg.Fields.Append "item", adInteger, , adFldRowID
    rsg.Fields.Append "idempresa", adChar, 8, adFldIsNullable
    rsg.Fields.Append "idsucursal", adChar, 8, adFldIsNullable
    rsg.Fields.Append "iddocumento", adChar, 2, adFldIsNullable
    rsg.Fields.Append "glsdocreferencia", adChar, 500, adFldIsNullable
    rsg.Fields.Append "abredocumento", adChar, 5, adFldIsNullable
    rsg.Fields.Append "simbolo", adChar, 5, adFldIsNullable
    rsg.Fields.Append "idserie", adChar, 4, adFldIsNullable
    rsg.Fields.Append "iddocventas", adChar, 8, adFldIsNullable
    rsg.Fields.Append "fecemision", adChar, 15, adFldIsNullable
    rsg.Fields.Append "estdocventas", adChar, 3, adFldIsNullable
    rsg.Fields.Append "idpercliente", adChar, 8, adFldIsNullable
    rsg.Fields.Append "glscliente", adVarChar, 100, adFldIsNullable
    rsg.Fields.Append "idmoneda", adChar, 5, adFldIsNullable
    rsg.Fields.Append "totalprecioventa", adDouble, adFldRowID
    rsg.Open
    
    If rst.RecordCount = 0 Then
        rsg.Fields("item") = 1
        rsg.Fields("idempresa") = ""
        rsg.Fields("idsucursal") = ""
        rsg.Fields("iddocumento") = ""
        rsg.Fields("glsdocreferencia") = ""
        rsg.Fields("abredocumento") = ""
        rsg.Fields("simbolo") = ""
        rsg.Fields("idserie") = ""
        rsg.Fields("iddocventas") = ""
        rsg.Fields("fecemision") = ""
        rsg.Fields("estdocventas") = ""
        rsg.Fields("idpercliente") = ""
        rsg.Fields("glscliente") = ""
        rsg.Fields("idmoneda") = ""
        rsg.Fields("totalprecioventa") = ""
                
    Else
        If Not rst.EOF Then
            rst.MoveFirst
            Do While Not rst.EOF
                rsg.AddNew
                rsg.Fields("item") = Val("" & rst.Fields("item"))
                rsg.Fields("idempresa") = Trim("" & rst.Fields("idempresa"))
                rsg.Fields("idsucursal") = Trim("" & rst.Fields("idsucursal"))
                rsg.Fields("iddocumento") = Trim("" & rst.Fields("iddocumento"))
                rsg.Fields("glsdocreferencia") = Trim("" & rst.Fields("glsdocreferencia"))
                rsg.Fields("abredocumento") = Trim("" & rst.Fields("abredocumento"))
                rsg.Fields("simbolo") = Trim("" & rst.Fields("simbolo"))
                rsg.Fields("idserie") = Trim("" & rst.Fields("idserie"))
                rsg.Fields("iddocventas") = Trim("" & rst.Fields("iddocventas"))
                rsg.Fields("fecemision") = "" & rst.Fields("fecemision")
                rsg.Fields("estdocventas") = "" & rst.Fields("estdocventas")
                rsg.Fields("idpercliente") = Trim("" & rst.Fields("idpercliente"))
                rsg.Fields("glscliente") = Trim("" & rst.Fields("glscliente"))
                rsg.Fields("idmoneda") = Trim("" & rst.Fields("idmoneda"))
                rsg.Fields("totalprecioventa") = Trim("" & rst.Fields("totalprecioventa"))
                rst.MoveNext
            Loop
        End If
        mostrarDatosGridSQL gLista3, rsg, StrMsgError
    End If
    
    csql = " select idempresa,tipodocreferencia,seriedocreferencia,numdocreferencia from docreferencia" & _
            " where idempresa = '" & glsEmpresa & "' and tipodocorigen = '" & txtCod_TipoDoc.Text & "' and seriedocorigen = '" & txt_serie.Text & "' and numdocorigen = '" & txtNum_Documento.Text & "' "
    If rst.State = 1 Then rst.Close
    rst.Open csql, Cn, adOpenKeyset, adLockOptimistic
     
    If rst.RecordCount > 0 Then
        item = 0
        rsdetallePed.Fields.Append "Item", adInteger, , adFldRowID
        rsdetallePed.Fields.Append "idempresa", adChar, 8, adFldRowID
        rsdetallePed.Fields.Append "idsucursal", adChar, 8, adFldRowID
        rsdetallePed.Fields.Append "iddocumento", adChar, 2, adFldRowID
        rsdetallePed.Fields.Append "glsdocreferencia", adChar, 500, adFldRowID
        rsdetallePed.Fields.Append "abredocumento", adChar, 5, adFldRowID
        rsdetallePed.Fields.Append "simbolo", adChar, 5, adFldRowID
        rsdetallePed.Fields.Append "idserie", adVarChar, 4, adFldRowID
        rsdetallePed.Fields.Append "iddocventas", adChar, 8, adFldRowID
        rsdetallePed.Fields.Append "fecEmision", adVarChar, 15, adFldRowID
        rsdetallePed.Fields.Append "estdocventas", adChar, 3, adFldRowID
        rsdetallePed.Fields.Append "idpercliente", adVarChar, 8, adFldRowID
        rsdetallePed.Fields.Append "glscliente", adVarChar, 100, adFldRowID
        rsdetallePed.Fields.Append "idmoneda", adVarChar, 3, adFldRowID
        rsdetallePed.Fields.Append "totalprecioventa", adDouble, adFldRowID
        rsdetallePed.Open
                
        rst.MoveFirst
        Do While Not rst.EOF
            If Len(Trim("" & rst.Fields("numdocreferencia"))) = 8 Then
                csql = " Select d.idempresa,d.idsucursal,d.iddocumento,d.glsdocreferencia,dc.abredocumento,m.simbolo,d.idserie,d.iddocventas,d.fecemision,d.estdocventas,d.idpercliente," & _
                        " d.glscliente,d.idmoneda,d.totalprecioventa from docventas d inner join documentos dc on d.iddocumento = dc.iddocumento " & _
                        " inner join monedas m on d.idmoneda = m.idmoneda " & _
                        " where d.idempresa = '" & Trim("" & rst.Fields("idempresa")) & "' and d.iddocumento = '" & Trim("" & rst.Fields("tipodocreferencia")) & "' " & _
                        " and d.idserie = '" & Trim("" & rst.Fields("seriedocreferencia")) & "' and d.iddocventas = '" & Trim("" & rst.Fields("numdocreferencia")) & "'"
                If rsDetalle.State = 1 Then rsDetalle.Close
                rsDetalle.Open csql, Cn, adOpenKeyset, adLockOptimistic
                
                rsDetalle.MoveFirst
                Do While Not rsDetalle.EOF
                    item = item + 1
                    rsdetallePed.AddNew
                    rsdetallePed!item = item
                    rsdetallePed!idempresa = "" & rsDetalle.Fields("idempresa")
                    rsdetallePed!idsucursal = "" & rsDetalle.Fields("idsucursal")
                    rsdetallePed!idDocumento = "" & rsDetalle.Fields("iddocumento")
                    rsdetallePed!glsdocreferencia = "" & rsDetalle.Fields("glsdocreferencia")
                    rsdetallePed!abredocumento = "" & rsDetalle.Fields("abredocumento")
                    rsdetallePed!simbolo = "" & rsDetalle.Fields("simbolo")
                    rsdetallePed!idserie = "" & rsDetalle.Fields("idserie")
                    rsdetallePed!idDocVentas = "" & rsDetalle.Fields("iddocventas")
                    rsdetallePed!fecEmision = "" & rsDetalle.Fields("fecemision")
                    rsdetallePed!estdocventas = "" & rsDetalle.Fields("estdocventas")
                    rsdetallePed!idpercliente = "" & rsDetalle.Fields("idpercliente")
                    rsdetallePed!GlsCliente = "" & rsDetalle.Fields("glscliente")
                    rsdetallePed!idMoneda = "" & rsDetalle.Fields("idmoneda")
                    rsdetallePed!totalprecioventa = "" & rsDetalle.Fields("totalprecioventa")
                    rsdetallePed.Update
               
                    rsDetalle.MoveNext
                Loop
            End If
            rst.MoveNext
        Loop
        mostrarDatosGridSQL gLista2, rsdetallePed, StrMsgError
    End If

    If rsdetallePed.RecordCount > 0 Then
        item = 0
        rsdetalleCot.Fields.Append "Item", adInteger, , adFldRowID
        rsdetalleCot.Fields.Append "idempresa", adChar, 8, adFldRowID
        rsdetalleCot.Fields.Append "idsucursal", adChar, 8, adFldRowID
        rsdetalleCot.Fields.Append "iddocumento", adChar, 2, adFldRowID
        rsdetalleCot.Fields.Append "glsdocreferencia", adChar, 500, adFldRowID
        rsdetalleCot.Fields.Append "abredocumento", adChar, 5, adFldRowID
        rsdetalleCot.Fields.Append "simbolo", adChar, 5, adFldRowID
        rsdetalleCot.Fields.Append "idserie", adVarChar, 4, adFldRowID
        rsdetalleCot.Fields.Append "iddocventas", adChar, 8, adFldRowID
        rsdetalleCot.Fields.Append "fecEmision", adVarChar, 15, adFldRowID
        rsdetalleCot.Fields.Append "estdocventas", adChar, 3, adFldRowID
        rsdetalleCot.Fields.Append "idpercliente", adVarChar, 8, adFldRowID
        rsdetalleCot.Fields.Append "glscliente", adVarChar, 100, adFldRowID
        rsdetalleCot.Fields.Append "idmoneda", adVarChar, 3, adFldRowID
        rsdetalleCot.Fields.Append "totalprecioventa", adDouble, adFldRowID
        rsdetalleCot.Open
        
        rsdetallePed.MoveFirst
        Do While Not rsdetallePed.EOF
            csql = " select idempresa,tipodocreferencia,seriedocreferencia,numdocreferencia from docreferencia" & _
                   " where idempresa = '" & Trim("" & rsdetallePed.Fields("idempresa")) & "' and tipodocorigen = '" & Trim("" & rsdetallePed.Fields("iddocumento")) & "' and seriedocorigen = '" & Trim("" & rsdetallePed.Fields("idserie")) & "' and numdocorigen = '" & Trim("" & rsdetallePed.Fields("iddocventas")) & "' "
            If rst.State = 1 Then rst.Close
            rst.Open csql, Cn, adOpenKeyset, adLockOptimistic
             
            If rst.RecordCount > 0 Then
                rst.MoveFirst
                Do While Not rst.EOF
                    If Len(Trim("" & rst.Fields("numdocreferencia"))) = 8 Then
                        csql = " Select d.idempresa,d.idsucursal,d.iddocumento,d.glsdocreferencia,dc.abredocumento,m.simbolo,d.idserie,d.iddocventas,d.fecemision,d.estdocventas,d.idpercliente," & _
                               " d.glscliente,d.idmoneda,d.totalprecioventa from docventas d inner join documentos dc on d.iddocumento = dc.iddocumento " & _
                               " inner join monedas m on d.idmoneda = m.idmoneda " & _
                               " where d.idempresa = '" & Trim("" & rst.Fields("idempresa")) & "' and d.iddocumento = '" & Trim("" & rst.Fields("tipodocreferencia")) & "' " & _
                               " and d.idserie = '" & Trim("" & rst.Fields("seriedocreferencia")) & "' and d.iddocventas = '" & Trim("" & rst.Fields("numdocreferencia")) & "'"
                        If rsDetalle.State = 1 Then rsDetalle.Close
                        rsDetalle.Open csql, Cn, adOpenKeyset, adLockOptimistic
            
                        rsDetalle.MoveFirst
                        Do While Not rsDetalle.EOF
                            item = item + 1
                            rsdetalleCot.AddNew
                            rsdetalleCot!item = item
                            rsdetalleCot!idempresa = "" & rsDetalle.Fields("idempresa")
                            rsdetalleCot!idsucursal = "" & rsDetalle.Fields("idsucursal")
                            rsdetalleCot!idDocumento = "" & rsDetalle.Fields("iddocumento")
                            rsdetalleCot!glsdocreferencia = "" & rsDetalle.Fields("glsdocreferencia")
                            rsdetalleCot!abredocumento = "" & rsDetalle.Fields("abredocumento")
                            rsdetalleCot!simbolo = "" & rsDetalle.Fields("simbolo")
                            rsdetalleCot!idserie = "" & rsDetalle.Fields("idserie")
                            rsdetalleCot!idDocVentas = "" & rsDetalle.Fields("iddocventas")
                            rsdetalleCot!fecEmision = "" & rsDetalle.Fields("fecemision")
                            rsdetalleCot!estdocventas = "" & rsDetalle.Fields("estdocventas")
                            rsdetalleCot!idpercliente = "" & rsDetalle.Fields("idpercliente")
                            rsdetalleCot!GlsCliente = "" & rsDetalle.Fields("glscliente")
                            rsdetalleCot!idMoneda = "" & rsDetalle.Fields("idmoneda")
                            rsdetalleCot!totalprecioventa = "" & rsDetalle.Fields("totalprecioventa")
                            rsdetalleCot.Update
                           
                            rsDetalle.MoveNext
                        Loop
                    End If
                    rst.MoveNext
                Loop
            End If
            rsdetallePed.MoveNext
        Loop
        mostrarDatosGridSQL gLista1, rsdetalleCot, StrMsgError
    End If
         
    csql = " select idempresa,tipodocorigen,seriedocorigen,numdocorigen from docreferencia" & _
           " where idempresa = '" & glsEmpresa & "' and tipodocreferencia= '" & Trim(txtCod_TipoDoc.Text) & "' and seriedocreferencia = '" & Trim(txt_serie.Text) & "' and numdocreferencia = '" & Trim(txtNum_Documento.Text) & "' "
    If rst.State = 1 Then rst.Close
    rst.Open csql, Cn, adOpenKeyset, adLockOptimistic
     
    If rst.RecordCount > 0 Then
        item = 0
        rstdetCot.Fields.Append "Item", adInteger, , adFldRowID
        rstdetCot.Fields.Append "idempresa", adChar, 8, adFldRowID
        rstdetCot.Fields.Append "idsucursal", adChar, 8, adFldRowID
        rstdetCot.Fields.Append "iddocumento", adChar, 2, adFldRowID
        rstdetCot.Fields.Append "glsdocreferencia", adChar, 500, adFldRowID
        rstdetCot.Fields.Append "abredocumento", adChar, 5, adFldRowID
        rstdetCot.Fields.Append "simbolo", adChar, 5, adFldRowID
        rstdetCot.Fields.Append "idserie", adVarChar, 4, adFldRowID
        rstdetCot.Fields.Append "iddocventas", adChar, 8, adFldRowID
        rstdetCot.Fields.Append "fecEmision", adVarChar, 15, adFldRowID
        rstdetCot.Fields.Append "estdocventas", adChar, 3, adFldRowID
        rstdetCot.Fields.Append "idpercliente", adVarChar, 8, adFldRowID
        rstdetCot.Fields.Append "glscliente", adVarChar, 100, adFldRowID
        rstdetCot.Fields.Append "idmoneda", adVarChar, 3, adFldRowID
        rstdetCot.Fields.Append "totalprecioventa", adDouble, adFldRowID
        rstdetCot.Open
            
        rst.MoveFirst
        Do While Not rst.EOF
            If Len(Trim("" & rst.Fields("numdocorigen"))) = 8 Then
                csql = " Select d.idempresa,d.idsucursal,d.iddocumento,d.glsdocreferencia,dc.abredocumento,m.simbolo,d.idserie,d.iddocventas,d.fecemision,d.estdocventas,d.idpercliente," & _
                       " d.glscliente,d.idmoneda,d.totalprecioventa from docventas d inner join documentos dc on d.iddocumento = dc.iddocumento " & _
                       " inner join monedas m on d.idmoneda = m.idmoneda " & _
                       " where d.idempresa = '" & Trim("" & rst.Fields("idempresa")) & "' and d.iddocumento = '" & Trim("" & rst.Fields("tipodocorigen")) & "' " & _
                       " and d.idserie = '" & Trim("" & rst.Fields("seriedocorigen")) & "' and d.iddocventas = '" & Trim("" & rst.Fields("numdocorigen")) & "'"
                If rsDetalle.State = 1 Then rsDetalle.Close
                rsDetalle.Open csql, Cn, adOpenKeyset, adLockOptimistic
                If rsDetalle.RecordCount > 0 Then
                    rsDetalle.MoveFirst
                    Do While Not rsDetalle.EOF
                        item = item + 1
                        rstdetCot.AddNew
                        rstdetCot!item = item
                        rstdetCot!idempresa = "" & rsDetalle.Fields("idempresa")
                        rstdetCot!idsucursal = "" & rsDetalle.Fields("idsucursal")
                        rstdetCot!idDocumento = "" & rsDetalle.Fields("iddocumento")
                        rstdetCot!glsdocreferencia = "" & rsDetalle.Fields("glsdocreferencia")
                        rstdetCot!abredocumento = "" & rsDetalle.Fields("abredocumento")
                        rstdetCot!simbolo = "" & rsDetalle.Fields("simbolo")
                        rstdetCot!idserie = "" & rsDetalle.Fields("idserie")
                        rstdetCot!idDocVentas = "" & rsDetalle.Fields("iddocventas")
                        rstdetCot!fecEmision = "" & rsDetalle.Fields("fecemision")
                        rstdetCot!estdocventas = "" & rsDetalle.Fields("estdocventas")
                        rstdetCot!idpercliente = "" & rsDetalle.Fields("idpercliente")
                        rstdetCot!GlsCliente = "" & rsDetalle.Fields("glscliente")
                        rstdetCot!idMoneda = "" & rsDetalle.Fields("idmoneda")
                        rstdetCot!totalprecioventa = "" & rsDetalle.Fields("totalprecioventa")
                        rstdetCot.Update
                          
                        rsDetalle.MoveNext
                    Loop
                End If
            End If
            rst.MoveNext
        Loop
        mostrarDatosGridSQL gLista4, rstdetCot, StrMsgError
    End If
         
End Sub

Private Sub SecuenciaFactura2()
Dim csql                 As String
Dim rst                  As New ADODB.Recordset
Dim rsg                  As New ADODB.Recordset
Dim rsDetalle            As New ADODB.Recordset
Dim rsdetalleCot         As New ADODB.Recordset
Dim rsdetallePed         As New ADODB.Recordset
Dim rsdetalleFacGuia     As New ADODB.Recordset
Dim rsdetFacGuia         As New ADODB.Recordset
Dim rstdetCot            As New ADODB.Recordset
Dim StrMsgError          As String
Dim item                 As Integer

    csql = " Select (@i:=@i +1) as item,d.idempresa,d.idsucursal,d.iddocumento,d.glsdocreferencia,dc.abredocumento,m.simbolo,d.idserie,d.iddocventas,d.fecemision,d.estdocventas,d.idpercliente," & _
               " d.glscliente,d.idmoneda,d.totalprecioventa from docventas d inner join documentos dc on d.iddocumento = dc.iddocumento " & _
               " inner join monedas m on d.idmoneda = m.idmoneda ,(select @i:=0) foo " & _
               " where d.idempresa = '" & glsEmpresa & "' and d.iddocumento = '" & txtCod_TipoDoc.Text & "' " & _
               " and d.idserie = '" & txt_serie.Text & "' and d.iddocventas = '" & txtNum_Documento.Text & "'"
               
    If rst.State = 1 Then rst.Close
    rst.Open csql, Cn, adOpenKeyset, adLockOptimistic
     
    rsg.Fields.Append "item", adInteger, , adFldRowID
    rsg.Fields.Append "idempresa", adChar, 8, adFldIsNullable
    rsg.Fields.Append "idsucursal", adChar, 8, adFldIsNullable
    rsg.Fields.Append "iddocumento", adChar, 2, adFldIsNullable
    rsg.Fields.Append "glsdocreferencia", adChar, 500, adFldIsNullable
    rsg.Fields.Append "abredocumento", adChar, 5, adFldIsNullable
    rsg.Fields.Append "simbolo", adChar, 5, adFldIsNullable
    rsg.Fields.Append "idserie", adChar, 4, adFldIsNullable
    rsg.Fields.Append "iddocventas", adChar, 8, adFldIsNullable
    rsg.Fields.Append "fecemision", adChar, 15, adFldIsNullable
    rsg.Fields.Append "estdocventas", adChar, 3, adFldIsNullable
    rsg.Fields.Append "idpercliente", adChar, 8, adFldIsNullable
    rsg.Fields.Append "glscliente", adVarChar, 100, adFldIsNullable
    rsg.Fields.Append "idmoneda", adChar, 5, adFldIsNullable
    rsg.Fields.Append "totalprecioventa", adDouble, adFldRowID
    rsg.Open
    
    If rst.RecordCount = 0 Then
        rsg.Fields("item") = 1
        rsg.Fields("idempresa") = ""
        rsg.Fields("idsucursal") = ""
        rsg.Fields("iddocumento") = ""
        rsg.Fields("glsdocreferencia") = ""
        rsg.Fields("abredocumento") = ""
        rsg.Fields("simbolo") = ""
        rsg.Fields("idserie") = ""
        rsg.Fields("iddocventas") = ""
        rsg.Fields("fecemision") = ""
        rsg.Fields("estdocventas") = ""
        rsg.Fields("idpercliente") = ""
        rsg.Fields("glscliente") = ""
        rsg.Fields("idmoneda") = ""
        rsg.Fields("totalprecioventa") = ""
                
    Else
        If Not rst.EOF Then
            rst.MoveFirst
            Do While Not rst.EOF
                rsg.AddNew
                rsg.Fields("item") = Val("" & rst.Fields("item"))
                rsg.Fields("idempresa") = Trim("" & rst.Fields("idempresa"))
                rsg.Fields("idsucursal") = Trim("" & rst.Fields("idsucursal"))
                rsg.Fields("iddocumento") = Trim("" & rst.Fields("iddocumento"))
                rsg.Fields("glsdocreferencia") = Trim("" & rst.Fields("glsdocreferencia"))
                rsg.Fields("abredocumento") = Trim("" & rst.Fields("abredocumento"))
                rsg.Fields("simbolo") = Trim("" & rst.Fields("simbolo"))
                rsg.Fields("idserie") = Trim("" & rst.Fields("idserie"))
                rsg.Fields("iddocventas") = Trim("" & rst.Fields("iddocventas"))
                rsg.Fields("fecemision") = "" & rst.Fields("fecemision")
                rsg.Fields("estdocventas") = "" & rst.Fields("estdocventas")
                rsg.Fields("idpercliente") = Trim("" & rst.Fields("idpercliente"))
                rsg.Fields("glscliente") = Trim("" & rst.Fields("glscliente"))
                rsg.Fields("idmoneda") = Trim("" & rst.Fields("idmoneda"))
                rsg.Fields("totalprecioventa") = Trim("" & rst.Fields("totalprecioventa"))
                rst.MoveNext
            Loop
        End If
        mostrarDatosGridSQL gLista4, rsg, StrMsgError
    End If
    
    csql = " select idempresa,tipodocreferencia,seriedocreferencia,numdocreferencia from docreferencia" & _
           " where idempresa = '" & glsEmpresa & "' and tipodocorigen = '" & txtCod_TipoDoc.Text & "'" & _
           " and seriedocorigen = '" & txt_serie.Text & "' " & _
           " and numdocorigen = '" & txtNum_Documento.Text & "' "
           
    If rst.State = 1 Then rst.Close
    rst.Open csql, Cn, adOpenKeyset, adLockOptimistic
     
    If rst.RecordCount > 0 Then
        item = 0
        rsdetallePed.Fields.Append "Item", adInteger, , adFldRowID
        rsdetallePed.Fields.Append "idempresa", adChar, 8, adFldRowID
        rsdetallePed.Fields.Append "idsucursal", adChar, 8, adFldRowID
        rsdetallePed.Fields.Append "iddocumento", adChar, 2, adFldRowID
        rsdetallePed.Fields.Append "glsdocreferencia", adChar, 500, adFldRowID
        rsdetallePed.Fields.Append "abredocumento", adChar, 5, adFldRowID
        rsdetallePed.Fields.Append "simbolo", adChar, 5, adFldRowID
        rsdetallePed.Fields.Append "idserie", adVarChar, 4, adFldRowID
        rsdetallePed.Fields.Append "iddocventas", adChar, 8, adFldRowID
        rsdetallePed.Fields.Append "fecEmision", adVarChar, 15, adFldRowID
        rsdetallePed.Fields.Append "estdocventas", adChar, 3, adFldRowID
        rsdetallePed.Fields.Append "idpercliente", adVarChar, 8, adFldRowID
        rsdetallePed.Fields.Append "glscliente", adVarChar, 100, adFldRowID
        rsdetallePed.Fields.Append "idmoneda", adVarChar, 3, adFldRowID
        rsdetallePed.Fields.Append "totalprecioventa", adDouble, adFldRowID
        rsdetallePed.Open
                
        rst.MoveFirst
        Do While Not rst.EOF
            If Len(Trim("" & rst.Fields("numdocreferencia"))) = 8 Then
                csql = " Select d.idempresa,d.idsucursal,d.iddocumento,d.glsdocreferencia,dc.abredocumento,m.simbolo,d.idserie,d.iddocventas,d.fecemision,d.estdocventas,d.idpercliente," & _
                       " d.glscliente,d.idmoneda,d.totalprecioventa from docventas d inner join documentos dc on d.iddocumento = dc.iddocumento " & _
                       " inner join monedas m on d.idmoneda = m.idmoneda " & _
                       " where d.idempresa = '" & Trim("" & rst.Fields("idempresa")) & "' and d.iddocumento = '" & Trim("" & rst.Fields("tipodocreferencia")) & "' " & _
                       " and d.idserie = '" & Trim("" & rst.Fields("seriedocreferencia")) & "' and d.iddocventas = '" & Trim("" & rst.Fields("numdocreferencia")) & "'"
                If rsDetalle.State = 1 Then rsDetalle.Close
                rsDetalle.Open csql, Cn, adOpenKeyset, adLockOptimistic
                rsDetalle.MoveFirst
                Do While Not rsDetalle.EOF
                    item = item + 1
                    rsdetallePed.AddNew
                    rsdetallePed!item = item
                    rsdetallePed!idempresa = "" & rsDetalle.Fields("idempresa")
                    rsdetallePed!idsucursal = "" & rsDetalle.Fields("idsucursal")
                    rsdetallePed!idDocumento = "" & rsDetalle.Fields("iddocumento")
                    rsdetallePed!glsdocreferencia = "" & rsDetalle.Fields("glsdocreferencia")
                    rsdetallePed!abredocumento = "" & rsDetalle.Fields("abredocumento")
                    rsdetallePed!simbolo = "" & rsDetalle.Fields("simbolo")
                    rsdetallePed!idserie = "" & rsDetalle.Fields("idserie")
                    rsdetallePed!idDocVentas = "" & rsDetalle.Fields("iddocventas")
                    rsdetallePed!fecEmision = "" & rsDetalle.Fields("fecemision")
                    rsdetallePed!estdocventas = "" & rsDetalle.Fields("estdocventas")
                    rsdetallePed!idpercliente = "" & rsDetalle.Fields("idpercliente")
                    rsdetallePed!GlsCliente = "" & rsDetalle.Fields("glscliente")
                    rsdetallePed!idMoneda = "" & rsDetalle.Fields("idmoneda")
                    rsdetallePed!totalprecioventa = "" & rsDetalle.Fields("totalprecioventa")
                    rsdetallePed.Update
               
                    rsDetalle.MoveNext
                Loop
            End If
            rst.MoveNext
         Loop
         mostrarDatosGridSQL gLista3, rsdetallePed, StrMsgError
    End If

    If rsdetallePed.RecordCount > 0 Then
        item = 0
        rsdetalleCot.Fields.Append "Item", adInteger, , adFldRowID
        rsdetalleCot.Fields.Append "idempresa", adChar, 8, adFldRowID
        rsdetalleCot.Fields.Append "idsucursal", adChar, 8, adFldRowID
        rsdetalleCot.Fields.Append "iddocumento", adChar, 2, adFldRowID
        rsdetalleCot.Fields.Append "glsdocreferencia", adChar, 500, adFldRowID
        rsdetalleCot.Fields.Append "abredocumento", adChar, 5, adFldRowID
        rsdetalleCot.Fields.Append "simbolo", adChar, 5, adFldRowID
        rsdetalleCot.Fields.Append "idserie", adVarChar, 4, adFldRowID
        rsdetalleCot.Fields.Append "iddocventas", adChar, 8, adFldRowID
        rsdetalleCot.Fields.Append "fecEmision", adVarChar, 15, adFldRowID
        rsdetalleCot.Fields.Append "estdocventas", adChar, 3, adFldRowID
        rsdetalleCot.Fields.Append "idpercliente", adVarChar, 8, adFldRowID
        rsdetalleCot.Fields.Append "glscliente", adVarChar, 100, adFldRowID
        rsdetalleCot.Fields.Append "idmoneda", adVarChar, 3, adFldRowID
        rsdetalleCot.Fields.Append "totalprecioventa", adDouble, adFldRowID
        rsdetalleCot.Open

        rsdetallePed.MoveFirst
        Do While Not rsdetallePed.EOF
            csql = " select idempresa,tipodocreferencia,seriedocreferencia,numdocreferencia from docreferencia" & _
                   " where idempresa = '" & Trim("" & rsdetallePed.Fields("idempresa")) & "' and tipodocorigen = '" & Trim("" & rsdetallePed.Fields("iddocumento")) & "' and seriedocorigen = '" & Trim("" & rsdetallePed.Fields("idserie")) & "' and numdocorigen = '" & Trim("" & rsdetallePed.Fields("iddocventas")) & "' "
            If rst.State = 1 Then rst.Close
            rst.Open csql, Cn, adOpenKeyset, adLockOptimistic
             
            If rst.RecordCount > 0 Then
                rst.MoveFirst
                Do While Not rst.EOF
                    If Len(Trim("" & rst.Fields("numdocreferencia"))) = 8 Then
                        csql = " Select d.idempresa,d.idsucursal,d.iddocumento,d.glsdocreferencia,dc.abredocumento,m.simbolo,d.idserie,d.iddocventas,d.fecemision,d.estdocventas,d.idpercliente," & _
                               " d.glscliente,d.idmoneda,d.totalprecioventa from docventas d inner join documentos dc on d.iddocumento = dc.iddocumento" & _
                               " inner join monedas m on d.idmoneda = m.idmoneda " & _
                               " where d.idempresa = '" & Trim("" & rst.Fields("idempresa")) & "' and d.iddocumento = '" & Trim("" & rst.Fields("tipodocreferencia")) & "' " & _
                               " and d.idserie = '" & Trim("" & rst.Fields("seriedocreferencia")) & "' and d.iddocventas = '" & Trim("" & rst.Fields("numdocreferencia")) & "'"
                        If rsDetalle.State = 1 Then rsDetalle.Close
                        rsDetalle.Open csql, Cn, adOpenKeyset, adLockOptimistic
        
                        rsDetalle.MoveFirst
                        Do While Not rsDetalle.EOF
                            item = item + 1
                            rsdetalleCot.AddNew
                            rsdetalleCot!item = item
                            rsdetalleCot!idempresa = "" & rsDetalle.Fields("idempresa")
                            rsdetalleCot!idsucursal = "" & rsDetalle.Fields("idsucursal")
                            rsdetalleCot!idDocumento = "" & rsDetalle.Fields("iddocumento")
                            rsdetalleCot!glsdocreferencia = "" & rsDetalle.Fields("glsdocreferencia")
                            rsdetalleCot!abredocumento = "" & rsDetalle.Fields("abredocumento")
                            rsdetalleCot!simbolo = "" & rsDetalle.Fields("simbolo")
                            rsdetalleCot!idserie = "" & rsDetalle.Fields("idserie")
                            rsdetalleCot!idDocVentas = "" & rsDetalle.Fields("iddocventas")
                            rsdetalleCot!fecEmision = "" & rsDetalle.Fields("fecemision")
                            rsdetalleCot!estdocventas = "" & rsDetalle.Fields("estdocventas")
                            rsdetalleCot!idpercliente = "" & rsDetalle.Fields("idpercliente")
                            rsdetalleCot!GlsCliente = "" & rsDetalle.Fields("glscliente")
                            rsdetalleCot!idMoneda = "" & rsDetalle.Fields("idmoneda")
                            rsdetalleCot!totalprecioventa = "" & rsDetalle.Fields("totalprecioventa")
                            rsdetalleCot.Update
                            
                            rsDetalle.MoveNext
                        Loop
                    End If
                    rst.MoveNext
                Loop
            End If
            rsdetallePed.MoveNext
        Loop
        mostrarDatosGridSQL gLista2, rsdetalleCot, StrMsgError
    End If

    If rsdetalleCot.RecordCount > 0 Then
        item = 0
        rstdetCot.Fields.Append "Item", adInteger, , adFldRowID
        rstdetCot.Fields.Append "idempresa", adChar, 8, adFldRowID
        rstdetCot.Fields.Append "idsucursal", adChar, 8, adFldRowID
        rstdetCot.Fields.Append "iddocumento", adChar, 2, adFldRowID
        rstdetCot.Fields.Append "glsdocreferencia", adChar, 500, adFldRowID
        rstdetCot.Fields.Append "abredocumento", adChar, 5, adFldRowID
        rstdetCot.Fields.Append "simbolo", adChar, 5, adFldRowID
        rstdetCot.Fields.Append "idserie", adVarChar, 4, adFldRowID
        rstdetCot.Fields.Append "iddocventas", adChar, 8, adFldRowID
        rstdetCot.Fields.Append "fecEmision", adVarChar, 15, adFldRowID
        rstdetCot.Fields.Append "estdocventas", adChar, 3, adFldRowID
        rstdetCot.Fields.Append "idpercliente", adVarChar, 8, adFldRowID
        rstdetCot.Fields.Append "glscliente", adVarChar, 100, adFldRowID
        rstdetCot.Fields.Append "idmoneda", adVarChar, 3, adFldRowID
        rstdetCot.Fields.Append "totalprecioventa", adDouble, adFldRowID
        rstdetCot.Open
    
        rsdetalleCot.MoveFirst
        Do While Not rsdetalleCot.EOF
            csql = " select idempresa,tipodocreferencia,seriedocreferencia,numdocreferencia from docreferencia" & _
                    " where idempresa = '" & Trim("" & rsdetalleCot.Fields("idempresa")) & "' and tipodocorigen = '" & Trim("" & rsdetalleCot.Fields("iddocumento")) & "' and seriedocorigen = '" & Trim("" & rsdetalleCot.Fields("idserie")) & "' and numdocorigen = '" & Trim("" & rsdetalleCot.Fields("iddocventas")) & "' "
            If rst.State = 1 Then rst.Close
            rst.Open csql, Cn, adOpenKeyset, adLockOptimistic
        
            If rst.RecordCount > 0 Then
                rst.MoveFirst
                Do While Not rst.EOF
                    If Len(Trim("" & rst.Fields("numdocreferencia"))) = 8 Then
                        csql = " Select d.idempresa,d.idsucursal,d.iddocumento,d.glsdocreferencia,dc.abredocumento,m.simbolo,d.idserie,d.iddocventas,d.fecemision,d.estdocventas,d.idpercliente," & _
                                " d.glscliente,d.idmoneda,d.totalprecioventa from docventas d inner join documentos dc on d.iddocumento = dc.iddocumento " & _
                                " inner join monedas m on d.idmoneda = m.idmoneda " & _
                                " where d.idempresa = '" & Trim("" & rst.Fields("idempresa")) & "' and d.iddocumento = '" & Trim("" & rst.Fields("tipodocreferencia")) & "' " & _
                                " and d.idserie = '" & Trim("" & rst.Fields("seriedocreferencia")) & "' and d.iddocventas = '" & Trim("" & rst.Fields("numdocreferencia")) & "'"
                        If rsDetalle.State = 1 Then rsDetalle.Close
                        rsDetalle.Open csql, Cn, adOpenKeyset, adLockOptimistic
                     
                        rsDetalle.MoveFirst
                        Do While Not rsDetalle.EOF
                            item = item + 1
                            rstdetCot.AddNew
                            rstdetCot!item = item
                            rstdetCot!idempresa = "" & rsDetalle.Fields("idempresa")
                            rstdetCot!idsucursal = "" & rsDetalle.Fields("idsucursal")
                            rstdetCot!idDocumento = "" & rsDetalle.Fields("iddocumento")
                            rstdetCot!glsdocreferencia = "" & rsDetalle.Fields("glsdocreferencia")
                            rstdetCot!abredocumento = "" & rsDetalle.Fields("abredocumento")
                            rstdetCot!simbolo = "" & rsDetalle.Fields("simbolo")
                            rstdetCot!idserie = "" & rsDetalle.Fields("idserie")
                            rstdetCot!idDocVentas = "" & rsDetalle.Fields("iddocventas")
                            rstdetCot!fecEmision = "" & rsDetalle.Fields("fecemision")
                            rstdetCot!estdocventas = "" & rsDetalle.Fields("estdocventas")
                            rstdetCot!idpercliente = "" & rsDetalle.Fields("idpercliente")
                            rstdetCot!GlsCliente = "" & rsDetalle.Fields("glscliente")
                            rstdetCot!idMoneda = "" & rsDetalle.Fields("idmoneda")
                            rstdetCot!totalprecioventa = "" & rsDetalle.Fields("totalprecioventa")
                            rstdetCot.Update
                            
                            rsDetalle.MoveNext
                        Loop
                    End If
                    rst.MoveNext
                Loop
            End If
            rsdetalleCot.MoveNext
        Loop
        mostrarDatosGridSQL gLista1, rstdetCot, StrMsgError
    End If
     
End Sub

Private Sub SecuenciaGuia1()
Dim csql                 As String
Dim rst                  As New ADODB.Recordset
Dim rsg                  As New ADODB.Recordset
Dim rsDetalle            As New ADODB.Recordset
Dim rsdetalleCot         As New ADODB.Recordset
Dim rsdetallePed         As New ADODB.Recordset
Dim rsdetalleFacGuia     As New ADODB.Recordset
Dim rsdetFacGuia         As New ADODB.Recordset
Dim rstdetCot            As New ADODB.Recordset
Dim StrMsgError          As String
Dim item                 As Integer


    csql = " Select (@i:=@i +1) as item,d.idempresa,d.idsucursal,d.iddocumento,d.glsdocreferencia,dc.abredocumento,m.simbolo,d.idserie,d.iddocventas,d.fecemision,d.estdocventas,d.idpercliente," & _
            " d.glscliente,d.idmoneda,d.totalprecioventa from docventas d inner join documentos dc on d.iddocumento = dc.iddocumento " & _
            " inner join monedas m on d.idmoneda = m.idmoneda,(select @i:=0) foo " & _
            " where d.idempresa = '" & glsEmpresa & "' and d.iddocumento = '" & txtCod_TipoDoc.Text & "' " & _
            " and d.idserie = '" & txt_serie.Text & "' and d.iddocventas = '" & txtNum_Documento.Text & "'"
               
    If rst.State = 1 Then rst.Close
    rst.Open csql, Cn, adOpenKeyset, adLockOptimistic
     
    rsg.Fields.Append "item", adInteger, , adFldRowID
    rsg.Fields.Append "idempresa", adChar, 8, adFldIsNullable
    rsg.Fields.Append "idsucursal", adChar, 8, adFldIsNullable
    rsg.Fields.Append "iddocumento", adChar, 2, adFldIsNullable
    rsg.Fields.Append "glsdocreferencia", adChar, 500, adFldIsNullable
    rsg.Fields.Append "abredocumento", adChar, 5, adFldIsNullable
    rsg.Fields.Append "simbolo", adChar, 5, adFldIsNullable
    rsg.Fields.Append "idserie", adChar, 4, adFldIsNullable
    rsg.Fields.Append "iddocventas", adChar, 8, adFldIsNullable
    rsg.Fields.Append "fecemision", adChar, 15, adFldIsNullable
    rsg.Fields.Append "estdocventas", adChar, 3, adFldIsNullable
    rsg.Fields.Append "idpercliente", adChar, 8, adFldIsNullable
    rsg.Fields.Append "glscliente", adVarChar, 100, adFldIsNullable
    rsg.Fields.Append "idmoneda", adChar, 5, adFldIsNullable
    rsg.Fields.Append "totalprecioventa", adDouble, adFldRowID
    rsg.Open
    
    If rst.RecordCount = 0 Then
        rsg.Fields("item") = 1
        rsg.Fields("idempresa") = ""
        rsg.Fields("idsucursal") = ""
        rsg.Fields("iddocumento") = ""
        rsg.Fields("glsdocreferencia") = ""
        rsg.Fields("abredocumento") = ""
        rsg.Fields("simbolo") = ""
        rsg.Fields("idserie") = ""
        rsg.Fields("iddocventas") = ""
        rsg.Fields("fecemision") = ""
        rsg.Fields("estdocventas") = ""
        rsg.Fields("idpercliente") = ""
        rsg.Fields("glscliente") = ""
        rsg.Fields("idmoneda") = ""
        rsg.Fields("totalprecioventa") = ""
                
    Else
        If Not rst.EOF Then
            rst.MoveFirst
            Do While Not rst.EOF
                rsg.AddNew
                rsg.Fields("item") = Val("" & rst.Fields("item"))
                rsg.Fields("idempresa") = Trim("" & rst.Fields("idempresa"))
                rsg.Fields("idsucursal") = Trim("" & rst.Fields("idsucursal"))
                rsg.Fields("iddocumento") = Trim("" & rst.Fields("iddocumento"))
                rsg.Fields("glsdocreferencia") = Trim("" & rst.Fields("glsdocreferencia"))
                rsg.Fields("abredocumento") = Trim("" & rst.Fields("abredocumento"))
                rsg.Fields("simbolo") = Trim("" & rst.Fields("simbolo"))
                rsg.Fields("idserie") = Trim("" & rst.Fields("idserie"))
                rsg.Fields("iddocventas") = Trim("" & rst.Fields("iddocventas"))
                rsg.Fields("fecemision") = "" & rst.Fields("fecemision")
                rsg.Fields("estdocventas") = "" & rst.Fields("estdocventas")
                rsg.Fields("idpercliente") = Trim("" & rst.Fields("idpercliente"))
                rsg.Fields("glscliente") = Trim("" & rst.Fields("glscliente"))
                rsg.Fields("idmoneda") = Trim("" & rst.Fields("idmoneda"))
                rsg.Fields("totalprecioventa") = Trim("" & rst.Fields("totalprecioventa"))
                rst.MoveNext
            Loop
        End If
        mostrarDatosGridSQL gLista4, rsg, StrMsgError
    End If
    
    csql = " select idempresa,tipodocreferencia,seriedocreferencia,numdocreferencia from docreferencia" & _
           " where idempresa = '" & glsEmpresa & "' and tipodocorigen = '" & txtCod_TipoDoc.Text & "' and seriedocorigen = '" & txt_serie.Text & "' and numdocorigen = '" & txtNum_Documento.Text & "' "
    If rst.State = 1 Then rst.Close
    rst.Open csql, Cn, adOpenKeyset, adLockOptimistic
     
    If rst.RecordCount > 0 Then
        item = 0
        rsdetallePed.Fields.Append "Item", adInteger, , adFldRowID
        rsdetallePed.Fields.Append "idempresa", adChar, 8, adFldRowID
        rsdetallePed.Fields.Append "idsucursal", adChar, 8, adFldRowID
        rsdetallePed.Fields.Append "iddocumento", adChar, 2, adFldRowID
        rsdetallePed.Fields.Append "glsdocreferencia", adChar, 500, adFldRowID
        rsdetallePed.Fields.Append "abredocumento", adChar, 5, adFldRowID
        rsdetallePed.Fields.Append "simbolo", adChar, 5, adFldRowID
        rsdetallePed.Fields.Append "idserie", adVarChar, 4, adFldRowID
        rsdetallePed.Fields.Append "iddocventas", adChar, 8, adFldRowID
        rsdetallePed.Fields.Append "fecEmision", adVarChar, 15, adFldRowID
        rsdetallePed.Fields.Append "estdocventas", adChar, 3, adFldRowID
        rsdetallePed.Fields.Append "idpercliente", adVarChar, 8, adFldRowID
        rsdetallePed.Fields.Append "glscliente", adVarChar, 100, adFldRowID
        rsdetallePed.Fields.Append "idmoneda", adVarChar, 3, adFldRowID
        rsdetallePed.Fields.Append "totalprecioventa", adDouble, adFldRowID
        rsdetallePed.Open
                
        rst.MoveFirst
        Do While Not rst.EOF
            If Len(Trim("" & rst.Fields("numdocreferencia"))) = 8 Then
                csql = " Select d.idempresa,d.idsucursal,d.iddocumento,d.glsdocreferencia,dc.abredocumento,m.simbolo,d.idserie,d.iddocventas,d.fecemision,d.estdocventas,d.idpercliente," & _
                        " d.glscliente,d.idmoneda,d.totalprecioventa from docventas d inner join documentos dc on d.iddocumento = dc.iddocumento " & _
                        " inner join monedas m on d.idmoneda = m.idmoneda" & _
                        " where d.idempresa = '" & Trim("" & rst.Fields("idempresa")) & "' and d.iddocumento = '" & Trim("" & rst.Fields("tipodocreferencia")) & "' " & _
                        " and d.idserie = '" & Trim("" & rst.Fields("seriedocreferencia")) & "' and d.iddocventas = '" & Trim("" & rst.Fields("numdocreferencia")) & "'"
                If rsDetalle.State = 1 Then rsDetalle.Close
                rsDetalle.Open csql, Cn, adOpenKeyset, adLockOptimistic
                
                rsDetalle.MoveFirst
                Do While Not rsDetalle.EOF
                    item = item + 1
                    rsdetallePed.AddNew
                    rsdetallePed!item = item
                    rsdetallePed!idempresa = "" & rsDetalle.Fields("idempresa")
                    rsdetallePed!idsucursal = "" & rsDetalle.Fields("idsucursal")
                    rsdetallePed!idDocumento = "" & rsDetalle.Fields("iddocumento")
                    rsdetallePed!glsdocreferencia = "" & rsDetalle.Fields("glsdocreferencia")
                    rsdetallePed!abredocumento = "" & rsDetalle.Fields("abredocumento")
                    rsdetallePed!simbolo = "" & rsDetalle.Fields("simbolo")
                    rsdetallePed!idserie = "" & rsDetalle.Fields("idserie")
                    rsdetallePed!idDocVentas = "" & rsDetalle.Fields("iddocventas")
                    rsdetallePed!fecEmision = "" & rsDetalle.Fields("fecemision")
                    rsdetallePed!estdocventas = "" & rsDetalle.Fields("estdocventas")
                    rsdetallePed!idpercliente = "" & rsDetalle.Fields("idpercliente")
                    rsdetallePed!GlsCliente = "" & rsDetalle.Fields("glscliente")
                    rsdetallePed!idMoneda = "" & rsDetalle.Fields("idmoneda")
                    rsdetallePed!totalprecioventa = "" & rsDetalle.Fields("totalprecioventa")
                    rsdetallePed.Update
                    
                    rsDetalle.MoveNext
                Loop
            End If
            rst.MoveNext
        Loop
        mostrarDatosGridSQL gLista3, rsdetallePed, StrMsgError
    End If

    If rsdetallePed.RecordCount > 0 Then
        item = 0
        rsdetalleCot.Fields.Append "Item", adInteger, , adFldRowID
        rsdetalleCot.Fields.Append "idempresa", adChar, 8, adFldRowID
        rsdetalleCot.Fields.Append "idsucursal", adChar, 8, adFldRowID
        rsdetalleCot.Fields.Append "iddocumento", adChar, 2, adFldRowID
        rsdetalleCot.Fields.Append "glsdocreferencia", adChar, 500, adFldRowID
        rsdetalleCot.Fields.Append "abredocumento", adChar, 5, adFldRowID
        rsdetalleCot.Fields.Append "simbolo", adChar, 5, adFldRowID
        rsdetalleCot.Fields.Append "idserie", adVarChar, 4, adFldRowID
        rsdetalleCot.Fields.Append "iddocventas", adChar, 8, adFldRowID
        rsdetalleCot.Fields.Append "fecEmision", adVarChar, 15, adFldRowID
        rsdetalleCot.Fields.Append "estdocventas", adChar, 3, adFldRowID
        rsdetalleCot.Fields.Append "idpercliente", adVarChar, 8, adFldRowID
        rsdetalleCot.Fields.Append "glscliente", adVarChar, 100, adFldRowID
        rsdetalleCot.Fields.Append "idmoneda", adVarChar, 3, adFldRowID
        rsdetalleCot.Fields.Append "totalprecioventa", adDouble, adFldRowID
        rsdetalleCot.Open

        rsdetallePed.MoveFirst
        Do While Not rsdetallePed.EOF
            csql = " select idempresa,tipodocreferencia,seriedocreferencia,numdocreferencia from docreferencia" & _
                   " where idempresa = '" & Trim("" & rsdetallePed.Fields("idempresa")) & "' and tipodocorigen = '" & Trim("" & rsdetallePed.Fields("iddocumento")) & "' and seriedocorigen = '" & Trim("" & rsdetallePed.Fields("idserie")) & "' and numdocorigen = '" & Trim("" & rsdetallePed.Fields("iddocventas")) & "' "
            If rst.State = 1 Then rst.Close
            rst.Open csql, Cn, adOpenKeyset, adLockOptimistic
             
            If rst.RecordCount > 0 Then
                rst.MoveFirst
                Do While Not rst.EOF
                    If Len(Trim("" & rst.Fields("numdocreferencia"))) = 8 Then
                        csql = " Select d.idempresa,d.idsucursal,d.iddocumento,d.glsdocreferencia,dc.abredocumento,m.simbolo,d.idserie,d.iddocventas,d.fecemision,d.estdocventas,d.idpercliente," & _
                                " d.glscliente,d.idmoneda,d.totalprecioventa from docventas d inner join documentos dc on d.iddocumento = dc.iddocumento " & _
                                " inner join monedas m on d.idmoneda = m.idmoneda" & _
                                " where d.idempresa = '" & Trim("" & rst.Fields("idempresa")) & "' and d.iddocumento = '" & Trim("" & rst.Fields("tipodocreferencia")) & "' " & _
                                " and d.idserie = '" & Trim("" & rst.Fields("seriedocreferencia")) & "' and d.iddocventas = '" & Trim("" & rst.Fields("numdocreferencia")) & "'"
                        If rsDetalle.State = 1 Then rsDetalle.Close
                        rsDetalle.Open csql, Cn, adOpenKeyset, adLockOptimistic
                            
                        rsDetalle.MoveFirst
                        Do While Not rsDetalle.EOF
                            item = item + 1
                            rsdetalleCot.AddNew
                            rsdetalleCot!item = item
                            rsdetalleCot!idempresa = "" & rsDetalle.Fields("idempresa")
                            rsdetalleCot!idsucursal = "" & rsDetalle.Fields("idsucursal")
                            rsdetalleCot!idDocumento = "" & rsDetalle.Fields("iddocumento")
                            rsdetalleCot!glsdocreferencia = "" & rsDetalle.Fields("glsdocreferencia")
                            rsdetalleCot!abredocumento = "" & rsDetalle.Fields("abredocumento")
                            rsdetalleCot!simbolo = "" & rsDetalle.Fields("simbolo")
                            rsdetalleCot!idserie = "" & rsDetalle.Fields("idserie")
                            rsdetalleCot!idDocVentas = "" & rsDetalle.Fields("iddocventas")
                            rsdetalleCot!fecEmision = "" & rsDetalle.Fields("fecemision")
                            rsdetalleCot!estdocventas = "" & rsDetalle.Fields("estdocventas")
                            rsdetalleCot!idpercliente = "" & rsDetalle.Fields("idpercliente")
                            rsdetalleCot!GlsCliente = "" & rsDetalle.Fields("glscliente")
                            rsdetalleCot!idMoneda = "" & rsDetalle.Fields("idmoneda")
                            rsdetalleCot!totalprecioventa = "" & rsDetalle.Fields("totalprecioventa")
                            rsdetalleCot.Update
                       
                            rsDetalle.MoveNext
                        Loop
                    End If
                    rst.MoveNext
                Loop
            End If
            rsdetallePed.MoveNext
        Loop
        mostrarDatosGridSQL gLista2, rsdetalleCot, StrMsgError
    End If
         
    If rsdetalleCot.RecordCount > 0 Then
        item = 0
        rstdetCot.Fields.Append "Item", adInteger, , adFldRowID
        rstdetCot.Fields.Append "idempresa", adChar, 8, adFldRowID
        rstdetCot.Fields.Append "idsucursal", adChar, 8, adFldRowID
        rstdetCot.Fields.Append "iddocumento", adChar, 2, adFldRowID
        rstdetCot.Fields.Append "glsdocreferencia", adChar, 500, adFldRowID
        rstdetCot.Fields.Append "abredocumento", adChar, 5, adFldRowID
        rstdetCot.Fields.Append "simbolo", adChar, 5, adFldRowID
        rstdetCot.Fields.Append "idserie", adVarChar, 4, adFldRowID
        rstdetCot.Fields.Append "iddocventas", adChar, 8, adFldRowID
        rstdetCot.Fields.Append "fecEmision", adVarChar, 15, adFldRowID
        rstdetCot.Fields.Append "estdocventas", adChar, 3, adFldRowID
        rstdetCot.Fields.Append "idpercliente", adVarChar, 8, adFldRowID
        rstdetCot.Fields.Append "glscliente", adVarChar, 100, adFldRowID
        rstdetCot.Fields.Append "idmoneda", adVarChar, 3, adFldRowID
        rstdetCot.Fields.Append "totalprecioventa", adDouble, adFldRowID
        rstdetCot.Open
        
        rsdetalleCot.MoveFirst
        Do While Not rsdetalleCot.EOF
            csql = " select idempresa,tipodocreferencia,seriedocreferencia,numdocreferencia from docreferencia" & _
                    " where idempresa = '" & Trim("" & rsdetalleCot.Fields("idempresa")) & "' and tipodocorigen = '" & Trim("" & rsdetalleCot.Fields("iddocumento")) & "' and seriedocorigen = '" & Trim("" & rsdetalleCot.Fields("idserie")) & "' and numdocorigen = '" & Trim("" & rsdetalleCot.Fields("iddocventas")) & "' "
            If rst.State = 1 Then rst.Close
            rst.Open csql, Cn, adOpenKeyset, adLockOptimistic
        
            If rst.RecordCount > 0 Then
                rst.MoveFirst
                Do While Not rst.EOF
                    If Len(Trim("" & rst.Fields("numdocreferencia"))) = 8 Then
                        csql = " Select d.idempresa,d.idsucursal,d.iddocumento,d.glsdocreferencia,dc.abredocumento,m.simbolo,d.idserie,d.iddocventas,d.fecemision,d.estdocventas,d.idpercliente," & _
                                " d.glscliente,d.idmoneda,d.totalprecioventa from docventas d inner join documentos dc on  d.iddocumento = dc.iddocumento " & _
                                " inner join monedas m  on d.idmoneda = m.idmoneda " & _
                                " where d.idempresa = '" & Trim("" & rst.Fields("idempresa")) & "' and d.iddocumento = '" & Trim("" & rst.Fields("tipodocreferencia")) & "' " & _
                                " and d.idserie = '" & Trim("" & rst.Fields("seriedocreferencia")) & "' and d.iddocventas = '" & Trim("" & rst.Fields("numdocreferencia")) & "'"
                        If rsDetalle.State = 1 Then rsDetalle.Close
                        rsDetalle.Open csql, Cn, adOpenKeyset, adLockOptimistic
                         
                        rsDetalle.MoveFirst
                        Do While Not rsDetalle.EOF
                            item = item + 1
                            rstdetCot.AddNew
                            rstdetCot!item = item
                            rstdetCot!idempresa = "" & rsDetalle.Fields("idempresa")
                            rstdetCot!idsucursal = "" & rsDetalle.Fields("idsucursal")
                            rstdetCot!idDocumento = "" & rsDetalle.Fields("iddocumento")
                            rstdetCot!glsdocreferencia = "" & rsDetalle.Fields("glsdocreferencia")
                            rstdetCot!abredocumento = "" & rsDetalle.Fields("abredocumento")
                            rstdetCot!simbolo = "" & rsDetalle.Fields("simbolo")
                            rstdetCot!idserie = "" & rsDetalle.Fields("idserie")
                            rstdetCot!idDocVentas = "" & rsDetalle.Fields("iddocventas")
                            rstdetCot!fecEmision = "" & rsDetalle.Fields("fecemision")
                            rstdetCot!estdocventas = "" & rsDetalle.Fields("estdocventas")
                            rstdetCot!idpercliente = "" & rsDetalle.Fields("idpercliente")
                            rstdetCot!GlsCliente = "" & rsDetalle.Fields("glscliente")
                            rstdetCot!idMoneda = "" & rsDetalle.Fields("idmoneda")
                            rstdetCot!totalprecioventa = "" & rsDetalle.Fields("totalprecioventa")
                            rstdetCot.Update
                            
                            rsDetalle.MoveNext
                        Loop
                    End If
                    rst.MoveNext
                Loop
            End If
            rsdetalleCot.MoveNext
        Loop
        mostrarDatosGridSQL gLista1, rstdetCot, StrMsgError
    End If
    
End Sub

Private Sub SecuenciaGuia2()
Dim csql                 As String
Dim rst                  As New ADODB.Recordset
Dim rsg                  As New ADODB.Recordset
Dim rsDetalle            As New ADODB.Recordset
Dim rsdetalleCot         As New ADODB.Recordset
Dim rsdetallePed         As New ADODB.Recordset
Dim rsdetalleFacGuia     As New ADODB.Recordset
Dim rsdetFacGuia         As New ADODB.Recordset
Dim rstdetCot            As New ADODB.Recordset
Dim StrMsgError          As String
Dim item                 As Integer

    csql = " Select (@i:=@i +1) as item,d.idempresa,d.idsucursal,d.iddocumento,d.glsdocreferencia,dc.abredocumento,m.simbolo,d.idserie,d.iddocventas,d.fecemision,d.estdocventas,d.idpercliente," & _
            " d.glscliente,d.idmoneda,d.totalprecioventa from docventas d inner join documentos dc on d.iddocumento = dc.iddocumento " & _
            " inner join monedas m on d.idmoneda = m.idmoneda,(select @i:=0) foo " & _
            " where d.idempresa = '" & glsEmpresa & "' and d.iddocumento = '" & txtCod_TipoDoc.Text & "' " & _
            " and d.idserie = '" & txt_serie.Text & "' and d.iddocventas = '" & txtNum_Documento.Text & "'"
               
    If rst.State = 1 Then rst.Close
    rst.Open csql, Cn, adOpenKeyset, adLockOptimistic
     
    rsg.Fields.Append "item", adInteger, , adFldRowID
    rsg.Fields.Append "idempresa", adChar, 8, adFldIsNullable
    rsg.Fields.Append "idsucursal", adChar, 8, adFldIsNullable
    rsg.Fields.Append "iddocumento", adChar, 2, adFldIsNullable
    rsg.Fields.Append "glsdocreferencia", adChar, 500, adFldIsNullable
    rsg.Fields.Append "abredocumento", adChar, 5, adFldIsNullable
    rsg.Fields.Append "simbolo", adChar, 5, adFldIsNullable
    rsg.Fields.Append "idserie", adChar, 4, adFldIsNullable
    rsg.Fields.Append "iddocventas", adChar, 8, adFldIsNullable
    rsg.Fields.Append "fecemision", adChar, 15, adFldIsNullable
    rsg.Fields.Append "estdocventas", adChar, 3, adFldIsNullable
    rsg.Fields.Append "idpercliente", adChar, 8, adFldIsNullable
    rsg.Fields.Append "glscliente", adVarChar, 100, adFldIsNullable
    rsg.Fields.Append "idmoneda", adChar, 5, adFldIsNullable
    rsg.Fields.Append "totalprecioventa", adDouble, adFldRowID
    rsg.Open
    
    If rst.RecordCount = 0 Then
        rsg.Fields("item") = 1
        rsg.Fields("idempresa") = ""
        rsg.Fields("idsucursal") = ""
        rsg.Fields("iddocumento") = ""
        rsg.Fields("glsdocreferencia") = ""
        rsg.Fields("abredocumento") = ""
        rsg.Fields("simbolo") = ""
        rsg.Fields("idserie") = ""
        rsg.Fields("iddocventas") = ""
        rsg.Fields("fecemision") = ""
        rsg.Fields("estdocventas") = ""
        rsg.Fields("idpercliente") = ""
        rsg.Fields("glscliente") = ""
        rsg.Fields("idmoneda") = ""
        rsg.Fields("totalprecioventa") = ""
                
    Else
        If Not rst.EOF Then
            rst.MoveFirst
            Do While Not rst.EOF
                rsg.AddNew
                rsg.Fields("item") = Val("" & rst.Fields("item"))
                rsg.Fields("idempresa") = Trim("" & rst.Fields("idempresa"))
                rsg.Fields("idsucursal") = Trim("" & rst.Fields("idsucursal"))
                rsg.Fields("iddocumento") = Trim("" & rst.Fields("iddocumento"))
                rsg.Fields("glsdocreferencia") = Trim("" & rst.Fields("glsdocreferencia"))
                rsg.Fields("abredocumento") = Trim("" & rst.Fields("abredocumento"))
                rsg.Fields("simbolo") = Trim("" & rst.Fields("simbolo"))
                rsg.Fields("idserie") = Trim("" & rst.Fields("idserie"))
                rsg.Fields("iddocventas") = Trim("" & rst.Fields("iddocventas"))
                rsg.Fields("fecemision") = "" & rst.Fields("fecemision")
                rsg.Fields("estdocventas") = "" & rst.Fields("estdocventas")
                rsg.Fields("idpercliente") = Trim("" & rst.Fields("idpercliente"))
                rsg.Fields("glscliente") = Trim("" & rst.Fields("glscliente"))
                rsg.Fields("idmoneda") = Trim("" & rst.Fields("idmoneda"))
                rsg.Fields("totalprecioventa") = Trim("" & rst.Fields("totalprecioventa"))
                rst.MoveNext
            Loop
        End If
    End If
    
    mostrarDatosGridSQL gLista3, rsg, StrMsgError

    If rsg.RecordCount > 0 Then
        item = 0
        rsdetalleCot.Fields.Append "Item", adInteger, , adFldRowID
        rsdetalleCot.Fields.Append "idempresa", adChar, 8, adFldRowID
        rsdetalleCot.Fields.Append "idsucursal", adChar, 8, adFldRowID
        rsdetalleCot.Fields.Append "iddocumento", adChar, 2, adFldRowID
        rsdetalleCot.Fields.Append "glsdocreferencia", adChar, 500, adFldRowID
        rsdetalleCot.Fields.Append "abredocumento", adChar, 5, adFldRowID
        rsdetalleCot.Fields.Append "simbolo", adChar, 5, adFldRowID
        rsdetalleCot.Fields.Append "idserie", adVarChar, 4, adFldRowID
        rsdetalleCot.Fields.Append "iddocventas", adChar, 8, adFldRowID
        rsdetalleCot.Fields.Append "fecEmision", adVarChar, 15, adFldRowID
        rsdetalleCot.Fields.Append "estdocventas", adChar, 3, adFldRowID
        rsdetalleCot.Fields.Append "idpercliente", adVarChar, 8, adFldRowID
        rsdetalleCot.Fields.Append "glscliente", adVarChar, 100, adFldRowID
        rsdetalleCot.Fields.Append "idmoneda", adVarChar, 3, adFldRowID
        rsdetalleCot.Fields.Append "totalprecioventa", adDouble, adFldRowID
        rsdetalleCot.Open
        
        rsg.MoveFirst
        Do While Not rsg.EOF
            csql = " select idempresa,tipodocreferencia,seriedocreferencia,numdocreferencia from docreferencia" & _
                   " where idempresa = '" & Trim("" & rsg.Fields("idempresa")) & "' and tipodocorigen = '" & Trim("" & rsg.Fields("iddocumento")) & "' and seriedocorigen = '" & Trim("" & rsg.Fields("idserie")) & "' and numdocorigen = '" & Trim("" & rsg.Fields("iddocventas")) & "' "
            If rst.State = 1 Then rst.Close
            rst.Open csql, Cn, adOpenKeyset, adLockOptimistic
             
            If rst.RecordCount > 0 Then
                rst.MoveFirst
                Do While Not rst.EOF
                    If Len(Trim("" & rst.Fields("numdocreferencia"))) = 8 Then
                        csql = " Select d.idempresa,d.idsucursal,d.iddocumento,d.glsdocreferencia,dc.abredocumento,m.simbolo,d.idserie,d.iddocventas,d.fecemision,d.estdocventas,d.idpercliente," & _
                                " d.glscliente,d.idmoneda,d.totalprecioventa from docventas d inner join documentos dc on d.iddocumento = dc.iddocumento " & _
                                " inner join monedas m on d.idmoneda = m.idmoneda " & _
                                " where d.idempresa = '" & Trim("" & rst.Fields("idempresa")) & "' and d.iddocumento = '" & Trim("" & rst.Fields("tipodocreferencia")) & "' " & _
                                " and d.idserie = '" & Trim("" & rst.Fields("seriedocreferencia")) & "' and d.iddocventas = '" & Trim("" & rst.Fields("numdocreferencia")) & "'"
                        If rsDetalle.State = 1 Then rsDetalle.Close
                        rsDetalle.Open csql, Cn, adOpenKeyset, adLockOptimistic
        
                        If rsDetalle.RecordCount > 0 Then
                            rsDetalle.MoveFirst
                            Do While Not rsDetalle.EOF
                                item = item + 1
                                rsdetalleCot.AddNew
                                rsdetalleCot!item = item
                                rsdetalleCot!idempresa = "" & rsDetalle.Fields("idempresa")
                                rsdetalleCot!idsucursal = "" & rsDetalle.Fields("idsucursal")
                                rsdetalleCot!idDocumento = "" & rsDetalle.Fields("iddocumento")
                                rsdetalleCot!glsdocreferencia = "" & rsDetalle.Fields("glsdocreferencia")
                                rsdetalleCot!abredocumento = "" & rsDetalle.Fields("abredocumento")
                                rsdetalleCot!simbolo = "" & rsDetalle.Fields("simbolo")
                                rsdetalleCot!idserie = "" & rsDetalle.Fields("idserie")
                                rsdetalleCot!idDocVentas = "" & rsDetalle.Fields("iddocventas")
                                rsdetalleCot!fecEmision = "" & rsDetalle.Fields("fecemision")
                                rsdetalleCot!estdocventas = "" & rsDetalle.Fields("estdocventas")
                                rsdetalleCot!idpercliente = "" & rsDetalle.Fields("idpercliente")
                                rsdetalleCot!GlsCliente = "" & rsDetalle.Fields("glscliente")
                                rsdetalleCot!idMoneda = "" & rsDetalle.Fields("idmoneda")
                                rsdetalleCot!totalprecioventa = "" & rsDetalle.Fields("totalprecioventa")
                                rsdetalleCot.Update
                                
                                rsDetalle.MoveNext
                            Loop
                        End If
                    End If
                    rst.MoveNext
                Loop
            End If
        rsg.MoveNext
        Loop
        mostrarDatosGridSQL gLista2, rsdetalleCot, StrMsgError
    End If

    If rsdetalleCot.RecordCount > 0 Then
        item = 0
        rstdetCot.Fields.Append "Item", adInteger, , adFldRowID
        rstdetCot.Fields.Append "idempresa", adChar, 8, adFldRowID
        rstdetCot.Fields.Append "idsucursal", adChar, 8, adFldRowID
        rstdetCot.Fields.Append "iddocumento", adChar, 2, adFldRowID
        rstdetCot.Fields.Append "glsdocreferencia", adChar, 500, adFldRowID
        rstdetCot.Fields.Append "abredocumento", adChar, 5, adFldRowID
        rstdetCot.Fields.Append "simbolo", adChar, 5, adFldRowID
        rstdetCot.Fields.Append "idserie", adVarChar, 4, adFldRowID
        rstdetCot.Fields.Append "iddocventas", adChar, 8, adFldRowID
        rstdetCot.Fields.Append "fecEmision", adVarChar, 15, adFldRowID
        rstdetCot.Fields.Append "estdocventas", adChar, 3, adFldRowID
        rstdetCot.Fields.Append "idpercliente", adVarChar, 8, adFldRowID
        rstdetCot.Fields.Append "glscliente", adVarChar, 100, adFldRowID
        rstdetCot.Fields.Append "idmoneda", adVarChar, 3, adFldRowID
        rstdetCot.Fields.Append "totalprecioventa", adDouble, adFldRowID
        rstdetCot.Open
    
        rsdetalleCot.MoveFirst
        Do While Not rsdetalleCot.EOF
            csql = " select idempresa,tipodocreferencia,seriedocreferencia,numdocreferencia from docreferencia" & _
                    " where idempresa = '" & Trim("" & rsdetalleCot.Fields("idempresa")) & "' and tipodocorigen = '" & Trim("" & rsdetalleCot.Fields("iddocumento")) & "' and seriedocorigen = '" & Trim("" & rsdetalleCot.Fields("idserie")) & "' and numdocorigen = '" & Trim("" & rsdetalleCot.Fields("iddocventas")) & "' "
            If rst.State = 1 Then rst.Close
            rst.Open csql, Cn, adOpenKeyset, adLockOptimistic
        
            If rst.RecordCount > 0 Then
                rst.MoveFirst
                Do While Not rst.EOF
                    If Len(Trim("" & rst.Fields("numdocreferencia"))) = 8 Then
                        csql = " Select d.idempresa,d.idsucursal,d.iddocumento,d.glsdocreferencia,dc.abredocumento,m.simbolo,d.idserie,d.iddocventas,d.fecemision,d.estdocventas,d.idpercliente," & _
                                " d.glscliente,d.idmoneda,d.totalprecioventa from docventas d inner join documentos dc on d.iddocumento = dc.iddocumento " & _
                                " inner join monedas m on d.idmoneda = m.idmoneda " & _
                                " where d.idempresa = '" & Trim("" & rst.Fields("idempresa")) & "' and d.iddocumento = '" & Trim("" & rst.Fields("tipodocreferencia")) & "' " & _
                                " and d.idserie = '" & Trim("" & rst.Fields("seriedocreferencia")) & "' and d.iddocventas = '" & Trim("" & rst.Fields("numdocreferencia")) & "'"
                        If rsDetalle.State = 1 Then rsDetalle.Close
                        rsDetalle.Open csql, Cn, adOpenKeyset, adLockOptimistic
                     
                        rsDetalle.MoveFirst
                        Do While Not rsDetalle.EOF
                            item = item + 1
                            rstdetCot.AddNew
                            rstdetCot!item = item
                            rstdetCot!idempresa = "" & rsDetalle.Fields("idempresa")
                            rstdetCot!idsucursal = "" & rsDetalle.Fields("idsucursal")
                            rstdetCot!idDocumento = "" & rsDetalle.Fields("iddocumento")
                            rstdetCot!glsdocreferencia = "" & rsDetalle.Fields("glsdocreferencia")
                            rstdetCot!abredocumento = "" & rsDetalle.Fields("abredocumento")
                            rstdetCot!simbolo = "" & rsDetalle.Fields("simbolo")
                            rstdetCot!idserie = "" & rsDetalle.Fields("idserie")
                            rstdetCot!idDocVentas = "" & rsDetalle.Fields("iddocventas")
                            rstdetCot!fecEmision = "" & rsDetalle.Fields("fecemision")
                            rstdetCot!estdocventas = "" & rsDetalle.Fields("estdocventas")
                            rstdetCot!idpercliente = "" & rsDetalle.Fields("idpercliente")
                            rstdetCot!GlsCliente = "" & rsDetalle.Fields("glscliente")
                            rstdetCot!idMoneda = "" & rsDetalle.Fields("idmoneda")
                            rstdetCot!totalprecioventa = "" & rsDetalle.Fields("totalprecioventa")
                            rstdetCot.Update
                  
                            rsDetalle.MoveNext
                        Loop
                    End If
                    rst.MoveNext
                Loop
            End If
            rsdetalleCot.MoveNext
        Loop
        mostrarDatosGridSQL gLista1, rstdetCot, StrMsgError
    End If
     
    csql = " select idempresa,tipodocorigen,seriedocorigen,numdocorigen from docreferencia" & _
            " where idempresa = '" & glsEmpresa & "' and tipodocreferencia = '" & txtCod_TipoDoc.Text & "' and seriedocreferencia = '" & txt_serie.Text & "' and numdocreferencia = '" & txtNum_Documento.Text & "' "
    If rst.State = 1 Then rst.Close
    rst.Open csql, Cn, adOpenKeyset, adLockOptimistic
     
    If rst.RecordCount > 0 Then
        item = 0
        rsdetallePed.Fields.Append "Item", adInteger, , adFldRowID
        rsdetallePed.Fields.Append "idempresa", adChar, 8, adFldRowID
        rsdetallePed.Fields.Append "idsucursal", adChar, 8, adFldRowID
        rsdetallePed.Fields.Append "iddocumento", adChar, 2, adFldRowID
        rsdetallePed.Fields.Append "glsdocreferencia", adChar, 500, adFldRowID
        rsdetallePed.Fields.Append "abredocumento", adChar, 5, adFldRowID
        rsdetallePed.Fields.Append "simbolo", adChar, 5, adFldRowID
        rsdetallePed.Fields.Append "idserie", adVarChar, 4, adFldRowID
        rsdetallePed.Fields.Append "iddocventas", adChar, 8, adFldRowID
        rsdetallePed.Fields.Append "fecEmision", adVarChar, 15, adFldRowID
        rsdetallePed.Fields.Append "estdocventas", adChar, 3, adFldRowID
        rsdetallePed.Fields.Append "idpercliente", adVarChar, 8, adFldRowID
        rsdetallePed.Fields.Append "glscliente", adVarChar, 100, adFldRowID
        rsdetallePed.Fields.Append "idmoneda", adVarChar, 3, adFldRowID
        rsdetallePed.Fields.Append "totalprecioventa", adDouble, adFldRowID
        rsdetallePed.Open
        
        rst.MoveFirst
        Do While Not rst.EOF
            If Len(Trim("" & rst.Fields("numdocorigen"))) = 8 And Trim("" & rst.Fields("tipodocorigen")) <> "99" Then
                csql = " Select d.idempresa,d.idsucursal,d.iddocumento,d.glsdocreferencia,dc.abredocumento,m.simbolo,d.idserie,d.iddocventas,d.fecemision,d.estdocventas,d.idpercliente," & _
                        " d.glscliente,d.idmoneda,d.totalprecioventa from docventas d inner join documentos dc on d.iddocumento = dc.iddocumento " & _
                        " inner join monedas m on d.idmoneda = m.idmoneda " & _
                        " where d.idempresa = '" & Trim("" & rst.Fields("idempresa")) & "' and d.iddocumento = '" & Trim("" & rst.Fields("tipodocorigen")) & "' " & _
                        " and d.idserie = '" & Trim("" & rst.Fields("seriedocorigen")) & "' and d.iddocventas = '" & Trim("" & rst.Fields("numdocorigen")) & "'"
                If rsDetalle.State = 1 Then rsDetalle.Close
                rsDetalle.Open csql, Cn, adOpenKeyset, adLockOptimistic
                
                rsDetalle.MoveFirst
                Do While Not rsDetalle.EOF
                    item = item + 1
                    rsdetallePed.AddNew
                    rsdetallePed!item = item
                    rsdetallePed!idempresa = "" & rsDetalle.Fields("idempresa")
                    rsdetallePed!idsucursal = "" & rsDetalle.Fields("idsucursal")
                    rsdetallePed!idDocumento = "" & rsDetalle.Fields("iddocumento")
                    rsdetallePed!glsdocreferencia = "" & rsDetalle.Fields("glsdocreferencia")
                    rsdetallePed!abredocumento = "" & rsDetalle.Fields("abredocumento")
                    rsdetallePed!simbolo = "" & rsDetalle.Fields("simbolo")
                    rsdetallePed!idserie = "" & rsDetalle.Fields("idserie")
                    rsdetallePed!idDocVentas = "" & rsDetalle.Fields("iddocventas")
                    rsdetallePed!fecEmision = "" & rsDetalle.Fields("fecemision")
                    rsdetallePed!estdocventas = "" & rsDetalle.Fields("estdocventas")
                    rsdetallePed!idpercliente = "" & rsDetalle.Fields("idpercliente")
                    rsdetallePed!GlsCliente = "" & rsDetalle.Fields("glscliente")
                    rsdetallePed!idMoneda = "" & rsDetalle.Fields("idmoneda")
                    rsdetallePed!totalprecioventa = "" & rsDetalle.Fields("totalprecioventa")
                    rsdetallePed.Update
                    
                    rsDetalle.MoveNext
                Loop
            End If
            rst.MoveNext
        Loop
        mostrarDatosGridSQL gLista4, rsdetallePed, StrMsgError
    End If

End Sub

Private Sub SecuenciaBoleta1()
Dim csql                 As String
Dim rst                  As New ADODB.Recordset
Dim rsg                  As New ADODB.Recordset
Dim rsDetalle            As New ADODB.Recordset
Dim rsdetalleCot         As New ADODB.Recordset
Dim rsdetallePed         As New ADODB.Recordset
Dim rsdetalleFacGuia     As New ADODB.Recordset
Dim rsdetFacGuia         As New ADODB.Recordset
Dim rstdetCot            As New ADODB.Recordset
Dim StrMsgError          As String
Dim item                 As Integer

    csql = " Select (@i:=@i +1) as item,d.idempresa,d.idsucursal,d.iddocumento,d.glsdocreferencia,dc.abredocumento,m.simbolo,d.idserie,d.iddocventas,d.fecemision,d.estdocventas,d.idpercliente," & _
            " d.glscliente,d.idmoneda,d.totalprecioventa from docventas d inner join documentos dc on d.iddocumento = dc.iddocumento " & _
            " inner join monedas m on d.idmoneda = m.idmoneda,(select @i:=0) foo " & _
            " where d.idempresa = '" & glsEmpresa & "' and d.iddocumento = '" & txtCod_TipoDoc.Text & "' " & _
            " and d.idserie = '" & txt_serie.Text & "' and d.iddocventas = '" & txtNum_Documento.Text & "'"
               
    If rst.State = 1 Then rst.Close
    rst.Open csql, Cn, adOpenKeyset, adLockOptimistic
     
    rsg.Fields.Append "item", adInteger, , adFldRowID
    rsg.Fields.Append "idempresa", adChar, 8, adFldIsNullable
    rsg.Fields.Append "idsucursal", adChar, 8, adFldIsNullable
    rsg.Fields.Append "iddocumento", adChar, 2, adFldIsNullable
    rsg.Fields.Append "glsdocreferencia", adChar, 500, adFldIsNullable
    rsg.Fields.Append "abredocumento", adChar, 5, adFldIsNullable
    rsg.Fields.Append "simbolo", adChar, 5, adFldIsNullable
    rsg.Fields.Append "idserie", adChar, 4, adFldIsNullable
    rsg.Fields.Append "iddocventas", adChar, 8, adFldIsNullable
    rsg.Fields.Append "fecemision", adChar, 15, adFldIsNullable
    rsg.Fields.Append "estdocventas", adChar, 3, adFldIsNullable
    rsg.Fields.Append "idpercliente", adChar, 8, adFldIsNullable
    rsg.Fields.Append "glscliente", adVarChar, 100, adFldIsNullable
    rsg.Fields.Append "idmoneda", adChar, 5, adFldIsNullable
    rsg.Fields.Append "totalprecioventa", adDouble, adFldRowID
    rsg.Open
    
    If rst.RecordCount = 0 Then
        rsg.Fields("item") = 1
        rsg.Fields("idempresa") = ""
        rsg.Fields("idsucursal") = ""
        rsg.Fields("iddocumento") = ""
        rsg.Fields("glsdocreferencia") = ""
        rsg.Fields("abredocumento") = ""
        rsg.Fields("simbolo") = ""
        rsg.Fields("idserie") = ""
        rsg.Fields("iddocventas") = ""
        rsg.Fields("fecemision") = ""
        rsg.Fields("estdocventas") = ""
        rsg.Fields("idpercliente") = ""
        rsg.Fields("glscliente") = ""
        rsg.Fields("idmoneda") = ""
        rsg.Fields("totalprecioventa") = ""
                
    Else
        If Not rst.EOF Then
            rst.MoveFirst
            Do While Not rst.EOF
                rsg.AddNew
                rsg.Fields("item") = Val("" & rst.Fields("item"))
                rsg.Fields("idempresa") = Trim("" & rst.Fields("idempresa"))
                rsg.Fields("idsucursal") = Trim("" & rst.Fields("idsucursal"))
                rsg.Fields("iddocumento") = Trim("" & rst.Fields("iddocumento"))
                rsg.Fields("glsdocreferencia") = Trim("" & rst.Fields("glsdocreferencia"))
                rsg.Fields("abredocumento") = Trim("" & rst.Fields("abredocumento"))
                rsg.Fields("simbolo") = Trim("" & rst.Fields("simbolo"))
                rsg.Fields("idserie") = Trim("" & rst.Fields("idserie"))
                rsg.Fields("iddocventas") = Trim("" & rst.Fields("iddocventas"))
                rsg.Fields("fecemision") = "" & rst.Fields("fecemision")
                rsg.Fields("estdocventas") = "" & rst.Fields("estdocventas")
                rsg.Fields("idpercliente") = Trim("" & rst.Fields("idpercliente"))
                rsg.Fields("glscliente") = Trim("" & rst.Fields("glscliente"))
                rsg.Fields("idmoneda") = Trim("" & rst.Fields("idmoneda"))
                rsg.Fields("totalprecioventa") = Trim("" & rst.Fields("totalprecioventa"))
                rst.MoveNext
            Loop
        End If
         mostrarDatosGridSQL gLista3, rsg, StrMsgError
    End If
    
    csql = " select idempresa,tipodocreferencia,seriedocreferencia,numdocreferencia from docreferencia" & _
            " where idempresa = '" & glsEmpresa & "' and tipodocorigen = '" & txtCod_TipoDoc.Text & "' and seriedocorigen = '" & txt_serie.Text & "' and numdocorigen = '" & txtNum_Documento.Text & "' "
    If rst.State = 1 Then rst.Close
    rst.Open csql, Cn, adOpenKeyset, adLockOptimistic
     
    If rst.RecordCount > 0 Then
        item = 0
        rsdetallePed.Fields.Append "Item", adInteger, , adFldRowID
        rsdetallePed.Fields.Append "idempresa", adChar, 8, adFldRowID
        rsdetallePed.Fields.Append "idsucursal", adChar, 8, adFldRowID
        rsdetallePed.Fields.Append "iddocumento", adChar, 2, adFldRowID
        rsdetallePed.Fields.Append "glsdocreferencia", adChar, 500, adFldRowID
        rsdetallePed.Fields.Append "abredocumento", adChar, 5, adFldRowID
        rsdetallePed.Fields.Append "simbolo", adChar, 5, adFldRowID
        rsdetallePed.Fields.Append "idserie", adVarChar, 4, adFldRowID
        rsdetallePed.Fields.Append "iddocventas", adChar, 8, adFldRowID
        rsdetallePed.Fields.Append "fecEmision", adVarChar, 15, adFldRowID
        rsdetallePed.Fields.Append "estdocventas", adChar, 3, adFldRowID
        rsdetallePed.Fields.Append "idpercliente", adVarChar, 8, adFldRowID
        rsdetallePed.Fields.Append "glscliente", adVarChar, 100, adFldRowID
        rsdetallePed.Fields.Append "idmoneda", adVarChar, 3, adFldRowID
        rsdetallePed.Fields.Append "totalprecioventa", adDouble, adFldRowID
        rsdetallePed.Open
        
        rst.MoveFirst
        Do While Not rst.EOF
            If Len(Trim("" & rst.Fields("numdocreferencia"))) = 8 Then
                csql = " Select d.idempresa,d.idsucursal,d.iddocumento,d.glsdocreferencia,dc.abredocumento,m.simbolo,d.idserie,d.iddocventas,d.fecemision,d.estdocventas,d.idpercliente," & _
                        " d.glscliente,d.idmoneda,d.totalprecioventa from docventas d inner join documentos dc on d.iddocumento = dc.iddocumento " & _
                        " inner join monedas m on d.idmoneda = m.idmoneda " & _
                        " where d.idempresa = '" & Trim("" & rst.Fields("idempresa")) & "' and d.iddocumento = '" & Trim("" & rst.Fields("tipodocreferencia")) & "' " & _
                        " and d.idserie = '" & Trim("" & rst.Fields("seriedocreferencia")) & "' and d.iddocventas = '" & Trim("" & rst.Fields("numdocreferencia")) & "'"
                If rsDetalle.State = 1 Then rsDetalle.Close
                rsDetalle.Open csql, Cn, adOpenKeyset, adLockOptimistic
                
                rsDetalle.MoveFirst
                Do While Not rsDetalle.EOF
                    item = item + 1
                    rsdetallePed.AddNew
                    rsdetallePed!item = item
                    rsdetallePed!idempresa = "" & rsDetalle.Fields("idempresa")
                    rsdetallePed!idsucursal = "" & rsDetalle.Fields("idsucursal")
                    rsdetallePed!idDocumento = "" & rsDetalle.Fields("iddocumento")
                    rsdetallePed!glsdocreferencia = "" & rsDetalle.Fields("glsdocreferencia")
                    rsdetallePed!abredocumento = "" & rsDetalle.Fields("abredocumento")
                    rsdetallePed!simbolo = "" & rsDetalle.Fields("simbolo")
                    rsdetallePed!idserie = "" & rsDetalle.Fields("idserie")
                    rsdetallePed!idDocVentas = "" & rsDetalle.Fields("iddocventas")
                    rsdetallePed!fecEmision = "" & rsDetalle.Fields("fecemision")
                    rsdetallePed!estdocventas = "" & rsDetalle.Fields("estdocventas")
                    rsdetallePed!idpercliente = "" & rsDetalle.Fields("idpercliente")
                    rsdetallePed!GlsCliente = "" & rsDetalle.Fields("glscliente")
                    rsdetallePed!idMoneda = "" & rsDetalle.Fields("idmoneda")
                    rsdetallePed!totalprecioventa = "" & rsDetalle.Fields("totalprecioventa")
                    rsdetallePed.Update
                    
                    rsDetalle.MoveNext
                Loop
            End If
            rst.MoveNext
        Loop
        mostrarDatosGridSQL gLista2, rsdetallePed, StrMsgError
    End If

    If rsdetallePed.RecordCount > 0 Then
        item = 0
        rsdetalleCot.Fields.Append "Item", adInteger, , adFldRowID
        rsdetalleCot.Fields.Append "idempresa", adChar, 8, adFldRowID
        rsdetalleCot.Fields.Append "idsucursal", adChar, 8, adFldRowID
        rsdetalleCot.Fields.Append "iddocumento", adChar, 2, adFldRowID
        rsdetalleCot.Fields.Append "glsdocreferencia", adChar, 500, adFldRowID
        rsdetalleCot.Fields.Append "abredocumento", adChar, 5, adFldRowID
        rsdetalleCot.Fields.Append "simbolo", adChar, 5, adFldRowID
        rsdetalleCot.Fields.Append "idserie", adVarChar, 4, adFldRowID
        rsdetalleCot.Fields.Append "iddocventas", adChar, 8, adFldRowID
        rsdetalleCot.Fields.Append "fecEmision", adVarChar, 15, adFldRowID
        rsdetalleCot.Fields.Append "estdocventas", adChar, 3, adFldRowID
        rsdetalleCot.Fields.Append "idpercliente", adVarChar, 8, adFldRowID
        rsdetalleCot.Fields.Append "glscliente", adVarChar, 100, adFldRowID
        rsdetalleCot.Fields.Append "idmoneda", adVarChar, 3, adFldRowID
        rsdetalleCot.Fields.Append "totalprecioventa", adDouble, adFldRowID
        rsdetalleCot.Open
        
        rsdetallePed.MoveFirst
        Do While Not rsdetallePed.EOF
            csql = " select idempresa,tipodocreferencia,seriedocreferencia,numdocreferencia from docreferencia" & _
                    " where idempresa = '" & Trim("" & rsdetallePed.Fields("idempresa")) & "' and tipodocorigen = '" & Trim("" & rsdetallePed.Fields("iddocumento")) & "' and seriedocorigen = '" & Trim("" & rsdetallePed.Fields("idserie")) & "' and numdocorigen = '" & Trim("" & rsdetallePed.Fields("iddocventas")) & "' "
            If rst.State = 1 Then rst.Close
            rst.Open csql, Cn, adOpenKeyset, adLockOptimistic
             
            If rst.RecordCount > 0 Then
                rst.MoveFirst
                Do While Not rst.EOF
                    If Len(Trim("" & rst.Fields("numdocreferencia"))) = 8 Then
                        csql = " Select d.idempresa,d.idsucursal,d.iddocumento,d.glsdocreferencia,dc.abredocumento,m.simbolo,d.idserie,d.iddocventas,d.fecemision,d.estdocventas,d.idpercliente," & _
                                " d.glscliente,d.idmoneda,d.totalprecioventa from docventas d inner join documentos dc on d.iddocumento = dc.iddocumento " & _
                                " inner join monedas m on d.idmoneda = m.idmoneda " & _
                                " where d.idempresa = '" & Trim("" & rst.Fields("idempresa")) & "' and d.iddocumento = '" & Trim("" & rst.Fields("tipodocreferencia")) & "' " & _
                                " and d.idserie = '" & Trim("" & rst.Fields("seriedocreferencia")) & "' and d.iddocventas = '" & Trim("" & rst.Fields("numdocreferencia")) & "'"
                        If rsDetalle.State = 1 Then rsDetalle.Close
                        rsDetalle.Open csql, Cn, adOpenKeyset, adLockOptimistic
        
                        rsDetalle.MoveFirst
                        Do While Not rsDetalle.EOF
                            item = item + 1
                            rsdetalleCot.AddNew
                            rsdetalleCot!item = item
                            rsdetalleCot!idempresa = "" & rsDetalle.Fields("idempresa")
                            rsdetalleCot!idsucursal = "" & rsDetalle.Fields("idsucursal")
                            rsdetalleCot!idDocumento = "" & rsDetalle.Fields("iddocumento")
                            rsdetalleCot!glsdocreferencia = "" & rsDetalle.Fields("glsdocreferencia")
                            rsdetalleCot!abredocumento = "" & rsDetalle.Fields("abredocumento")
                            rsdetalleCot!simbolo = "" & rsDetalle.Fields("simbolo")
                            rsdetalleCot!idserie = "" & rsDetalle.Fields("idserie")
                            rsdetalleCot!idDocVentas = "" & rsDetalle.Fields("iddocventas")
                            rsdetalleCot!fecEmision = "" & rsDetalle.Fields("fecemision")
                            rsdetalleCot!estdocventas = "" & rsDetalle.Fields("estdocventas")
                            rsdetalleCot!idpercliente = "" & rsDetalle.Fields("idpercliente")
                            rsdetalleCot!GlsCliente = "" & rsDetalle.Fields("glscliente")
                            rsdetalleCot!idMoneda = "" & rsDetalle.Fields("idmoneda")
                            rsdetalleCot!totalprecioventa = "" & rsDetalle.Fields("totalprecioventa")
                            rsdetalleCot.Update
                       
                            rsDetalle.MoveNext
                        Loop
                    End If
                    rst.MoveNext
                Loop
            End If
            rsdetallePed.MoveNext
        Loop
        mostrarDatosGridSQL gLista1, rsdetalleCot, StrMsgError
    End If
         
    csql = " select idempresa,tipodocorigen,seriedocorigen,numdocorigen from docreferencia" & _
            " where idempresa = '" & glsEmpresa & "' and tipodocreferencia= '" & Trim(txtCod_TipoDoc.Text) & "' and seriedocreferencia = '" & Trim(txt_serie.Text) & "' and numdocreferencia = '" & Trim(txtNum_Documento.Text) & "' "
    If rst.State = 1 Then rst.Close
    rst.Open csql, Cn, adOpenKeyset, adLockOptimistic
     
    If rst.RecordCount > 0 Then
        item = 0
        rstdetCot.Fields.Append "Item", adInteger, , adFldRowID
        rstdetCot.Fields.Append "idempresa", adChar, 8, adFldRowID
        rstdetCot.Fields.Append "idsucursal", adChar, 8, adFldRowID
        rstdetCot.Fields.Append "iddocumento", adChar, 2, adFldRowID
        rstdetCot.Fields.Append "glsdocreferencia", adChar, 500, adFldRowID
        rstdetCot.Fields.Append "abredocumento", adChar, 5, adFldRowID
        rstdetCot.Fields.Append "simbolo", adChar, 5, adFldRowID
        rstdetCot.Fields.Append "idserie", adVarChar, 4, adFldRowID
        rstdetCot.Fields.Append "iddocventas", adChar, 8, adFldRowID
        rstdetCot.Fields.Append "fecEmision", adVarChar, 15, adFldRowID
        rstdetCot.Fields.Append "estdocventas", adChar, 3, adFldRowID
        rstdetCot.Fields.Append "idpercliente", adVarChar, 8, adFldRowID
        rstdetCot.Fields.Append "glscliente", adVarChar, 100, adFldRowID
        rstdetCot.Fields.Append "idmoneda", adVarChar, 3, adFldRowID
        rstdetCot.Fields.Append "totalprecioventa", adDouble, adFldRowID
        rstdetCot.Open
                
        rst.MoveFirst
        Do While Not rst.EOF
            If Len(Trim("" & rst.Fields("numdocorigen"))) = 8 Then
                csql = " Select d.idempresa,d.idsucursal,d.iddocumento,d.glsdocreferencia,dc.abredocumento,m.simbolo,d.idserie,d.iddocventas,d.fecemision,d.estdocventas,d.idpercliente," & _
                       " d.glscliente,d.idmoneda,d.totalprecioventa from docventas d inner join documentos dc on d.iddocumento = dc.iddocumento " & _
                       " inner join monedas m on d.idmoneda = m.idmoneda " & _
                       " where d.idempresa = '" & Trim("" & rst.Fields("idempresa")) & "' and d.iddocumento = '" & Trim("" & rst.Fields("tipodocorigen")) & "' " & _
                       " and d.idserie = '" & Trim("" & rst.Fields("seriedocorigen")) & "' and d.iddocventas = '" & Trim("" & rst.Fields("numdocorigen")) & "'"
                If rsDetalle.State = 1 Then rsDetalle.Close
                rsDetalle.Open csql, Cn, adOpenKeyset, adLockOptimistic
                      
                rsDetalle.MoveFirst
                Do While Not rsDetalle.EOF
                    item = item + 1
                    rstdetCot.AddNew
                    rstdetCot!item = item
                    rstdetCot!idempresa = "" & rsDetalle.Fields("idempresa")
                    rstdetCot!idsucursal = "" & rsDetalle.Fields("idsucursal")
                    rstdetCot!idDocumento = "" & rsDetalle.Fields("iddocumento")
                    rstdetCot!glsdocreferencia = "" & rsDetalle.Fields("glsdocreferencia")
                    rstdetCot!abredocumento = "" & rsDetalle.Fields("abredocumento")
                    rstdetCot!simbolo = "" & rsDetalle.Fields("simbolo")
                    rstdetCot!idserie = "" & rsDetalle.Fields("idserie")
                    rstdetCot!idDocVentas = "" & rsDetalle.Fields("iddocventas")
                    rstdetCot!fecEmision = "" & rsDetalle.Fields("fecemision")
                    rstdetCot!estdocventas = "" & rsDetalle.Fields("estdocventas")
                    rstdetCot!idpercliente = "" & rsDetalle.Fields("idpercliente")
                    rstdetCot!GlsCliente = "" & rsDetalle.Fields("glscliente")
                    rstdetCot!idMoneda = "" & rsDetalle.Fields("idmoneda")
                    rstdetCot!totalprecioventa = "" & rsDetalle.Fields("totalprecioventa")
                    rstdetCot.Update
                    
                    rsDetalle.MoveNext
                Loop
            End If
            rst.MoveNext
        Loop
        mostrarDatosGridSQL gLista4, rstdetCot, StrMsgError
    End If
         
End Sub

Private Sub SecuenciaBoleta2()
Dim csql                 As String
Dim rst                  As New ADODB.Recordset
Dim rsg                  As New ADODB.Recordset
Dim rsDetalle            As New ADODB.Recordset
Dim rsdetalleCot         As New ADODB.Recordset
Dim rsdetallePed         As New ADODB.Recordset
Dim rsdetalleFacGuia     As New ADODB.Recordset
Dim rsdetFacGuia         As New ADODB.Recordset
Dim rstdetCot            As New ADODB.Recordset
Dim StrMsgError          As String
Dim item                 As Integer

    csql = " Select (@i:=@i +1) as item,d.idempresa,d.idsucursal,d.iddocumento,d.glsdocreferencia,dc.abredocumento,m.simbolo,d.idserie,d.iddocventas,d.fecemision,d.estdocventas,d.idpercliente," & _
               " d.glscliente,d.idmoneda,d.totalprecioventa from docventas d inner join documentos dc on d.iddocumento = dc.iddocumento " & _
               " inner join monedas m on d.idmoneda = m.idmoneda ,(select @i:=0) foo " & _
               " where d.idempresa = '" & glsEmpresa & "' and d.iddocumento = '" & txtCod_TipoDoc.Text & "' " & _
               " and d.idserie = '" & txt_serie.Text & "' and d.iddocventas = '" & txtNum_Documento.Text & "'"
               
    If rst.State = 1 Then rst.Close
    rst.Open csql, Cn, adOpenKeyset, adLockOptimistic
     
    rsg.Fields.Append "item", adInteger, , adFldRowID
    rsg.Fields.Append "idempresa", adChar, 8, adFldIsNullable
    rsg.Fields.Append "idsucursal", adChar, 8, adFldIsNullable
    rsg.Fields.Append "iddocumento", adChar, 2, adFldIsNullable
    rsg.Fields.Append "glsdocreferencia", adChar, 500, adFldIsNullable
    rsg.Fields.Append "abredocumento", adChar, 5, adFldIsNullable
    rsg.Fields.Append "simbolo", adChar, 5, adFldIsNullable
    rsg.Fields.Append "idserie", adChar, 4, adFldIsNullable
    rsg.Fields.Append "iddocventas", adChar, 8, adFldIsNullable
    rsg.Fields.Append "fecemision", adChar, 15, adFldIsNullable
    rsg.Fields.Append "estdocventas", adChar, 3, adFldIsNullable
    rsg.Fields.Append "idpercliente", adChar, 8, adFldIsNullable
    rsg.Fields.Append "glscliente", adVarChar, 100, adFldIsNullable
    rsg.Fields.Append "idmoneda", adChar, 5, adFldIsNullable
    rsg.Fields.Append "totalprecioventa", adDouble, adFldRowID
    rsg.Open
    
    If rst.RecordCount = 0 Then
        rsg.Fields("item") = 1
        rsg.Fields("idempresa") = ""
        rsg.Fields("idsucursal") = ""
        rsg.Fields("iddocumento") = ""
        rsg.Fields("glsdocreferencia") = ""
        rsg.Fields("abredocumento") = ""
        rsg.Fields("simbolo") = ""
        rsg.Fields("idserie") = ""
        rsg.Fields("iddocventas") = ""
        rsg.Fields("fecemision") = ""
        rsg.Fields("estdocventas") = ""
        rsg.Fields("idpercliente") = ""
        rsg.Fields("glscliente") = ""
        rsg.Fields("idmoneda") = ""
        rsg.Fields("totalprecioventa") = ""
                
    Else
        If Not rst.EOF Then
            rst.MoveFirst
            Do While Not rst.EOF
                rsg.AddNew
                rsg.Fields("item") = Val("" & rst.Fields("item"))
                rsg.Fields("idempresa") = Trim("" & rst.Fields("idempresa"))
                rsg.Fields("idsucursal") = Trim("" & rst.Fields("idsucursal"))
                rsg.Fields("iddocumento") = Trim("" & rst.Fields("iddocumento"))
                rsg.Fields("glsdocreferencia") = Trim("" & rst.Fields("glsdocreferencia"))
                rsg.Fields("abredocumento") = Trim("" & rst.Fields("abredocumento"))
                rsg.Fields("simbolo") = Trim("" & rst.Fields("simbolo"))
                rsg.Fields("idserie") = Trim("" & rst.Fields("idserie"))
                rsg.Fields("iddocventas") = Trim("" & rst.Fields("iddocventas"))
                rsg.Fields("fecemision") = "" & rst.Fields("fecemision")
                rsg.Fields("estdocventas") = "" & rst.Fields("estdocventas")
                rsg.Fields("idpercliente") = Trim("" & rst.Fields("idpercliente"))
                rsg.Fields("glscliente") = Trim("" & rst.Fields("glscliente"))
                rsg.Fields("idmoneda") = Trim("" & rst.Fields("idmoneda"))
                rsg.Fields("totalprecioventa") = Trim("" & rst.Fields("totalprecioventa"))
                rst.MoveNext
            Loop
        End If
        mostrarDatosGridSQL gLista4, rsg, StrMsgError
    End If

    csql = " select idempresa,tipodocreferencia,seriedocreferencia,numdocreferencia from docreferencia" & _
           " where idempresa = '" & glsEmpresa & "' and tipodocorigen = '" & txtCod_TipoDoc.Text & "'" & _
           " and seriedocorigen = '" & txt_serie.Text & "' " & _
           " and numdocorigen = '" & txtNum_Documento.Text & "' "
    If rst.State = 1 Then rst.Close
    rst.Open csql, Cn, adOpenKeyset, adLockOptimistic
     
    If rst.RecordCount > 0 Then
        item = 0
        rsdetallePed.Fields.Append "Item", adInteger, , adFldRowID
        rsdetallePed.Fields.Append "idempresa", adChar, 8, adFldRowID
        rsdetallePed.Fields.Append "idsucursal", adChar, 8, adFldRowID
        rsdetallePed.Fields.Append "iddocumento", adChar, 2, adFldRowID
        rsdetallePed.Fields.Append "glsdocreferencia", adChar, 500, adFldRowID
        rsdetallePed.Fields.Append "abredocumento", adChar, 5, adFldRowID
        rsdetallePed.Fields.Append "simbolo", adChar, 5, adFldRowID
        rsdetallePed.Fields.Append "idserie", adVarChar, 4, adFldRowID
        rsdetallePed.Fields.Append "iddocventas", adChar, 8, adFldRowID
        rsdetallePed.Fields.Append "fecEmision", adVarChar, 15, adFldRowID
        rsdetallePed.Fields.Append "estdocventas", adChar, 3, adFldRowID
        rsdetallePed.Fields.Append "idpercliente", adVarChar, 8, adFldRowID
        rsdetallePed.Fields.Append "glscliente", adVarChar, 100, adFldRowID
        rsdetallePed.Fields.Append "idmoneda", adVarChar, 3, adFldRowID
        rsdetallePed.Fields.Append "totalprecioventa", adDouble, adFldRowID
        rsdetallePed.Open
        
        rst.MoveFirst
        Do While Not rst.EOF
            If Len(Trim("" & rst.Fields("numdocreferencia"))) = 8 Then
                csql = " Select d.idempresa,d.idsucursal,d.iddocumento,d.glsdocreferencia,dc.abredocumento,m.simbolo,d.idserie,d.iddocventas,d.fecemision,d.estdocventas,d.idpercliente," & _
                        " d.glscliente,d.idmoneda,d.totalprecioventa from docventas d inner join documentos dc on d.iddocumento = dc.iddocumento " & _
                        " inner join monedas m on d.idmoneda = m.idmoneda " & _
                        " where d.idempresa = '" & Trim("" & rst.Fields("idempresa")) & "' and d.iddocumento = '" & Trim("" & rst.Fields("tipodocreferencia")) & "' " & _
                        " and d.idserie = '" & Trim("" & rst.Fields("seriedocreferencia")) & "' and d.iddocventas = '" & Trim("" & rst.Fields("numdocreferencia")) & "'"
                If rsDetalle.State = 1 Then rsDetalle.Close
                rsDetalle.Open csql, Cn, adOpenKeyset, adLockOptimistic
                
                rsDetalle.MoveFirst
                Do While Not rsDetalle.EOF
                    item = item + 1
                    rsdetallePed.AddNew
                    rsdetallePed!item = item
                    rsdetallePed!idempresa = "" & rsDetalle.Fields("idempresa")
                    rsdetallePed!idsucursal = "" & rsDetalle.Fields("idsucursal")
                    rsdetallePed!idDocumento = "" & rsDetalle.Fields("iddocumento")
                    rsdetallePed!glsdocreferencia = "" & rsDetalle.Fields("glsdocreferencia")
                    rsdetallePed!abredocumento = "" & rsDetalle.Fields("abredocumento")
                    rsdetallePed!simbolo = "" & rsDetalle.Fields("simbolo")
                    rsdetallePed!idserie = "" & rsDetalle.Fields("idserie")
                    rsdetallePed!idDocVentas = "" & rsDetalle.Fields("iddocventas")
                    rsdetallePed!fecEmision = "" & rsDetalle.Fields("fecemision")
                    rsdetallePed!estdocventas = "" & rsDetalle.Fields("estdocventas")
                    rsdetallePed!idpercliente = "" & rsDetalle.Fields("idpercliente")
                    rsdetallePed!GlsCliente = "" & rsDetalle.Fields("glscliente")
                    rsdetallePed!idMoneda = "" & rsDetalle.Fields("idmoneda")
                    rsdetallePed!totalprecioventa = "" & rsDetalle.Fields("totalprecioventa")
                    rsdetallePed.Update
                    
                    rsDetalle.MoveNext
                Loop
            End If
            rst.MoveNext
        Loop
        mostrarDatosGridSQL gLista3, rsdetallePed, StrMsgError
    End If

    If rsdetallePed.RecordCount > 0 Then
        item = 0
        rsdetalleCot.Fields.Append "Item", adInteger, , adFldRowID
        rsdetalleCot.Fields.Append "idempresa", adChar, 8, adFldRowID
        rsdetalleCot.Fields.Append "idsucursal", adChar, 8, adFldRowID
        rsdetalleCot.Fields.Append "iddocumento", adChar, 2, adFldRowID
        rsdetalleCot.Fields.Append "glsdocreferencia", adChar, 500, adFldRowID
        rsdetalleCot.Fields.Append "abredocumento", adChar, 5, adFldRowID
        rsdetalleCot.Fields.Append "simbolo", adChar, 5, adFldRowID
        rsdetalleCot.Fields.Append "idserie", adVarChar, 4, adFldRowID
        rsdetalleCot.Fields.Append "iddocventas", adChar, 8, adFldRowID
        rsdetalleCot.Fields.Append "fecEmision", adVarChar, 15, adFldRowID
        rsdetalleCot.Fields.Append "estdocventas", adChar, 3, adFldRowID
        rsdetalleCot.Fields.Append "idpercliente", adVarChar, 8, adFldRowID
        rsdetalleCot.Fields.Append "glscliente", adVarChar, 100, adFldRowID
        rsdetalleCot.Fields.Append "idmoneda", adVarChar, 3, adFldRowID
        rsdetalleCot.Fields.Append "totalprecioventa", adDouble, adFldRowID
        rsdetalleCot.Open
        
        rsdetallePed.MoveFirst
        Do While Not rsdetallePed.EOF
            csql = " select idempresa,tipodocreferencia,seriedocreferencia,numdocreferencia from docreferencia" & _
                   " where idempresa = '" & Trim("" & rsdetallePed.Fields("idempresa")) & "' and tipodocorigen = '" & Trim("" & rsdetallePed.Fields("iddocumento")) & "' and seriedocorigen = '" & Trim("" & rsdetallePed.Fields("idserie")) & "' and numdocorigen = '" & Trim("" & rsdetallePed.Fields("iddocventas")) & "' "
            If rst.State = 1 Then rst.Close
            rst.Open csql, Cn, adOpenKeyset, adLockOptimistic
             
            If rst.RecordCount > 0 Then
                rst.MoveFirst
                Do While Not rst.EOF
                    If Len(Trim("" & rst.Fields("numdocreferencia"))) = 8 Then
                        csql = " Select d.idempresa,d.idsucursal,d.iddocumento,d.glsdocreferencia,dc.abredocumento,m.simbolo,d.idserie,d.iddocventas,d.fecemision,d.estdocventas,d.idpercliente," & _
                                " d.glscliente,d.idmoneda,d.totalprecioventa from docventas d inner join documentos dc on d.iddocumento = dc.iddocumento" & _
                                " inner join monedas m on d.idmoneda = m.idmoneda " & _
                                " where d.idempresa = '" & Trim("" & rst.Fields("idempresa")) & "' and d.iddocumento = '" & Trim("" & rst.Fields("tipodocreferencia")) & "' " & _
                                " and d.idserie = '" & Trim("" & rst.Fields("seriedocreferencia")) & "' and d.iddocventas = '" & Trim("" & rst.Fields("numdocreferencia")) & "'"
                        If rsDetalle.State = 1 Then rsDetalle.Close
                        rsDetalle.Open csql, Cn, adOpenKeyset, adLockOptimistic
                            
                        rsDetalle.MoveFirst
                        Do While Not rsDetalle.EOF
                            item = item + 1
                            rsdetalleCot.AddNew
                            rsdetalleCot!item = item
                            rsdetalleCot!idempresa = "" & rsDetalle.Fields("idempresa")
                            rsdetalleCot!idsucursal = "" & rsDetalle.Fields("idsucursal")
                            rsdetalleCot!idDocumento = "" & rsDetalle.Fields("iddocumento")
                            rsdetalleCot!glsdocreferencia = "" & rsDetalle.Fields("glsdocreferencia")
                            rsdetalleCot!abredocumento = "" & rsDetalle.Fields("abredocumento")
                            rsdetalleCot!simbolo = "" & rsDetalle.Fields("simbolo")
                            rsdetalleCot!idserie = "" & rsDetalle.Fields("idserie")
                            rsdetalleCot!idDocVentas = "" & rsDetalle.Fields("iddocventas")
                            rsdetalleCot!fecEmision = "" & rsDetalle.Fields("fecemision")
                            rsdetalleCot!estdocventas = "" & rsDetalle.Fields("estdocventas")
                            rsdetalleCot!idpercliente = "" & rsDetalle.Fields("idpercliente")
                            rsdetalleCot!GlsCliente = "" & rsDetalle.Fields("glscliente")
                            rsdetalleCot!idMoneda = "" & rsDetalle.Fields("idmoneda")
                            rsdetalleCot!totalprecioventa = "" & rsDetalle.Fields("totalprecioventa")
                            rsdetalleCot.Update
                       
                            rsDetalle.MoveNext
                        Loop
                    End If
                    rst.MoveNext
                Loop
            End If
            rsdetallePed.MoveNext
        Loop
        mostrarDatosGridSQL gLista2, rsdetalleCot, StrMsgError
    End If
 

    If rsdetalleCot.RecordCount > 0 Then
        item = 0
        rstdetCot.Fields.Append "Item", adInteger, , adFldRowID
        rstdetCot.Fields.Append "idempresa", adChar, 8, adFldRowID
        rstdetCot.Fields.Append "idsucursal", adChar, 8, adFldRowID
        rstdetCot.Fields.Append "iddocumento", adChar, 2, adFldRowID
        rstdetCot.Fields.Append "glsdocreferencia", adChar, 500, adFldRowID
        rstdetCot.Fields.Append "abredocumento", adChar, 5, adFldRowID
        rstdetCot.Fields.Append "simbolo", adChar, 5, adFldRowID
        rstdetCot.Fields.Append "idserie", adVarChar, 4, adFldRowID
        rstdetCot.Fields.Append "iddocventas", adChar, 8, adFldRowID
        rstdetCot.Fields.Append "fecEmision", adVarChar, 15, adFldRowID
        rstdetCot.Fields.Append "estdocventas", adChar, 3, adFldRowID
        rstdetCot.Fields.Append "idpercliente", adVarChar, 8, adFldRowID
        rstdetCot.Fields.Append "glscliente", adVarChar, 100, adFldRowID
        rstdetCot.Fields.Append "idmoneda", adVarChar, 3, adFldRowID
        rstdetCot.Fields.Append "totalprecioventa", adDouble, adFldRowID
        rstdetCot.Open
        
        rsdetalleCot.MoveFirst
        Do While Not rsdetalleCot.EOF
            csql = " select idempresa,tipodocreferencia,seriedocreferencia,numdocreferencia from docreferencia" & _
                " where idempresa = '" & Trim("" & rsdetalleCot.Fields("idempresa")) & "' and tipodocorigen = '" & Trim("" & rsdetalleCot.Fields("iddocumento")) & "' and seriedocorigen = '" & Trim("" & rsdetalleCot.Fields("idserie")) & "' and numdocorigen = '" & Trim("" & rsdetalleCot.Fields("iddocventas")) & "' "
            If rst.State = 1 Then rst.Close
            rst.Open csql, Cn, adOpenKeyset, adLockOptimistic
        
            If rst.RecordCount > 0 Then
                rst.MoveFirst
                Do While Not rst.EOF
                    If Len(Trim("" & rst.Fields("numdocreferencia"))) = 8 Then
                        csql = " Select d.idempresa,d.idsucursal,d.iddocumento,d.glsdocreferencia,dc.abredocumento,m.simbolo,d.idserie,d.iddocventas,d.fecemision,d.estdocventas,d.idpercliente," & _
                                " d.glscliente,d.idmoneda,d.totalprecioventa from docventas d inner join documentos dc on d.iddocumento = dc.iddocumento " & _
                                " inner join monedas m on d.idmoneda = m.idmoneda " & _
                                " where d.idempresa = '" & Trim("" & rst.Fields("idempresa")) & "' and d.iddocumento = '" & Trim("" & rst.Fields("tipodocreferencia")) & "' " & _
                                " and d.idserie = '" & Trim("" & rst.Fields("seriedocreferencia")) & "' and d.iddocventas = '" & Trim("" & rst.Fields("numdocreferencia")) & "'"
                        If rsDetalle.State = 1 Then rsDetalle.Close
                        rsDetalle.Open csql, Cn, adOpenKeyset, adLockOptimistic
                     
                        rsDetalle.MoveFirst
                        Do While Not rsDetalle.EOF
                            item = item + 1
                            rstdetCot.AddNew
                            rstdetCot!item = item
                            rstdetCot!idempresa = "" & rsDetalle.Fields("idempresa")
                            rstdetCot!idsucursal = "" & rsDetalle.Fields("idsucursal")
                            rstdetCot!idDocumento = "" & rsDetalle.Fields("iddocumento")
                            rstdetCot!glsdocreferencia = "" & rsDetalle.Fields("glsdocreferencia")
                            rstdetCot!abredocumento = "" & rsDetalle.Fields("abredocumento")
                            rstdetCot!simbolo = "" & rsDetalle.Fields("simbolo")
                            rstdetCot!idserie = "" & rsDetalle.Fields("idserie")
                            rstdetCot!idDocVentas = "" & rsDetalle.Fields("iddocventas")
                            rstdetCot!fecEmision = "" & rsDetalle.Fields("fecemision")
                            rstdetCot!estdocventas = "" & rsDetalle.Fields("estdocventas")
                            rstdetCot!idpercliente = "" & rsDetalle.Fields("idpercliente")
                            rstdetCot!GlsCliente = "" & rsDetalle.Fields("glscliente")
                            rstdetCot!idMoneda = "" & rsDetalle.Fields("idmoneda")
                            rstdetCot!totalprecioventa = "" & rsDetalle.Fields("totalprecioventa")
                            rstdetCot.Update
                            
                            rsDetalle.MoveNext
                        Loop
                    End If
                    rst.MoveNext
                Loop
            End If
            rsdetalleCot.MoveNext
        Loop
        mostrarDatosGridSQL gLista1, rstdetCot, StrMsgError
    End If
     
End Sub

Private Sub cmbAyudaTipoDoc_Click()
    
    mostrarAyuda "DOCUMENTOS", txtCod_TipoDoc, txtGls_TipoDoc, "and iddocumento in ('01','03','40','86','92','90')"

End Sub

Private Sub Form_Load()

    Me.top = 0
    Me.left = 0
    ConfGrid1 gLista1, False, False, False, False
    ConfGrid1 gLista2, False, False, False, False
    ConfGrid1 gLista3, False, False, False, False
    ConfGrid1 gLista4, False, False, False, False

End Sub
 
Private Sub gdetalle3_OnKeyPress(Key As Integer)
    
    Select Case Key
        Case 27:
            FraLista3.Visible = False
            gLista3.SetFocus
    End Select

End Sub

Private Sub gdetalle4_OnKeyPress(Key As Integer)
    
    Select Case Key
        Case 27:
            FraLista4.Visible = False
            gLista4.SetFocus
    End Select

End Sub

Private Sub gdetalle2_OnKeyPress(Key As Integer)
    
    Select Case Key
        Case 27:
            FraLista2.Visible = False
            gLista2.SetFocus
    End Select

End Sub

Private Sub gdetalle1_OnKeyPress(Key As Integer)
    
    Select Case Key
        Case 27:
            FraLista1.Visible = False
            gLista1.SetFocus
    End Select

End Sub

Private Sub gLista1_OnClick()
Dim StrMsgError     As String
Dim rst             As New ADODB.Recordset
Dim rsg             As New ADODB.Recordset
Dim csql            As String
    
    If gLista1.Count <> 0 Then
        ConfGrid1 gdetalle1, False, False, False, False
        FraLista1.Visible = True
        FraLista2.Visible = False
        FraLista3.Visible = False
        FraLista4.Visible = False
        csql = "Select (@i:=@i +1) as item, idproducto,glsproducto,glsum,cantidad,vvunit,dctopv,igvunit,pvunit from docventasdet ,(select @i:=0) foo " & _
               "where idempresa =  '" & gLista1.Columns.ColumnByName("idempresa").Value & "' and " & _
               "iddocumento = '" & gLista1.Columns.ColumnByName("iddocumento").Value & "' and " & _
               "idserie = '" & gLista1.Columns.ColumnByName("idserie").Value & "' and " & _
               "iddocventas = '" & gLista1.Columns.ColumnByName("iddocventas").Value & "' "
               
        rst.Open csql, Cn, adOpenStatic, adLockReadOnly
        If rst.State = 1 Then rst.Close
        rst.Open csql, Cn, adOpenKeyset, adLockOptimistic
        rsg.Fields.Append "item", adInteger, , adFldRowID
        rsg.Fields.Append "idproducto", adChar, 8, adFldIsNullable
        rsg.Fields.Append "glsproducto", adChar, 150, adFldIsNullable
        rsg.Fields.Append "cantidad", adDouble, adFldIsNullable
        rsg.Fields.Append "glsum", adChar, 10, adFldIsNullable
        rsg.Fields.Append "vvunit", adDouble, adFldIsNullable
        rsg.Fields.Append "igvunit", adDouble, adFldIsNullable
        rsg.Fields.Append "dctopv", adDouble, adFldIsNullable
        rsg.Fields.Append "pvunit", adDouble, adFldIsNullable
        rsg.Open
        
        If rst.RecordCount = 0 Then
            rsg.AddNew
            rsg.Fields("item") = 1
            rsg.Fields("idproducto") = ""
            rsg.Fields("glsproducto") = ""
            rsg.Fields("cantidad") = 0#
            rsg.Fields("glsum") = ""
            rsg.Fields("vvunit") = 0#
            rsg.Fields("igvunit") = 0#
            rsg.Fields("dctopv") = 0#
            rsg.Fields("pvunit") = 0#
                    
        Else
            If Not rst.EOF Then
                rst.MoveFirst
                Do While Not rst.EOF
                    rsg.AddNew
                    rsg.Fields("item") = Val("" & rst.Fields("item"))
                    rsg.Fields("idproducto") = Trim("" & rst.Fields("idproducto"))
                    rsg.Fields("glsproducto") = Trim("" & rst.Fields("glsproducto"))
                    rsg.Fields("cantidad") = Trim("" & rst.Fields("cantidad"))
                    rsg.Fields("glsum") = Trim("" & rst.Fields("glsum"))
                    rsg.Fields("vvunit") = Trim("" & rst.Fields("vvunit"))
                    rsg.Fields("igvunit") = Trim("" & rst.Fields("igvunit"))
                    rsg.Fields("dctopv") = Trim("" & rst.Fields("dctopv"))
                    rsg.Fields("pvunit") = Trim("" & rst.Fields("pvunit"))
                    rst.MoveNext
                Loop
            End If
        End If
        mostrarDatosGridSQL gdetalle1, rsg, StrMsgError
    End If

End Sub

Private Sub gLista2_OnClick()
Dim StrMsgError     As String
Dim rst             As New ADODB.Recordset
Dim rsg             As New ADODB.Recordset
Dim csql            As String
    
    If gLista2.Count <> 0 Then
        ConfGrid1 gdetalle2, False, False, False, False
        FraLista2.Visible = True
        FraLista1.Visible = False
        FraLista3.Visible = False
        FraLista4.Visible = False
        
        csql = "Select (@i:=@i +1) as item, idproducto,glsproducto,glsum,cantidad,vvunit,igvunit,dctopv,pvunit from docventasdet ,(select @i:=0) foo " & _
               "where idempresa =  '" & gLista2.Columns.ColumnByName("idempresa").Value & "' and " & _
               "iddocumento = '" & gLista2.Columns.ColumnByName("iddocumento").Value & "' and " & _
               "idserie = '" & gLista2.Columns.ColumnByName("idserie").Value & "' and " & _
               "iddocventas = '" & gLista2.Columns.ColumnByName("iddocventas").Value & "' "
             rst.Open csql, Cn, adOpenStatic, adLockReadOnly
             
        If rst.State = 1 Then rst.Close
        rst.Open csql, Cn, adOpenKeyset, adLockOptimistic
         
        rsg.Fields.Append "item", adInteger, , adFldRowID
        rsg.Fields.Append "idproducto", adChar, 8, adFldIsNullable
        rsg.Fields.Append "glsproducto", adChar, 150, adFldIsNullable
        rsg.Fields.Append "cantidad", adDouble, adFldIsNullable
        rsg.Fields.Append "glsum", adChar, 10, adFldIsNullable
        rsg.Fields.Append "vvunit", adDouble, adFldIsNullable
        rsg.Fields.Append "igvunit", adDouble, adFldIsNullable
        rsg.Fields.Append "dctopv", adDouble, adFldIsNullable
        rsg.Fields.Append "pvunit", adDouble, adFldIsNullable
        rsg.Open
        
        If rst.RecordCount = 0 Then
            rsg.Fields("item") = 1
            rsg.Fields("idproducto") = ""
            rsg.Fields("glsproducto") = ""
            rsg.Fields("cantidad") = ""
            rsg.Fields("glsum") = ""
            rsg.Fields("vvunit") = ""
            rsg.Fields("igvunit") = ""
            rsg.Fields("dctopv") = ""
            rsg.Fields("pvunit") = ""
                    
        Else
            If Not rst.EOF Then
                rst.MoveFirst
                Do While Not rst.EOF
                    rsg.AddNew
                    rsg.Fields("item") = Val("" & rst.Fields("item"))
                    rsg.Fields("idproducto") = Trim("" & rst.Fields("idproducto"))
                    rsg.Fields("glsproducto") = Trim("" & rst.Fields("glsproducto"))
                    rsg.Fields("cantidad") = Trim("" & rst.Fields("cantidad"))
                    rsg.Fields("glsum") = Trim("" & rst.Fields("glsum"))
                    rsg.Fields("vvunit") = Trim("" & rst.Fields("vvunit"))
                    rsg.Fields("igvunit") = Trim("" & rst.Fields("igvunit"))
                    rsg.Fields("dctopv") = Trim("" & rst.Fields("dctopv"))
                    rsg.Fields("pvunit") = Trim("" & rst.Fields("pvunit"))
                    rst.MoveNext
                Loop
            End If
        End If
        mostrarDatosGridSQL gdetalle2, rsg, StrMsgError
    End If

End Sub

Private Sub gLista3_OnClick()
Dim StrMsgError     As String
Dim rst             As New ADODB.Recordset
Dim rsg             As New ADODB.Recordset
Dim csql            As String
    
    If gLista3.Count <> 0 Then
        ConfGrid1 gdetalle3, False, False, False, False
        FraLista3.Visible = True
        FraLista2.Visible = False
        FraLista1.Visible = False
        FraLista4.Visible = False
       
        csql = "Select (@i:=@i +1) as item, idproducto,glsproducto,glsum,cantidad,vvunit,igvunit,dctopv,pvunit from docventasdet ,(select @i:=0) foo " & _
                "where idempresa =  '" & gLista3.Columns.ColumnByName("idempresa").Value & "' and " & _
                "iddocumento = '" & gLista3.Columns.ColumnByName("iddocumento").Value & "' and " & _
                "idserie = '" & gLista3.Columns.ColumnByName("idserie").Value & "' and " & _
                "iddocventas = '" & gLista3.Columns.ColumnByName("iddocventas").Value & "' "
        rst.Open csql, Cn, adOpenStatic, adLockReadOnly
            
        If rst.State = 1 Then rst.Close
        rst.Open csql, Cn, adOpenKeyset, adLockOptimistic
        rsg.Fields.Append "item", adInteger, , adFldRowID
        rsg.Fields.Append "idproducto", adChar, 8, adFldIsNullable
        rsg.Fields.Append "glsproducto", adChar, 150, adFldIsNullable
        rsg.Fields.Append "cantidad", adDouble, adFldIsNullable
        rsg.Fields.Append "glsum", adChar, 10, adFldIsNullable
        rsg.Fields.Append "vvunit", adDouble, adFldIsNullable
        rsg.Fields.Append "igvunit", adDouble, adFldIsNullable
        rsg.Fields.Append "dctopv", adDouble, adFldIsNullable
        rsg.Fields.Append "pvunit", adDouble, adFldIsNullable
        rsg.Open
       
        If rst.RecordCount = 0 Then
            rsg.Fields("item") = 1
            rsg.Fields("idproducto") = ""
            rsg.Fields("glsproducto") = ""
            rsg.Fields("cantidad") = ""
            rsg.Fields("glsum") = ""
            rsg.Fields("vvunit") = ""
            rsg.Fields("igvunit") = ""
            rsg.Fields("dctopv") = ""
            rsg.Fields("pvunit") = ""
                   
        Else
            If Not rst.EOF Then
                rst.MoveFirst
                Do While Not rst.EOF
                    rsg.AddNew
                    rsg.Fields("item") = Val("" & rst.Fields("item"))
                    rsg.Fields("idproducto") = Trim("" & rst.Fields("idproducto"))
                    rsg.Fields("glsproducto") = Trim("" & rst.Fields("glsproducto"))
                    rsg.Fields("cantidad") = Trim("" & rst.Fields("cantidad"))
                    rsg.Fields("glsum") = Trim("" & rst.Fields("glsum"))
                    rsg.Fields("vvunit") = Trim("" & rst.Fields("vvunit"))
                    rsg.Fields("igvunit") = Trim("" & rst.Fields("igvunit"))
                    rsg.Fields("dctopv") = Trim("" & rst.Fields("dctopv"))
                    rsg.Fields("pvunit") = Trim("" & rst.Fields("pvunit"))
                    rst.MoveNext
                Loop
            End If
        End If
        mostrarDatosGridSQL gdetalle3, rsg, StrMsgError
    End If

End Sub

Private Sub gLista4_OnClick()
Dim StrMsgError     As String
Dim rst             As New ADODB.Recordset
Dim rsg             As New ADODB.Recordset
Dim csql            As String
    
    If gLista4.Count <> 0 Then
        ConfGrid1 gDetalle4, False, False, False, False
        If Trim("" & gLista4.Columns.ColumnByFieldName("iddocumento").Value = "94") Then
            FraLista4.Visible = False
        Else
            FraLista4.Visible = True
        End If
        FraLista3.Visible = False
        FraLista2.Visible = False
        FraLista1.Visible = False
    
        csql = "Select (@i:=@i +1) as item, idproducto,glsproducto,glsum,cantidad,vvunit,igvunit,dctopv,pvunit from docventasdet ,(select @i:=0) foo " & _
               "where idempresa =  '" & gLista4.Columns.ColumnByName("idempresa").Value & "' and " & _
               "iddocumento = '" & gLista4.Columns.ColumnByName("iddocumento").Value & "' and " & _
               "idserie = '" & gLista4.Columns.ColumnByName("idserie").Value & "' and " & _
               "iddocventas = '" & gLista4.Columns.ColumnByName("iddocventas").Value & "' "
        rst.Open csql, Cn, adOpenStatic, adLockReadOnly
        
        If rst.State = 1 Then rst.Close
        rst.Open csql, Cn, adOpenKeyset, adLockOptimistic
        rsg.Fields.Append "item", adInteger, , adFldRowID
        rsg.Fields.Append "idproducto", adChar, 8, adFldIsNullable
        rsg.Fields.Append "glsproducto", adChar, 150, adFldIsNullable
        rsg.Fields.Append "cantidad", adDouble, adFldIsNullable
        rsg.Fields.Append "glsum", adChar, 10, adFldIsNullable
        rsg.Fields.Append "vvunit", adDouble, adFldIsNullable
        rsg.Fields.Append "igvunit", adDouble, adFldIsNullable
        rsg.Fields.Append "dctopv", adDouble, adFldIsNullable
        rsg.Fields.Append "pvunit", adDouble, adFldIsNullable
        rsg.Open
        
        If rst.RecordCount = 0 Then
            rsg.AddNew
            rsg.Fields("item") = 1
            rsg.Fields("idproducto") = ""
            rsg.Fields("glsproducto") = ""
            rsg.Fields("cantidad") = 0#
            rsg.Fields("glsum") = ""
            rsg.Fields("vvunit") = 0#
            rsg.Fields("dctopv") = 0#
            rsg.Fields("igvunit") = 0#
            rsg.Fields("pvunit") = 0#
                    
        Else
            If Not rst.EOF Then
                rst.MoveFirst
                Do While Not rst.EOF
                    rsg.AddNew
                    rsg.Fields("item") = Val("" & rst.Fields("item"))
                    rsg.Fields("idproducto") = Trim("" & rst.Fields("idproducto"))
                    rsg.Fields("glsproducto") = Trim("" & rst.Fields("glsproducto"))
                    rsg.Fields("cantidad") = Trim("" & rst.Fields("cantidad"))
                    rsg.Fields("glsum") = Trim("" & rst.Fields("glsum"))
                    rsg.Fields("vvunit") = Trim("" & rst.Fields("vvunit"))
                    rsg.Fields("igvunit") = Trim("" & rst.Fields("igvunit"))
                    rsg.Fields("dctopv") = Trim("" & rst.Fields("dctopv"))
                    rsg.Fields("pvunit") = Trim("" & rst.Fields("pvunit"))
                    rst.MoveNext
                Loop
            End If
        End If
        mostrarDatosGridSQL gDetalle4, rsg, StrMsgError
    End If

End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
On Error GoTo Err
Dim StrMsgError As String
    
    Select Case Button.Index
        Case 1 'Imprime Costo
            If StrMsgError <> "" Then GoTo Err
            MostrarSecuencia
            If StrMsgError <> "" Then GoTo Err
        Case 2 'Salir
            gLista1.m.ExportToXLS App.Path & "\Temporales\Listado.xls"
            ShellEx App.Path & "\Temporales\Listado.xls", essSW_MAXIMIZE, , , "open", Me.hwnd
        Case 3 'Salir
            Unload Me
    End Select
    
    Exit Sub
    
Err:
    If StrMsgError = "" Then StrMsgError = Err.Description
    MsgBox StrMsgError, vbInformation, App.Title
End Sub

Private Sub MostrarSecuencia()
Dim csql                As String
Dim rst                 As New ADODB.Recordset
Dim StrMsgError         As String
Dim valor               As String

    gLista1.Dataset.Close
    gLista2.Dataset.Close
    gLista3.Dataset.Close
    gLista4.Dataset.Close

    txt_serie.Text = Format(Trim(txt_serie.Text), "000")
    txtNum_Documento.Text = Format(Trim(txtNum_Documento.Text), "00000000")
    valor = traerCampo("docventas", "iddocventas", "idserie", txt_serie.Text, True, "iddocventas = '" & txtNum_Documento.Text & "' and iddocumento = '" & txtCod_TipoDoc.Text & "'")
    
    If valor <> "" Then
        If txtCod_TipoDoc.Text = "92" Then
            SecuenciaCotizacion
        End If
             
        If txtCod_TipoDoc.Text = "40" Or txtCod_TipoDoc.Text = "90" Then
            SecuenciaPedido
        End If
             
        If txtCod_TipoDoc.Text = "03" Then
            csql = " Select tipodocreferencia,seriedocreferencia,numdocreferencia from docreferencia Where idempresa = '" & glsEmpresa & "' and tipodocorigen = '" & txtCod_TipoDoc.Text & "' " & _
                    " and seriedocorigen = '" & txt_serie.Text & "' and numdocorigen = '" & txtNum_Documento.Text & "'"
            If rst.State = 1 Then rst.Close
            rst.Open csql, Cn, adOpenKeyset, adLockOptimistic
                         
            If rst.RecordCount > 0 Then
                rst.MoveFirst
                Do While Not rst.EOF
                    If Trim("" & rst.Fields("tipodocreferencia")) = "40" Then
                        SecuenciaBoleta1
                        Exit Do
                    ElseIf Trim("" & rst.Fields("tipodocreferencia")) = "86" Then
                        SecuenciaBoleta2
                        Exit Do
                    End If
                    rst.MoveNext
                Loop
            End If
        End If
             
        If txtCod_TipoDoc.Text = "01" Then
            csql = " Select tipodocreferencia,seriedocreferencia,numdocreferencia from docreferencia Where idempresa = '" & glsEmpresa & "' and tipodocorigen = '" & txtCod_TipoDoc.Text & "' " & _
                    " and seriedocorigen = '" & txt_serie.Text & "' and numdocorigen = '" & txtNum_Documento.Text & "'"
            If rst.State = 1 Then rst.Close
            rst.Open csql, Cn, adOpenKeyset, adLockOptimistic
                         
            If rst.RecordCount > 0 Then
                rst.MoveFirst
                Do While Not rst.EOF
                    If Trim("" & rst.Fields("tipodocreferencia")) = "40" Then
                        SecuenciaFactura1
                        Exit Do
                        ElseIf Trim("" & rst.Fields("tipodocreferencia")) = "86" Then
                        SecuenciaFactura2
                        Exit Do
                    End If
                    rst.MoveNext
                Loop
            End If
        End If
            
        If txtCod_TipoDoc.Text = "86" Then
            csql = " Select tipodocreferencia,seriedocreferencia,numdocreferencia from docreferencia Where idempresa = '" & glsEmpresa & "' and tipodocorigen = '" & txtCod_TipoDoc.Text & "' " & _
                    " and seriedocorigen = '" & txt_serie.Text & "' and numdocorigen = '" & txtNum_Documento.Text & "'"
            If rst.State = 1 Then rst.Close
            rst.Open csql, Cn, adOpenKeyset, adLockOptimistic
            If rst.RecordCount > 0 Then
                rst.MoveFirst
                Do While Not rst.EOF
                    If Trim("" & rst.Fields("tipodocreferencia")) = "01" Or Trim("" & rst.Fields("tipodocreferencia")) = "03" Then
                        SecuenciaGuia1
                        Exit Do
                    ElseIf Trim("" & rst.Fields("tipodocreferencia")) = "40" Then
                        SecuenciaGuia2
                        Exit Do
                    End If
                    rst.MoveNext
                Loop
            End If
        End If
    Else
        MsgBox "El Documento no Existe ", vbInformation, App.Title
    End If

End Sub

Private Sub Toolbar1_Click()

    FraLista1.Visible = False
    FraLista2.Visible = False
    FraLista3.Visible = False
    FraLista4.Visible = False

End Sub

Private Sub txtCod_TipoDoc_KeyPress(KeyAscii As Integer)
    
    If KeyAscii = 13 Then
        txtCod_TipoDoc.Text = Format(Trim(txtCod_TipoDoc.Text), "00")
        If txtCod_TipoDoc.Text = "01" Or txtCod_TipoDoc.Text = "03" Or txtCod_TipoDoc.Text = "40" Or txtCod_TipoDoc.Text = "86" Or txtCod_TipoDoc.Text = "90" Or txtCod_TipoDoc.Text = "92" Then
            txtGls_TipoDoc.Text = Format(txtCod_TipoDoc.Text, "00")
            txtGls_TipoDoc.Text = traerCampo("documentos", "glsdocumento", "iddocumento", Format(txtCod_TipoDoc.Text, "00"), False)
            KeyAscii = 0
        Else
            Exit Sub
        End If
    Else
        txtGls_TipoDoc.Text = ""
        If txtGls_TipoDoc.Text <> "" Then SendKeys "{tab}"
    End If
    
End Sub
