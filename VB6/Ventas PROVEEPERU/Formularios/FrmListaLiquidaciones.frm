VERSION 5.00
Object = "{6A24B331-7634-11D3-A5B0-0050044A7E1A}#1.5#0"; "DXDBGrid.dll"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{F41D1D30-7878-4923-8CB3-6CCACDC9C9DE}#1.0#0"; "catcontrols.ocx"
Begin VB.Form FrmListaLiquidaciones 
   Appearance      =   0  'Flat
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Lista de Guias Madres"
   ClientHeight    =   5550
   ClientLeft      =   3315
   ClientTop       =   2370
   ClientWidth     =   10050
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5550
   ScaleWidth      =   10050
   Begin VB.Frame fraListado 
      Appearance      =   0  'Flat
      Caption         =   "Listado"
      ForeColor       =   &H00000000&
      Height          =   4965
      Left            =   0
      TabIndex        =   1
      Top             =   600
      Width           =   10020
      Begin VB.Frame Frame1 
         Appearance      =   0  'Flat
         ForeColor       =   &H80000008&
         Height          =   735
         Left            =   120
         TabIndex        =   2
         Top             =   240
         Width           =   9705
         Begin CATControls.CATTextBox txtGls_Origen 
            Height          =   315
            Left            =   1575
            TabIndex        =   3
            Top             =   240
            Width           =   4080
            _ExtentX        =   7197
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
            Container       =   "FrmListaLiquidaciones.frx":0000
            Vacio           =   -1  'True
         End
         Begin MSComCtl2.DTPicker dtp_Desde 
            Height          =   315
            Left            =   6420
            TabIndex        =   6
            Top             =   225
            Width           =   1290
            _ExtentX        =   2275
            _ExtentY        =   556
            _Version        =   393216
            Format          =   103809025
            CurrentDate     =   38955
         End
         Begin MSComCtl2.DTPicker dtp_Hasta 
            Height          =   315
            Left            =   8340
            TabIndex        =   7
            Top             =   225
            Width           =   1290
            _ExtentX        =   2275
            _ExtentY        =   556
            _Version        =   393216
            Format          =   103809025
            CurrentDate     =   38955
         End
         Begin VB.Label Label2 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            Caption         =   "Hasta:"
            ForeColor       =   &H80000007&
            Height          =   195
            Left            =   7785
            TabIndex        =   9
            Top             =   285
            Width           =   465
         End
         Begin VB.Label Label3 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            Caption         =   "Desde:"
            ForeColor       =   &H80000007&
            Height          =   195
            Left            =   5760
            TabIndex        =   8
            Top             =   285
            Width           =   510
         End
         Begin VB.Label lbl_Cliente 
            Appearance      =   0  'Flat
            Caption         =   "Unidad Producción"
            ForeColor       =   &H80000007&
            Height          =   240
            Left            =   60
            TabIndex        =   4
            Top             =   300
            Width           =   1395
         End
      End
      Begin DXDBGRIDLibCtl.dxDBGrid gLista 
         Height          =   3555
         Left            =   135
         OleObjectBlob   =   "FrmListaLiquidaciones.frx":001C
         TabIndex        =   5
         Top             =   1035
         Width           =   9705
      End
   End
   Begin MSComctlLib.ImageList imgDocVentas 
      Left            =   750
      Top             =   4500
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
            Picture         =   "FrmListaLiquidaciones.frx":4B3D
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmListaLiquidaciones.frx":4ED7
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmListaLiquidaciones.frx":5329
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmListaLiquidaciones.frx":56C3
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmListaLiquidaciones.frx":5A5D
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmListaLiquidaciones.frx":5DF7
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmListaLiquidaciones.frx":6191
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmListaLiquidaciones.frx":652B
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmListaLiquidaciones.frx":68C5
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmListaLiquidaciones.frx":6C5F
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmListaLiquidaciones.frx":6FF9
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmListaLiquidaciones.frx":7CBB
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
      Width           =   10050
      _ExtentX        =   17727
      _ExtentY        =   1164
      ButtonWidth     =   1270
      ButtonHeight    =   1005
      Appearance      =   1
      ImageList       =   "imgDocVentas"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   3
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Aceptar"
            Object.ToolTipText     =   "Grabar"
            ImageIndex      =   9
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
Attribute VB_Name = "FrmListaLiquidaciones"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private rsg                 As New ADODB.Recordset
Private strTDExportar       As String
Dim indNuevoDoc             As Boolean
Dim StrOrigen               As String

Private Sub dtp_Desde_Change()
Dim StrMsgError As String

On Error GoTo Err
    
    listaDocVentas StrMsgError
    If StrMsgError <> "" Then GoTo Err

Exit Sub
Err:
If StrMsgError = "" Then StrMsgError = Err.Description
MsgBox StrMsgError, vbInformation, App.Title
End Sub

Private Sub dtp_Hasta_Change()
Dim StrMsgError As String

On Error GoTo Err
    
listaDocVentas StrMsgError
If StrMsgError <> "" Then GoTo Err

Exit Sub
Err:
If StrMsgError = "" Then StrMsgError = Err.Description
MsgBox StrMsgError, vbInformation, App.Title
End Sub

Private Sub Form_Load()
Dim StrMsgError As String
On Error GoTo Err

ConfGrid gLista, True, False, False, False

Set gLista.DataSource = Nothing

dtp_Hasta.Value = Format(getFechaSistema, "dd/mm/yyyy")
dtp_Desde.Value = Format(getFechaSistema, "dd/mm/yyyy")
    
txtGls_Origen.Text = Trim("" & traerCampo("unidadproduccion", "DescUnidad", "CodUnidProd", StrOrigen, False))

listaDocVentas StrMsgError
If StrMsgError <> "" Then GoTo Err

Exit Sub
Err:
If StrMsgError = "" Then StrMsgError = Err.Description
MsgBox StrMsgError, vbInformation, App.Title
End Sub

Private Sub listaDocVentas(ByRef StrMsgError As String)
Dim rst             As New ADODB.Recordset
Dim strCond         As String
Dim strFiltroAprob  As String
Dim item            As Integer
On Error GoTo Err

    '********FORMATO GRILLA
    Set gLista.DataSource = Nothing
    
    Set rsg = Nothing
    If rsg.State = 1 Then rsg.Close
    Set rsg = Nothing
    
    item = 0
    
    'Formato cabecera****************************************************
    rsg.Fields.Append "Item", adChar, 13, adFldRowID
    rsg.Fields.Append "chkMarca", adChar, 1, adFldIsNullable
    rsg.Fields.Append "DescUnidada", adChar, 255, adFldIsNullable
    rsg.Fields.Append "idupp", adChar, 8, adFldIsNullable
    rsg.Fields.Append "SerieGuia", adChar, 3, adFldIsNullable
    rsg.Fields.Append "NumGuia", adChar, 8, adFldIsNullable
    rsg.Fields.Append "FechaGuia", adVarChar, 30, adFldIsNullable
    rsg.Fields.Append "GlsPartida", adVarChar, 255, adFldIsNullable
    rsg.Fields.Append "ValCantidad", adDouble, 14, adFldIsNullable
    rsg.Fields.Append "Valpeso", adDouble, 14, adFldIsNullable
    rsg.Fields.Append "GlsPersona", adChar, 250, adFldIsNullable
    rsg.Fields.Append "idPersona", adChar, 8, adFldIsNullable
    rsg.Fields.Append "ValEX", adDouble, 14, adFldIsNullable
    '********************************************************************+
    rsg.Open , , adOpenKeyset, adLockOptimistic
    '********************************************************************

    csql = "SELECT a.ValEX,a.idupp,c.CodUnidProd as idpersona,c.DescUnidad as GlsPersona,b.DescUnidad,a.SerieGuia,a.NumGuia,a.FechaGuia,a.GlsPartida,a.ValCantidad,a.Valpeso FROM docventasguiasm a inner join unidadproduccion b " & _
           "on a.idupp = b.CodUnidProd " & _
           "and a.idempresa = b.idempresa " & _
           "inner join unidadproduccion c on a.idproveedor = c.CodUnidProd and a.idempresa = b.idempresa " & _
           "where a.idempresa = '" & glsEmpresa & "' and IndImportado = 0 " & _
           "and a.idproveedor = '" & StrOrigen & "' " & _
           "and a.FechaGuia between '" & Format(dtp_Desde.Value, "yyyy-mm-dd") & "' and '" & Format(dtp_Hasta.Value, "yyyy-mm-dd") & "' "
           
            
    rst.Open csql, Cn, adOpenForwardOnly, adLockReadOnly
    If Not rst.EOF And Not rst.BOF Then
        Do While Not rst.EOF
            rsg.AddNew
            item = item + 1
            rsg.Fields("Item") = item
            rsg.Fields("chkMarca") = 0
            rsg.Fields("DescUnidada") = rst.Fields("DescUnidad")
            rsg.Fields("idupp") = rst.Fields("idupp")
            rsg.Fields("SerieGuia") = rst.Fields("SerieGuia")
            rsg.Fields("NumGuia") = rst.Fields("NumGuia")
            rsg.Fields("FechaGuia") = Format(rst.Fields("FechaGuia"), "dd/mm/yyyy")
            rsg.Fields("GlsPartida") = rst.Fields("GlsPartida")
            rsg.Fields("ValCantidad") = Val(Format(rst.Fields("ValCantidad"), "0.00"))
            rsg.Fields("Valpeso") = Val(Format(rst.Fields("Valpeso"), "0.00"))
            rsg.Fields("ValEX") = Val(Format(rst.Fields("ValEX"), "0.00"))
            rsg.Fields("GlsPersona") = rst.Fields("GlsPersona")
            rsg.Fields("idPersona") = rst.Fields("idPersona")
            rst.MoveNext
        Loop

        mostrarDatosGridSQL gLista, rsg, StrMsgError
        If StrMsgError <> "" Then GoTo Err
    End If
    
Me.Refresh
If rst.State = 1 Then rst.Close
Set rst = Nothing
Exit Sub
Err:
If rst.State = 1 Then rst.Close
Set rst = Nothing
If StrMsgError = "" Then StrMsgError = Err.Description
End Sub

Private Sub Form_Unload(Cancel As Integer)

If rsg.State = 1 Then rsg.Close
Set rsg = Nothing

End Sub

Private Sub gLista_OnCheckEditToggleClick(ByVal Column As DXDBGRIDLibCtl.IdxGridColumn, ByVal Node As DXDBGRIDLibCtl.IdxGridNode, ByVal Text As String, ByVal State As DXDBGRIDLibCtl.ExCheckBoxState)
    If gLista.Dataset.State = dsEdit Then
        gLista.Dataset.Post
    End If
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
Dim StrMsgError     As String
Dim strNumDoc       As String
Dim strSerie        As String
Dim PAR             As String
On Error GoTo Err

    Select Case Button.Index
        Case 1 'Aceptar
            Me.Hide
        Case 3 'Salir
            Me.Hide
    End Select

Exit Sub
Err:
If StrMsgError = "" Then StrMsgError = Err.Description
MsgBox StrMsgError, vbInformation, App.Title
End Sub



Private Sub gLista_OnDblClick()
Dim StrMsgError As String
On Error GoTo Err

    If gLista.Count > 0 Then
        
        gLista.Dataset.Edit
            gLista.Columns.ColumnByFieldName("chkMarca").Value = 1
        gLista.Dataset.Post
        
        Me.Hide
    End If

Exit Sub
Err:
MsgBox StrMsgError, vbInformation, App.Title
End Sub

Public Sub MostrarForm(ByVal StrOrigenO As String, ByRef rscd As ADODB.Recordset, ByRef StrMsgError As String)

On Error GoTo Err
        
    StrOrigen = StrOrigenO
        
    FrmListaLiquidaciones.Show 1
    
    'Quitamos Filtros existentes
    gLista.Dataset.Filter = ""
    gLista.Dataset.Filtered = True
    
    Set gLista.DataSource = Nothing
    
    If TypeName(rsg) = "Nothing" Then
        Exit Sub
    Else
        If rsg.State = 0 Then
            Exit Sub
        End If
    End If
    
    'Eliminamos los registros q no estan marcados
    If rsg.RecordCount > 0 Then
        rsg.MoveFirst
        Do While Not rsg.EOF
            If rsg.Fields("chkMarca") = "0" Then
                rsg.Delete adAffectCurrent
                rsg.Update
            End If
            rsg.MoveNext
        Loop
    End If
       
    Set rscd = rsg.Clone(adLockReadOnly)
        
    If rsg.State = 1 Then rsg.Close
    Set rsg = Nothing
    
    Unload Me
Exit Sub
Err:
If StrMsgError = "" Then StrMsgError = Err.Description
End Sub


