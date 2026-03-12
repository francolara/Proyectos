VERSION 5.00
Object = "{6A24B331-7634-11D3-A5B0-0050044A7E1A}#1.5#0"; "DXDBGrid.dll"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "mscomctl.OCX"
Object = "{F41D1D30-7878-4923-8CB3-6CCACDC9C9DE}#1.0#0"; "CATControls.ocx"
Begin VB.Form Frm_Det_Formula 
   Appearance      =   0  'Flat
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Detalle de Productos por Fórmula"
   ClientHeight    =   6660
   ClientLeft      =   2355
   ClientTop       =   1875
   ClientWidth     =   11490
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6660
   ScaleWidth      =   11490
   ShowInTaskbar   =   0   'False
   Begin VB.Frame fraGeneral 
      Appearance      =   0  'Flat
      ForeColor       =   &H80000008&
      Height          =   5940
      Left            =   45
      TabIndex        =   1
      Top             =   675
      Width           =   11385
      Begin CATControls.CATTextBox txtCod_Almacen 
         Height          =   315
         Left            =   1695
         TabIndex        =   2
         Tag             =   "TidAlmacen"
         Top             =   1020
         Width           =   915
         _ExtentX        =   1614
         _ExtentY        =   556
         BackColor       =   16777152
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
         Container       =   "Frm_Det_Formula.frx":0000
         Estilo          =   1
         EnterTab        =   -1  'True
      End
      Begin CATControls.CATTextBox txtGls_Almacen 
         Height          =   315
         Left            =   2700
         TabIndex        =   3
         Top             =   1020
         Width           =   7365
         _ExtentX        =   12991
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
         Container       =   "Frm_Det_Formula.frx":001C
         Vacio           =   -1  'True
      End
      Begin CATControls.CATTextBox txtCod_Combo 
         Height          =   315
         Index           =   0
         Left            =   1695
         TabIndex        =   5
         Tag             =   "TidNivelPred"
         Top             =   645
         Width           =   915
         _ExtentX        =   1614
         _ExtentY        =   556
         BackColor       =   16777152
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
         Container       =   "Frm_Det_Formula.frx":0038
         Estilo          =   1
         Vacio           =   -1  'True
         EnterTab        =   -1  'True
      End
      Begin CATControls.CATTextBox txtGls_Combo 
         Height          =   315
         Index           =   0
         Left            =   2700
         TabIndex        =   6
         Top             =   660
         Width           =   7365
         _ExtentX        =   12991
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
         Container       =   "Frm_Det_Formula.frx":0054
         Vacio           =   -1  'True
      End
      Begin DXDBGRIDLibCtl.dxDBGrid g 
         Height          =   4320
         Left            =   180
         OleObjectBlob   =   "Frm_Det_Formula.frx":0070
         TabIndex        =   8
         Top             =   1440
         Width           =   11025
      End
      Begin VB.Label lblNivel 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         Caption         =   "Combo"
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
         Index           =   0
         Left            =   735
         TabIndex        =   7
         Top             =   720
         Width           =   495
      End
      Begin VB.Label lbl_Almacen 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         Caption         =   "Almacén"
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
         Left            =   735
         TabIndex        =   4
         Top             =   1065
         Width           =   630
      End
   End
   Begin MSComctlLib.ImageList imgDocVentas 
      Left            =   10305
      Top             =   165
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   13
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Frm_Det_Formula.frx":2E53
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Frm_Det_Formula.frx":31ED
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Frm_Det_Formula.frx":363F
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Frm_Det_Formula.frx":39D9
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Frm_Det_Formula.frx":3D73
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Frm_Det_Formula.frx":410D
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Frm_Det_Formula.frx":44A7
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Frm_Det_Formula.frx":4841
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Frm_Det_Formula.frx":4BDB
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Frm_Det_Formula.frx":4F75
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Frm_Det_Formula.frx":530F
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Frm_Det_Formula.frx":5FD1
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Frm_Det_Formula.frx":636B
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
      Width           =   11490
      _ExtentX        =   20267
      _ExtentY        =   1164
      ButtonWidth     =   1402
      ButtonHeight    =   1005
      Appearance      =   1
      ImageList       =   "imgDocVentas"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   2
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Aceptar"
            Object.ToolTipText     =   "Grabar"
            ImageIndex      =   4
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Cancelar"
            Object.ToolTipText     =   "Cancelar"
            ImageIndex      =   9
         EndProperty
      EndProperty
      BorderStyle     =   1
   End
End
Attribute VB_Name = "Frm_Det_Formula"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim PRsVentas                       As ADODB.Recordset
Dim SwIngreso                       As Boolean
Dim CItemDocVentas                  As Integer
Dim CIdLista                        As String

Public Sub MostrarForm(StrMsgError As String, PIdCombo As String, PIdAlmacen As String, PIdLista As String, PRsDetProductos As ADODB.Recordset, PSwIngreso As Boolean, PItemDocVentas As Integer, PVVUnit As Double)
On Error GoTo Err

    Set PRsVentas = PRsDetProductos.Clone(adLockReadOnly)
    CItemDocVentas = PItemDocVentas
    CIdLista = PIdLista
    If PRsVentas.RecordCount > 0 Then
        PRsVentas.MoveFirst
        PRsVentas.Filter = "ItemDocVentas = " & PItemDocVentas & ""
    End If
    txtCod_Combo(0).Text = PIdCombo
    txtCod_Almacen.Text = PIdAlmacen
    SwIngreso = False
    
    Me.Show 1
    
    PRsVentas.Filter = ""
    PRsVentas.Filter = adFilterNone
        
    PSwIngreso = SwIngreso
    If PSwIngreso Then
        With PRsDetProductos
            If .RecordCount > 0 Then
                .MoveFirst
                .Filter = "ItemDocVentas = " & PItemDocVentas & ""
                Do While Not .EOF
                    .Delete adAffectCurrent
                    .Update
                    .MoveNext
                Loop
                .Filter = ""
                .Filter = adFilterNone
            End If
        End With
        With g
            If Not .Dataset.EOF Then
            
                PVVUnit = Val("" & .Columns.ColumnByFieldName("TotalVVenta").SummaryFooterValue)
                
                .Dataset.First
        
                Do While Not .Dataset.EOF
                    PRsDetProductos.AddNew
                    PRsDetProductos.Fields("Item") = .Columns.ColumnByFieldName("Item").Value
                    PRsDetProductos.Fields("IdProducto") = .Columns.ColumnByFieldName("IdProducto").Value
                    PRsDetProductos.Fields("GlsProducto") = .Columns.ColumnByFieldName("GlsProducto").Value
                    PRsDetProductos.Fields("CantidadStock") = 0
                    PRsDetProductos.Fields("Cantidad") = Val(.Columns.ColumnByFieldName("Cantidad").Value)
                    PRsDetProductos.Fields("ItemDocVentas") = PItemDocVentas
                    PRsDetProductos.Fields("IdComboCab") = txtCod_Combo(0).Text
                    PRsDetProductos.Fields("VVUnit") = Val("" & .Columns.ColumnByFieldName("VVUnit").Value)
                    PRsDetProductos.Fields("PorcDcto") = Val("" & .Columns.ColumnByFieldName("PorcDcto").Value)
                    PRsDetProductos.Fields("TotalVVenta") = Val("" & .Columns.ColumnByFieldName("TotalVVenta").Value)
                    .Dataset.Next
                Loop
                
            End If
        End With
    End If
    
    Unload Me
    
    Exit Sub
Err:
    If StrMsgError = "" Then StrMsgError = Err.Description
    Exit Sub
End Sub

Private Sub Form_Load()
On Error GoTo Err
Dim StrMsgError As String
    
    ConfGrid g, True, False, False, False
    llena_grid StrMsgError
    If StrMsgError <> "" Then GoTo Err
    
    Exit Sub
    
Err:
    If StrMsgError = "" Then StrMsgError = Err.Description
    MsgBox StrMsgError, vbInformation, App.Title
End Sub

Private Sub llena_grid(StrMsgError As String)
On Error GoTo Err
Dim RsGrid  As New ADODB.Recordset
Dim X       As Integer
    
    With PRsVentas
        RsGrid.Fields.Append "Item", adDouble, 14, adFldIsNullable
        RsGrid.Fields.Append "IdProducto", adVarChar, 10, adFldIsNullable
        RsGrid.Fields.Append "GlsProducto", adVarChar, 240, adFldIsNullable
        RsGrid.Fields.Append "CantidadStock", adDouble, 14, adFldIsNullable
        RsGrid.Fields.Append "Cantidad", adDouble, 14, adFldIsNullable
        RsGrid.Fields.Append "ItemDocventas", adDouble, 14, adFldIsNullable
        RsGrid.Fields.Append "IdComboCab", adVarChar, 10, adFldIsNullable
        RsGrid.Fields.Append "VVUnit", adDouble, 14, adFldIsNullable
        RsGrid.Fields.Append "PorcDcto", adInteger, , adFldIsNullable
        RsGrid.Fields.Append "TotalVVenta", adDouble, 14, adFldIsNullable
        RsGrid.Open
        
        If Not .EOF Then
            .MoveFirst
            X = 0
            Do While Not .EOF
                If .Fields("ItemDocVentas") = CItemDocVentas Then
                    X = X + 1
                    RsGrid.AddNew
                    RsGrid.Fields("Item") = X
                    RsGrid.Fields("IdProducto") = "" & .Fields("IdProducto")
                    RsGrid.Fields("GlsProducto") = "" & .Fields("GlsProducto")
                    RsGrid.Fields("CantidadStock") = "0"
                    RsGrid.Fields("Cantidad") = Val("" & .Fields("Cantidad"))
                    RsGrid.Fields("ItemDocVentas") = Val("" & .Fields("ItemDocVentas"))
                    RsGrid.Fields("IdComboCab") = "" & .Fields("IdComboCab")
                    RsGrid.Fields("VVUnit") = Val("" & .Fields("VVUnit"))
                    RsGrid.Fields("PorcDcto") = Val("" & .Fields("PorcDcto"))
                    RsGrid.Fields("TotalVVenta") = Val("" & .Fields("TotalVVenta"))
                End If
                .MoveNext
            Loop
        Else
            RsGrid.AddNew
            RsGrid.Fields("Item") = 1
            RsGrid.Fields("IdProducto") = ""
            RsGrid.Fields("GlsProducto") = ""
            RsGrid.Fields("CantidadStock") = "0"
            RsGrid.Fields("Cantidad") = "0"
            RsGrid.Fields("ItemDocVentas") = "0"
            RsGrid.Fields("IdComboCab") = ""
            RsGrid.Fields("VVUnit") = "0"
            RsGrid.Fields("PorcDcto") = "0"
            RsGrid.Fields("TotalVVenta") = "0"
        End If
    End With
    
    With g
        .DefaultFields = False
        .Dataset.ADODataset.CursorLocation = clUseClient
        .Dataset.Active = False
        Set .DataSource = RsGrid
        .Dataset.Active = True
        .KeyField = "Item"
        .Dataset.Edit
        .Dataset.Post
        .Dataset.RecNo = 1
    End With
    
    Exit Sub
Err:
    If StrMsgError = "" Then StrMsgError = Err.Description
    Exit Sub
End Sub

Private Sub Aceptar(StrMsgError As String)
On Error GoTo Err

    If g.Count >= 1 Then
        If g.Count = 1 And (g.Columns.ColumnByFieldName("idProducto").Value = "" Or g.Columns.ColumnByFieldName("Cantidad").Value <= 0) Then
            StrMsgError = "Falta Ingresar Detalle"
            GoTo Err
        End If
    End If
    SwIngreso = True
    Me.Hide
    
    Exit Sub
    
Err:
    If StrMsgError = "" Then StrMsgError = Err.Description
End Sub

Private Sub eliminaNulosGrilla()
Dim indWhile As Boolean
Dim indEntro As Boolean
Dim i As Integer

    indWhile = True
    Do While indWhile = True
        If g.Count >= 1 Then
            g.Dataset.First
            indEntro = False
            Do While Not g.Dataset.EOF
                If Trim(g.Columns.ColumnByFieldName("idProducto").Value) = "" Or g.Columns.ColumnByFieldName("Cantidad").Value <= 0 Then
                    g.Dataset.Delete
                    indEntro = True
                    Exit Do
                End If
                g.Dataset.Next
            Loop
            indWhile = indEntro
        Else
            indWhile = False
        End If
    Loop
    
    If g.Count >= 1 Then
        g.Dataset.First
        i = 0
        Do While Not g.Dataset.EOF
            i = i + 1
            g.Dataset.Edit
            g.Columns.ColumnByFieldName("item").Value = i
            If g.Dataset.State = dsEdit Then g.Dataset.Post
            g.Dataset.Next
        Loop
    Else
        g.Dataset.Append
    End If
    
End Sub

Private Sub g_OnAfterDatasetAction(ByVal Action As DXDBGRIDLibCtl.ExDatasetAction)
On Error GoTo Err
Dim StrMsgError                     As String
Dim X                               As Double

    If Action = daInsert Then
        With g
            .Dataset.Edit
            .Columns.ColumnByFieldName("Item").Value = .Count
            .Columns.ColumnByFieldName("IdProducto").Value = ""
            .Columns.ColumnByFieldName("GlsProducto").Value = ""
            .Columns.ColumnByFieldName("IdUM").Value = ""
            .Columns.ColumnByFieldName("Cantidad").Value = "0"
            .Columns.ColumnByFieldName("VVUnit").Value = "0"
            .Columns.ColumnByFieldName("PorcDcto").Value = "0"
            .Columns.ColumnByFieldName("TotalVVenta").Value = "0"
            .Dataset.Post
        End With
    End If
    
    Exit Sub
Err:
    If StrMsgError = "" Then StrMsgError = Err.Description
    MsgBox StrMsgError, vbInformation, App.Title
    Exit Sub
End Sub

Private Sub g_OnBeforeDatasetAction(ByVal Action As DXDBGRIDLibCtl.ExDatasetAction, Allow As Boolean)
On Error GoTo Err
Dim StrMsgError As String

    If Action = daInsert Then
        With g
        End With
    End If
    
    Exit Sub

Err:
    Allow = False
    If StrMsgError = "" Then StrMsgError = Err.Description
    MsgBox StrMsgError, vbInformation, App.Title
    Exit Sub
End Sub

Private Sub g_OnChangeNode(ByVal OldNode As DXDBGRIDLibCtl.IdxGridNode, ByVal Node As DXDBGRIDLibCtl.IdxGridNode)
On Error GoTo Err
Dim StrMsgError As String

    If traerCampo("Productos V", "V.IdTipoProducto", "V.IdProducto", g.Columns.ColumnByFieldName("IdProducto").Value, True) = "06002" Then
        g.Columns.ColumnByFieldName("GlsProducto").DisableEditor = False
        g.Columns.ColumnByFieldName("VVUnit").DisableEditor = False
    Else
        g.Columns.ColumnByFieldName("GlsProducto").DisableEditor = True
        If Val("" & traerCampo("PreciosVenta V", "V.VVUnit", "V.IdProducto", g.Columns.ColumnByFieldName("IdProducto").Value, True, "V.IdLista = '" & CIdLista & "'")) > 0 Then
            g.Columns.ColumnByFieldName("VVUnit").DisableEditor = True
        Else
            g.Columns.ColumnByFieldName("VVUnit").DisableEditor = False
        End If
    End If
    
    Exit Sub
Err:
    If StrMsgError = "" Then StrMsgError = Err.Description
    MsgBox StrMsgError, vbInformation, App.Title
End Sub

Private Sub G_OnEdited(ByVal Node As DXDBGRIDLibCtl.IdxGridNode)
On Error GoTo Err
Dim StrMsgError                 As String
Dim X                           As Double
Dim CArrays(3)                  As String
    
    With g
        If .Dataset.Modified = False Then Exit Sub

        Select Case .Columns.FocusedColumn.Index
            Case .Columns.ColumnByFieldName("IdProducto").Index
                If Len(Trim(.Columns.ColumnByFieldName("IdProducto").Value)) > 0 Then
                    traerCampos "PreciosVenta V Inner Join Productos P On V.IdEmpresa = P.IdEmpresa And V.IdProducto = P.IdProducto", "P.GlsProducto,P.IdUMVenta,V.VVUnit", "V.IdProducto", .Columns.ColumnByFieldName("IdProducto").Value, 3, CArrays, False, "V.IdEmpresa = '" & glsEmpresa & "' And V.IdLista = '" & CIdLista & "'"
                    .Columns.ColumnByFieldName("GlsProducto").Value = CArrays(0)
                    .Columns.ColumnByFieldName("IdUM").Value = CArrays(1)
                    .Columns.ColumnByFieldName("Cantidad").Value = "0"
                    .Columns.ColumnByFieldName("VVUnit").Value = Val(CArrays(2))
                    .Columns.ColumnByFieldName("PorcDcto").Value = "0"
                    .Columns.ColumnByFieldName("TotalVVenta").Value = "0"
                    
                Else
                    .Columns.ColumnByFieldName("GlsProducto").Value = ""
                    .Columns.ColumnByFieldName("IdUM").Value = ""
                    .Columns.ColumnByFieldName("Cantidad").Value = "0"
                    .Columns.ColumnByFieldName("VVUnit").Value = "0"
                    .Columns.ColumnByFieldName("PorcDcto").Value = "0"
                    .Columns.ColumnByFieldName("TotalVVenta").Value = "0"
                End If
            
            Case .Columns.ColumnByFieldName("Cantidad").Index, .Columns.ColumnByFieldName("VVUnit").Index, .Columns.ColumnByFieldName("PorcDcto").Index
                .Columns.ColumnByFieldName("Cantidad").Value = Val(.Columns.ColumnByFieldName("Cantidad").Value & "")
                .Columns.ColumnByFieldName("VVUnit").Value = Val(.Columns.ColumnByFieldName("VVUnit").Value & "")
                .Columns.ColumnByFieldName("PorcDcto").Value = Val(.Columns.ColumnByFieldName("PorcDcto").Value & "")
                
                If Val(.Columns.ColumnByFieldName("PorcDcto").Value & "") > Val("" & traerCampo("PreciosVenta", "MaxDcto", "IdProducto", .Columns.ColumnByFieldName("IdProducto").Value, True, "IdLista = '" & CIdLista & "'")) Then
                    .Columns.ColumnByFieldName("PorcDcto").Value = "0"
                    StrMsgError = "El Descuento ingresado es mayor al Máximo descuento asignado a producto,Verifique."
                End If
                
                If .Columns.ColumnByFieldName("PorcDcto").Value > "0" Then
                    .Columns.ColumnByFieldName("TotalVVenta").Value = (Val(.Columns.ColumnByFieldName("VVUnit").Value & "") - (Val(.Columns.ColumnByFieldName("VVUnit").Value & "") * (Val(.Columns.ColumnByFieldName("PorcDcto").Value & "") / 100))) * Val(.Columns.ColumnByFieldName("Cantidad").Value & "")
                Else
                    .Columns.ColumnByFieldName("TotalVVenta").Value = Val(.Columns.ColumnByFieldName("VVUnit").Value & "") * Val(.Columns.ColumnByFieldName("Cantidad").Value & "")
                End If
                If StrMsgError <> "" Then GoTo Err
        End Select
    End With
    
    Exit Sub
    
Err:
    If StrMsgError = "" Then StrMsgError = Err.Description
    MsgBox StrMsgError, vbInformation, App.Title
    Exit Sub
End Sub

Private Sub g_OnKeyDown(KeyCode As Integer, ByVal Shift As Long)
On Error GoTo Err
Dim StrMsgError As String

    If KeyCode = 46 Then
        If g.Count > 0 Then
            If MsgBox("¿Seguro de eliminar el registro?", vbInformation + vbYesNo, App.Title) = vbYes Then
                g.Dataset.Delete
            End If
        End If
    End If
    
    If KeyCode = 13 Then
        If g.Dataset.State = dsEdit Or g.Dataset.State = dsInsert Then
              g.Dataset.Post
        End If
    End If
    
    Exit Sub
Err:
    If StrMsgError = "" Then StrMsgError = Err.Description
    MsgBox StrMsgError, vbInformation, App.Title
    Exit Sub
End Sub

Private Sub g_OnEditButtonClick(ByVal Column As DXDBGRIDLibCtl.IdxGridColumn, ByVal Node As DXDBGRIDLibCtl.IdxGridNode)
On Error GoTo Err
Dim StrMsgError As String
Dim StrCod As String
Dim StrDes As String
Dim strCodUM   As String
Dim strDesUM   As String
Dim intFila As Integer
Dim indPedido As Boolean
Dim rspa        As New ADODB.Recordset
    
    intFila = Node.Index + 1

    Select Case Column.Index
        Case g.Columns.ColumnByFieldName("idProducto").Index
            StrCod = g.Columns.ColumnByFieldName("idProducto").Value
            StrDes = g.Columns.ColumnByFieldName("GlsProducto").Value
            strCodUM = g.Columns.ColumnByFieldName("idUM").Value
            
            indPedido = False
            FrmAyudaProdOC.ExecuteReturnTextAlm txtCod_Almacen.Text, rspa, StrCod, StrDes, "", glsValidaStock, CIdLista, False, True, indPedido, StrMsgError
            g.SetFocus
            
            If rspa.RecordCount <> 0 Then
                mostrarProdImp rspa, StrMsgError
                If StrMsgError <> "" Then GoTo Err
            End If
    End Select
    
    Exit Sub

Err:
    If StrMsgError = "" Then StrMsgError = Err.Description
    MsgBox StrMsgError, vbInformation, App.Title
End Sub

Private Sub mostrarProdImp(ByVal rsdd As ADODB.Recordset, ByRef StrMsgError As String)
On Error GoTo Err
Dim rsg As New ADODB.Recordset
Dim rsd As New ADODB.Recordset
Dim strSerieDocVentas As String
Dim dblTC  As Double
Dim strCodFabri As String
Dim strCodMar As String
Dim strDesMar As String
Dim intAfecto As Integer
Dim strTipoProd As String
Dim strMoneda As String
Dim strCodUM   As String
Dim strDesUM   As String
Dim dblVVUnit  As Double
Dim dblIGVUnit  As Double
Dim dblPVUnit  As Double
Dim dblFactor  As Double
Dim intFila As Integer
Dim i As Integer
Dim indExisteDocRef As Boolean
Dim primero As Boolean
Dim strInserta As Boolean

    primero = True
    rsdd.MoveFirst
    Do While Not rsdd.EOF
        strInserta = True
        If strInserta = True Then
        
            If primero = True Then
                primero = False
            Else
                g.Dataset.Insert
            End If
        
            g.SetFocus
            g.Dataset.Edit
            g.Columns.ColumnByFieldName("idProducto").Value = "" & rsdd.Fields("idProducto")
            g.Columns.ColumnByFieldName("GlsProducto").Value = "" & rsdd.Fields("GlsProducto")
            
            If Trim("" & rsdd.Fields("idProducto")) = "" Then Exit Sub
            
            g.Columns.ColumnByFieldName("Cantidad").Value = 1
            g.Columns.ColumnByFieldName("VVUnit").Value = Val("" & traerCampo("PreciosVenta", "VVUnit", "IdProducto", "" & rsdd.Fields("idProducto"), True, "IdLista = '" & CIdLista & "'"))
            g.Columns.ColumnByFieldName("PorcDcto").Value = 1
            g.Columns.ColumnByFieldName("TotalVVenta").Value = Val(g.Columns.ColumnByFieldName("VVUnit").Value & "")
            
            g.Dataset.Post
            g.Dataset.RecNo = intFila
            g.Dataset.Edit
                            
            g.Dataset.Post
            If "" & rsdd.Fields("idProducto") <> "" Then
                g.Columns.FocusedIndex = g.Columns.ColumnByFieldName("Cantidad").Index
            End If
        End If
        rsdd.MoveNext
    Loop

    Exit Sub
    
Err:
    If StrMsgError = "" Then StrMsgError = Err.Description
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
On Error GoTo Err
Dim StrMsgError As String
Dim strCodUltProd As String

    Select Case Button.Index
        Case 1 'Grabar
            Aceptar StrMsgError
            If StrMsgError <> "" Then GoTo Err
        Case 2 'Cancelar
            SwIngreso = False
            Me.Hide
            
    End Select
    
    Exit Sub
    
Err:
    MsgBox StrMsgError, vbInformation, App.Title
End Sub

Private Sub txtCod_Almacen_Change()
    
    txtGls_Almacen.Text = traerCampo("almacenes", "GlsAlmacen", "idAlmacen", txtCod_Almacen.Text, True)

End Sub

Private Sub txtCod_Combo_Change(Index As Integer)
    
    txtGls_Combo(0).Text = traerCampo("productos", "glsProducto", "idProducto", txtCod_Combo(0).Text, True)

End Sub
