VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{F41D1D30-7878-4923-8CB3-6CCACDC9C9DE}#1.0#0"; "catcontrols.ocx"
Begin VB.Form FrmDiasPromediar 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Reporte: Dias a Promediar"
   ClientHeight    =   3270
   ClientLeft      =   5400
   ClientTop       =   2085
   ClientWidth     =   7080
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3270
   ScaleWidth      =   7080
   Begin VB.Frame Frame1 
      Height          =   2625
      Left            =   0
      TabIndex        =   0
      Top             =   630
      Width           =   7080
      Begin VB.Frame Frame2 
         Appearance      =   0  'Flat
         Caption         =   "Dias:"
         ForeColor       =   &H00C00000&
         Height          =   645
         Left            =   90
         TabIndex        =   12
         Top             =   180
         Width           =   2670
         Begin CATControls.CATTextBox txtDias 
            Height          =   285
            Left            =   1575
            TabIndex        =   13
            Tag             =   "TidMoneda"
            Top             =   225
            Width           =   915
            _ExtentX        =   1614
            _ExtentY        =   503
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
            Alignment       =   1
            FontName        =   "MS Sans Serif"
            FontSize        =   8.25
            ForeColor       =   -2147483640
            MaxLength       =   8
            Container       =   "FrmDiasPromediar.frx":0000
            Estilo          =   1
            EnterTab        =   -1  'True
         End
         Begin VB.Label Label3 
            Appearance      =   0  'Flat
            Caption         =   "Dias a Promediar:"
            ForeColor       =   &H80000007&
            Height          =   240
            Left            =   225
            TabIndex        =   14
            Top             =   270
            Width           =   1350
         End
      End
      Begin VB.Frame fraReportes 
         Appearance      =   0  'Flat
         Caption         =   "Fechas:"
         ForeColor       =   &H00C00000&
         Height          =   765
         Index           =   1
         Left            =   90
         TabIndex        =   7
         Top             =   1710
         Width           =   6915
         Begin MSComCtl2.DTPicker dtpfInicio 
            Height          =   315
            Left            =   1575
            TabIndex        =   8
            Top             =   300
            Width           =   1365
            _ExtentX        =   2408
            _ExtentY        =   556
            _Version        =   393216
            Format          =   42598401
            CurrentDate     =   38667
         End
         Begin MSComCtl2.DTPicker dtpFFinal 
            Height          =   315
            Left            =   4515
            TabIndex        =   9
            Top             =   300
            Width           =   1365
            _ExtentX        =   2408
            _ExtentY        =   556
            _Version        =   393216
            Format          =   42598401
            CurrentDate     =   38667
         End
         Begin VB.Label Label1 
            Appearance      =   0  'Flat
            Caption         =   "del"
            ForeColor       =   &H80000008&
            Height          =   240
            Left            =   990
            TabIndex        =   11
            Top             =   375
            Width           =   390
         End
         Begin VB.Label Label2 
            Appearance      =   0  'Flat
            Caption         =   "al"
            ForeColor       =   &H80000008&
            Height          =   240
            Left            =   4140
            TabIndex        =   10
            Top             =   375
            Width           =   240
         End
      End
      Begin VB.Frame fraReportes 
         Appearance      =   0  'Flat
         Caption         =   "Sucursal:"
         ForeColor       =   &H00C00000&
         Height          =   765
         Index           =   4
         Left            =   90
         TabIndex        =   2
         Top             =   855
         Width           =   6915
         Begin VB.CommandButton cmbAyudaSucursal 
            Height          =   315
            Left            =   6285
            Picture         =   "FrmDiasPromediar.frx":001C
            Style           =   1  'Graphical
            TabIndex        =   3
            Top             =   300
            Width           =   390
         End
         Begin CATControls.CATTextBox txtCod_Sucursal 
            Height          =   285
            Left            =   1560
            TabIndex        =   4
            Tag             =   "TidMoneda"
            Top             =   315
            Width           =   915
            _ExtentX        =   1614
            _ExtentY        =   503
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
            Container       =   "FrmDiasPromediar.frx":03A6
            Estilo          =   1
            EnterTab        =   -1  'True
         End
         Begin CATControls.CATTextBox txtGls_Sucursal 
            Height          =   285
            Left            =   2535
            TabIndex        =   5
            Top             =   315
            Width           =   3690
            _ExtentX        =   6509
            _ExtentY        =   503
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
            Container       =   "FrmDiasPromediar.frx":03C2
            Vacio           =   -1  'True
         End
         Begin VB.Label Label4 
            Appearance      =   0  'Flat
            Caption         =   "Sucursal:"
            ForeColor       =   &H80000007&
            Height          =   240
            Left            =   240
            TabIndex        =   6
            Top             =   360
            Width           =   765
         End
      End
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   660
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   7080
      _ExtentX        =   12488
      _ExtentY        =   1164
      ButtonWidth     =   1402
      ButtonHeight    =   1005
      Appearance      =   1
      ImageList       =   "imgDocVentas"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   2
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Imprimir"
            Object.ToolTipText     =   "Nuevo"
            ImageIndex      =   11
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Salir"
            Object.ToolTipText     =   "Grabar"
            ImageIndex      =   2
         EndProperty
      EndProperty
      BorderStyle     =   1
   End
   Begin MSComctlLib.ImageList imgDocVentas 
      Left            =   6660
      Top             =   630
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
            Picture         =   "FrmDiasPromediar.frx":03DE
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmDiasPromediar.frx":0778
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmDiasPromediar.frx":0BCA
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmDiasPromediar.frx":0F64
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmDiasPromediar.frx":12FE
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmDiasPromediar.frx":1698
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmDiasPromediar.frx":1A32
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmDiasPromediar.frx":1DCC
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmDiasPromediar.frx":2166
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmDiasPromediar.frx":2500
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmDiasPromediar.frx":289A
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmDiasPromediar.frx":355C
            Key             =   ""
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "FrmDiasPromediar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmbAyudaSucursal_Click()
    mostrarAyuda "SUCURSAL", txtCod_Sucursal, txtGls_Sucursal
End Sub

Private Sub Form_Load()
    txtGls_Sucursal.Text = "TODAS LAS SUCURSALES"
    dtpfInicio.Value = Format(Date, "dd/mm/yyyy")
    dtpFFinal.Value = Format(Date, "dd/mm/yyyy")
    txtDias.Text = ""
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
Dim StrMsgError As String
On Error GoTo Err
Select Case Button.Index
    Case 1 'Imprimir
        imprimir App.Path & "\Temporales\Reporte_Dias_a_Promediar.xlt", StrMsgError
        If StrMsgError <> "" Then GoTo Err
    Case 2 'Salir
        Unload Me
End Select

Exit Sub
Err:
MsgBox StrMsgError, vbInformation, App.Title
End Sub

Private Sub txtCod_Sucursal_Click()
    If txtCod_Sucursal.Text <> "" Then
        txtGls_Sucursal.Text = traerCampo("personas", "GlsPersona", "idPersona", txtCod_Sucursal.Text, False)
    Else
        txtGls_Sucursal.Text = "TODOS LAS SUCURSALES"
    End If
End Sub

Private Sub imprimir(ByRef xArchivo As String, ByRef StrMsgError As String)
    Dim xl As Excel.Application
    Dim wb As Workbook
    Dim shRxn As Worksheet
    
    Dim fIni        As String
    Dim Ffin        As String
    Dim Hoja        As String
    Dim cHoja       As Integer
    Dim fila        As Integer
    Dim columna     As Integer
    Dim i           As Integer
    Dim contador    As Integer
    
    Dim rsCabecera As New ADODB.Recordset
    Dim rsDetalle As New ADODB.Recordset
    Dim rsRangos As New ADODB.Recordset
    
    MousePointer = 13
    
    On Error GoTo ExcelNoAbierto
    Set xl = GetObject(, "Excel.Application")
    GoTo YaEstabaAbierto
ExcelNoAbierto:
    Set xl = New Excel.Application
YaEstabaAbierto:
    On Error GoTo 0
    Set wb = xl.Workbooks.Open(xArchivo)
    xl.Visible = False
    cHoja = 1
    Hoja = "Hoja" & CStr(cHoja)
    Set shRxn = wb.Worksheets(Hoja)
    wb.Worksheets(Hoja).Select
    
    fIni = Format(dtpfInicio.Value, "yyyy-mm-dd")
    Ffin = Format(dtpFFinal.Value, "yyyy-mm-dd")
    
    csql = "CALL spu_ListaDiasPromediar('" & glsEmpresa & "', '" & Trim(txtCod_Sucursal.Text) & "', " & CInt(Trim(txtDias.Text)) & ", '" & glsListaVentas & "', '" & fIni & "', '" & Ffin & "')"
    Cn.Execute csql
    
    csql = "select GlsGrupo, idProducto, GlsProducto, PVUnit from tmpdiaspromediar group by idProducto order by GlsGrupo, GlsProducto"
    If rsCabecera.State = 1 Then rsCabecera.Close
    rsCabecera.Open csql, Cn, adOpenStatic, adLockReadOnly
    
    fila = 7
    columna = 3
    contador = 0
    
    xl.Cells(2, 3).Value = traerCampo("empresas", "GlsEmpresa", "idEmpresa", glsEmpresa, False)
    xl.Cells(3, 3).Value = txtGls_Sucursal.Text
    xl.Cells(2, 9).Value = Format(dtpfInicio.Value, "dd/mm/yyyy")
    xl.Cells(3, 9).Value = Format(dtpFFinal.Value, "dd/mm/yyyy")
    xl.Cells(5, 2).Value = "Nro. Dias: " & CStr(txtDias.Text)
    
    csql = "Select distinct Rango From tmpdiaspromediar ORDER BY Rango"
    If rsRangos.State = 1 Then rsRangos.Close
    rsRangos.Open csql, Cn, adOpenStatic, adLockReadOnly
    If rsRangos.RecordCount <> 0 Then
        Do While Not rsRangos.EOF
            xl.Cells(13 + contador, 2).Value = "" & rsRangos.Fields("Rango")
            contador = contador + 1
            rsRangos.MoveNext
        Loop
        rsRangos.Close
        Set rsRangos = Nothing
    End If
    contador = 0

    If rsCabecera.RecordCount <> 0 Then
        Do While Not rsCabecera.EOF
            If contador <= 252 Then
                xl.Cells(fila + 1, columna).Value = "" & rsCabecera.Fields("GlsGrupo")
                xl.Cells(fila + 2, columna).Value = "" & rsCabecera.Fields("idProducto")
                xl.Cells(fila + 3, columna).Value = "" & rsCabecera.Fields("GlsProducto")
                xl.Cells(fila + 4, columna).Value = "" & rsCabecera.Fields("PVUnit")
                
                csql = "select distinct a.Rango,(select Unidades from tmpdiaspromediar where a.Rango = Rango and idProducto = '" & rsCabecera.Fields("idProducto") & "') as Unidades from tmpdiaspromediar a ORDER BY a.Rango"
                If rsDetalle.State = 1 Then rsDetalle.Close
                rsDetalle.Open csql, Cn, adOpenStatic, adLockReadOnly
                If rsDetalle.RecordCount <> 0 Then
                    i = 6
                    Do While Not rsDetalle.EOF
                        xl.Cells(fila + i, columna).Value = "" & rsDetalle.Fields("Unidades")
                        i = i + 1
                        rsDetalle.MoveNext
                    Loop
                    rsDetalle.Close
                    Set rsDetalle = Nothing
                End If
                columna = columna + 1
            Else
                cHoja = cHoja + 1
                Hoja = "Hoja" & CStr(cHoja)
                Set shRxn = wb.Worksheets(Hoja)
                wb.Worksheets(Hoja).Select
                
                xl.Cells(2, 3).Value = traerCampo("empresas", "GlsEmpresa", "idEmpresa", glsEmpresa, False)
                xl.Cells(3, 3).Value = txtGls_Sucursal.Text
                xl.Cells(2, 9).Value = Format(dtpfInicio.Value, "dd/mm/yyyy")
                xl.Cells(3, 9).Value = Format(dtpFFinal.Value, "dd/mm/yyyy")
                xl.Cells(5, 2).Value = "Nro. Dias: " & CStr(txtDias.Text)
                
                csql = "Select distinct Rango From tmpdiaspromediar ORDER BY Rango "
                If rsRangos.State = 1 Then rsRangos.Close
                rsRangos.Open csql, Cn, adOpenStatic, adLockReadOnly
                If rsRangos.RecordCount <> 0 Then
                    contador = 0
                    Do While Not rsRangos.EOF
                        xl.Cells(13 + contador, 2).Value = "" & rsRangos.Fields("Rango")
                        contador = contador + 1
                        rsRangos.MoveNext
                    Loop
                    rsRangos.Close
                    Set rsRangos = Nothing
                End If
                contador = -1
                columna = 3
            End If
            contador = contador + 1
            rsCabecera.MoveNext
        Loop
    End If
    
   MousePointer = 1
    
   xl.Visible = True

End Sub
