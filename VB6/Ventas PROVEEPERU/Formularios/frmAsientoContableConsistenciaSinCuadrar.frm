VERSION 5.00
Object = "{F41D1D30-7878-4923-8CB3-6CCACDC9C9DE}#1.0#0"; "catcontrols.ocx"
Begin VB.Form frmAsientoContableConsistenciaSinCuadrar 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Asientos Contables Sin Cuadrar"
   ClientHeight    =   2175
   ClientLeft      =   4185
   ClientTop       =   3315
   ClientWidth     =   5910
   LinkTopic       =   "Generar Asientos Contables"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   2175
   ScaleWidth      =   5910
   Begin VB.Frame Frame5 
      Height          =   1590
      Left            =   45
      TabIndex        =   5
      Top             =   45
      Width           =   5820
      Begin VB.CommandButton cmbAyudaCtaCorriente 
         Height          =   315
         Left            =   5310
         Picture         =   "frmAsientoContableConsistenciaSinCuadrar.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   960
         Width           =   390
      End
      Begin VB.ComboBox cbxMes 
         Height          =   315
         ItemData        =   "frmAsientoContableConsistenciaSinCuadrar.frx":038A
         Left            =   3375
         List            =   "frmAsientoContableConsistenciaSinCuadrar.frx":03B2
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   315
         Width           =   2340
      End
      Begin VB.ComboBox cbxAno 
         Height          =   315
         ItemData        =   "frmAsientoContableConsistenciaSinCuadrar.frx":069A
         Left            =   900
         List            =   "frmAsientoContableConsistenciaSinCuadrar.frx":06B6
         Style           =   2  'Dropdown List
         TabIndex        =   0
         Top             =   315
         Width           =   1350
      End
      Begin CATControls.CATTextBox txtCod_Moneda 
         Height          =   315
         Left            =   900
         TabIndex        =   2
         Top             =   945
         Width           =   555
         _ExtentX        =   979
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
         Container       =   "frmAsientoContableConsistenciaSinCuadrar.frx":06EA
         Estilo          =   1
         EnterTab        =   -1  'True
      End
      Begin CATControls.CATTextBox txtGls_Moneda 
         Height          =   315
         Left            =   1485
         TabIndex        =   9
         Top             =   945
         Width           =   3810
         _ExtentX        =   6720
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
         Locked          =   -1  'True
         Container       =   "frmAsientoContableConsistenciaSinCuadrar.frx":0706
         Estilo          =   1
         Vacio           =   -1  'True
      End
      Begin VB.Label lbl_CtaCorriente 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         Caption         =   "Moneda"
         ForeColor       =   &H80000007&
         Height          =   195
         Left            =   180
         TabIndex        =   10
         Top             =   1005
         Width           =   585
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Año"
         Height          =   195
         Left            =   180
         TabIndex        =   7
         Top             =   405
         Width           =   285
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Mes"
         Height          =   195
         Left            =   2970
         TabIndex        =   6
         Top             =   360
         Width           =   300
      End
   End
   Begin VB.CommandButton cmbCancelar 
      Caption         =   "&Salir"
      Height          =   390
      Left            =   3090
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   1710
      Width           =   1095
   End
   Begin VB.CommandButton cmbOperar 
      Caption         =   "Aceptar"
      Height          =   390
      Left            =   1935
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   1710
      Width           =   1095
   End
End
Attribute VB_Name = "frmAsientoContableConsistenciaSinCuadrar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmbAyudaCtaCorriente_Click()
    
    mostrarAyuda "MONEDA", txtCod_Moneda, txtGls_Moneda
    
End Sub

Private Sub cmbCancelar_Click()
    
    Unload Me
    
End Sub

Private Sub cmbOperar_Click()
Dim StrMsgError As String
Dim cCompro As String
Dim xOrigen As String
Dim xFecha As Variant
Dim nDebes As Double, nHabers As Double
Dim nDebeD As Double, nHaberD As Double
Dim RsCuadrar           As New ADODB.Recordset
Dim vistaPrevia     As New frmReportePreview
Dim aplicacion      As New CRAXDRT.Application
Dim reporte         As CRAXDRT.Report
Dim GlsReporte      As String
Dim GlsForm         As String
On Error GoTo Err
    
    If Len(Trim(txtCod_Moneda.Text)) = 0 Then
        MsgBox "Ingrese Moneda", vbInformation, App.Title
        Exit Sub
    End If
    
    If rsAsientosContables.State = 1 Then
        If rsAsientosContables.RecordCount <> 0 Then
            If RsCuadrar.State = 1 Then RsCuadrar.Close
        
            RsCuadrar.Fields.Append "idComprobante", adVarChar, 9, adFldRowID
            RsCuadrar.Fields.Append "IdOrigen", adVarChar, 2, adFldIsNullable
            RsCuadrar.Fields.Append "FecCompro", adVarChar, 30, adFldRowID
            RsCuadrar.Fields.Append "GlsEmpresa", adVarChar, 250, adFldIsNullable
            RsCuadrar.Fields.Append "ValDebe", adDouble, adFldRowID
            RsCuadrar.Fields.Append "ValHaber", adDouble, adFldIsNullable
            RsCuadrar.Fields.Append "GlsTitulo", adVarChar, 250, adFldRowID
            RsCuadrar.Open
            
            rsAsientosContables.MoveFirst
            If Trim(rsAsientosContables.Fields("IdPeriodo") & "") = Trim(cbxAno.Text & Format(cbxMes.ListIndex + 1, "00")) Then
                Do While Not rsAsientosContables.EOF
                    nDebes = 0#: nHabers = 0#: nDebeD = 0#: nHaberD = 0#
                    xFecha = rsAsientosContables.Fields("FecCompro")
                    xOrigen = rsAsientosContables.Fields("IdOrigen") & ""
                    cCompro = rsAsientosContables.Fields("IdComprobante")
                    Do While cCompro = rsAsientosContables.Fields("IdComprobante") And Not rsAsientosContables.EOF
                        If rsAsientosContables.Fields("IdTipoDH") = "D" Then
                            nDebes = rsAsientosContables.Fields("TotalImporteS") + nDebes
                            nDebeD = rsAsientosContables.Fields("TotalImporteD") + nDebeD
                        Else
                            nHabers = rsAsientosContables.Fields("TotalImporteS") + nHabers
                            nHaberD = rsAsientosContables.Fields("TotalImporteD") + nHaberD
                        End If
                        rsAsientosContables.MoveNext
                        If rsAsientosContables.EOF Then Exit Do
                        If rsAsientosContables.Fields("IdComprobante") <> cCompro Then Exit Do
                    Loop
                    If txtCod_Moneda.Text = "PEN" Then
                        If Format(nDebes, "0.00") <> Format(nHabers, "0.00") Then
                            RsCuadrar.AddNew
                            RsCuadrar.Fields("IdComprobante") = cCompro
                            RsCuadrar.Fields("FecCompro") = xFecha
                            RsCuadrar.Fields("IdOrigen") = xOrigen
                            RsCuadrar.Fields("ValDebe") = nDebes
                            RsCuadrar.Fields("ValHaber") = nHabers
                            RsCuadrar.Fields("GlsTitulo") = "COMPROBANTES DE " & UCase(Trim(Mid(cbxMes.Text, 1, Len(cbxMes.Text) - 2))) & " DESCUADRADOS EN SOLES"
                            RsCuadrar.Update
                        End If
                    End If
                    If txtCod_Moneda.Text = "USD" Then
                        If Format(nDebeD, "0.00") <> Format(nHaberD, "0.00") Then
                            RsCuadrar.AddNew
                            RsCuadrar.Fields("idComprobante") = cCompro
                            RsCuadrar.Fields("FecCompro") = xFecha
                            RsCuadrar.Fields("IdOrigen") = xOrigen
                            RsCuadrar.Fields("ValDebe") = nDebeD
                            RsCuadrar.Fields("ValHaber") = nHaberD
                            RsCuadrar.Fields("GlsTitulo") = "COMPROBANTES DE " & UCase(Trim(Mid(cbxMes.Text, 1, Len(cbxMes.Text) - 2))) & " DESCUADRADOS EN DOLARES"
                            RsCuadrar.Update
                        End If
                    End If

                Loop
    
                Screen.MousePointer = 11
            
                GlsReporte = "RptConsisSinCuadrar.rpt"
                GlsForm = "Asientos Contables Sin Cuadrar"
                Set reporte = aplicacion.OpenReport(gStrRutaRpts & GlsReporte)
                If Not RsCuadrar.EOF And Not RsCuadrar.BOF Then
                    reporte.Database.SetDataSource RsCuadrar, 3
                    vistaPrevia.CRViewer91.ReportSource = reporte
                    vistaPrevia.CRViewer91.ViewReport
                    vistaPrevia.CRViewer91.DisplayGroupTree = False
                    Screen.MousePointer = 0
                    vistaPrevia.WindowState = 2
                    vistaPrevia.Show
                Else
                    MsgBox "No se reportan inconsistencias.", vbInformation, App.Title
                End If
            Else
                MsgBox "El mes no corresponde con los movimientos generados.", vbInformation, App.Title
            End If
        Else
            MsgBox "No existen movimientos generados para la consistencia.", vbInformation, App.Title
        End If
    Else
        MsgBox "No existen movimientos generados para la consistencia.", vbInformation, App.Title
    End If
    cmbCancelar.SetFocus
    
    Screen.MousePointer = 0
    Exit Sub
    
Err:
    Screen.MousePointer = 0
        
    If TypeName(RsCuadrar) = "Recordset" Then
        If RsCuadrar.State = 1 Then RsCuadrar.Close
        Set RsCuadrar = Nothing
    End If
    MsgBox Err.Description, vbInformation, App.Title
    
    Exit Sub
    
End Sub

Private Sub Form_Load()
Dim fecha As Date
Dim i As Integer
    
    fecha = Format(getFechaSistema, "dd/mm/yyyy")
    
    Me.Height = 2745
    Me.Width = 6150
    
    strAno = Format(Year(fecha), "0000")
    strMes = Format(Month(fecha), "00")
    
    cbxAno.Clear
    For i = 2008 To Val(strAno)
        cbxAno.AddItem i
    Next
    
    cbxAno.Clear
    For i = 2008 To Val(strAno)
        cbxAno.AddItem i
    Next
    
    For i = 0 To cbxAno.ListCount - 1
        cbxAno.ListIndex = i
        If cbxAno.Text = Format(Year(fecha), "0000") Then Exit For
    Next
    
    cbxMes.ListIndex = Val(Format(Month(fecha), "00")) - 1
    
    txtCod_Moneda.Text = "PEN"
    txtGls_Moneda.Text = "NUEVOS SOLES"
    
End Sub
