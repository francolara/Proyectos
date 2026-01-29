VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{F41D1D30-7878-4923-8CB3-6CCACDC9C9DE}#1.0#0"; "catcontrols.ocx"
Begin VB.Form FrmImportaVales 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Importar Vales"
   ClientHeight    =   2565
   ClientLeft      =   4830
   ClientTop       =   3795
   ClientWidth     =   8295
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   2565
   ScaleWidth      =   8295
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdSalir 
      Caption         =   "&Salir"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   4275
      TabIndex        =   8
      Top             =   1980
      Width           =   1320
   End
   Begin VB.CommandButton CmdAceptar 
      Caption         =   "&Aceptar"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   2835
      TabIndex        =   7
      Top             =   1980
      Width           =   1320
   End
   Begin VB.Frame Frame1 
      Height          =   1725
      Left            =   90
      TabIndex        =   0
      Top             =   90
      Width           =   8070
      Begin VB.TextBox TxtCod_Concepto 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFC0&
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   1710
         Locked          =   -1  'True
         TabIndex        =   5
         Top             =   720
         Width           =   825
      End
      Begin VB.TextBox TxtGls_Concepto 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFC0&
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   2610
         Locked          =   -1  'True
         TabIndex        =   4
         Top             =   720
         Width           =   4740
      End
      Begin VB.CommandButton CmbAyudaExcel 
         Height          =   315
         Left            =   7455
         Picture         =   "FrmImportaVales.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   270
         Width           =   390
      End
      Begin MSComDlg.CommonDialog CdExcel 
         Left            =   1035
         Top             =   630
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
      Begin CATControls.CATTextBox TxtGls_Excel 
         Height          =   330
         Left            =   1710
         TabIndex        =   2
         Top             =   270
         Width           =   5640
         _ExtentX        =   9948
         _ExtentY        =   582
         BackColor       =   16777215
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial Narrow"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         FontName        =   "Arial Narrow"
         FontSize        =   9.75
         ForeColor       =   -2147483640
         Container       =   "FrmImportaVales.frx":038A
      End
      Begin CATControls.CATTextBox Txt_GlsObservacion 
         Height          =   330
         Left            =   1710
         TabIndex        =   9
         Top             =   1170
         Width           =   5640
         _ExtentX        =   9948
         _ExtentY        =   582
         BackColor       =   16777215
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial Narrow"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         FontName        =   "Arial Narrow"
         FontSize        =   9.75
         ForeColor       =   -2147483640
         Container       =   "FrmImportaVales.frx":03A6
      End
      Begin VB.Label Label2 
         Caption         =   "Observación"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   180
         TabIndex        =   10
         Top             =   1245
         Width           =   1515
      End
      Begin VB.Label lblMotivo 
         Caption         =   "Concepto"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   180
         TabIndex        =   6
         Top             =   765
         Width           =   960
      End
      Begin VB.Label Label1 
         Caption         =   "Archivo Excel Vales"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   180
         TabIndex        =   3
         Top             =   345
         Width           =   1515
      End
   End
End
Attribute VB_Name = "FrmImportaVales"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmbAyudaExcel_Click()
On Error GoTo Err
Dim StrMsgError                     As String

    CdExcel.Filter = "Microsoft Excel (*.xls)|*.xls"
    CdExcel.ShowOpen
    TxtGls_Excel.Text = CdExcel.FileName
    
    Exit Sub
    
Err:
    If StrMsgError = "" Then StrMsgError = Err.Description
    MsgBox StrMsgError, vbInformation, App.Title
End Sub

Private Sub cmdaceptar_Click()
On Error GoTo Err
Dim StrMsgError                     As String

    ImportaValesXL TxtGls_Excel.Text, "", StrMsgError
    If StrMsgError <> "" Then GoTo Err
    
    MsgBox "Fin de Proceso", vbInformation, App.Title
    
    Exit Sub

Err:
    If StrMsgError = "" Then StrMsgError = Err.Description
    MsgBox StrMsgError, vbInformation, App.Title
End Sub

Private Sub ImportaValesXL(PArchivo As String, Msg1 As String, StrMsgError As String)
Dim xl                              As New Excel.Application
Dim wb                              As Workbook
Dim shRxn                           As Worksheet
Dim NFil                            As Integer
Dim cselect                         As String
Dim CCodValeTempAux                 As String
Dim CCodValeTemp                    As String
Dim rstemp                          As New ADODB.Recordset
Dim CIdValesCab                     As String
On Error GoTo Err
    
    MousePointer = 13
    
    Set xl = New Excel.Application
    Set wb = xl.Workbooks.Open(PArchivo)

    xl.Cells.Select
    With xl.Selection
        .VerticalAlignment = xlTop
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
    
    wb.Worksheets("RESUMEN").Select
    NFil = 2
    
    xl.Visible = False
    StrMsgError = ""

    cselect = "Delete From TempImportaVales"
    Cn.Execute cselect
    
    CCodValeTempAux = generaCorrelativo("ValesCab", "IdValeTemp", 8, , True)
    
    rstemp.Fields.Append "IdAlmacen", adVarChar, 8, adFldIsNullable
    rstemp.Fields.Append "Fecha", adVarChar, 10, adFldIsNullable
    rstemp.Fields.Append "IdValeTemp", adVarChar, 8, adFldIsNullable
    rstemp.Fields.Append "IdCentroCosto", adVarChar, 8, adFldIsNullable
    rstemp.Open
    
    cselect = ""
    
    Do While Len(xl.Cells(NFil, 1).Value) > 0
        
        If rstemp.RecordCount > 0 Then
            rstemp.MoveFirst
            rstemp.Filter = "IdAlmacen = '" & xl.Cells(NFil, 1) & "' And Fecha = '" & xl.Cells(NFil, 3) & "' And IdCentroCosto = '" & xl.Cells(NFil, 2) & "'"
            If Not rstemp.EOF Then
                CCodValeTemp = rstemp.Fields("IdValeTemp")
            End If
            rstemp.Filter = ""
            rstemp.Filter = adFilterNone
        End If
        
        If Len(Trim(CCodValeTemp)) = 0 Then
            CCodValeTemp = CCodValeTempAux
            CCodValeTempAux = Format(CCodValeTempAux + 1, "00000000")
            
            rstemp.AddNew
            rstemp.Fields("IdAlmacen") = xl.Cells(NFil, 1)
            rstemp.Fields("Fecha") = "" & xl.Cells(NFil, 3)
            rstemp.Fields("IdValeTemp") = CCodValeTemp
            rstemp.Fields("IdCentroCosto") = xl.Cells(NFil, 2)
        End If
        
        If Val(xl.Cells(NFil, 5)) > 0 Then
            If Len(Trim("" & traerCampo("Productos", "IdProducto", "IdProducto", xl.Cells(NFil, 4), True))) > 0 Then
                cselect = cselect & "('" & xl.Cells(NFil, 1) & "','" & xl.Cells(NFil, 2) & "','" & Format(xl.Cells(NFil, 3), "yyyy-mm-dd") & "'," & _
                          "'" & xl.Cells(NFil, 4) & "'," & xl.Cells(NFil, 5) & "," & xl.Cells(NFil, 6) & ",'" & CCodValeTemp & "'),"
            End If
        End If
        
        NFil = NFil + 1
        CCodValeTemp = ""
        
    Loop
    rstemp.Close: Set rstemp = Nothing
    
    cselect = "Insert Into TempImportaVales(IdAlmacen,IdCentroCosto,Fecha,IdProducto,Kilos,Unidades,IdValeTemp)Values" & left(cselect, Len(cselect) - 1)
    Cn.Execute (cselect)
    
    CIdValesCab = generaCorrelativoAnoMes_Vale("ValesCab", "IdValesCab", "S", True)
    
    cselect = "Insert Into ValesCab(IdValesCab,TipoVale,FechaEmision,IdConcepto,IdAlmacen,ObsValesCab,IdMoneda,TipoCambio,IdEmpresa,IdSucursal,EstValeCab," & _
                "IdPeriodoInv,IdCentroCosto,IdValeTemp,FechaRegistro,IdUsuarioRegistro)" & _
                "Select (@i:=@i +1),A.*,SysDate(),'" & glsUser & "' " & _
                "From (Select @i:='" & CIdValesCab & "' - 1) Foo,(" & _
                    "Select 'S',A.Fecha,'" & txtCod_Concepto.Text & "',A.IdAlmacen,'" & Txt_GlsObservacion.Text & "','PEN',IfNull(B.TcVenta,0)," & _
                    "'" & glsEmpresa & "','" & glsSucursal & "','GEN','" & glsCodPeriodoINV & "',A.IdCentroCosto,A.IdValeTemp " & _
                    "From TempImportaVales A " & _
                    "Left Join TiposDeCambio B " & _
                        "On A.Fecha = B.Fecha " & _
                    "Group By A.IdValeTemp" & _
                ") A"
    Cn.Execute (cselect)
    
    cselect = "Insert Into ValesDet(IdValesCab,Item,IdProducto,GlsProducto,IdUM,Factor,Afecto,Cantidad,IdMoneda,IdEmpresa,IdSucursal,Cantidad2," & _
                "IdSucursalOrigen,TipoVale)" & _
                "Select A.IdValesCab,A.Item,A.IdProducto,A.GlsProducto,A.IdUMVenta,1,A.AfectoIgv,A.Kilos,'PEN','" & glsEmpresa & "'," & _
                "'" & glsSucursal & "',A.Unidades,'" & glsSucursal & "','S' " & _
                "From(" & _
                "Select A.IdValesCab,if(@Previo <> A.IdValesCab, @i:=1, (@i:=@i +1)) As Item," & _
                "If(@Previo <> A.IdValesCab,@Previo:=A.IdValesCab, @Previo:=@Previo ) As Ant,B.IdProducto,C.GlsProducto,C.IdUMVenta,C.AfectoIgv," & _
                "B.Kilos,B.Unidades " & _
                "From (Select @i:=0, @Previo:='') Foo,ValesCab A " & _
                "Inner Join TempImportaVales B " & _
                "On A.IdValeTemp = B.IdValeTemp " & _
                "Inner Join Productos C " & _
                "On A.IdEmpresa = C.IdEmpresa And B.IdProducto = C.IdProducto " & _
                "Where A.IdEmpresa = '" & glsEmpresa & "' And A.TipoVale = 'S' " & _
                "Order By A.IdValeTemp,B.IdProducto" & _
                ") A"
    Cn.Execute (cselect)
    
    cselect = "Select A.IdAlmacen,A.IdValesCab " & _
                "From ValesCab A " & _
                "Inner Join TempImportaVales B " & _
                    "On A.IdValeTemp = B.IdValeTemp " & _
                "Where A.IdEmpresa = '" & glsEmpresa & "' And A.TipoVale = 'S' " & _
                "Group By A.IdValesCab"
    rstemp.Open cselect, Cn, adOpenStatic, adLockReadOnly
    
    Do While Not rstemp.EOF
        'actualizaStock "" & rstemp.Fields("IdValesCab"), 0, StrMsgError, "S", False
        Actualiza_Stock_Nuevo StrMsgError, "I", glsSucursal, "S", Trim("" & rstemp.Fields("IdValesCab")), Trim("" & rstemp.Fields("IdAlmacen"))
        If StrMsgError <> "" Then GoTo Err
        rstemp.MoveNext
    Loop
    rstemp.Close: Set rstemp = Nothing
    
    MousePointer = 1
    Clipboard.Clear
    xl.ActiveWorkbook.Close False, False, False
    xl.Quit
    
    Exit Sub
Err:
    If StrMsgError = "" Then StrMsgError = Err.Description
End Sub

Private Sub cmdsalir_Click()
    
    Unload Me
    
End Sub

Private Sub Form_Load()
On Error GoTo Err
Dim StrMsgError As String
    
    Me.top = 0
    Me.left = 0
    txtCod_Concepto.Text = traerCampo("Parametros", "ValParametro", "GlsParametro", "CONCEPTO_CONSUMO_SALIDA", True)
    
    Exit Sub

Err:
    If StrMsgError = "" Then StrMsgError = Err.Description
    MsgBox StrMsgError, vbInformation, App.Title
End Sub

Private Sub txtCod_Concepto_Change()
On Error GoTo Err
Dim StrMsgError                         As String
    
    txtGls_Concepto.Text = traerCampo("Conceptos", "GlsConcepto", "IdConcepto", txtCod_Concepto.Text, False)
    
    Exit Sub

Err:
    If StrMsgError = "" Then StrMsgError = Err.Description
    MsgBox StrMsgError, vbInformation, App.Title
End Sub
