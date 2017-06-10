VERSION 5.00
Object = "{FB992564-9055-42B5-B433-FEA84CEA93C4}#11.0#0"; "crviewer.dll"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmVendasPublico 
   Caption         =   "Relatório de Vendas Total Publico"
   ClientHeight    =   7995
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   12600
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   7995
   ScaleWidth      =   12600
   WindowState     =   2  'Maximized
   Begin VB.CommandButton cmdCinema 
      Height          =   330
      Left            =   6360
      Picture         =   "frmVendasPublico.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   14
      TabStop         =   0   'False
      ToolTipText     =   "Pesquisa"
      Top             =   690
      Width           =   330
   End
   Begin VB.ComboBox cboCinema 
      Height          =   315
      Left            =   1620
      TabIndex        =   13
      Top             =   690
      Width           =   4725
   End
   Begin VB.CommandButton cmdEmpr 
      Height          =   330
      Left            =   6360
      Picture         =   "frmVendasPublico.frx":038A
      Style           =   1  'Graphical
      TabIndex        =   12
      TabStop         =   0   'False
      ToolTipText     =   "Pesquisa"
      Top             =   360
      Width           =   330
   End
   Begin VB.ComboBox cboEmpr 
      Height          =   315
      Left            =   1620
      TabIndex        =   11
      Top             =   360
      Width           =   4725
   End
   Begin VB.CommandButton cmdPerioido 
      Height          =   330
      Left            =   6360
      Picture         =   "frmVendasPublico.frx":0714
      Style           =   1  'Graphical
      TabIndex        =   10
      TabStop         =   0   'False
      ToolTipText     =   "Pesquisa"
      Top             =   30
      Width           =   330
   End
   Begin VB.CommandButton cmdCancela 
      Cancel          =   -1  'True
      Caption         =   "Cancela"
      Height          =   420
      Left            =   10815
      TabIndex        =   4
      Top             =   570
      Width           =   1650
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "Ok"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   10800
      TabIndex        =   3
      Top             =   105
      Width           =   1650
   End
   Begin CrystalActiveXReportViewerLib11Ctl.CrystalActiveXReportViewer CRVBordero 
      Height          =   6885
      Left            =   30
      TabIndex        =   5
      Top             =   1080
      Width           =   12555
      _cx             =   22146
      _cy             =   12144
      DisplayGroupTree=   -1  'True
      DisplayToolbar  =   -1  'True
      EnableGroupTree =   -1  'True
      EnableNavigationControls=   -1  'True
      EnableStopButton=   -1  'True
      EnablePrintButton=   -1  'True
      EnableZoomControl=   -1  'True
      EnableCloseButton=   -1  'True
      EnableProgressControl=   0   'False
      EnableSearchControl=   -1  'True
      EnableRefreshButton=   0   'False
      EnableDrillDown =   -1  'True
      EnableAnimationControl=   0   'False
      EnableSelectExpertButton=   0   'False
      EnableToolbar   =   -1  'True
      DisplayBorder   =   -1  'True
      DisplayTabs     =   -1  'True
      DisplayBackgroundEdge=   -1  'True
      SelectionFormula=   ""
      EnablePopupMenu =   -1  'True
      EnableExportButton=   -1  'True
      EnableSearchExpertButton=   0   'False
      EnableHelpButton=   0   'False
      LaunchHTTPHyperlinksInNewBrowser=   -1  'True
      EnableLogonPrompts=   -1  'True
      LocaleID        =   1046
   End
   Begin MSComCtl2.DTPicker dtpDtIni 
      Height          =   315
      Left            =   1965
      TabIndex        =   6
      Top             =   15
      Width           =   1395
      _ExtentX        =   2461
      _ExtentY        =   556
      _Version        =   393216
      Format          =   59965441
      CurrentDate     =   38606
   End
   Begin MSComCtl2.DTPicker dtpDtFim 
      Height          =   315
      Left            =   3660
      TabIndex        =   7
      Top             =   15
      Width           =   1395
      _ExtentX        =   2461
      _ExtentY        =   556
      _Version        =   393216
      Format          =   59965441
      CurrentDate     =   38606
   End
   Begin VB.Label lblA 
      AutoSize        =   -1  'True
      Caption         =   "a"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   3450
      TabIndex        =   9
      Top             =   45
      Width           =   120
   End
   Begin VB.Label lblDe 
      AutoSize        =   -1  'True
      Caption         =   "De"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   1635
      TabIndex        =   8
      Top             =   45
      Width           =   255
   End
   Begin VB.Label lblEmpr 
      AutoSize        =   -1  'True
      Caption         =   "Empresa:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   90
      TabIndex        =   2
      Top             =   375
      Width           =   795
   End
   Begin VB.Label lblDtMovto 
      AutoSize        =   -1  'True
      Caption         =   "Data Movimento:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   90
      TabIndex        =   1
      Top             =   45
      Width           =   1455
   End
   Begin VB.Label lblCinema 
      AutoSize        =   -1  'True
      Caption         =   "Cinema:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   90
      TabIndex        =   0
      Top             =   690
      Width           =   690
   End
End
Attribute VB_Name = "frmVendasPublico"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim tbEmpr    As New ADODB.Recordset
Dim tbCinema  As New ADODB.Recordset

Dim carregando As Boolean
    
Private Sub cboEmpr_Click()
    Call carregaCinema
End Sub

Private Sub cmdCancela_Click()
    Unload Me
End Sub

Private Sub cmdCinema_Click()
    Call carregaCinema
End Sub

Private Sub cmdEmpr_Click()
     Call carregaEmpr
End Sub

Private Sub cmdOK_Click()
    Dim sqlStr   As String
    Dim oRs      As New ADODB.Recordset
    Dim m_Report As New VendasTotalPublico
    
    sqlStr = "SELECT * FROM vw_rel_tot_publico "
    sqlStr = sqlStr & "WHERE vw_rel_tot_publico.sre_data between CONVERT(DATETIME, " & Format(dtpDtIni.Value, "'dd/mm/yyyy'") & ", 103) and CONVERT(DATETIME, " & Format(dtpDtFim.Value, "'dd/mm/yyyy'") & ", 103) "
    
    If cboEmpr.ListIndex <> -1 Then
        If cboEmpr.ItemData(cboEmpr.ListIndex) <> 0 Then
            sqlStr = sqlStr & "AND   vw_rel_tot_publico.emp_cd     = " & cboEmpr.ItemData(cboEmpr.ListIndex) & " "
        End If
    End If

    If cboCinema.ListIndex <> -1 Then
        If cboCinema.ItemData(cboCinema.ListIndex) <> 0 Then
            sqlStr = sqlStr & "AND   vw_rel_tot_publico.cin_cd     = " & cboCinema.ItemData(cboCinema.ListIndex) & " "
        End If
    End If
    
    sqlStr = sqlStr & "ORDER BY vw_rel_tot_publico.emp_nm, "
    sqlStr = sqlStr & "vw_rel_tot_publico.cin_nm, "
    sqlStr = sqlStr & "vw_rel_tot_publico.sre_data "
    
    Set oRs = dbConnect.Execute(sqlStr)
    
    m_Report.Database.SetDataSource oRs
    Call m_Report.perTitulo.SetText("Data Exibição: de " & Format(dtpDtIni.Value, "dd/mm/yyyy") & " ate " & Format(dtpDtFim.Value, "dd/mm/yyyy"))
    
    CRVBordero.ReportSource = m_Report
    CRVBordero.ViewReport
End Sub

Private Sub cmdPerioido_Click()
     Call carregaEmpr
End Sub

Private Sub dtpDtFim_Change()
     If carregando Then
        Exit Sub
     End If
     
     Call carregaEmpr
End Sub

Private Sub dtpDtIni_Change()
     If carregando Then
        Exit Sub
     End If
     
     Call carregaEmpr

End Sub

Private Sub Form_Load()
    carregando = True
    
    dtpDtIni.Value = Date
    dtpDtFim.Value = Date
    
    carregando = False
End Sub

Private Sub Form_Resize()
    CRVBordero.Left = 0
    CRVBordero.Width = Me.Width - 165
    CRVBordero.Height = Me.Height - 2400 - 165
    
    lblDtMovto.Left = (Me.Width - 6570 - 165) / 2
    lblEmpr.Left = lblDtMovto.Left
    lblCinema.Left = lblDtMovto.Left
    
    lblDe.Left = lblDtMovto.Left + 1545
    dtpDtIni.Left = lblDtMovto.Left + 1875
    lblA.Left = lblDtMovto.Left + 3360
    dtpDtFim.Left = lblDtMovto.Left + 3570
    cmdPerioido.Left = lblDtMovto.Left + 6240
    
    cboEmpr.Left = lblDtMovto.Left + 1515
    cboCinema.Left = cboEmpr.Left
    
    cmdEmpr.Left = cboEmpr.Left + cboEmpr.Width
    cmdCinema.Left = cboCinema.Left + cboCinema.Width
    
    cmdOK.Left = cmdEmpr.Left + cmdEmpr.Width + 30
    cmdCancela.Left = cmdOK.Left

End Sub

Private Sub carregaEmpr()
    Dim strSql As String
    
    cboEmpr.Clear
    
    If tbEmpr.State = adStateOpen Then
        tbEmpr.Close
    End If
    
    strSql = "SELECT DISTINCT emp_cd, emp_nm FROM tb_bol_empr "
    strSql = strSql & "WHERE bol_dt_mov between CONVERT(DATETIME, " & Format(dtpDtIni.Value, "'dd/mm/yyyy'") & ", 103) AND CONVERT(DATETIME, " & Format(dtpDtFim.Value, "'dd/mm/yyyy'") & ", 103) "
    strSql = strSql & "ORDER BY emp_nm"
    
    tbEmpr.Open strSql, dbConnect, adOpenDynamic
    
    cboEmpr.AddItem "Todas", 0
    cboEmpr.ItemData(cboEmpr.NewIndex) = 0
    
    Do While Not tbEmpr.EOF
        cboEmpr.AddItem tbEmpr.Fields("emp_nm").Value
        cboEmpr.ItemData(cboEmpr.NewIndex) = tbEmpr.Fields("emp_cd").Value
        
        tbEmpr.MoveNext
    Loop
    
    cboEmpr.ListIndex = 0
End Sub

Private Sub carregaCinema()
    Dim strSql As String
    
    cboCinema.Clear
    
    If tbCinema.State = adStateOpen Then
        tbCinema.Close
    End If
    
    strSql = "SELECT DISTINCT emp_cd, cin_cd, cin_nm FROM tb_bol_cin "
    strSql = strSql & "WHERE bol_dt_mov between CONVERT(DATETIME, " & Format(dtpDtIni.Value, "'dd/mm/yyyy'") & ", 103) AND CONVERT(DATETIME, " & Format(dtpDtFim.Value, "'dd/mm/yyyy'") & ", 103) "
    
    If cboEmpr.ListIndex <> -1 Then
        If cboEmpr.ItemData(cboEmpr.ListIndex) <> 0 Then
            strSql = strSql & " AND emp_cd = " & cboEmpr.ItemData(cboEmpr.ListIndex)
        End If
    End If
    
    strSql = strSql & "ORDER BY cin_nm"
    
    tbCinema.Open strSql, dbConnect, adOpenDynamic
    
    cboCinema.AddItem "Todos", 0
    cboCinema.ItemData(cboCinema.NewIndex) = 0
    
    Do While Not tbCinema.EOF
        cboCinema.AddItem tbCinema.Fields("cin_nm").Value
        cboCinema.ItemData(cboCinema.NewIndex) = tbCinema.Fields("cin_cd").Value
        
        tbCinema.MoveNext
    Loop
    
    cboCinema.ListIndex = 0
End Sub

