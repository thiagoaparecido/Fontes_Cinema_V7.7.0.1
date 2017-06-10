VERSION 5.00
Object = "{FB992564-9055-42B5-B433-FEA84CEA93C4}#11.0#0"; "crviewer.dll"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmVendasFilme 
   Caption         =   "Relatório de Vendas por Filme"
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
   Begin VB.CommandButton cmdFilme 
      Height          =   330
      Left            =   6360
      Picture         =   "frmVendasFilme.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   20
      TabStop         =   0   'False
      ToolTipText     =   "Pesquisa"
      Top             =   1350
      Width           =   330
   End
   Begin VB.ComboBox cboFilme 
      Height          =   315
      Left            =   1620
      TabIndex        =   19
      Top             =   1350
      Width           =   4725
   End
   Begin VB.CommandButton cmdCinema 
      Height          =   330
      Left            =   6360
      Picture         =   "frmVendasFilme.frx":038A
      Style           =   1  'Graphical
      TabIndex        =   18
      TabStop         =   0   'False
      ToolTipText     =   "Pesquisa"
      Top             =   1005
      Width           =   330
   End
   Begin VB.ComboBox cboCinema 
      Height          =   315
      Left            =   1620
      TabIndex        =   17
      Top             =   1005
      Width           =   4725
   End
   Begin VB.CommandButton cmdEmpr 
      Height          =   330
      Left            =   6360
      Picture         =   "frmVendasFilme.frx":0714
      Style           =   1  'Graphical
      TabIndex        =   16
      TabStop         =   0   'False
      ToolTipText     =   "Pesquisa"
      Top             =   675
      Width           =   330
   End
   Begin VB.ComboBox cboEmpr 
      Height          =   315
      Left            =   1620
      TabIndex        =   15
      Top             =   675
      Width           =   4725
   End
   Begin VB.CommandButton cmdDistrib 
      Height          =   330
      Left            =   6360
      Picture         =   "frmVendasFilme.frx":0A9E
      Style           =   1  'Graphical
      TabIndex        =   14
      TabStop         =   0   'False
      ToolTipText     =   "Pesquisa"
      Top             =   345
      Width           =   330
   End
   Begin VB.ComboBox cboDistrib 
      Height          =   315
      Left            =   1620
      TabIndex        =   13
      Top             =   345
      Width           =   4725
   End
   Begin VB.CommandButton cmdPerioido 
      Height          =   330
      Left            =   6360
      Picture         =   "frmVendasFilme.frx":0E28
      Style           =   1  'Graphical
      TabIndex        =   12
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
      TabIndex        =   5
      Top             =   975
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
      TabIndex        =   4
      Top             =   465
      Width           =   1650
   End
   Begin CrystalActiveXReportViewerLib11Ctl.CrystalActiveXReportViewer CRVBordero 
      Height          =   6225
      Left            =   30
      TabIndex        =   6
      Top             =   1710
      Width           =   12555
      _cx             =   22146
      _cy             =   10980
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
      TabIndex        =   8
      Top             =   15
      Width           =   1395
      _ExtentX        =   2461
      _ExtentY        =   556
      _Version        =   393216
      Format          =   16515073
      CurrentDate     =   38606
   End
   Begin MSComCtl2.DTPicker dtpDtFim 
      Height          =   315
      Left            =   3660
      TabIndex        =   9
      Top             =   15
      Width           =   1395
      _ExtentX        =   2461
      _ExtentY        =   556
      _Version        =   393216
      Format          =   16515073
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
      TabIndex        =   11
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
      TabIndex        =   10
      Top             =   45
      Width           =   255
   End
   Begin VB.Label lblDistrib 
      AutoSize        =   -1  'True
      Caption         =   "Distribuidora:"
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
      TabIndex        =   7
      Top             =   360
      Width           =   1140
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
      TabIndex        =   3
      Top             =   690
      Width           =   795
   End
   Begin VB.Label lblFilme 
      AutoSize        =   -1  'True
      Caption         =   "Filme:"
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
      Top             =   1320
      Width           =   510
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
      Top             =   1005
      Width           =   690
   End
End
Attribute VB_Name = "frmVendasFilme"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim tbDistrib As New ADODB.Recordset
Dim tbEmpr    As New ADODB.Recordset
Dim tbCinema  As New ADODB.Recordset
Dim tbFilme   As New ADODB.Recordset

Dim carregando As Boolean
    
Private Sub cboCinema_Click()
    Call carregaFilme
End Sub

Private Sub cboDistrib_Click()
     Call carregaFilme
End Sub

Private Sub cboEmpr_Click()
    Call carregaCinema
End Sub

Private Sub cmdCancela_Click()
    Unload Me
End Sub

Private Sub cmdCinema_Click()
    Call carregaCinema
End Sub

Private Sub cmdDistrib_Click()
     Call carregaDistrib
End Sub

Private Sub cmdEmpr_Click()
     Call carregaEmpr
End Sub

Private Sub cmdFilme_Click()
    Call carregaFilme
End Sub

Private Sub cmdOK_Click()
    Dim sqlStr   As String
    Dim oRs      As New ADODB.Recordset
    Dim m_Report As New VendasFilme
        
    sqlStr = ""
    sqlStr = sqlStr & "SELECT tb_bol_empr.emp_cd          AS emp_cd, "
    sqlStr = sqlStr & "       tb_bol_empr.emp_nm          AS emp_nm, "
    sqlStr = sqlStr & "       tb_bol_cin.cin_cd           AS cin_cd, "
    sqlStr = sqlStr & "       tb_bol_cin.cin_nm           AS cin_nm, "
    sqlStr = sqlStr & "       tb_bol_filme.fil_cd         AS fil_cd, "
    sqlStr = sqlStr & "       tb_bol_filme.fil_nm         AS fil_nm, "
    sqlStr = sqlStr & "       tb_bol_distrib.dis_cd       AS dis_cd, "
    sqlStr = sqlStr & "       tb_bol_distrib.dis_nm       AS dis_nm, "
    sqlStr = sqlStr & "       SUM(tb_bol_ingre.bin_qtde)  AS publico, "
    sqlStr = sqlStr & "       SUM(tb_bol_ingre.ing_valor * tb_bol_ingre.bin_qtde) AS renda, "
    sqlStr = sqlStr & "       SUM(CASE "
    sqlStr = sqlStr & "              WHEN tb_bol_ingre.igt_cd = 9 THEN tb_bol_ingre.bin_qtde "
    sqlStr = sqlStr & "              ELSE 0 "
    sqlStr = sqlStr & "           END)                    AS cortesias, "
    sqlStr = sqlStr & "       SUM(CASE "
    sqlStr = sqlStr & "              WHEN tb_bol_ingre.igt_cd = 1  "
    sqlStr = sqlStr & "              OR   tb_bol_ingre.igt_cd = 3 THEN tb_bol_ingre.bin_qtde "
    sqlStr = sqlStr & "              ELSE 0 "
    sqlStr = sqlStr & "           END)                    AS interias, "
    sqlStr = sqlStr & "       SUM(CASE "
    sqlStr = sqlStr & "              WHEN tb_bol_ingre.igt_cd = 2  "
    sqlStr = sqlStr & "              OR   tb_bol_ingre.igt_cd = 4 THEN tb_bol_ingre.bin_qtde "
    sqlStr = sqlStr & "              ELSE 0 "
    sqlStr = sqlStr & "           END)                    AS meias, "
    sqlStr = sqlStr & "       CASE "
    sqlStr = sqlStr & "          WHEN DATEPART(WEEKDAY, tb_bol_filme.fil_dt_ini) < 6  THEN "
    sqlStr = sqlStr & "             DATEDIFF(WEEK, DATEADD(DAY, -1 * (DATEPART(WEEKDAY, tb_bol_filme.fil_dt_ini) - 1) - 2, tb_bol_filme.fil_dt_ini), tb_bol_ingre.sre_data) "
    sqlStr = sqlStr & "          ELSE "
    sqlStr = sqlStr & "             DATEDIFF(WEEK, DATEADD(DAY, 6 - DATEPART(WEEKDAY, tb_bol_filme.fil_dt_ini), tb_bol_filme.fil_dt_ini), tb_bol_ingre.sre_data) "
    sqlStr = sqlStr & "       END                        AS semana "
    sqlStr = sqlStr & "FROM tb_boletim, "
    sqlStr = sqlStr & "     tb_bol_empr,  "
    sqlStr = sqlStr & "     tb_bol_cin, "
    sqlStr = sqlStr & "     tb_bol_filme, "
    sqlStr = sqlStr & "     tb_bol_distrib, "
    sqlStr = sqlStr & "     tb_bol_ingre "
    sqlStr = sqlStr & "WHERE tb_boletim.bol_dt_mov   = tb_bol_empr.bol_dt_mov "
    sqlStr = sqlStr & "AND   tb_boletim.emp_cd       = tb_bol_empr.emp_cd "
    sqlStr = sqlStr & "AND   tb_boletim.bol_dt_mov   = tb_bol_cin.bol_dt_mov "
    sqlStr = sqlStr & "AND   tb_boletim.emp_cd       = tb_bol_cin.emp_cd "
    sqlStr = sqlStr & "AND   tb_boletim.cin_cd       = tb_bol_cin.cin_cd "
    sqlStr = sqlStr & "AND   tb_boletim.bol_dt_mov   = tb_bol_filme.bol_dt_mov "
    sqlStr = sqlStr & "AND   tb_boletim.emp_cd       = tb_bol_filme.emp_cd "
    sqlStr = sqlStr & "AND   tb_boletim.cin_cd       = tb_bol_filme.cin_cd "
    sqlStr = sqlStr & "AND   tb_boletim.fil_cd       = tb_bol_filme.fil_cd "
    sqlStr = sqlStr & "AND   tb_boletim.bol_dt_mov   = tb_bol_ingre.bol_dt_mov "
    sqlStr = sqlStr & "AND   tb_boletim.emp_cd       = tb_bol_ingre.emp_cd "
    sqlStr = sqlStr & "AND   tb_boletim.cin_cd       = tb_bol_ingre.cin_cd "
    sqlStr = sqlStr & "AND   tb_boletim.sal_cd       = tb_bol_ingre.sal_cd "
    sqlStr = sqlStr & "AND   tb_boletim.fil_cd       = tb_bol_ingre.fil_cd "
    sqlStr = sqlStr & "AND   tb_bol_filme.bol_dt_mov = tb_bol_distrib.bol_dt_mov "
    sqlStr = sqlStr & "AND   tb_bol_filme.dis_cd     = tb_bol_distrib.dis_cd "
    sqlStr = sqlStr & "AND   tb_bol_ingre.bin_dev    = 'N' "
    sqlStr = sqlStr & "AND   tb_bol_ingre.sre_data   = tb_bol_ingre.bol_dt_mov "
    sqlStr = sqlStr & "AND   tb_bol_ingre.sre_data between CONVERT(DATETIME, " & Format(dtpDtIni.Value, "'dd/mm/yyyy'") & ", 103) and CONVERT(DATETIME, " & Format(dtpDtFim.Value, "'dd/mm/yyyy'") & ", 103) "
    
    If cboDistrib.ListIndex <> -1 Then
        If cboDistrib.ItemData(cboDistrib.ListIndex) <> 0 Then
            sqlStr = sqlStr & "AND   tb_bol_filme.dis_cd     = " & cboDistrib.ItemData(cboDistrib.ListIndex) & " "
        End If
    End If
    
    If cboEmpr.ListIndex <> -1 Then
        If cboEmpr.ItemData(cboEmpr.ListIndex) <> 0 Then
            sqlStr = sqlStr & "AND   tb_bol_ingre.emp_cd     = " & cboEmpr.ItemData(cboEmpr.ListIndex) & " "
        End If
    End If
    
    If cboCinema.ListIndex <> -1 Then
        If cboCinema.ItemData(cboCinema.ListIndex) <> 0 Then
            sqlStr = sqlStr & "AND   tb_bol_ingre.cin_cd     = " & cboCinema.ItemData(cboCinema.ListIndex) & " "
        End If
    End If
    
    If cboFilme.ListIndex <> -1 Then
        If cboFilme.ItemData(cboFilme.ListIndex) <> 0 Then
            sqlStr = sqlStr & "AND   tb_bol_ingre.fil_cd     = " & cboFilme.ItemData(cboFilme.ListIndex) & " "
        End If
    End If
    
    sqlStr = sqlStr & "GROUP BY tb_bol_empr.emp_cd, "
    sqlStr = sqlStr & "         tb_bol_empr.emp_nm, "
    sqlStr = sqlStr & "         tb_bol_cin.cin_cd, "
    sqlStr = sqlStr & "         tb_bol_cin.cin_nm, "
    sqlStr = sqlStr & "         tb_bol_filme.fil_cd, "
    sqlStr = sqlStr & "         tb_bol_filme.fil_nm, "
    sqlStr = sqlStr & "         tb_bol_distrib.dis_cd, "
    sqlStr = sqlStr & "         tb_bol_distrib.dis_nm, "
    sqlStr = sqlStr & "         CASE "
    sqlStr = sqlStr & "            WHEN DATEPART(WEEKDAY, tb_bol_filme.fil_dt_ini) < 6  THEN "
    sqlStr = sqlStr & "               DATEDIFF(WEEK, DATEADD(DAY, -1 * (DATEPART(WEEKDAY, tb_bol_filme.fil_dt_ini) - 1) - 2, tb_bol_filme.fil_dt_ini), tb_bol_ingre.sre_data) "
    sqlStr = sqlStr & "            ELSE "
    sqlStr = sqlStr & "               DATEDIFF(WEEK, DATEADD(DAY, 6 - DATEPART(WEEKDAY, tb_bol_filme.fil_dt_ini), tb_bol_filme.fil_dt_ini), tb_bol_ingre.sre_data) "
    sqlStr = sqlStr & "         END "
    sqlStr = sqlStr & "ORDER BY tb_bol_empr.emp_nm, "
    sqlStr = sqlStr & "         tb_bol_filme.fil_nm, "
    sqlStr = sqlStr & "         tb_bol_cin.cin_nm "
    
    Set oRs = dbConnect.Execute(sqlStr)
    
    m_Report.Database.SetDataSource oRs
    Call m_Report.perTitulo.SetText("Data Exibição: de " & Format(dtpDtIni.Value, "dd/mm/yyyy") & " ate " & Format(dtpDtFim.Value, "dd/mm/yyyy"))
    
    CRVBordero.ReportSource = m_Report
    CRVBordero.ViewReport
End Sub

Private Sub cmdPerioido_Click()
     Call carregaDistrib
     Call carregaEmpr
End Sub

Private Sub dtpDtFim_Change()
     If carregando Then
        Exit Sub
     End If
     
     Call carregaDistrib
     Call carregaEmpr
End Sub

Private Sub dtpDtIni_Change()
     If carregando Then
        Exit Sub
     End If
     
     Call carregaDistrib
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
    lblDistrib.Left = lblDtMovto.Left
    lblEmpr.Left = lblDtMovto.Left
    lblCinema.Left = lblDtMovto.Left
    lblFilme.Left = lblDtMovto.Left
    
    lblDe.Left = lblDtMovto.Left + 1545
    dtpDtIni.Left = lblDtMovto.Left + 1875
    lblA.Left = lblDtMovto.Left + 3360
    dtpDtFim.Left = lblDtMovto.Left + 3570
    cmdPerioido.Left = lblDtMovto.Left + 6240
    
    cboDistrib.Left = lblDtMovto.Left + 1515
    cboEmpr.Left = cboDistrib.Left
    cboCinema.Left = cboDistrib.Left
    cboFilme.Left = cboDistrib.Left
    
    cmdDistrib.Left = cboDistrib.Left + cboDistrib.Width
    cmdEmpr.Left = cboEmpr.Left + cboEmpr.Width
    cmdCinema.Left = cboCinema.Left + cboCinema.Width
    cmdFilme.Left = cboFilme.Left + cboFilme.Width
    
    cmdOK.Left = cmdDistrib.Left + cmdDistrib.Width + 30
    cmdCancela.Left = cmdOK.Left

End Sub

Private Sub carregaDistrib()
    Dim strSql As String
    
    cboDistrib.Clear
    
    If tbDistrib.State = adStateOpen Then
        tbDistrib.Close
    End If
    
    strSql = "SELECT DISTINCT dis_cd, dis_nm FROM tb_bol_distrib "
    strSql = strSql & "WHERE bol_dt_mov between CONVERT(DATETIME, " & Format(dtpDtIni.Value, "'dd/mm/yyyy'") & ", 103) AND CONVERT(DATETIME, " & Format(dtpDtFim.Value, "'dd/mm/yyyy'") & ", 103) "
    strSql = strSql & "ORDER BY dis_nm"
    
    tbDistrib.Open strSql, dbConnect, adOpenDynamic
    
    cboDistrib.AddItem "Todas", 0
    cboDistrib.ItemData(cboDistrib.NewIndex) = 0
    
    Do While Not tbDistrib.EOF
        cboDistrib.AddItem tbDistrib.Fields("dis_nm").Value
        cboDistrib.ItemData(cboDistrib.NewIndex) = tbDistrib.Fields("dis_cd").Value
        
        tbDistrib.MoveNext
    Loop
    
    cboDistrib.ListIndex = 0
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

Private Sub carregaFilme()
    Dim strSql As String
    
    cboFilme.Clear
    
    If tbFilme.State = adStateOpen Then
        tbFilme.Close
    End If
    
    strSql = "SELECT DISTINCT emp_cd, cin_cd, fil_cd, fil_nm FROM vw_bol_filme "
    strSql = strSql & "WHERE bol_dt_mov between CONVERT(DATETIME, " & Format(dtpDtIni.Value, "'dd/mm/yyyy'") & ", 103) AND CONVERT(DATETIME, " & Format(dtpDtFim.Value, "'dd/mm/yyyy'") & ", 103) "
    
    If cboEmpr.ListIndex <> -1 Then
        If cboEmpr.ItemData(cboEmpr.ListIndex) <> 0 Then
            strSql = strSql & " AND emp_cd = " & cboEmpr.ItemData(cboEmpr.ListIndex)
        End If
    End If
    
    If cboCinema.ListIndex <> -1 Then
        If cboCinema.ItemData(cboCinema.ListIndex) <> 0 Then
            strSql = strSql & " AND cin_cd = " & cboCinema.ItemData(cboCinema.ListIndex)
        End If
    End If
    
    If cboDistrib.ListIndex <> -1 Then
        If cboDistrib.ItemData(cboDistrib.ListIndex) <> 0 Then
            strSql = strSql & " AND dis_cd = " & cboDistrib.ItemData(cboDistrib.ListIndex)
        End If
    End If
    
    strSql = strSql & "ORDER BY fil_nm"
    
    tbFilme.Open strSql, dbConnect, adOpenDynamic
    
    cboFilme.AddItem "Todos", 0
    cboFilme.ItemData(cboFilme.NewIndex) = 0
    
    Do While Not tbFilme.EOF
        cboFilme.AddItem tbFilme.Fields("fil_nm").Value
        cboFilme.ItemData(cboFilme.NewIndex) = tbFilme.Fields("fil_cd").Value
        
        tbFilme.MoveNext
    Loop
    
    If cboFilme.ListCount > 0 Then
        cboFilme.ListIndex = 0
    End If
End Sub

