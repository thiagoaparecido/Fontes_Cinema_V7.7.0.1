VERSION 5.00
Object = "{FB992564-9055-42B5-B433-FEA84CEA93C4}#11.0#0"; "crviewer.dll"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmBordero 
   Caption         =   "Relatório de Borderô"
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
      Picture         =   "frmBordero.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   23
      TabStop         =   0   'False
      ToolTipText     =   "Pesquisa"
      Top             =   1665
      Width           =   330
   End
   Begin VB.ComboBox cboFilme 
      Height          =   315
      Left            =   1620
      TabIndex        =   22
      Top             =   1665
      Width           =   4725
   End
   Begin VB.CommandButton cmdSala 
      Height          =   330
      Left            =   6360
      Picture         =   "frmBordero.frx":038A
      Style           =   1  'Graphical
      TabIndex        =   21
      TabStop         =   0   'False
      ToolTipText     =   "Pesquisa"
      Top             =   1335
      Width           =   330
   End
   Begin VB.ComboBox cboSala 
      Height          =   315
      Left            =   1620
      TabIndex        =   20
      Top             =   1335
      Width           =   4725
   End
   Begin VB.CommandButton cmdCinema 
      Height          =   330
      Left            =   6360
      Picture         =   "frmBordero.frx":0714
      Style           =   1  'Graphical
      TabIndex        =   19
      TabStop         =   0   'False
      ToolTipText     =   "Pesquisa"
      Top             =   1005
      Width           =   330
   End
   Begin VB.ComboBox cboCinema 
      Height          =   315
      Left            =   1620
      TabIndex        =   18
      Top             =   1005
      Width           =   4725
   End
   Begin VB.CommandButton cmdEmpr 
      Height          =   330
      Left            =   6360
      Picture         =   "frmBordero.frx":0A9E
      Style           =   1  'Graphical
      TabIndex        =   17
      TabStop         =   0   'False
      ToolTipText     =   "Pesquisa"
      Top             =   675
      Width           =   330
   End
   Begin VB.ComboBox cboEmpr 
      Height          =   315
      Left            =   1620
      TabIndex        =   16
      Top             =   675
      Width           =   4725
   End
   Begin VB.CommandButton cmdDistrib 
      Height          =   330
      Left            =   6360
      Picture         =   "frmBordero.frx":0E28
      Style           =   1  'Graphical
      TabIndex        =   15
      TabStop         =   0   'False
      ToolTipText     =   "Pesquisa"
      Top             =   345
      Width           =   330
   End
   Begin VB.ComboBox cboDistrib 
      Height          =   315
      Left            =   1620
      TabIndex        =   14
      Top             =   345
      Width           =   4725
   End
   Begin VB.CommandButton cmdPerioido 
      Height          =   330
      Left            =   6360
      Picture         =   "frmBordero.frx":11B2
      Style           =   1  'Graphical
      TabIndex        =   13
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
      TabIndex        =   6
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
      TabIndex        =   5
      Top             =   465
      Width           =   1650
   End
   Begin CrystalActiveXReportViewerLib11Ctl.CrystalActiveXReportViewer CRVBordero 
      Height          =   5925
      Left            =   30
      TabIndex        =   7
      Top             =   2010
      Width           =   12555
      _cx             =   22146
      _cy             =   10451
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
      TabIndex        =   9
      Top             =   15
      Width           =   1395
      _ExtentX        =   2461
      _ExtentY        =   556
      _Version        =   393216
      Format          =   59834369
      CurrentDate     =   38606
   End
   Begin MSComCtl2.DTPicker dtpDtFim 
      Height          =   315
      Left            =   3660
      TabIndex        =   10
      Top             =   15
      Width           =   1395
      _ExtentX        =   2461
      _ExtentY        =   556
      _Version        =   393216
      Format          =   59834369
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
      TabIndex        =   12
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
      TabIndex        =   11
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
      TabIndex        =   8
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
      TabIndex        =   4
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
      TabIndex        =   3
      Top             =   1635
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
      TabIndex        =   2
      Top             =   45
      Width           =   1455
   End
   Begin VB.Label lblSala 
      AutoSize        =   -1  'True
      Caption         =   "Sala:"
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
      Top             =   1320
      Width           =   450
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
Attribute VB_Name = "frmBordero"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim tbDistrib As New ADODB.Recordset
Dim tbEmpr    As New ADODB.Recordset
Dim tbCinema  As New ADODB.Recordset
Dim tbSala    As New ADODB.Recordset
Dim tbFilme   As New ADODB.Recordset

Dim carregando As Boolean
    
Private Sub cboCinema_Click()

    Call carregaSala
        
End Sub

Private Sub cboDistrib_Click()
     
     Call carregaFilme

End Sub

Private Sub cboEmpr_Click()
     
    Call carregaCinema
     
End Sub

Private Sub cboSala_Click()
    
    Call carregaFilme
    
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
    Dim m_Report As New Bordero1
    
'    If cboDtMovto.codigo = "" Then
'        MsgBox "É necessário selecionar uma data!", vbExclamation, App.ProductName
'        Exit Sub
'   End If
    
    sqlStr = "SELECT * FROM vw_bordero_aux "
    sqlStr = sqlStr & "WHERE vw_bordero_aux.bol_dt_mov between CONVERT(DATETIME, " & Format(dtpDtIni.Value, "'dd/mm/yyyy'") & ", 103) and CONVERT(DATETIME, " & Format(dtpDtFim.Value, "'dd/mm/yyyy'") & ", 103) "
    
    If cboDistrib.ListIndex <> -1 Then
        If cboDistrib.ItemData(cboDistrib.ListIndex) <> 0 Then
            sqlStr = sqlStr & "AND   vw_bordero_aux.dis_cd     = " & cboDistrib.ItemData(cboDistrib.ListIndex) & " "
        End If
    End If
    
    If cboEmpr.ListIndex <> -1 Then
        If cboEmpr.ItemData(cboEmpr.ListIndex) <> 0 Then
            sqlStr = sqlStr & "AND   vw_bordero_aux.emp_cd     = " & cboEmpr.ItemData(cboEmpr.ListIndex) & " "
        End If
    End If
    
    If cboCinema.ListIndex <> -1 Then
        If cboCinema.ItemData(cboCinema.ListIndex) <> 0 Then
            sqlStr = sqlStr & "AND   vw_bordero_aux.cin_cd     = " & cboCinema.ItemData(cboCinema.ListIndex) & " "
        End If
    End If
    
    If cboSala.ListIndex <> -1 Then
        If cboSala.ItemData(cboSala.ListIndex) <> 0 Then
            sqlStr = sqlStr & "AND   vw_bordero_aux.sal_cd     = " & cboSala.ItemData(cboSala.ListIndex) & " "
        End If
    End If
    
    If cboFilme.ListIndex <> -1 Then
        If cboFilme.ItemData(cboFilme.ListIndex) <> 0 Then
            sqlStr = sqlStr & "AND   vw_bordero_aux.fil_cd     = " & cboFilme.ItemData(cboFilme.ListIndex) & " "
        End If
    End If
    
    sqlStr = sqlStr & "ORDER BY vw_bordero_aux.bol_dt_mov, "
    sqlStr = sqlStr & "vw_bordero_aux.emp_nm, "
    sqlStr = sqlStr & "vw_bordero_aux.cin_nm, "
    sqlStr = sqlStr & "vw_bordero_aux.sal_desc, "
    sqlStr = sqlStr & "vw_bordero_aux.fil_nm, "
    sqlStr = sqlStr & "vw_bordero_aux.igt_cd "
    
    Set oRs = dbConnect.Execute(sqlStr)
    
    m_Report.Database.SetDataSource oRs
    
    CRVBordero.ReportSource = m_Report
    CRVBordero.ViewReport
End Sub

Private Sub cmdPerioido_Click()
     Call carregaDistrib
     Call carregaEmpr
End Sub

Private Sub cmdSala_Click()
    Call carregaSala
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
    lblSala.Left = lblDtMovto.Left
    lblFilme.Left = lblDtMovto.Left
    
    lblDe.Left = lblDtMovto.Left + 1545
    dtpDtIni.Left = lblDtMovto.Left + 1875
    lblA.Left = lblDtMovto.Left + 3360
    dtpDtFim.Left = lblDtMovto.Left + 3570
    cmdPerioido.Left = lblDtMovto.Left + 6240
    
    cboDistrib.Left = lblDtMovto.Left + 1515
    cboEmpr.Left = cboDistrib.Left
    cboCinema.Left = cboDistrib.Left
    cboSala.Left = cboDistrib.Left
    cboFilme.Left = cboDistrib.Left
    
    cmdDistrib.Left = cboDistrib.Left + cboDistrib.Width
    cmdEmpr.Left = cboEmpr.Left + cboEmpr.Width
    cmdCinema.Left = cboCinema.Left + cboCinema.Width
    cmdSala.Left = cboSala.Left + cboSala.Width
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

Private Sub carregaSala()
    Dim strSql As String
    
    cboSala.Clear
    
    If tbSala.State = adStateOpen Then
        tbSala.Close
    End If
    
    strSql = "SELECT DISTINCT emp_cd, cin_cd, sal_cd, sal_desc FROM tb_bol_sala "
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
    
    strSql = strSql & "ORDER BY sal_desc"
    
    tbSala.Open strSql, dbConnect, adOpenDynamic
    
    cboSala.AddItem "Todas", 0
    cboSala.ItemData(cboSala.NewIndex) = 0
    
    Do While Not tbSala.EOF
        cboSala.AddItem tbSala.Fields("sal_desc").Value
        cboSala.ItemData(cboSala.NewIndex) = tbSala.Fields("sal_cd").Value
        
        tbSala.MoveNext
    Loop
    
    cboSala.ListIndex = 0
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
    
    If cboSala.ListIndex <> -1 Then
        If cboSala.ItemData(cboSala.ListIndex) <> 0 Then
            strSql = strSql & " AND sal_cd = " & cboSala.ItemData(cboSala.ListIndex)
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

