VERSION 5.00
Begin VB.MDIForm MDIFrmAdmCine 
   BackColor       =   &H00FFFFFF&
   Caption         =   "Administração"
   ClientHeight    =   13050
   ClientLeft      =   465
   ClientTop       =   750
   ClientWidth     =   21735
   LinkTopic       =   "MDIForm1"
   LockControls    =   -1  'True
   Picture         =   "MDIFrmAdmCine.frx":0000
   StartUpPosition =   2  'CenterScreen
   Begin VB.Menu mnCadastro 
      Caption         =   "Cadastro"
      Begin VB.Menu mnEmpresa 
         Caption         =   "Empresa"
      End
      Begin VB.Menu mnCinema 
         Caption         =   "Cinema"
      End
      Begin VB.Menu mnSalas 
         Caption         =   "Salas"
      End
      Begin VB.Menu mnPoltronas 
         Caption         =   "Poltronas"
      End
      Begin VB.Menu mnuDistrib 
         Caption         =   "Distribuidoras"
      End
      Begin VB.Menu mnuLugaresSala 
         Caption         =   "Alteração de Lugares"
      End
      Begin VB.Menu mnuSep1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFeriados 
         Caption         =   "Feriados"
      End
      Begin VB.Menu mnuParametros 
         Caption         =   "Parâmetros"
      End
      Begin VB.Menu mnuCaixas 
         Caption         =   "Caixas"
      End
      Begin VB.Menu mnuUsuarios 
         Caption         =   "Usuários"
      End
      Begin VB.Menu mnuCatracas 
         Caption         =   "Catracas"
      End
      Begin VB.Menu mnuAcesso 
         Caption         =   "Perfil Acesso"
      End
      Begin VB.Menu mnuSep2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuSair 
         Caption         =   "Sair"
      End
   End
   Begin VB.Menu mnuProg 
      Caption         =   "Programação"
      Begin VB.Menu mnuFilmes 
         Caption         =   "Filmes"
      End
      Begin VB.Menu mnuSessoes 
         Caption         =   "Sessões"
         Begin VB.Menu mnuProgramacao 
            Caption         =   "Período Sessões"
         End
         Begin VB.Menu mnuCadSessoes 
            Caption         =   "Cadastro Sessões"
         End
      End
      Begin VB.Menu mnuPrecos 
         Caption         =   "Preços"
         Begin VB.Menu mnuProgPreco 
            Caption         =   "Período Preços"
         End
         Begin VB.Menu mnuCadPrecos 
            Caption         =   "Cadastro Preços"
         End
      End
      Begin VB.Menu mnuComb 
         Caption         =   "Combos"
         Begin VB.Menu mnuProgCombo 
            Caption         =   "Período Combo"
         End
         Begin VB.Menu mnuCombos 
            Caption         =   "Cadastro Combo"
         End
      End
   End
   Begin VB.Menu mnuImpressos 
      Caption         =   "&Impressos"
      Begin VB.Menu mnuFechamentoCaixa 
         Caption         =   "&Fechamento Adm. de Caixa"
      End
      Begin VB.Menu mnuVendaCombo 
         Caption         =   "&Relatório de Combos"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuVendaIngresso 
         Caption         =   "&Relatório de Vendas"
      End
      Begin VB.Menu mnuBoletimAdm 
         Caption         =   "&Boletim Parcial (Detalhado)"
      End
      Begin VB.Menu mnuBoletimAdm2 
         Caption         =   "&Boletim Administrativo"
      End
      Begin VB.Menu mnuPosCaixas 
         Caption         =   "&Posição Caixas"
      End
   End
   Begin VB.Menu mnuVideo 
      Caption         =   "&VideoHall"
      Begin VB.Menu mnuImagens 
         Caption         =   "&Imagens"
      End
      Begin VB.Menu mnuTrailers 
         Caption         =   "&Trailers"
      End
      Begin VB.Menu mnuParametrosVideo 
         Caption         =   "&Parâmetros"
      End
   End
   Begin VB.Menu mnuFerramentas 
      Caption         =   "&Ferramentas"
      Begin VB.Menu mnuBackup 
         Caption         =   "&Backup base de dados"
      End
      Begin VB.Menu mnuRestor 
         Caption         =   "&Restor base de dados"
      End
      Begin VB.Menu mnuEnvioMovto 
         Caption         =   "&Envio de Movimentos"
      End
      Begin VB.Menu mnuLog 
         Caption         =   "&Log Sistema"
      End
   End
   Begin VB.Menu mnuSobre 
      Caption         =   "&Sobre"
      Begin VB.Menu mnuVersao 
         Caption         =   "Versao"
      End
   End
End
Attribute VB_Name = "MDIFrmAdmCine"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private clsTB_USUARIO As New Cine2005.clsTB_USUARIO
Private acesso        As New clsAcesso

'Dim WithEvents oBackupEvent As SQLDMO.Backup

Private Sub MDIForm_Load()
    Dim log As New clsLog
    
    Set log.ConexaoADO = dbConnect
    log.usu_nm = strUsuario
    log.slg_descricao = "Entrou no módulo " & App.ProductName
    
    log.insereLog
    
    mnuVersao.Caption = "TickeMidia - ADMCINE - Versão " + str(App.Major) + "." + str(App.Minor)
End Sub

Private Sub MDIForm_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    Dim log As New clsLog
    
    Set log.ConexaoADO = dbConnect
    log.usu_nm = strUsuario
    log.slg_descricao = "Saiu do módulo " & App.ProductName
    
    log.insereLog
    
    Call backupBase
End Sub

Private Sub mnCinema_Click()
    frmCinema.Show
    frmCinema.ZOrder 0
End Sub

Private Sub mnEmpresa_Click()
    frmEmpresa.Show
    frmEmpresa.ZOrder 0
End Sub

Private Sub mnPoltronas_Click()
    frmPoltronas.Show
    frmPoltronas.ZOrder 0
End Sub

Private Sub mnSalas_Click()
    frmSalas.Show
    frmSalas.ZOrder 0
End Sub

Private Sub mnuAcesso_Click()

    frmSenhaSupervisor.mod_cd = 1
    frmSenhaSupervisor.fun_cd = 11
    frmSenhaSupervisor.Show vbModal

    If pbsupervisor Then
    
        frmAcesso.Show
        frmAcesso.ZOrder 0
    
    End If


    'Set clsTB_USUARIO.ConexaoADO = dbConnect
    'Set acesso.ConexaoADO = dbConnect
    
    'mod_cd, fun_cd, per_cd'
   '     If Not acesso.VerificaAcesso(frmLogin.mod_cd, 11, frmLogin.per_cd) Then
   '         MsgBox "Usuário não tem o perfil para entrar nesse módulo! Contate o administrador.", vbExclamation, App.ProductName
   '         Exit Sub
   '     End If
   '' frmAcesso.Show
    'frmAcesso.ZOrder 0
End Sub

Private Sub mnuBackup_Click()
    frmBackup.Show vbModal
End Sub

Private Sub mnuBoletimAdm_Click()
    frmBoletimAdm1.Show vbModal
End Sub

Private Sub mnuBoletimAdm2_Click()
    frmBoletimAdm2.Show vbModal
End Sub

Private Sub mnuCaixas_Click()
    frmCaixas.Show
    frmCaixas.ZOrder 0
End Sub

Private Sub mnuCatracas_Click()
    frmCatraca.Show
    frmCatraca.ZOrder 0
End Sub

Private Sub mnuCombos_Click()
    frmCombo.Show
    frmCombo.ZOrder 0
End Sub

Private Sub mnuDistrib_Click()
    frmDistribuidora.Show
    frmDistribuidora.ZOrder 0
End Sub

Private Sub mnuEnvioMovto_Click()
    frmExport.Show vbModal
End Sub

Private Sub mnuFechamentoCaixa_Click()
    frmFechamentoCaixa.Show vbModal
End Sub

Private Sub mnuFeriados_Click()
    frmFeriados.Show
    frmFeriados.ZOrder 0
End Sub

Private Sub mnuFilmes_Click()
    frmFilmes.Show
    frmFilmes.ZOrder 0
End Sub

Private Sub mnuImagens_Click()
    frmImagem.Show
End Sub

Private Sub mnuLog_Click()
    frmLog.Show vbModal
End Sub

Private Sub mnuLugaresSala_Click()
    frmLugaresSala.Show
    frmLugaresSala.ZOrder 0
End Sub

Private Sub mnuParametros_Click()
    frmParametros.Show
    frmParametros.ZOrder 0
End Sub

Private Sub mnuParametrosVideo_Click()
        frmParametrosVideoHall.Show
End Sub

Private Sub mnuPosCaixas_Click()
    frmPosicaoCaixa.Show vbModal
End Sub

Private Sub mnuCadPrecos_Click()
    frmPreco1.Show
    frmPreco1.ZOrder 0
End Sub

Private Sub mnuProgCombo_Click()
    frmProgCombo.Show
    frmProgCombo.ZOrder 0
End Sub

Private Sub mnuProgPreco_Click()
    frmProgPreco.Show
    frmProgPreco.ZOrder 0
End Sub

Private Sub mnuProgramacao_Click()
    frmProgramacao.Show
    frmProgramacao.ZOrder 0
End Sub

Private Sub mnuRestor_Click()
    frmRestor.Show vbModal
End Sub

Private Sub mnuSair_Click()
    Dim log As New clsLog
    
    Set log.ConexaoADO = dbConnect
    log.usu_nm = strUsuario
    log.slg_descricao = "Saiu do módulo " & App.ProductName
    
    log.insereLog
    
    Call backupBase
   
   End
End Sub

Private Sub mnuCadSessoes_Click()
    frmSessoes.Show
    frmSessoes.ZOrder 0
End Sub

Private Sub mnuTrailers_Click()
        frmTrailer.Show
End Sub

Private Sub mnuUsuarios_Click()
    frmUsuarios.Show
    frmUsuarios.ZOrder 0
End Sub

Private Sub backupBaseOld()
'    'Dim gSQLServer As SQLDMO.SQLServer
'    'Dim oBackup    As SQLDMO.Backup
'    Dim bConnect   As Boolean
'    Dim fullArq    As String
'    Dim Registry As New Cine2005.ManipulaRegistry 'Variável para permitir a leitura do Registry
'
'    bConnect = True
'
'    If CInt(pDiasBackup) > 0 And DateDiff("d", CVDate(pUltimoBackup), Date) >= CInt(pDiasBackup) Then
'        fullArq = App.Path & "\"
'
'        fullArq = gGetShortPathName(fullArq) & "Cinema" & Format(Date, "ddmmyyyy") & ".bak"
'
'        'Set gSQLServer = New SQLDMO.SQLServer
'
'        ' Set the login timeout.
'        gSQLServer.LoginTimeout = 15
'
'        'gSQLServer.Connect servidor, usuarioDB, senhaDB
'        bConnect = True
'
'        'Set oBackup = New SQLDMO.Backup
'        'Set oBackupEvent = oBackup ' enable events
'
'        'oBackup.Database = baseDados
'        'oBackup.Files = fullArq
'
'        If Len(Dir(fullArq)) > 0 Then
'            Call Kill(fullArq)
'        End If
'
'        ' Change mousepointer while trying to connect.
'        Screen.MousePointer = vbHourglass
'
'        ' Backup the database.
'        oBackup.SQLBackup gSQLServer
'
'        ' Change mousepointer back to the default after connect.
'        Screen.MousePointer = vbDefault
'
'        Set oBackupEvent = Nothing ' disable events
'        Set oBackup = Nothing
'
'        Call gSQLServer.Disconnect
'        Set gSQLServer = Nothing
'
'        pUltimoBackup = Format(Date, "dd/mm/yyyy")
'        Registry.GravaString RegKey_LOCAL_MACHINE, "Software\" & App.ProductName, "UltimoBackup", pUltimoBackup
'    End If
'
'
'    fullArq = App.Path & "\"
'
'    fullArq = gGetShortPathName(fullArq) & "Cinema" & DatePart("w", Date) & ".bak"
'
'    'Set gSQLServer = New SQLDMO.SQLServer
'
'    ' Set the login timeout.
'    gSQLServer.LoginTimeout = 15
'
'    gSQLServer.Connect servidor, usuarioDB, senhaDB
'    bConnect = True
'
'    'Set oBackup = New SQLDMO.Backup
'    Set oBackupEvent = oBackup ' enable events
'
'    oBackup.Database = baseDados
'    oBackup.Files = fullArq
'
'    If Len(Dir(fullArq)) > 0 Then
'        Call Kill(fullArq)
'    End If
'
'    ' Change mousepointer while trying to connect.
'    Screen.MousePointer = vbHourglass
'
'    ' Backup the database.
'    oBackup.SQLBackup gSQLServer
'
'    ' Change mousepointer back to the default after connect.
'    Screen.MousePointer = vbDefault
'
'    Set oBackupEvent = Nothing ' disable events
'    Set oBackup = Nothing
'
'    Call gSQLServer.Disconnect
'    Set gSQLServer = Nothing
'
'    Exit Sub
'
'TrataErro:
'    Screen.MousePointer = vbDefault
'    MsgBox Err.Description, vbCritical, App.ProductName
'
'    If bConnect Then
'        Call gSQLServer.Disconnect
'    End If
'
'    Set oBackupEvent = Nothing
'    Set oBackup = Nothing
'    Set gSQLServer = Nothing

End Sub

Private Sub mnuVendaCombo_Click()
    frmVendaCombo.Show vbModal
End Sub

Private Sub mnuVendaIngresso_Click()
    frmVendaIngresso.Show vbModal
End Sub

' VB will create the right prototypes for you, if you select the oBackupEvent in
' the drop down listbox of your editor
Private Sub oBackupEvent_Complete(ByVal Message As String)
    MsgBox "Backup completo!", vbInformation, App.ProductName
End Sub

Private Sub oBackupEvent_PercentComplete(ByVal Message As String, ByVal Percent As Long)
    'Dim pos1    As Integer
    'Dim pos2    As Integer
    'Dim perct As Integer
    
    'pos1 = InStr(Message, "'")
    'pos2 = InStr(pos1 + 1, Message, "'")
    
    'perct = CInt(Mid(Message, pos1 + 1, pos2 - pos1 - 1))
    
    'ProgressBar1.Value = perct
End Sub


Private Sub backupBase()
        On Error GoTo TrataErro
        Dim Registry As New Cine2005.ManipulaRegistry 'Variável para permitir a leitura do Registry
    
        Dim wArquivo As String
        Dim rs As ADODB.Recordset
        Set rs = New ADODB.Recordset
        
        If CInt(pDiasBackup) > 0 And DateDiff("d", CVDate(pUltimoBackup), Date) >= CInt(pDiasBackup) Then
            rs.Open "Backup Database Cinema to Disk = '" + "Cinema" & Format(Date, "ddmmyyyy") & ".bak" + "' With Init", "file Name=" + Trim(App.Path) + "\Cinema.udl", adOpenDynamic, adLockPessimistic
            pUltimoBackup = Format(Date, "dd/mm/yyyy")
            Registry.GravaString RegKey_LOCAL_MACHINE, "Software\" & App.ProductName, "UltimoBackup", pUltimoBackup
        End If
        
    rs.Open "Backup Database Cinema to Disk = '" + "Cinema" & DatePart("w", Date) & ".bak" + "' With Init", "file Name=" + Trim(App.Path) + "\Cinema.udl", adOpenDynamic, adLockPessimistic
        
    Set rs = Nothing
    Exit Sub
    
TrataErro:
    Set rs = Nothing
    MsgBox "Erro " & Format$(Err.Number) & " ao criar o arquivo." & vbCrLf & Err.Description, vbCritical + vbOKOnly, "Atenção"

End Sub
