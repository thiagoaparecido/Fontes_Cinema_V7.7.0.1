VERSION 5.00
Begin VB.MDIForm MDIFrmAdmCentral 
   BackColor       =   &H00FFFFFF&
   Caption         =   "Administração Central"
   ClientHeight    =   8490
   ClientLeft      =   60
   ClientTop       =   750
   ClientWidth     =   11880
   LinkTopic       =   "MDIForm1"
   Picture         =   "MDIFrmAdmCentral.frx":0000
   StartUpPosition =   2  'CenterScreen
   Begin VB.Menu mnuRealtorios 
      Caption         =   "&Relatórios"
      Begin VB.Menu mnuBordero 
         Caption         =   "&Borderô"
      End
      Begin VB.Menu mnuOcupacao 
         Caption         =   "&Ocupação"
      End
      Begin VB.Menu mnuVendasFilme 
         Caption         =   "Vendas por &Filme"
      End
      Begin VB.Menu mnuVendasTotalPublico 
         Caption         =   "Vendas Total Publico"
      End
   End
   Begin VB.Menu mnuFerramentas 
      Caption         =   "&Ferramentas"
      Begin VB.Menu mnuImpotMovto 
         Caption         =   "&Importa Movimento"
      End
      Begin VB.Menu mnuLog 
         Caption         =   "&Log Sistema"
      End
      Begin VB.Menu mnuExpurgo 
         Caption         =   "&Expurgo"
      End
      Begin VB.Menu mnuBackup 
         Caption         =   "&Backup base de dados"
      End
      Begin VB.Menu mnuRestor 
         Caption         =   "&Restore base de dados"
      End
   End
End
Attribute VB_Name = "MDIFrmAdmCentral"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'Dim WithEvents oBackupEvent As SQLDMO.Backup

Private Sub MDIForm_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    Call Backupbase
End Sub

Private Sub mnuBackup_Click()
    frmBackup.Show vbModal
End Sub

Private Sub mnuBordero_Click()
    frmBordero.Show
End Sub

Private Sub mnuExpurgo_Click()
    frmExpurgo.Show vbModal
End Sub

Private Sub mnuImpotMovto_Click()
    frmImporta.Show vbModal
End Sub

Private Sub mnuLog_Click()
    frmLog.Show vbModal
End Sub

Private Sub mnuOcupacao_Click()
    frmOcupacao.Show
End Sub

Private Sub mnuRestor_Click()
    frmRestor.Show vbModal
End Sub

Private Sub mnuVendasFilme_Click()
    frmVendasFilme.Show
End Sub

Private Sub mnuVendasTotalPublico_Click()
    frmVendasPublico.Show
End Sub


Private Sub Backupbase()
    On Error GoTo TrataErro
    
    Dim Registry As New Cine2005.ManipulaRegistry 'Variável para permitir a leitura do Registry
    
    If CInt(pDiasBackup) <= 0 Then
        Exit Sub
    End If
    
    If DateDiff("d", CVDate(pUltimoBackup), Date) < CInt(pDiasBackup) Then
        Exit Sub
    End If
   
    Dim rs As ADODB.Recordset
    Set rs = New ADODB.Recordset
         
    rs.Open "Backup Database CentralADM to Disk = 'CentralAdm" & Format(Date, "ddmmyyyy") & ".bak' With Init", "file Name=" + Trim(App.Path) + "\AdmCentral.udl", adOpenDynamic, adLockPessimistic
        
    Set rs = Nothing
    
    pUltimoBackup = Format(Date, "dd/mm/yyyy")
    Registry.GravaString RegKey_LOCAL_MACHINE, "Software\" & App.ProductName, "UltimoBackup", pUltimoBackup
    
    Unload Me
    Exit Sub
    
TrataErro:
    Set rs = Nothing
    MsgBox "Erro " & Format$(Err.Number) & " ao criar o arquivo." & vbCrLf & Err.Description, vbCritical + vbOKOnly, "Atenção"

End Sub
'Private Sub backupBase_old()
'    Dim gSQLServer As SQLDMO.SQLServer
'    Dim oBackup    As SQLDMO.Backup
'    Dim bConnect   As Boolean
'    Dim fullArq    As String
'    Dim Registry As New Cine2005.ManipulaRegistry 'Variável para permitir a leitura do Registry
'
'    If CInt(pDiasBackup) <= 0 Then
'        Exit Sub
'    End If
'
'    If DateDiff("d", CVDate(pUltimoBackup), Date) < CInt(pDiasBackup) Then
'        Exit Sub
'    End If
'
'    bConnect = True
'
'    fullArq = App.Path & "\"
'
'    fullArq = gGetShortPathName(fullArq) & "CentralAdm" & Format(Date, "ddmmyyyy") & ".bak"
'
'    Set gSQLServer = New SQLDMO.SQLServer
'
'    ' Set the login timeout.
'    gSQLServer.LoginTimeout = 15
'
'    gSQLServer.Connect servidor, usuarioDB, senhaDB
'    bConnect = True
'
'    Set oBackup = New SQLDMO.Backup
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
'    Call gSQLServer.DisConnect
'    Set gSQLServer = Nothing
'
'    pUltimoBackup = Format(Date, "dd/mm/yyyy")
'    Registry.GravaString RegKey_LOCAL_MACHINE, "Software\" & App.ProductName, "UltimoBackup", pUltimoBackup
'
'    Exit Sub
'
'TrataErro:
'    Screen.MousePointer = vbDefault
'    MsgBox Err.Description, vbCritical, App.ProductName
'
'    If bConnect Then
'        Call gSQLServer.DisConnect
'    End If
'
'    Set oBackupEvent = Nothing
'    Set oBackup = Nothing
'    Set gSQLServer = Nothing
'
'End Sub

Private Sub oBackupEvent_Complete(ByVal Message As String)
    MsgBox "Backup completo!", vbInformation, App.ProductName
End Sub



