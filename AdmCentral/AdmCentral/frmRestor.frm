VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmRestor 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Restor base de dados"
   ClientHeight    =   1350
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   8220
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1350
   ScaleWidth      =   8220
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdCancela 
      Cancel          =   -1  'True
      Caption         =   "Cancela"
      Height          =   420
      Left            =   4185
      TabIndex        =   3
      Top             =   675
      Width           =   1590
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "Ok"
      Default         =   -1  'True
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
      Left            =   2265
      TabIndex        =   2
      Top             =   675
      Width           =   1590
   End
   Begin VB.TextBox txtArq 
      Height          =   315
      Left            =   1005
      Locked          =   -1  'True
      TabIndex        =   1
      Top             =   150
      Width           =   6450
   End
   Begin VB.CommandButton cmdArq 
      Caption         =   "..."
      Height          =   315
      Left            =   7500
      TabIndex        =   0
      Top             =   150
      Width           =   315
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   6360
      Top             =   630
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Label lblArq 
      AutoSize        =   -1  'True
      Caption         =   "Diretório:"
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
      Left            =   165
      TabIndex        =   4
      Top             =   195
      Width           =   795
   End
End
Attribute VB_Name = "frmRestor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim fullArqBackup As String

'Dim WithEvents oRestoreEvent As SQLDMO.Restore

Private Sub cmdCancela_Click()
    Unload Me
End Sub

Private Sub cmdArq_Click()
    Dim arqBackup  As String
    
    On Error GoTo TrataErro
    
    CommonDialog1.Flags = cdlOFNHideReadOnly Or cdlOFNLongNames Or cdlOFNPathMustExist Or cdlOFNFileMustExist
    CommonDialog1.CancelError = True
    CommonDialog1.Filter = "Backup (*.bak)|*.bak"
    CommonDialog1.FilterIndex = 1
    CommonDialog1.ShowOpen
    
    fullArqBackup = CommonDialog1.FileTitle ' .FileName
    
    txtArq.Text = fullArqBackup
    
    fullArqBackup = gGetShortPathName(fullArqBackup)
    
    Exit Sub
TrataErro:
    If Err.Number <> cdlCancel Then
        MsgBox Err.Description, vbCritical, App.ProductName
    End If
End Sub
Private Sub cmdOK_Click()

        If Len(Dir(Trim(txtArq.Text))) = 0 Or Len(Trim(txtArq.Text)) = 0 Then
            MsgBox "Arquivo não existe", vbCritical, App.ProductName
            Exit Sub
        End If
             
        'On Error Resume Next
        
        Dim wArquivo As String
        Dim rs As ADODB.Recordset
        Set rs = New ADODB.Recordset
        
        txtArq.Text = Trim(txtArq.Text)
        
        DoEvents
        
        Dim wConect As String
                
        rs.Open "SELECT * FROM master.dbo.sysdatabases where name = 'CentralAdm'", "FILE NAME=" & App.Path & "\AdmCentral.udl", adOpenDynamic, adLockReadOnly, adCmdText
        rs.Close
        On Error GoTo TrataErro
        
        wConect = Replace(rs.ActiveConnection, "Initial Catalog=CentralAdm;", "Initial Catalog=master;")
        
        rs.Open "SELECT * FROM master.dbo.sysdatabases where name = 'CENTRALADM'", wConect, adOpenDynamic, adLockReadOnly, adCmdText
        
        If Not rs.EOF And Not rs.BOF Then  'Existe -> Apaga-lo
            
            Dim wSql As String
            wSql = "USE master "
            wSql = wSql + "ALTER DATABASE CentralAdm SET SINGLE_USER WITH ROLLBACK IMMEDIATE "
            
            If rs.State = 1 Then
                rs.Close
            End If
            rs.Open (wSql)
            wSql = wSql + "ALTER DATABASE CentralAdm SET SINGLE_USER "
            If rs.State = 1 Then
                rs.Close
            End If
            rs.Open (wSql)
            If rs.State = 1 Then
                rs.Close
            End If
            wSql = "USE MASTER DROP DATABASE CentralAdm"
            rs.Open (wSql)
        End If
            
            If rs.State = 1 Then
                rs.Close
            End If
            
            wSql = "USE master "
            wSql = wSql + "Restore Database CentralAdm "
            wSql = wSql + "From Disk ='" + txtArq.Text + "' "
            rs.Open (wSql)
            
            If rs.State = 1 Then
                rs.Close
            End If
            wSql = wSql + "ALTER DATABASE CentralAdm SET MULTI_USER "
            rs.Open (wSql)
            
            MsgBox "O Backup do sistema foi restaurado com sucesso.", vbInformation + vbOKOnly, "Atenção"
            
            On Error GoTo 0
            Set rs = Nothing
            End
        
        Exit Sub
        
TrataErro:
    
    MsgBox "Erro " & Format$(Err.Number) & " ao restaurar o arquivo." & vbCrLf & Err.Description, vbCritical + vbOKOnly, "Atenção"
    On Error Resume Next
    
    If rs.State = 1 Then
        rs.Close
    End If
    rs.Open ("SELECT * FROM [master].[dbo].[sysdatabases] where name = 'CentralAdm'")
    
    If rs.RecordCount = 1 Then 'Existe -> Apaga-lo
        wSql = "USE master ALTER DATABASE CentralAdm SET MULTI_USER "
        rs.Open (wSql)
    End If
    
End Sub

'Private Sub cmdOKOld_Click()
'    'Dim gSQLServer  As SQLDMO.SQLServer
'    'Dim oRestore    As SQLDMO.Restore
'    Dim bConnect    As Boolean
'    Dim bDisconnect As Boolean
'
'    On Error GoTo TrataErro
'
'    If Len(Dir(Trim(txtArq.Text))) = 0 Or Len(Trim(txtArq.Text)) = 0 Then
'        MsgBox "Arquivo não existe", vbCritical, App.ProductName
'        Exit Sub
'    End If
'
'    bConnect = False
'    bDisconnect = False
'
'    'Set gSQLServer = New SQLDMO.SQLServer
'
'    ' Set the login timeout.
'    'gSQLServer.LoginTimeout = 15
'
'    gSQLServer.Connect servidor, usuarioDB, senhaDB
'    bConnect = True
'
'    'Set oRestore = New SQLDMO.Restore
'    Set oRestoreEvent = oRestore        ' enable events
'
'    oRestore.Database = baseDados
'    oRestore.Files = fullArqBackup
'    oRestore.ReplaceDatabase = True
'
'    ' Change mousepointer while trying to connect.
'    Screen.MousePointer = vbHourglass
'
'    dbConnect.Close
'    Set dbConnect = Nothing
'    DoEvents
'
'    bDisconnect = True
'
'    ' Backup the database.
'    oRestore.SQLRestore gSQLServer
'
'    ' Change mousepointer back to the default after connect.
'    Screen.MousePointer = vbDefault
'
'    Set oRestoreEvent = Nothing         ' disable events
'    Set oRestore = Nothing
'
'    Call gSQLServer.Disconnect
'    Set gSQLServer = Nothing
'
'    End
'
'    Unload Me
'
'    Exit Sub
'TrataErro:
'    Screen.MousePointer = vbDefault
'    MsgBox Err.Description, vbCritical, App.ProductName
'
'    If bConnect Then
'        Call gSQLServer.Disconnect
'    End If
'
'    Set oRestoreEvent = Nothing
'    Set oRestore = Nothing
'    Set gSQLServer = Nothing
'
'    If bDisconnect Then
'        End
'    End If
'End Sub

Private Sub oRestoreEvent_Complete(ByVal Message As String)
    MsgBox "Restor completo. O sistema será encerrado, reinicie o sistema!", vbInformation, App.ProductName
End Sub

Private Sub oRestoreEvent_PercentComplete(ByVal Message As String, ByVal Percent As Long)
    Dim pos1    As Integer
    Dim pos2    As Integer
    Dim perct As Integer
    
    'pos1 = InStr(Message, "'")
    'pos2 = InStr(pos1 + 1, Message, "'")
    
    'perct = CInt(Mid(Message, pos1 + 1, pos2 - pos1 - 1))
    
    'ProgressBar1.Value = perct
End Sub

