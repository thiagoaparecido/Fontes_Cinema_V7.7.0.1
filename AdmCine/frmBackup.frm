VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmBackup 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Backup base de dados"
   ClientHeight    =   1425
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   8220
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1425
   ScaleWidth      =   8220
   StartUpPosition =   2  'CenterScreen
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   6360
      Top             =   735
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton cmdCancela 
      Cancel          =   -1  'True
      Caption         =   "Cancela"
      Height          =   420
      Left            =   4185
      TabIndex        =   2
      Top             =   780
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
      TabIndex        =   1
      Top             =   780
      Width           =   1590
   End
   Begin VB.CommandButton cmdArq 
      Caption         =   "..."
      Height          =   315
      Left            =   7500
      TabIndex        =   0
      Top             =   150
      Visible         =   0   'False
      Width           =   315
   End
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      Height          =   645
      Left            =   855
      TabIndex        =   4
      Top             =   0
      Width           =   7170
      Begin VB.TextBox txtArq 
         Height          =   315
         Left            =   90
         Locked          =   -1  'True
         TabIndex        =   5
         Top             =   180
         Width           =   6450
      End
   End
   Begin VB.Label lblArq 
      AutoSize        =   -1  'True
      Caption         =   "Arquivo:"
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
      TabIndex        =   3
      Top             =   195
      Width           =   720
   End
End
Attribute VB_Name = "frmBackup"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim fullArqBackup As String

'Dim WithEvents oBackupEvent As SQLDMO.Backup

Private Sub cmdCancela_Click()
    Unload Me
End Sub

Private Sub cmdArq_Click()
    Dim arqBackup As String
    Dim strAux    As String
    
    
    On Error GoTo TrataErro
    
    arqBackup = Format(Now, "yyyy_mm_dd") & "cine.bak"
    CommonDialog1.Flags = cdlOFNHideReadOnly Or cdlOFNLongNames Or cdlOFNPathMustExist
    CommonDialog1.CancelError = True
    CommonDialog1.FileName = arqBackup
    CommonDialog1.ShowOpen
    
    fullArqBackup = CommonDialog1.FileName
    
    fullArqBackup = Replace(fullArqBackup, CommonDialog1.FileTitle, arqBackup)
    strAux = Mid(fullArqBackup, 1, InStr(1, fullArqBackup, "\" + arqBackup))
    strAux = gGetShortPathName(strAux)
    fullArqBackup = strAux + arqBackup
    
    txtArq.Text = fullArqBackup
    
    Exit Sub
TrataErro:
    If Err.Number <> cdlCancel Then
        MsgBox Err.Description, vbCritical, App.ProductName
    End If
End Sub

Private Sub cmdOK_OLD_Click()
'    'Dim gSQLServer As SQLDMO.SQLServer
'    'Dim oBackup    As SQLDMO.Backup
'    Dim bConnect   As Boolean
'
'    On Error GoTo TrataErro
'
'    If Len(Trim(txtArq.Text)) = 0 Then
'        MsgBox "Arquivo invalido", vbCritical, App.ProductName
'        Exit Sub
'    End If
'
'    bConnect = True
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
'    oBackup.Files = fullArqBackup
'
'    If Len(Dir(fullArqBackup)) > 0 Then
'        Call Kill(fullArqBackup)
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
'    Set oBackupEvent = Nothing
'    Set oBackup = Nothing
'    Set gSQLServer = Nothing
End Sub

Private Sub Form_Load()
       txtArq.Text = Format(Now, "yyyy_mm_dd") & "cine.bak"
       'txtArq.Enabled = False
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


Private Sub cmdOK_Click()
    On Error GoTo TrataErro
    If Len(Trim(txtArq.Text)) = 0 Then
        MsgBox "Arquivo invalido", vbCritical, App.ProductName
        Exit Sub
    End If
    
    
        On Error GoTo TrataErro
    
        Dim wArquivo As String
        Dim rs As ADODB.Recordset
        Set rs = New ADODB.Recordset
         
    rs.Open "Backup Database Cinema to Disk = '" + txtArq.Text + "' With Init", "file Name=" + Trim(App.Path) + "\Cinema.udl", adOpenDynamic, adLockPessimistic
        
    Set rs = Nothing
    
    MsgBox "Backup Realizado com Sucesso!", vbInformation + vbOKOnly, "Atenção"
    
    Unload Me
    Exit Sub
    
TrataErro:
    Set rs = Nothing
    MsgBox "Erro " & Format$(Err.Number) & " ao criar o arquivo." & vbCrLf & Err.Description, vbCritical + vbOKOnly, "Atenção"

End Sub
