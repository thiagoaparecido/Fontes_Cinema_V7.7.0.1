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
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Enabled         =   0   'False
      Height          =   465
      Left            =   945
      TabIndex        =   4
      Top             =   135
      Width           =   6540
      Begin VB.TextBox txtArq 
         Height          =   315
         Left            =   0
         Locked          =   -1  'True
         TabIndex        =   5
         Top             =   45
         Width           =   6450
      End
   End
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
      Left            =   7515
      TabIndex        =   0
      Top             =   180
      Visible         =   0   'False
      Width           =   315
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
    
    arqBackup = Format(Now, "yyyy_mm_dd") & "central.bak"
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

Private Sub cmdok_click()
    On Error GoTo TrataErro
    
    If Len(Trim(txtArq.Text)) = 0 Then
        MsgBox "Arquivo invalido", vbCritical, App.ProductName
        Exit Sub
    End If
    
  Dim rs As ADODB.Recordset
    Set rs = New ADODB.Recordset
         
    rs.Open "Backup Database CentralADM to Disk = '" + txtArq.Text + "' With Init", "file Name=" + Trim(App.Path) + "\AdmCentral.udl", adOpenDynamic, adLockPessimistic
        
    Set rs = Nothing
    Me.Refresh
    DoEvents
    
    MsgBox "Backup Realizado com Sucesso!", vbInformation + vbOKOnly, "Atenção"
        
    Unload Me
    
    Exit Sub
TrataErro:
    Set rs = Nothing
    MsgBox "Erro " & Format$(Err.Number) & " ao criar o arquivo." & vbCrLf & Err.Description, vbCritical + vbOKOnly, "Atenção"
End Sub

Private Sub Form_Load()
        txtArq.Text = Format(Now, "yyyy_mm_dd") & "central.bak"
End Sub

Private Sub oBackupEvent_Complete(ByVal Message As String)
    MsgBox "Backup completo!", vbInformation, App.ProductName
End Sub

