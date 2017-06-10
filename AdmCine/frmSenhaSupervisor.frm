VERSION 5.00
Begin VB.Form frmSenhaSupervisor 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Login"
   ClientHeight    =   1260
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4755
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1260
   ScaleWidth      =   4755
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdCancela 
      Cancel          =   -1  'True
      Caption         =   "&Cancela"
      Height          =   390
      Left            =   3180
      TabIndex        =   6
      Top             =   600
      Width           =   1438
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "&OK"
      Default         =   -1  'True
      Height          =   390
      Left            =   3180
      TabIndex        =   5
      Top             =   120
      Width           =   1438
   End
   Begin VB.Frame fraLogin 
      Height          =   1095
      Left            =   120
      TabIndex        =   0
      Top             =   60
      Width           =   2955
      Begin VB.TextBox txtSenha 
         Height          =   315
         IMEMode         =   3  'DISABLE
         Left            =   990
         MaxLength       =   20
         PasswordChar    =   "*"
         TabIndex        =   2
         Top             =   645
         Width           =   1800
      End
      Begin VB.TextBox txtLogin 
         Height          =   315
         IMEMode         =   3  'DISABLE
         Left            =   990
         MaxLength       =   10
         TabIndex        =   1
         Top             =   300
         Width           =   1800
      End
      Begin VB.Label lblLabels 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "&Login:"
         Height          =   195
         Index           =   0
         Left            =   465
         TabIndex        =   4
         Top             =   330
         Width           =   435
      End
      Begin VB.Label lblLabels 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "&Senha:"
         Height          =   195
         Index           =   1
         Left            =   405
         TabIndex        =   3
         Top             =   660
         Width           =   510
      End
   End
End
Attribute VB_Name = "frmSenhaSupervisor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public TipoUsuario    As eTipoUsuario
Private clsTB_USUARIO As New Cine2005.clsTB_USUARIO
Private acesso        As New clsAcesso
Private bUsuarioOk    As Boolean

Private sSenhaBcoCript As String
Private sSenhaCript    As String

'Private cCryptoVB    As New CineCrypto.cnCrypto
Private sSenhaBranco As String
Private nTentativas  As Integer

Public mod_cd As Integer
Public fun_cd As Integer

Private Sub txtLogin_Change()
    txtSenha = ""
    bUsuarioOk = False
    txtSenha.Enabled = True
End Sub

Private Sub cmdCancela_Click()
    Unload Me
End Sub

Private Sub cmdOK_Click()
    
    nTentativas = nTentativas + 1
    
    If nTentativas > 3 Then
        MsgBox "Nº de tentativas esgotou!", vbCritical, App.ProductName
        nTentativas = 0
        Unload Me
        Exit Sub
    End If
      
    If nTentativas <= 3 Then
        pbSupervisor = SenhaOk()
        If pbSupervisor Then
            Unload Me
        End If
    End If
    
End Sub

Private Sub Form_Activate()
    txtLogin.SetFocus
End Sub

Private Sub Form_Load()
    
    TipoUsuario = GERENTE
    pbSupervisor = False
    
    Set clsTB_USUARIO.ConexaoADO = dbConnect
    Set acesso.ConexaoADO = dbConnect
    
'    cCryptoVB.DadoEntrada = ""
'    cCryptoVB.Decrypt
'    sSenhaBranco = cCryptoVB.Resultado
    sSenhaBranco = ""
    
    nTentativas = 0
    
End Sub

Private Function SenhaOk() As Boolean
    
    If Not UsuarioOk Then
        Exit Function
    End If
    
    txtSenha.Enabled = (sSenhaBcoCript <> "")
    
'    cCryptoVB.DadoEntrada = Trim(txtSenha.Text)
'    cCryptoVB.Decrypt
'    sSenhaCript = cCryptoVB.Resultado
    sSenhaCript = Trim(txtSenha.Text)
    
    If sSenhaBcoCript <> "" Then
        If sSenhaCript <> sSenhaBcoCript Then
            MsgBox "Senha inválida, tente novamente!", vbCritical, App.ProductName
            txtSenha.SetFocus
            SendKeys "{Home}+{End}"
            Exit Function
        End If
    End If

    SenhaOk = True
    
End Function

Private Function UsuarioOk() As Boolean

    Dim oRs    As New ADODB.Recordset
    Dim per_cd As Integer
    
    If txtLogin.Text = "" Then
        MsgBox "Gerente deve ser informado!", vbCritical, App.ProductName
        If txtLogin.Enabled Then txtLogin.SetFocus
        GoTo Fim
    End If
    
    If Not bUsuarioOk Then
    
        clsTB_USUARIO.usu_login = txtLogin.Text
        
        If Not clsTB_USUARIO.PegaSenhaLogin(oRs) Then
            MsgBox clsTB_USUARIO.MensagemErro, vbCritical, App.ProductName
            GoTo Fim
        End If
        
        If oRs.EOF Then
            MsgBox "Gerente não encontrado!", vbCritical, App.ProductName
            txtLogin.SetFocus
            SendKeys "{Home}+{End}"
            GoTo Fim
        End If
        
        piSupervisor = oRs.Fields("usu_cd")
        psSupervisor = txtLogin.Text
        per_cd = oRs.Fields("per_cd")
        
        ' Verifica se usuário tem o perfil necessário para se logar
        
        If Not acesso.VerificaAcesso(mod_cd, fun_cd, per_cd) Then
            MsgBox "Usuário não tem o perfil para acessar essa funcionalidade! Contate o administrador.", vbExclamation, App.ProductName
            GoTo Fim
        End If
        
        sSenhaBcoCript = Trim(IIf(IsNull(oRs.Fields("usu_senha")), "", oRs.Fields("usu_senha")))
        
    End If
    
    bUsuarioOk = True
    UsuarioOk = True
    
Fim:
    If oRs.State = 1 Then oRs.Close
    Set oRs = Nothing
    
End Function

Private Sub Form_Unload(Cancel As Integer)
'    Set cCryptoVB = Nothing
    Set clsTB_USUARIO = Nothing
End Sub

Private Sub txtLogin_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

