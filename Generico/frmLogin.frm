VERSION 5.00
Begin VB.Form frmLogin 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Login"
   ClientHeight    =   1545
   ClientLeft      =   2835
   ClientTop       =   3480
   ClientWidth     =   5400
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   912.837
   ScaleMode       =   0  'User
   ScaleWidth      =   5070.307
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame fraAlteraSenha 
      Caption         =   "Altera Senha"
      Height          =   1095
      Left            =   840
      TabIndex        =   10
      Top             =   1200
      Visible         =   0   'False
      Width           =   2835
      Begin VB.TextBox txtAltSenha 
         Height          =   315
         IMEMode         =   3  'DISABLE
         Left            =   870
         MaxLength       =   20
         PasswordChar    =   "*"
         TabIndex        =   2
         Top             =   300
         Width           =   1800
      End
      Begin VB.TextBox txtConfSenha 
         Height          =   315
         IMEMode         =   3  'DISABLE
         Left            =   870
         MaxLength       =   20
         PasswordChar    =   "*"
         TabIndex        =   3
         Top             =   660
         Width           =   1800
      End
      Begin VB.Label lblLabels 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "&Senha:"
         Height          =   195
         Index           =   2
         Left            =   300
         TabIndex        =   12
         Top             =   330
         Width           =   510
      End
      Begin VB.Label lblConfSenha 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Con&firma:"
         Height          =   195
         Left            =   120
         TabIndex        =   11
         Top             =   675
         Width           =   690
         WordWrap        =   -1  'True
      End
   End
   Begin VB.Frame fraLogin 
      Caption         =   "Login"
      Height          =   1095
      Left            =   840
      TabIndex        =   7
      Top             =   60
      Width           =   2835
      Begin VB.TextBox txtLogin 
         Height          =   315
         IMEMode         =   3  'DISABLE
         Left            =   870
         MaxLength       =   10
         TabIndex        =   0
         Top             =   300
         Width           =   1800
      End
      Begin VB.TextBox txtSenha 
         Height          =   315
         IMEMode         =   3  'DISABLE
         Left            =   870
         MaxLength       =   20
         PasswordChar    =   "*"
         TabIndex        =   1
         Top             =   645
         Width           =   1800
      End
      Begin VB.Label lblLabels 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "&Senha:"
         Height          =   195
         Index           =   1
         Left            =   285
         TabIndex        =   9
         Top             =   660
         Width           =   510
      End
      Begin VB.Label lblLabels 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "&Usuário:"
         Height          =   195
         Index           =   0
         Left            =   180
         TabIndex        =   8
         Top             =   330
         Width           =   585
      End
   End
   Begin VB.CommandButton cmdAlteraSenha 
      Caption         =   "&Altera Senha >>>"
      Height          =   390
      Left            =   3840
      TabIndex        =   6
      Top             =   1080
      Width           =   1438
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "&OK"
      Default         =   -1  'True
      Height          =   390
      Left            =   3840
      TabIndex        =   4
      Top             =   120
      Width           =   1438
   End
   Begin VB.CommandButton cmdCancela 
      Cancel          =   -1  'True
      Caption         =   "&Cancela"
      Height          =   390
      Left            =   3840
      TabIndex        =   5
      Top             =   600
      Width           =   1438
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   180
      Picture         =   "frmLogin.frx":0000
      Top             =   360
      Width           =   480
   End
End
Attribute VB_Name = "frmLogin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public TipoUsuario As eTipoUsuario
Public FormQueChamou As Form

Private clsTB_USUARIO As New Cine2005.clsTB_USUARIO
Private acesso        As New clsAcesso
Private bAlteraSenha As Boolean
Private bUsuarioOk As Boolean

Private sSenhaBcoCript As String
Private sSenhaCript As String

'Private cCryptoVB As New CineCrypto.cnCrypto
Private sSenhaBranco As String
Private nTentativas As Integer

Public mod_cd As Integer
Public fun_cd As Integer

Private Sub txtLogin_Change()
    txtSenha = ""
    txtAltSenha = ""
    txtConfSenha = ""
    bAlteraSenha = False
    bUsuarioOk = False
    Call AlteraSenha(False)
    txtSenha.Enabled = True
End Sub

Private Sub cmdAlteraSenha_Click()
    Call AlteraSenha(SenhaOk())
    bAlteraSenha = True
End Sub

Private Sub cmdCancela_Click()
    If Not FormQueChamou Is Nothing Then
        Unload Me
    Else
        End
    End If
End Sub

Private Sub cmdOK_Click()
    
    Dim sMens As String
    
    nTentativas = nTentativas + 1
    
    If nTentativas > 3 Then
        MsgBox "Nº de tentativas esgotou!", vbCritical, App.ProductName
        nTentativas = 0
        If Not FormQueChamou Is Nothing Then
            Unload Me
            Exit Sub
        Else
            End
        End If
    End If
    
    If Not SenhaOk() Then
        Exit Sub
    Else
        If Not GravaSenha Then
            Exit Sub
        End If
    End If
    
    If FormQueChamou Is Nothing Then

        If CDbl(Time) > 0.25 And CDbl(Time) <= 0.5 Then
            sMens = "Tenha um bom dia "
        ElseIf CDbl(Time) > 0.5 And CDbl(Time) < 0.75 Then
            sMens = "Tenha uma boa tarde "
        Else
            sMens = "Tenha uma boa noite "
        End If

        MsgBox sMens & strUsuario & ".", vbInformation, App.ProductName
    
    End If
    
    strLogin = txtLogin.Text
    
    Unload Me
    
End Sub

Private Sub Form_Activate()
'    Call UsuarioOk
'    If sSenhaBcoCript = "" Then
'        txtSenha.Enabled = False
'        If txtAltSenha.Visible Then txtAltSenha.SetFocus
'    Else
'        txtSenha.SetFocus
'    End If
    txtLogin.SetFocus
End Sub

Private Sub Form_Load()
    
    Set clsTB_USUARIO.ConexaoADO = dbConnect
    Set acesso.ConexaoADO = dbConnect
    
    txtLogin.Text = strLogin
    
'    cCryptoVB.DadoEntrada = ""
'    cCryptoVB.Decrypt
'    sSenhaBranco = cCryptoVB.Resultado
    sSenhaBranco = ""
    
    nTentativas = 0
    
End Sub

Private Function UsuarioOk() As Boolean

    Dim oRs    As New ADODB.Recordset
    Dim per_cd As Integer
    
    If txtLogin.Text = "" Then
        MsgBox "Usuário deve ser informado!", vbCritical, App.ProductName
        txtLogin.SetFocus
        GoTo Fim
    End If
    
    If Not bUsuarioOk Then
    
        clsTB_USUARIO.usu_login = txtLogin.Text
        
        If Not clsTB_USUARIO.PegaSenhaLogin(oRs) Then
            MsgBox clsTB_USUARIO.MensagemErro, vbCritical, App.ProductName
            GoTo Fim
        End If
        
        If oRs.EOF Then
            MsgBox "Usuário não encontrado!", vbCritical, App.ProductName
            txtLogin.SetFocus
            SendKeys "{Home}+{End}"
            GoTo Fim
        End If
        
        intUsuario = oRs.Fields("usu_cd")
        strUsuario = oRs.Fields("usu_nm")
        per_cd = oRs.Fields("per_cd")
        
        ' Verifica se usuário tem o perfil necessário para se logar
        
        'If oRs.Fields("per_cd") < TipoUsuario Then
        If Not acesso.VerificaAcesso(mod_cd, fun_cd, per_cd) Then
            MsgBox "Usuário não tem o perfil para entrar nesse módulo! Contate o administrador.", vbExclamation, App.ProductName
            GoTo Fim
        End If
        
        sSenhaBcoCript = Trim(IIf(IsNull(oRs.Fields("usu_senha")), "", oRs.Fields("usu_senha")))
        
        bAlteraSenha = (sSenhaBcoCript = "")
        Call AlteraSenha(sSenhaBcoCript = "")
        
    End If
    
    bUsuarioOk = True
    UsuarioOk = True
    
Fim:
    If oRs.State = 1 Then oRs.Close
    Set oRs = Nothing
    
End Function

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
        
        If sSenhaCript = sSenhaBranco Or sSenhaBcoCript = sSenhaBranco Then
            MsgBox "Senha inválida, tente novamente!", vbCritical, App.ProductName
            txtSenha.SetFocus
            SendKeys "{Home}+{End}"
            Exit Function
        End If
    
        If Trim(sSenhaBcoCript) <> Trim(sSenhaCript) Then
            If bAlteraSenha Then
                MsgBox "Senha não confere, portanto não pode ser alterada!", vbCritical, App.ProductName
            Else
                MsgBox "Senha inválida, tente novamente!", vbCritical, App.ProductName
            End If
            txtSenha.SetFocus
            SendKeys "{Home}+{End}"
            Exit Function
        End If
    
    End If
    
    If fraAlteraSenha.Visible Then
        If bAlteraSenha Then
            If Trim(txtAltSenha.Text) = "" Then
                MsgBox "Senha não pode ser branca, tente novamente!", vbCritical, App.ProductName
                txtAltSenha.SetFocus
                SendKeys "{Home}+{End}"
                Exit Function
            End If
            If Trim(txtAltSenha.Text) <> Trim(txtConfSenha) Then
                MsgBox "Senha não confere, tente novamente!", vbCritical, App.ProductName
                txtConfSenha.SetFocus
                SendKeys "{Home}+{End}"
                Exit Function
            End If
        End If
    End If
    
    SenhaOk = True
    
End Function

Private Function GravaSenha() As Boolean

    If bAlteraSenha Then
'        cCryptoVB.DadoEntrada = Trim(txtAltSenha.Text)
'        cCryptoVB.Encrypt
'        clsTB_USUARIO.usu_senha = cCryptoVB.Resultado
        clsTB_USUARIO.usu_senha = Trim(txtAltSenha.Text)
        clsTB_USUARIO.usu_cd = intUsuario
        If Not clsTB_USUARIO.AlterarSenha Then
            MsgBox clsTB_USUARIO.MensagemErro, vbCritical, App.ProductName
            GoTo Fim
        Else
            MsgBox "Senha alterada com sucesso!", vbInformation, App.ProductName
        End If
    End If
    
    GravaSenha = True
    
Fim:
    
End Function

Private Sub AlteraSenha(ByVal bAltera As Boolean)

    fraAlteraSenha.Visible = bAltera
    Me.Height = IIf(fraAlteraSenha.Visible, 2850, 2025)
    cmdAlteraSenha.Visible = Not bAltera
    
    If bAltera Then
        txtSenha.Enabled = False
        txtAltSenha.SetFocus
    End If
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
'    Set cCryptoVB = Nothing
    Set clsTB_USUARIO = Nothing
    Set acesso = Nothing
End Sub

Private Sub txtLogin_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub
