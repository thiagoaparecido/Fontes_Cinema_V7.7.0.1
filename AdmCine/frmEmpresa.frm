VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Object = "{234FCAFF-8D53-4DC2-9CD1-F90F4F6CB524}#22.0#0"; "Comandos.ocx"
Begin VB.Form frmEmpresa 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Cadastro - Empresa"
   ClientHeight    =   5610
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6960
   Icon            =   "frmEmpresa.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5610
   ScaleWidth      =   6960
   Begin VB.Frame fraManut 
      Enabled         =   0   'False
      Height          =   4785
      Left            =   15
      TabIndex        =   11
      Top             =   30
      Width           =   6840
      Begin VB.TextBox txt_cd_emp 
         Height          =   315
         Left            =   1260
         Locked          =   -1  'True
         MaxLength       =   4
         TabIndex        =   24
         Top             =   210
         Width           =   780
      End
      Begin MSMask.MaskEdBox msk_emp_cnpj 
         Height          =   315
         Left            =   1260
         TabIndex        =   1
         Top             =   960
         Width           =   1875
         _ExtentX        =   3307
         _ExtentY        =   556
         _Version        =   393216
         MaxLength       =   18
         Mask            =   "99.999.999/9999-99"
         PromptChar      =   "_"
      End
      Begin VB.TextBox txt_emp_tel 
         Height          =   315
         Left            =   1260
         MaxLength       =   20
         TabIndex        =   10
         Top             =   4320
         Width           =   2250
      End
      Begin VB.TextBox txt_emp_uf_end 
         Height          =   315
         Left            =   1260
         MaxLength       =   2
         TabIndex        =   8
         Top             =   3570
         Width           =   375
      End
      Begin VB.TextBox txt_emp_nm 
         Height          =   315
         Left            =   1260
         MaxLength       =   50
         TabIndex        =   0
         Top             =   585
         Width           =   5400
      End
      Begin VB.TextBox txt_emp_cid_end 
         Height          =   315
         Left            =   1260
         MaxLength       =   50
         TabIndex        =   7
         Top             =   3195
         Width           =   5400
      End
      Begin VB.TextBox txt_emp_brr_end 
         Height          =   315
         Left            =   1260
         MaxLength       =   50
         TabIndex        =   6
         Top             =   2820
         Width           =   5400
      End
      Begin VB.TextBox txt_emp_cmp_end 
         Height          =   315
         Left            =   1260
         MaxLength       =   20
         TabIndex        =   5
         Top             =   2445
         Width           =   2250
      End
      Begin VB.TextBox txt_emp_num_end 
         Height          =   315
         Left            =   1260
         MaxLength       =   5
         TabIndex        =   4
         Top             =   2070
         Width           =   915
      End
      Begin VB.TextBox txt_emp_end 
         Height          =   315
         Left            =   1260
         MaxLength       =   50
         TabIndex        =   3
         Top             =   1695
         Width           =   5400
      End
      Begin MSMask.MaskEdBox msk_emp_cep_end 
         Height          =   315
         Left            =   1260
         TabIndex        =   9
         Top             =   3930
         Width           =   1035
         _ExtentX        =   1826
         _ExtentY        =   556
         _Version        =   393216
         MaxLength       =   9
         Mask            =   "99999-999"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox msk_emp_inscricao 
         Height          =   315
         Left            =   1260
         TabIndex        =   2
         Top             =   1335
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   556
         _Version        =   393216
         MaxLength       =   15
         Mask            =   "999.999.999.999"
         PromptChar      =   "_"
      End
      Begin VB.Label lbl_cd_emp 
         AutoSize        =   -1  'True
         Caption         =   "Código:"
         Height          =   195
         Left            =   150
         TabIndex        =   25
         Top             =   210
         Width           =   540
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Inscriçao:"
         Height          =   195
         Left            =   150
         TabIndex        =   23
         Top             =   1335
         Width           =   690
      End
      Begin VB.Label lblTelefone 
         AutoSize        =   -1  'True
         Caption         =   "Telefone:"
         Height          =   195
         Left            =   150
         TabIndex        =   21
         Top             =   4320
         Width           =   675
      End
      Begin VB.Label lblCEP 
         AutoSize        =   -1  'True
         Caption         =   "CEP:"
         Height          =   195
         Left            =   150
         TabIndex        =   20
         Top             =   3945
         Width           =   360
      End
      Begin VB.Label lblEstado 
         AutoSize        =   -1  'True
         Caption         =   "Estado:"
         Height          =   195
         Left            =   150
         TabIndex        =   19
         Top             =   3570
         Width           =   540
      End
      Begin VB.Label lblNome 
         AutoSize        =   -1  'True
         Caption         =   "Nome:"
         Height          =   195
         Left            =   150
         TabIndex        =   18
         Top             =   585
         Width           =   465
      End
      Begin VB.Label lblCidade 
         AutoSize        =   -1  'True
         Caption         =   "Cidade:"
         Height          =   195
         Left            =   150
         TabIndex        =   17
         Top             =   3195
         Width           =   540
      End
      Begin VB.Label lblBairro 
         AutoSize        =   -1  'True
         Caption         =   "Bairro:"
         Height          =   195
         Left            =   150
         TabIndex        =   16
         Top             =   2820
         Width           =   450
      End
      Begin VB.Label lblComplemento 
         AutoSize        =   -1  'True
         Caption         =   "Complemento:"
         Height          =   195
         Left            =   150
         TabIndex        =   15
         Top             =   2445
         Width           =   1005
      End
      Begin VB.Label lblNumero 
         AutoSize        =   -1  'True
         Caption         =   "Número:"
         Height          =   195
         Left            =   150
         TabIndex        =   14
         Top             =   2070
         Width           =   600
      End
      Begin VB.Label lblEndereco 
         AutoSize        =   -1  'True
         Caption         =   "Endereço:"
         Height          =   195
         Left            =   150
         TabIndex        =   13
         Top             =   1695
         Width           =   735
      End
      Begin VB.Label lblCNPJ 
         AutoSize        =   -1  'True
         Caption         =   "CNPJ:"
         Height          =   195
         Left            =   150
         TabIndex        =   12
         Top             =   960
         Width           =   450
      End
   End
   Begin Comandos.cmdComandos cmdComandos 
      Align           =   2  'Align Bottom
      Height          =   765
      Left            =   0
      TabIndex        =   22
      Top             =   4845
      Width           =   6960
      _ExtentX        =   12277
      _ExtentY        =   1349
      EnabledAltera   =   -1  'True
      VisibleNovo     =   0   'False
      VisibleExclui   =   0   'False
   End
End
Attribute VB_Name = "frmEmpresa"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private iCodEmpresa As Integer

Private Sub cmdComandos_Click(ByVal iButtonClicked As Comandos.eButtonClicked, Cancel As Boolean)
   
   Select Case iButtonClicked
    
        Case ButtonAltera
            Call HabilitaManut(True)
            
        Case ButtonGrava
        
            Dim bRet As Boolean
            
            bRet = GravaDados
            
            If Not bRet Then
                Cancel = True
            End If
            
            Call HabilitaManut(Not bRet)
    
        Case ButtonFecha
            Unload Me
            
        Case ButtonCancela
            Call HabilitaManut(False)
            Call CarregaDados
            
    End Select
End Sub

Private Sub HabilitaManut(ByVal bHabilita As Boolean)
    fraManut.Enabled = bHabilita
End Sub

Private Sub Form_Load()

    Dim CineAx As New CineAux.CineAx
    Call CineAx.CentralizaTela(MDIFrmAdmCine, Me)
    
    Set CineAx = Nothing
    
    If Not CarregaDados Then
        MsgBox "Erro ao carregar dados da Empresa!", vbCritical, "Erro"
    End If
    
End Sub

Private Function CarregaDados() As Boolean

    On Error GoTo CarregaDados_Erro
    
    Dim oRs As New ADODB.Recordset
    Dim clsTB_EMPRESA As New Cine2005.clsTB_EMPRESA
    
    Set clsTB_EMPRESA.ConexaoADO = dbConnect
    
    If Not clsTB_EMPRESA.Selecionar(oRs) Then
        Exit Function
    End If
    
    If Not oRs.EOF() Then
        'iCodEmpresa = oRs.Fields("emp_cd")
        txt_cd_emp.Text = oRs.Fields("emp_cd")
        txt_emp_nm.Text = oRs.Fields("emp_nm")
        If Trim(oRs.Fields("emp_cnpj")) <> "" Then
            msk_emp_cnpj.Text = Format(oRs.Fields("emp_cnpj"), "@@.@@@.@@@/@@@@-@@")
        End If
        If Trim(oRs.Fields("emp_inscricao")) <> "" Then
            msk_emp_inscricao.Text = Format(oRs.Fields("emp_inscricao"), "@@@.@@@.@@@.@@@")
        End If
        txt_emp_end.Text = oRs.Fields("emp_end")
        txt_emp_num_end.Text = oRs.Fields("emp_num_end")
        txt_emp_cmp_end.Text = oRs.Fields("emp_cmp_end")
        txt_emp_brr_end.Text = oRs.Fields("emp_brr_end")
        txt_emp_cid_end.Text = oRs.Fields("emp_cid_end")
        txt_emp_uf_end.Text = oRs.Fields("emp_uf_end")
        If Trim(oRs.Fields("emp_cep_end")) <> "" Then
            msk_emp_cep_end.Text = Format(oRs.Fields("emp_cep_end"), "@@@@@-@@@")
        End If
        txt_emp_tel.Text = oRs.Fields("emp_tel")
    Else
        txt_cd_emp.Locked = False
    End If
    
    CarregaDados = True
    
CarregaDados_Erro:
    If oRs.State = 1 Then oRs.Close
    
    Set oRs = Nothing
    Set clsTB_EMPRESA = Nothing

End Function

Private Function GravaDados() As Boolean

    On Error GoTo GravaDados_Fim
    
    Dim sMensErro As String
    
    If Not IsNumeric(txt_cd_emp.Text) Then
        sMensErro = sMensErro & "Código do cinema invalido!" & vbCrLf
    End If
    
    If txt_emp_nm.Text = "" Then
        sMensErro = sMensErro & "Nome da Empresa deve ser informado!" & vbCrLf
    End If
    
    If msk_emp_cnpj.ClipText = "" Then
        sMensErro = sMensErro & "CNPJ da Empresa deve ser informado!" & vbCrLf
    End If
    
    If sMensErro <> "" Then
        MsgBox sMensErro, vbInformation, "Alerta"
        GoTo GravaDados_Fim
    End If
    
    Dim clsTB_EMPRESA As New Cine2005.clsTB_EMPRESA
    
    Set clsTB_EMPRESA.ConexaoADO = dbConnect
    
    'clsTB_EMPRESA.emp_cd = iCodEmpresa
    clsTB_EMPRESA.emp_cd = CInt(txt_cd_emp.Text)
    clsTB_EMPRESA.emp_nm = IIf(txt_emp_nm.Text = "", Empty, txt_emp_nm.Text)
    clsTB_EMPRESA.emp_cnpj = IIf(Trim(msk_emp_cnpj.ClipText) = "", Empty, msk_emp_cnpj.ClipText)
    clsTB_EMPRESA.emp_inscricao = IIf(Trim(msk_emp_inscricao.ClipText) = "", Empty, msk_emp_inscricao.ClipText)
    clsTB_EMPRESA.emp_end = IIf(txt_emp_end.Text = "", Empty, txt_emp_end.Text)
    clsTB_EMPRESA.emp_num_end = IIf(Val(txt_emp_num_end.Text) = 0, Empty, txt_emp_num_end.Text)
    clsTB_EMPRESA.emp_cmp_end = IIf(txt_emp_cmp_end.Text = "", Empty, txt_emp_cmp_end.Text)
    clsTB_EMPRESA.emp_brr_end = IIf(txt_emp_brr_end.Text = "", Empty, txt_emp_brr_end.Text)
    clsTB_EMPRESA.emp_cid_end = IIf(txt_emp_cid_end.Text = "", Empty, txt_emp_cid_end.Text)
    clsTB_EMPRESA.emp_uf_end = IIf(txt_emp_uf_end.Text = "", Empty, txt_emp_uf_end.Text)
    clsTB_EMPRESA.emp_cep_end = IIf(Trim(msk_emp_cep_end.ClipText) = "", Empty, msk_emp_cep_end.ClipText)
    clsTB_EMPRESA.emp_tel = IIf(txt_emp_tel.Text = "", Empty, txt_emp_tel.Text)
    
    GravaDados = clsTB_EMPRESA.Alterar()

    If Not GravaDados Then
        MsgBox "Erro ao gravar dados da Empresa!", vbCritical, "Erro"
    Else
        MsgBox "Dados gravados com sucesso!", vbInformation, App.ProductName
    End If

GravaDados_Fim:
    Set clsTB_EMPRESA = Nothing

End Function

Private Sub txt_emp_num_end_KeyPress(KeyAscii As Integer)
    If Not IsNumeric(Chr(KeyAscii)) And KeyAscii <> 8 Then
        KeyAscii = 0
    End If
End Sub

Private Sub txt_cd_emp_KeyPress(KeyAscii As Integer)
    If Not IsNumeric(Chr(KeyAscii)) And KeyAscii <> 8 Then
        KeyAscii = 0
    End If
End Sub

