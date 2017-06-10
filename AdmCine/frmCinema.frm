VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Object = "{234FCAFF-8D53-4DC2-9CD1-F90F4F6CB524}#22.0#0"; "Comandos.ocx"
Object = "{1753B334-0119-4B34-9134-D2B3CD181550}#21.0#0"; "Combo.ocx"
Begin VB.Form frmCinema 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Cadastro - Cinema"
   ClientHeight    =   5985
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6975
   Icon            =   "frmCinema.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5985
   ScaleWidth      =   6975
   Begin Comandos.cmdComandos cmdComandos 
      Align           =   2  'Align Bottom
      Height          =   765
      Left            =   0
      TabIndex        =   24
      Top             =   5220
      Width           =   6975
      _ExtentX        =   12303
      _ExtentY        =   1349
      EnabledAltera   =   -1  'True
      VisibleNovo     =   0   'False
      VisibleExclui   =   0   'False
   End
   Begin VB.Frame fraManut 
      Enabled         =   0   'False
      Height          =   5115
      Left            =   60
      TabIndex        =   12
      Top             =   60
      Width           =   6840
      Begin VB.TextBox txt_cd_cin 
         Height          =   315
         Left            =   1260
         Locked          =   -1  'True
         MaxLength       =   4
         TabIndex        =   26
         Top             =   195
         Width           =   780
      End
      Begin Combo.cboCodDesc ccd_emp_cd 
         Height          =   315
         Left            =   1260
         TabIndex        =   0
         Top             =   570
         Width           =   5415
         _ExtentX        =   9551
         _ExtentY        =   556
         NomeTabela      =   "tb_empresa"
         NomeCampoCodigo =   "emp_cd"
         NomeCampoDescricao=   "emp_nm"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         TipoCampoCodigo =   2
         MostraBotaoNovo =   0   'False
         CodigoVisible   =   0   'False
         Filtro          =   "emp_dt_des is null"
         MostraBotaoAtualiza=   0   'False
      End
      Begin VB.TextBox txt_cin_end 
         Height          =   315
         Left            =   1260
         MaxLength       =   50
         TabIndex        =   4
         Top             =   2055
         Width           =   5400
      End
      Begin VB.TextBox txt_cin_num_end 
         Height          =   315
         Left            =   1260
         MaxLength       =   5
         TabIndex        =   5
         Top             =   2430
         Width           =   915
      End
      Begin VB.TextBox txt_cin_cmp_end 
         Height          =   315
         Left            =   1260
         MaxLength       =   20
         TabIndex        =   6
         Top             =   2805
         Width           =   2250
      End
      Begin VB.TextBox txt_cin_brr_end 
         Height          =   315
         Left            =   1260
         MaxLength       =   50
         TabIndex        =   7
         Top             =   3180
         Width           =   5400
      End
      Begin VB.TextBox txt_cin_cid_end 
         Height          =   315
         Left            =   1260
         MaxLength       =   50
         TabIndex        =   8
         Top             =   3555
         Width           =   5400
      End
      Begin VB.TextBox txt_cin_nm 
         Height          =   315
         Left            =   1260
         MaxLength       =   50
         TabIndex        =   1
         Top             =   945
         Width           =   5400
      End
      Begin VB.TextBox txt_cin_uf_end 
         Height          =   315
         Left            =   1260
         MaxLength       =   2
         TabIndex        =   9
         Top             =   3930
         Width           =   375
      End
      Begin VB.TextBox txt_cin_tel 
         Height          =   315
         Left            =   1260
         MaxLength       =   20
         TabIndex        =   11
         Top             =   4680
         Width           =   2250
      End
      Begin MSMask.MaskEdBox msk_cin_cnpj 
         Height          =   315
         Left            =   1260
         TabIndex        =   2
         Top             =   1320
         Width           =   1875
         _ExtentX        =   3307
         _ExtentY        =   556
         _Version        =   393216
         MaxLength       =   18
         Mask            =   "99.999.999/9999-99"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox msk_cin_cep_end 
         Height          =   315
         Left            =   1260
         TabIndex        =   10
         Top             =   4290
         Width           =   1035
         _ExtentX        =   1826
         _ExtentY        =   556
         _Version        =   393216
         MaxLength       =   9
         Mask            =   "99999-999"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox msk_cin_inscricao 
         Height          =   315
         Left            =   1260
         TabIndex        =   3
         Top             =   1695
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   556
         _Version        =   393216
         MaxLength       =   15
         Mask            =   "999.999.999.999"
         PromptChar      =   "_"
      End
      Begin VB.Label lbl_cd_cin 
         AutoSize        =   -1  'True
         Caption         =   "Código:"
         Height          =   195
         Left            =   150
         TabIndex        =   27
         Top             =   195
         Width           =   540
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Inscriçao:"
         Height          =   195
         Left            =   150
         TabIndex        =   25
         Top             =   1695
         Width           =   690
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Empresa:"
         Height          =   195
         Left            =   150
         TabIndex        =   23
         Top             =   630
         Width           =   660
      End
      Begin VB.Label lblCNPJ 
         AutoSize        =   -1  'True
         Caption         =   "CNPJ:"
         Height          =   195
         Left            =   150
         TabIndex        =   22
         Top             =   1320
         Width           =   450
      End
      Begin VB.Label lblEndereco 
         AutoSize        =   -1  'True
         Caption         =   "Endereço:"
         Height          =   195
         Left            =   150
         TabIndex        =   21
         Top             =   2055
         Width           =   735
      End
      Begin VB.Label lblNumero 
         AutoSize        =   -1  'True
         Caption         =   "Número:"
         Height          =   195
         Left            =   150
         TabIndex        =   20
         Top             =   2430
         Width           =   600
      End
      Begin VB.Label lblComplemento 
         AutoSize        =   -1  'True
         Caption         =   "Complemento:"
         Height          =   195
         Left            =   150
         TabIndex        =   19
         Top             =   2805
         Width           =   1005
      End
      Begin VB.Label lblBairro 
         AutoSize        =   -1  'True
         Caption         =   "Bairro:"
         Height          =   195
         Left            =   150
         TabIndex        =   18
         Top             =   3180
         Width           =   450
      End
      Begin VB.Label lblCidade 
         AutoSize        =   -1  'True
         Caption         =   "Cidade:"
         Height          =   195
         Left            =   150
         TabIndex        =   17
         Top             =   3555
         Width           =   540
      End
      Begin VB.Label lblNome 
         AutoSize        =   -1  'True
         Caption         =   "Nome:"
         Height          =   195
         Left            =   150
         TabIndex        =   16
         Top             =   945
         Width           =   465
      End
      Begin VB.Label lblEstado 
         AutoSize        =   -1  'True
         Caption         =   "Estado:"
         Height          =   195
         Left            =   150
         TabIndex        =   15
         Top             =   3930
         Width           =   540
      End
      Begin VB.Label lblCEP 
         AutoSize        =   -1  'True
         Caption         =   "CEP:"
         Height          =   195
         Left            =   150
         TabIndex        =   14
         Top             =   4305
         Width           =   360
      End
      Begin VB.Label lblTelefone 
         AutoSize        =   -1  'True
         Caption         =   "Telefone:"
         Height          =   195
         Left            =   150
         TabIndex        =   13
         Top             =   4680
         Width           =   675
      End
   End
End
Attribute VB_Name = "frmCinema"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private iCodCinema As Integer

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
    
    Set ccd_emp_cd.ConexaoADO = dbConnect
    
    If Not CarregaDados Then
        MsgBox "Erro ao carregar dados do Cinema!", vbCritical, "Erro"
    End If
    
End Sub

Private Function CarregaDados() As Boolean

    On Error GoTo CarregaDados_Erro
    
    Dim oRs As New ADODB.Recordset
    Dim clsTB_CINEMA As New Cine2005.clsTB_CINEMA
    
    Set clsTB_CINEMA.ConexaoADO = dbConnect
    
    If Not clsTB_CINEMA.Selecionar(oRs) Then
        Exit Function
    End If
    
    If Not oRs.EOF() Then
        txt_cd_cin.Locked = True
        'iCodCinema = oRs.Fields("cin_cd")
        txt_cd_cin.Text = oRs.Fields("cin_cd")
        ccd_emp_cd.codigo = oRs.Fields("emp_cd")
        ccd_emp_cd.Refresh
        txt_cin_nm.Text = oRs.Fields("cin_nm")
        If Trim(oRs.Fields("cin_cnpj")) <> "" Then
            msk_cin_cnpj.Text = Format(oRs.Fields("cin_cnpj"), "@@.@@@.@@@/@@@@-@@")
        End If
        If Trim(oRs.Fields("cin_inscricao")) <> "" Then
            msk_cin_inscricao.Text = Format(oRs.Fields("cin_inscricao"), "@@@.@@@.@@@.@@@")
        End If
        txt_cin_end.Text = oRs.Fields("cin_end")
        txt_cin_num_end.Text = oRs.Fields("cin_num_end")
        txt_cin_cmp_end.Text = oRs.Fields("cin_cmp_end")
        txt_cin_brr_end.Text = oRs.Fields("cin_brr_end")
        txt_cin_cid_end.Text = oRs.Fields("cin_cid_end")
        txt_cin_uf_end.Text = oRs.Fields("cin_uf_end")
        If Trim(oRs.Fields("cin_cep_end")) <> "" Then
            msk_cin_cep_end.Text = Format(oRs.Fields("cin_cep_end"), "@@@@@-@@@")
        End If
        txt_cin_tel.Text = oRs.Fields("cin_tel")
    Else
        txt_cd_cin.Locked = False
    End If
    
    CarregaDados = True
    
CarregaDados_Erro:
    If oRs.State = 1 Then oRs.Close
    
    Set oRs = Nothing
    Set clsTB_CINEMA = Nothing

End Function

Private Function GravaDados() As Boolean

    On Error GoTo GravaDados_Fim
    
    Dim sMensErro As String
    
    If Not IsNumeric(txt_cd_cin.Text) Then
        sMensErro = sMensErro & "Código do cinema invalido!" & vbCrLf
    End If
    
    If ccd_emp_cd.codigo = "" Then
        sMensErro = sMensErro & "Empresa deve ser informada!" & vbCrLf
    End If
    
    If txt_cin_nm.Text = "" Then
        sMensErro = sMensErro & "Nome do Cinema deve ser informado!" & vbCrLf
    End If
    
    If msk_cin_cnpj.ClipText = "" Then
        sMensErro = sMensErro & "CNPJ do Cinema deve ser informado!" & vbCrLf
    End If
    
    If sMensErro <> "" Then
        MsgBox sMensErro, vbInformation, App.ProductName
        GoTo GravaDados_Fim
    End If
    
    Dim clsTB_CINEMA As New Cine2005.clsTB_CINEMA
    
    Set clsTB_CINEMA.ConexaoADO = dbConnect
    
    'clsTB_CINEMA.cin_cd = iCodCinema
    clsTB_CINEMA.cin_cd = CInt(txt_cd_cin.Text)
    clsTB_CINEMA.emp_cd = ccd_emp_cd.codigo
    clsTB_CINEMA.cin_nm = IIf(txt_cin_nm.Text = "", Empty, txt_cin_nm.Text)
    clsTB_CINEMA.cin_cnpj = IIf(Trim(msk_cin_cnpj.ClipText) = "", Empty, msk_cin_cnpj.ClipText)
    clsTB_CINEMA.cin_inscricao = IIf(Trim(msk_cin_inscricao.ClipText) = "", Empty, msk_cin_inscricao.ClipText)
    clsTB_CINEMA.cin_end = IIf(txt_cin_end.Text = "", Empty, txt_cin_end.Text)
    clsTB_CINEMA.cin_num_end = IIf(Val(txt_cin_num_end.Text) = 0, Empty, txt_cin_num_end.Text)
    clsTB_CINEMA.cin_cmp_end = IIf(txt_cin_cmp_end.Text = "", Empty, txt_cin_cmp_end.Text)
    clsTB_CINEMA.cin_brr_end = IIf(txt_cin_brr_end.Text = "", Empty, txt_cin_brr_end.Text)
    clsTB_CINEMA.cin_cid_end = IIf(txt_cin_cid_end.Text = "", Empty, txt_cin_cid_end.Text)
    clsTB_CINEMA.cin_uf_end = IIf(txt_cin_uf_end.Text = "", Empty, txt_cin_uf_end.Text)
    clsTB_CINEMA.cin_cep_end = IIf(Trim(msk_cin_cep_end.ClipText) = "", Empty, msk_cin_cep_end.ClipText)
    clsTB_CINEMA.cin_tel = IIf(txt_cin_tel.Text = "", Empty, txt_cin_tel.Text)
    
    GravaDados = clsTB_CINEMA.Alterar()
    
    If Not GravaDados Then
        MsgBox "Erro ao gravar dados do Cinema!", vbCritical, App.ProductName
    Else
        MsgBox "Dados gravados com sucesso!", vbInformation, App.ProductName
        Call CarregaDados
    End If

GravaDados_Fim:
    Set clsTB_CINEMA = Nothing

End Function

Private Sub txt_cd_cin_KeyPress(KeyAscii As Integer)
    If Not IsNumeric(Chr(KeyAscii)) And KeyAscii <> 8 Then
        KeyAscii = 0
    End If
End Sub

Private Sub txt_cin_num_end_KeyPress(KeyAscii As Integer)
    If Not IsNumeric(Chr(KeyAscii)) And KeyAscii <> 8 Then
        KeyAscii = 0
    End If
End Sub

