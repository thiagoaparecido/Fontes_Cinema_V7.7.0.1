VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{234FCAFF-8D53-4DC2-9CD1-F90F4F6CB524}#32.0#0"; "Comandos.ocx"
Object = "{91C13016-45DE-491B-BAC0-37F755626532}#31.0#0"; "Spin.ocx"
Begin VB.Form frmParametros 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Cadastro - Parâmetros"
   ClientHeight    =   5640
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   9255
   Icon            =   "frmParametros.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5640
   ScaleWidth      =   9255
   ShowInTaskbar   =   0   'False
   Begin VB.Frame fraManut 
      Enabled         =   0   'False
      Height          =   4770
      Left            =   60
      TabIndex        =   17
      Top             =   60
      Width           =   9135
      Begin VB.TextBox flt_par_perc_cortesias 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   7560
         TabIndex        =   45
         Text            =   "0,000"
         Top             =   2205
         Width           =   1215
      End
      Begin VB.TextBox flt_par_perc_meias 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   7560
         TabIndex        =   44
         Text            =   "0,000"
         Top             =   1845
         Width           =   1215
      End
      Begin VB.TextBox flt_par_outros 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   7560
         TabIndex        =   43
         Text            =   "0,000"
         Top             =   1485
         Width           =   1215
      End
      Begin VB.TextBox flt_par_direitos_aut 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   7560
         TabIndex        =   42
         Text            =   "0,000"
         Top             =   1125
         Width           =   1215
      End
      Begin VB.TextBox flt_par_imposto_mun 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   7560
         TabIndex        =   41
         Text            =   "0,000"
         Top             =   765
         Width           =   1215
      End
      Begin VB.TextBox flt_par_custo_ingresso 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   7560
         TabIndex        =   40
         Text            =   "0,000"
         Top             =   405
         Width           =   1215
      End
      Begin VB.CheckBox chk_par_imp_MFIM 
         Alignment       =   1  'Right Justify
         Caption         =   "Altera descrição Ingresso"
         Height          =   315
         Left            =   6540
         TabIndex        =   39
         Top             =   2775
         Width           =   2130
      End
      Begin VB.TextBox txtMsg3 
         Height          =   315
         Left            =   3210
         MaxLength       =   40
         TabIndex        =   15
         Top             =   4290
         Width           =   4515
      End
      Begin VB.TextBox txtMsg2 
         Height          =   315
         Left            =   3210
         MaxLength       =   40
         TabIndex        =   14
         Top             =   3945
         Width           =   4515
      End
      Begin VB.TextBox txtMsg1 
         Height          =   315
         Left            =   3210
         MaxLength       =   40
         TabIndex        =   13
         Top             =   3615
         Width           =   4515
      End
      Begin VB.CheckBox chk_par_imp_endereco 
         Alignment       =   1  'Right Justify
         Caption         =   "Imprime Endereço"
         Height          =   315
         Left            =   4455
         TabIndex        =   11
         Top             =   2775
         Width           =   1635
      End
      Begin VB.CheckBox chk_par_imp_CNPJ 
         Alignment       =   1  'Right Justify
         Caption         =   "Imprime CNPJ"
         Height          =   315
         Left            =   2880
         TabIndex        =   9
         Top             =   2775
         Width           =   1335
      End
      Begin VB.CheckBox chk_par_imp_TCK 
         Alignment       =   1  'Right Justify
         Caption         =   "Imprime Canhoto"
         Height          =   315
         Left            =   4515
         TabIndex        =   12
         Top             =   3075
         Width           =   1575
      End
      Begin VB.CheckBox chk_par_imp_IE 
         Alignment       =   1  'Right Justify
         Caption         =   "Imprime Inscrição"
         Height          =   315
         Left            =   2640
         TabIndex        =   10
         Top             =   3075
         Width           =   1575
      End
      Begin VB.CheckBox chk_par_imp_lotacao 
         Alignment       =   1  'Right Justify
         Caption         =   "Imprime Lotação"
         Height          =   315
         Left            =   765
         TabIndex        =   8
         Top             =   3075
         Width           =   1515
      End
      Begin VB.CheckBox chk_par_imp_cod_barra 
         Alignment       =   1  'Right Justify
         Caption         =   "Imprime Código de Barra"
         Height          =   315
         Left            =   165
         TabIndex        =   7
         Top             =   2775
         Width           =   2115
      End
      Begin Spin.SpinNumber spn_par_tmp_ses 
         Height          =   315
         Left            =   3465
         TabIndex        =   0
         Top             =   195
         Width           =   735
         _ExtentX        =   1296
         _ExtentY        =   556
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Max             =   59
         Min             =   1
         Value           =   "1"
      End
      Begin MSComCtl2.DTPicker dtp_par_hora_max_ses 
         Height          =   315
         Left            =   3465
         TabIndex        =   1
         Top             =   570
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   556
         _Version        =   393216
         CustomFormat    =   "hh:mm"
         Format          =   57278466
         CurrentDate     =   38483
      End
      Begin MSComCtl2.DTPicker dtp_par_hora_limite_12 
         Height          =   315
         Left            =   3465
         TabIndex        =   2
         Top             =   930
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   556
         _Version        =   393216
         CustomFormat    =   "hh:mm"
         Format          =   57278466
         CurrentDate     =   38483
      End
      Begin MSComCtl2.DTPicker dtp_par_hora_limite_23 
         Height          =   315
         Left            =   3465
         TabIndex        =   3
         Top             =   1245
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   556
         _Version        =   393216
         CustomFormat    =   "hh:mm"
         Format          =   57278466
         CurrentDate     =   38483
      End
      Begin MSComCtl2.DTPicker dtp_par_hora_limite_34 
         Height          =   315
         Left            =   3465
         TabIndex        =   4
         Top             =   1575
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   556
         _Version        =   393216
         CustomFormat    =   "hh:mm"
         Format          =   57278466
         CurrentDate     =   38483
      End
      Begin MSComCtl2.DTPicker dtp_par_hora_limite_45 
         Height          =   315
         Left            =   3465
         TabIndex        =   5
         Top             =   1905
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   556
         _Version        =   393216
         CustomFormat    =   "hh:mm"
         Format          =   57278466
         CurrentDate     =   38483
      End
      Begin MSComCtl2.DTPicker dtp_par_hora_limite_56 
         Height          =   315
         Left            =   3465
         TabIndex        =   6
         Top             =   2235
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   556
         _Version        =   393216
         CustomFormat    =   "hh:mm"
         Format          =   57278466
         CurrentDate     =   38483
      End
      Begin VB.Label lblMsg1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Mensagem Ingresso linha 1:"
         Height          =   195
         Left            =   1125
         TabIndex        =   38
         Top             =   3615
         Width           =   1980
      End
      Begin VB.Label lblMsg2 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Mensagem Ingresso linha 2:"
         Height          =   195
         Left            =   1125
         TabIndex        =   37
         Top             =   3945
         Width           =   1980
      End
      Begin VB.Label lblMsg3 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Mensagem Ingresso linha 3:"
         Height          =   195
         Left            =   1125
         TabIndex        =   36
         Top             =   4290
         Width           =   1980
      End
      Begin VB.Line Line3 
         X1              =   5025
         X2              =   5025
         Y1              =   120
         Y2              =   2640
      End
      Begin VB.Line Line2 
         X1              =   0
         X2              =   9105
         Y1              =   3500
         Y2              =   3500
      End
      Begin VB.Line Line1 
         X1              =   0
         X2              =   9105
         Y1              =   2655
         Y2              =   2655
      End
      Begin VB.Label Label10 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Percentual máximo de cortesias:"
         Height          =   195
         Left            =   5220
         TabIndex        =   35
         Top             =   2220
         Width           =   2280
      End
      Begin VB.Label Label9 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "%"
         Height          =   195
         Left            =   8835
         TabIndex        =   34
         Top             =   2235
         Width           =   120
      End
      Begin VB.Label Label8 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "%"
         Height          =   195
         Left            =   8835
         TabIndex        =   33
         Top             =   1905
         Width           =   120
      End
      Begin VB.Label Labe8 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Percentual máximo de meias:"
         Height          =   195
         Left            =   5445
         TabIndex        =   32
         Top             =   1890
         Width           =   2055
      End
      Begin VB.Label lbl_par_hora_limite_56 
         Alignment       =   1  'Right Justify
         Caption         =   "Horário de mudança de faixa de preços 5 - 6:"
         Height          =   195
         Left            =   90
         TabIndex        =   31
         Top             =   2280
         Width           =   3255
      End
      Begin VB.Label lbl_par_hora_limite_45 
         Alignment       =   1  'Right Justify
         Caption         =   "Horário de mudança de faixa de preços 4 - 5:"
         Height          =   195
         Left            =   90
         TabIndex        =   30
         Top             =   1950
         Width           =   3255
      End
      Begin VB.Label lbl_par_hora_limite_34 
         Alignment       =   1  'Right Justify
         Caption         =   "Horário de mudança de faixa de preços 3 - 4:"
         Height          =   195
         Left            =   90
         TabIndex        =   29
         Top             =   1620
         Width           =   3255
      End
      Begin VB.Label lbl_par_hora_limite_23 
         Alignment       =   1  'Right Justify
         Caption         =   "Horário de mudança de faixa de preços 2 - 3:"
         Height          =   195
         Left            =   90
         TabIndex        =   28
         Top             =   1290
         Width           =   3255
      End
      Begin VB.Label Label7 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "%"
         Height          =   195
         Left            =   8835
         TabIndex        =   27
         Top             =   1560
         Width           =   120
      End
      Begin VB.Label Label6 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Outros:"
         Height          =   195
         Left            =   6990
         TabIndex        =   26
         Top             =   1515
         Width           =   510
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "%"
         Height          =   195
         Left            =   8835
         TabIndex        =   25
         Top             =   1200
         Width           =   120
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "%"
         Height          =   195
         Left            =   8835
         TabIndex        =   24
         Top             =   840
         Width           =   120
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Direitos Autorais:"
         Height          =   195
         Left            =   6315
         TabIndex        =   23
         Top             =   1155
         Width           =   1185
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Imposto Municipal:"
         Height          =   195
         Left            =   6195
         TabIndex        =   22
         Top             =   795
         Width           =   1320
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Custo do Ingresso:"
         Height          =   195
         Left            =   6195
         TabIndex        =   21
         Top             =   435
         Width           =   1320
      End
      Begin VB.Label lblCNPJ 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Hora da Última Sessão:"
         Height          =   195
         Left            =   1665
         TabIndex        =   20
         Top             =   600
         Width           =   1665
      End
      Begin VB.Label lbl_par_hora_limite_12 
         Alignment       =   1  'Right Justify
         Caption         =   "Horário de mudança de faixa de preços 1 - 2:"
         Height          =   195
         Left            =   90
         TabIndex        =   19
         Top             =   960
         Width           =   3255
      End
      Begin VB.Label lblNome 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Tempo de Venda ( em minutos ):"
         Height          =   195
         Left            =   1035
         TabIndex        =   18
         Top             =   210
         Width           =   2295
      End
   End
   Begin Comandos.cmdComandos cmdComandos 
      Align           =   2  'Align Bottom
      Height          =   765
      Left            =   0
      TabIndex        =   16
      Top             =   4875
      Width           =   9255
      _ExtentX        =   16325
      _ExtentY        =   1349
      EnabledAltera   =   -1  'True
      VisibleNovo     =   0   'False
      VisibleExclui   =   0   'False
   End
End
Attribute VB_Name = "frmParametros"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim miSelStart          As Integer
Dim miLastKey           As Integer
Dim mbFocus             As Boolean

Dim msLostFormat        As String

Dim msSeparadorDecimal  As String
Dim msMilhar            As String
Dim miDigitosGrupo      As Integer
Dim miPosDecimal        As Integer
Dim m_QtdeDecimais      As Integer
Dim m_QtdeInteiros      As Integer

Const m_def_QtdeDecimais = 3
Const m_def_QtdeInteiros = 9

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
            
            Call CarregaParametros
    
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
        MsgBox "Erro ao carregar Parâmetros!", vbCritical, "Erro"
    End If
    
    Dim log As New clsLog
    
    Set log.ConexaoADO = dbConnect
    log.usu_nm = strUsuario
    log.slg_descricao = "Entrou na tela de " & Me.Caption & " do módulo " & App.ProductName
    
    log.insereLog
    
End Sub
Private Function CarregaDados() As Boolean

    On Error GoTo CarregaDados_Erro
    
    Dim oRs As New ADODB.Recordset
    Dim clsTB_PARAMETRO As New Cine2005.clsTB_PARAMETRO
    
    Set clsTB_PARAMETRO.ConexaoADO = dbConnect
    
    If Not clsTB_PARAMETRO.Selecionar(oRs) Then
        Exit Function
    End If
    
    If Not oRs.EOF() Then
        spn_par_tmp_ses.Value = oRs.Fields("par_tmp_ses")
        dtp_par_hora_max_ses.Value = oRs.Fields("par_hora_max_ses")
        
        dtp_par_hora_limite_12.Value = oRs.Fields("par_hora_limite12")
        dtp_par_hora_limite_23.Value = oRs.Fields("par_hora_limite23")
        dtp_par_hora_limite_34.Value = oRs.Fields("par_hora_limite34")
        dtp_par_hora_limite_45.Value = oRs.Fields("par_hora_limite45")
        dtp_par_hora_limite_56.Value = oRs.Fields("par_hora_limite56")
        
        chk_par_imp_cod_barra.Value = IIf(IsNull(oRs.Fields("par_imp_cod_barra")), 0, IIf(oRs.Fields("par_imp_cod_barra"), 1, 0))
        chk_par_imp_lotacao.Value = IIf(IsNull(oRs.Fields("par_imp_lotacao")), 0, IIf(oRs.Fields("par_imp_lotacao"), 1, 0))
        
        If IsNull(oRs.Fields("par_imp_endereco")) Then
            chk_par_imp_endereco.Value = False
        Else
            chk_par_imp_endereco.Value = IIf(oRs.Fields("par_imp_endereco"), vbChecked, vbUnchecked)
        End If
        
        If IsNull(oRs.Fields("par_imp_CNPJ")) Then
            chk_par_imp_CNPJ.Value = False
        Else
            chk_par_imp_CNPJ.Value = IIf(oRs.Fields("par_imp_CNPJ"), vbChecked, vbUnchecked)
        End If
        
        If IsNull(oRs.Fields("par_imp_IE")) Then
            chk_par_imp_IE.Value = False
        Else
            chk_par_imp_IE.Value = IIf(oRs.Fields("par_imp_IE"), vbChecked, vbUnchecked)
        End If
        
        If IsNull(oRs.Fields("par_imp_tck")) Then
            chk_par_imp_TCK.Value = False
        Else
            chk_par_imp_TCK.Value = IIf(oRs.Fields("par_imp_tck"), vbChecked, vbUnchecked)
        End If
        
        flt_par_custo_ingresso.Text = Format(IIf(IsNull(oRs.Fields("par_custo_ingresso")), 0, oRs.Fields("par_custo_ingresso")), "#0.000")
        flt_par_direitos_aut.Text = Format(IIf(IsNull(oRs.Fields("par_direitos_aut")), 0, oRs.Fields("par_direitos_aut")), "#0.000")
        flt_par_imposto_mun.Text = Format(IIf(IsNull(oRs.Fields("par_imposto_mun")), 0, oRs.Fields("par_imposto_mun")), "#0.000")
        flt_par_outros.Text = Format(IIf(IsNull(oRs.Fields("par_outros")), 0, oRs.Fields("par_outros")), "#0.000")
        flt_par_perc_meias.Text = Format(IIf(IsNull(oRs.Fields("par_perc_meias")), 0, oRs.Fields("par_perc_meias")), "#0.000")
        flt_par_perc_cortesias.Text = Format(IIf(IsNull(oRs.Fields("par_perc_cortesias")), 0, oRs.Fields("par_perc_cortesias")), "#0.000")
        
        txtMsg1.Text = IIf(IsNull(oRs.Fields("par_msg1")), "", oRs.Fields("par_msg1"))
        txtMsg2.Text = IIf(IsNull(oRs.Fields("par_msg2")), "", oRs.Fields("par_msg2"))
        txtMsg3.Text = IIf(IsNull(oRs.Fields("par_msg3")), "", oRs.Fields("par_msg3"))
        
        If IsNull(oRs.Fields("par_imp_MFIM")) Then
            chk_par_imp_MFIM.Value = False
        Else
            chk_par_imp_MFIM.Value = IIf(oRs.Fields("par_imp_MFIM"), vbChecked, vbUnchecked)
        End If
        
    End If
    
    CarregaDados = True
    
CarregaDados_Erro:
    If oRs.State = 1 Then oRs.Close
    
    Set oRs = Nothing
    Set clsTB_PARAMETRO = Nothing

End Function
Private Function GravaDados() As Boolean

    On Error GoTo GravaDados_Fim
    
    Dim sMensErro As String
    
    If sMensErro <> "" Then
        MsgBox sMensErro, vbInformation, "Alerta"
        GoTo GravaDados_Fim
    End If
    
    Dim clsTB_PARAMETRO As New Cine2005.clsTB_PARAMETRO
    
    Set clsTB_PARAMETRO.ConexaoADO = dbConnect
    
    clsTB_PARAMETRO.par_tmp_ses = spn_par_tmp_ses.Value
    clsTB_PARAMETRO.par_hora_max_ses = Format(dtp_par_hora_max_ses.Value, "hh:mm")
    
    clsTB_PARAMETRO.par_hora_limite = Format("00:00", "hh:mm")
    
    clsTB_PARAMETRO.par_hora_limite12 = Format(dtp_par_hora_limite_12.Value, "hh:mm")
    clsTB_PARAMETRO.par_hora_limite23 = Format(dtp_par_hora_limite_23.Value, "hh:mm")
    clsTB_PARAMETRO.par_hora_limite34 = Format(dtp_par_hora_limite_34.Value, "hh:mm")
    clsTB_PARAMETRO.par_hora_limite45 = Format(dtp_par_hora_limite_45.Value, "hh:mm")
    clsTB_PARAMETRO.par_hora_limite56 = Format(dtp_par_hora_limite_56.Value, "hh:mm")
    
    clsTB_PARAMETRO.par_imp_cod_barra = chk_par_imp_cod_barra.Value
    clsTB_PARAMETRO.par_imp_lotacao = chk_par_imp_lotacao.Value
    clsTB_PARAMETRO.par_imp_endereco = chk_par_imp_endereco.Value
    clsTB_PARAMETRO.par_imp_CNPJ = chk_par_imp_CNPJ.Value
    clsTB_PARAMETRO.par_imp_IE = chk_par_imp_IE.Value
    clsTB_PARAMETRO.par_imp_tck = chk_par_imp_TCK.Value
    clsTB_PARAMETRO.par_imposto_mun = flt_par_imposto_mun.Text
    clsTB_PARAMETRO.par_custo_ingresso = flt_par_custo_ingresso.Text
    clsTB_PARAMETRO.par_direitos_aut = flt_par_direitos_aut.Text
    clsTB_PARAMETRO.par_outros = flt_par_outros.Text
    clsTB_PARAMETRO.par_perc_meias = flt_par_perc_meias.Text
    clsTB_PARAMETRO.par_perc_cortesias = flt_par_perc_cortesias.Text
    
    clsTB_PARAMETRO.par_msg1 = txtMsg1.Text
    clsTB_PARAMETRO.par_msg2 = txtMsg2.Text
    clsTB_PARAMETRO.par_msg3 = txtMsg3.Text
    
    clsTB_PARAMETRO.par_imp_MFIM = chk_par_imp_MFIM.Value
    
    GravaDados = clsTB_PARAMETRO.Alterar()

    If Not GravaDados Then
        MsgBox "Erro ao gravar Parâmetros!", vbCritical, "Erro"
    Else
        intTempoEntreSessoes = clsTB_PARAMETRO.par_tmp_ses
        dtHoraMaxSessao = clsTB_PARAMETRO.par_hora_max_ses
        dtHoraLimite = clsTB_PARAMETRO.par_hora_limite
        MsgBox "Dados gravados com sucesso!", vbInformation, App.ProductName
    End If

GravaDados_Fim:
    Set clsTB_PARAMETRO = Nothing

End Function

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    Dim log As New clsLog
    
    Set log.ConexaoADO = dbConnect
    log.usu_nm = strUsuario
    log.slg_descricao = "Saiu da tela de " & Me.Caption & " do módulo " & App.ProductName
    
    log.insereLog
End Sub

Private Sub f_Change(Controle As TextBox)
    On Error Resume Next
    
    Dim iPosDecimal As Integer
    Dim iTamValor As Integer
    Dim sValor As String
    
    If Not mbFocus Then
        sValor = Controle.Text
        If InStr(Len(Controle.Text) - m_QtdeDecimais - 1, Controle.Text, ",") <> 0 And msSeparadorDecimal = "." Then
            iPosDecimal = InStr(Len(Controle.Text) - m_QtdeDecimais - 1, Controle.Text, ",")
            Mid(sValor, iPosDecimal) = msSeparadorDecimal
            'Controle.Text = sValor
        ElseIf InStr(Len(Controle.Text) - m_QtdeDecimais - 1, Controle.Text, ".") <> 0 And msSeparadorDecimal = "," Then
            iPosDecimal = InStr(Len(Controle.Text) - m_QtdeDecimais - 1, ".")
            Mid(sValor, iPosDecimal) = msSeparadorDecimal
            'Controle.Text = sValor
        End If
        
        Controle.Text = Format(CDbl(sValor), "0.00")
        
        Exit Sub
    End If
        
    mbFocus = False
    
    iTamValor = Len(Controle.Text)
    
    If Not IsNumeric(Controle.Text) Then
        Controle.Text = "0" & msSeparadorDecimal & String$(m_QtdeDecimais, "0")
    End If

    Controle.Text = Format$(Val(CStr(fusConverteValor(Controle.Text))), "0." & String$(m_QtdeDecimais, "0"))
    
    If Len(Controle.Text) <> iTamValor Then
        If miPosDecimal > miSelStart And iTamValor <> 1 Then
            miSelStart = miSelStart - 1
        End If
        iPosDecimal = InStr(Controle.Text, msSeparadorDecimal)
        miPosDecimal = iPosDecimal
    End If
    
    mbFocus = True
    
    iPosDecimal = InStr(Controle.Text, msSeparadorDecimal)
    
    Controle.Text = Left$(Controle.Text, iPosDecimal) & Mid$(Controle.Text, iPosDecimal + 1, m_QtdeDecimais)
    
    If Controle.Text = "" Then
        mbFocus = False
        Controle.Text = "0" & msSeparadorDecimal & String(m_QtdeDecimais, "0")
        iPosDecimal = InStr(Controle.Text, msSeparadorDecimal)
        miPosDecimal = iPosDecimal
        miSelStart = 1
        mbFocus = True
    End If
        
    If Controle.Text = "0" & msSeparadorDecimal & String(m_QtdeDecimais, "0") Then
        If miLastKey = vbKeyDelete Or miLastKey = vbKeyBack Then
            Controle.SelStart = iPosDecimal - 1
            Exit Sub
        End If
    End If
     
    If miLastKey = vbKeyBack Then
        Controle.SelStart = miSelStart - 1
    ElseIf miLastKey = vbKeyDelete Then
        Controle.SelStart = miSelStart
    Else
        Controle.SelStart = miSelStart + 1
    End If
    
    Controle.Text = Format(CDbl(Controle.Text), "0.00")
    
End Sub

Private Sub f_GotFocus(Controle As TextBox)
    Dim iCount As Integer
    Dim sValor As String
    Dim sChar As String
    Dim n As Integer

    Controle.MaxLength = m_QtdeDecimais + m_QtdeInteiros + 1
    n = Len(Controle.Text)
    For iCount = 1 To n
        sChar = Mid$(Controle.Text, iCount, 1)
        If IsNumeric(sChar) Or sChar = msSeparadorDecimal Then
            sValor = sValor & sChar
        End If
    Next

    Controle.Text = sValor
    
    Controle.SelStart = 0
    Controle.SelLength = Len(Controle.Text)

    mbFocus = True
End Sub

Private Sub f_KeyDown(Controle As TextBox, ByRef KeyCode As Integer, ByRef Shift As Integer)
    If Controle.Locked Then
        KeyCode = 0
        Shift = 0
    End If
    
    miPosDecimal = InStr(Controle.Text, msSeparadorDecimal)
    
    miSelStart = Controle.SelStart
    
    If KeyCode = vbKeyDelete And miSelStart = miPosDecimal - 1 Then
        KeyCode = 0
        Controle.SelStart = miPosDecimal - 1
    End If
    
    miSelStart = Controle.SelStart
    miLastKey = KeyCode
End Sub

Private Sub f_KeyPress(Controle As TextBox, ByRef KeyAscii As Integer)
    If Controle.Locked Then
        KeyAscii = 0
    End If
    
    Const vbKeyMenos = 45
    
    miSelStart = Controle.SelStart
    miPosDecimal = InStr(Controle.Text, msSeparadorDecimal)
    
    miLastKey = KeyAscii
    
    If KeyAscii = vbKeyMenos And InStr(Controle.Text, Chr(vbKeyMenos)) = 0 Then
        If Len(Controle.Text) = Controle.MaxLength Then
            KeyAscii = 0
            Exit Sub
        End If
        Controle.Text = "-" & Controle.Text
        KeyAscii = 0
    End If
    
    If KeyAscii <> 48 Then
                'KeyAscii <> Asc(msSeparadorDecimal) And
        If Not IsNumeric(Chr$(KeyAscii)) And _
                KeyAscii <> Asc(",") And _
                KeyAscii <> Asc(".") And _
                KeyAscii <> vbKeyEscape And _
                KeyAscii <> vbKeyReturn And _
                KeyAscii <> vbKeyBack Then
            KeyAscii = 0
        ElseIf IsNumeric(Chr$(KeyAscii)) And miSelStart = Len(Controle.Text) Then
            KeyAscii = 0
        ElseIf IsNumeric(Chr$(KeyAscii)) And miSelStart >= miPosDecimal Then
            If Controle.SelLength = 0 Then
                Controle.Text = Left$(Controle.Text, miSelStart) & Chr$(KeyAscii) & Right$(Controle.Text, Len(Controle.Text) - miSelStart - 1)
                Controle.SelStart = miSelStart + 1
            Else
                Controle.Text = Left$(Controle.Text, miPosDecimal) & Chr$(KeyAscii) & String$(Controle.SelLength - 1, "0") & Right$(Controle.Text, m_QtdeDecimais - Controle.SelLength)
            End If
            KeyAscii = 0
        End If
    End If
    
    'If miPosDecimal >= 1 And (KeyAscii = Asc(msSeparadorDecimal) Or KeyAscii = Asc(msSeparadorDecimal)) Then
    If miPosDecimal >= 1 And (KeyAscii = Asc(",") Or KeyAscii = Asc(".")) Then
        KeyAscii = 0
        Controle.SelStart = miPosDecimal
    End If
    
    If KeyAscii = vbKeyBack And miSelStart = miPosDecimal Then
        KeyAscii = 0
        Controle.SelStart = miPosDecimal - 1
    End If
End Sub

Private Sub f_LostFocus(Controle As TextBox)
    Dim iCount As Integer
    Dim sValor As String
    Dim sChar As String
    Dim n     As Integer
    
    mbFocus = False

    Controle.MaxLength = Len(msLostFormat)
    n = Len(Controle.Text)
    For iCount = 1 To n
        sChar = Mid$(Controle.Text, iCount, 1)
        If IsNumeric(sChar) Or sChar = msSeparadorDecimal Then
            sValor = sValor & sChar
        End If
    Next
    
    Controle.Text = Format$(Val(fusConverteValor(sValor)), msLostFormat)
End Sub

Private Function fusConverteValor(ByVal sValor As String) As String

    Dim iPosDecimal As Integer
    
    iPosDecimal = InStr(sValor, msSeparadorDecimal)
    
    If iPosDecimal <> 0 Then
        Mid(sValor, iPosDecimal, 1) = "."
    End If
    fusConverteValor = sValor
    
End Function

Private Function fusLostFormat() As String

    Dim sInteiros As String
    Dim sDecimais As String
    
    Dim iDig As Integer
    Dim iPos As Integer
    Dim sFormato As String
    
    sFormato = String$(m_QtdeInteiros - 1, "#") & "0"
    
    For iPos = Len(sFormato) To 1 Step -1
        sInteiros = Mid$(sFormato, iPos, 1) & sInteiros
        iDig = iDig + 1
        If iDig = miDigitosGrupo Then
            sInteiros = "," & sInteiros
            iDig = 0
        End If
    Next
    
    If Left$(sInteiros, 1) = "," Then
        sInteiros = Right$(sInteiros, Len(sInteiros) - 1)
    End If
    
    sDecimais = String(m_QtdeDecimais, "0")
    
    If Len(sDecimais) > 0 Then
        sDecimais = "." & sDecimais
    End If
    
    fusLostFormat = sInteiros & sDecimais
    
End Function

