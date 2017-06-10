VERSION 5.00
Object = "{234FCAFF-8D53-4DC2-9CD1-F90F4F6CB524}#32.0#0"; "Comandos.ocx"
Object = "{1753B334-0119-4B34-9134-D2B3CD181550}#36.0#0"; "Combo.ocx"
Object = "{91C13016-45DE-491B-BAC0-37F755626532}#31.0#0"; "Spin.ocx"
Object = "{B02F3647-766B-11CE-AF28-C3A2FBE76A13}#3.0#0"; "SPR32X30.ocx"
Begin VB.Form frmPoltronas 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Cadastro - Poltronas"
   ClientHeight    =   8520
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   10260
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   8520
   ScaleWidth      =   10260
   Begin VB.Frame fraPoltronas 
      Caption         =   "Cadastro Poltronas"
      Height          =   7665
      Left            =   45
      TabIndex        =   0
      Top             =   60
      Width           =   10170
      Begin VB.TextBox txtTotPoltronas 
         Height          =   315
         Left            =   9120
         Locked          =   -1  'True
         TabIndex        =   26
         Top             =   210
         Width           =   825
      End
      Begin VB.TextBox txtLotacao 
         Height          =   315
         Left            =   6420
         Locked          =   -1  'True
         TabIndex        =   25
         Top             =   210
         Width           =   825
      End
      Begin VB.Frame fraManutencao 
         Height          =   6930
         Left            =   135
         TabIndex        =   4
         Top             =   570
         Width           =   9900
         Begin VB.Frame fraLeganda 
            Caption         =   "Legenda"
            Height          =   1995
            Left            =   7155
            TabIndex        =   8
            Top             =   150
            Width           =   2595
            Begin VB.Shape shp_F4 
               FillColor       =   &H0000FFFF&
               FillStyle       =   0  'Solid
               Height          =   195
               Left            =   405
               Shape           =   1  'Square
               Top             =   1170
               Width           =   195
            End
            Begin VB.Label Label1 
               Appearance      =   0  'Flat
               Caption         =   "F3 - Poltrona Especial"
               ForeColor       =   &H80000008&
               Height          =   195
               Left            =   630
               TabIndex        =   29
               Top             =   780
               Width           =   1665
            End
            Begin VB.Shape shp_F3 
               FillColor       =   &H000080FF&
               FillStyle       =   0  'Solid
               Height          =   195
               Left            =   405
               Shape           =   1  'Square
               Top             =   780
               Width           =   195
            End
            Begin VB.Label Label6 
               Appearance      =   0  'Flat
               Caption         =   "F2 - Poltrona Ativa"
               ForeColor       =   &H80000008&
               Height          =   195
               Left            =   630
               TabIndex        =   24
               Top             =   390
               Width           =   1665
            End
            Begin VB.Shape shp_F5 
               FillColor       =   &H00C0C0C0&
               FillStyle       =   0  'Solid
               Height          =   195
               Left            =   405
               Shape           =   1  'Square
               Top             =   1560
               Width           =   195
            End
            Begin VB.Label Label7 
               Appearance      =   0  'Flat
               Caption         =   "F5 - Sem Poltrona"
               ForeColor       =   &H80000008&
               Height          =   195
               Left            =   630
               TabIndex        =   23
               Top             =   1560
               Width           =   1665
            End
            Begin VB.Label Label8 
               Appearance      =   0  'Flat
               Caption         =   "F4 - Poltrona Inativa"
               ForeColor       =   &H80000008&
               Height          =   195
               Left            =   630
               TabIndex        =   22
               Top             =   1170
               Width           =   1665
            End
            Begin VB.Shape shp_F2 
               FillColor       =   &H00C0C000&
               FillStyle       =   0  'Solid
               Height          =   195
               Left            =   405
               Shape           =   1  'Square
               Top             =   390
               Width           =   195
            End
         End
         Begin VB.Frame fraOriebtacaoHoriz 
            Caption         =   "Numeração Fila/Coluna"
            Height          =   1995
            Left            =   2895
            TabIndex        =   7
            Top             =   150
            Width           =   4095
            Begin VB.Frame Frame2 
               Height          =   855
               Left            =   2100
               TabIndex        =   18
               Top             =   990
               Width           =   1725
               Begin VB.OptionButton optOrientacaoVert 
                  Caption         =   "Cima/Baixo"
                  Height          =   330
                  Index           =   0
                  Left            =   75
                  TabIndex        =   20
                  Top             =   150
                  Width           =   1455
               End
               Begin VB.OptionButton optOrientacaoVert 
                  Caption         =   "Baixo/Cima"
                  Height          =   330
                  Index           =   1
                  Left            =   75
                  TabIndex        =   19
                  Top             =   420
                  Width           =   1455
               End
            End
            Begin VB.Frame Frame1 
               Height          =   855
               Left            =   240
               TabIndex        =   16
               Top             =   990
               Width           =   1755
               Begin VB.OptionButton optOrientacaoHoriz 
                  Caption         =   "Esquerda/Direita"
                  Height          =   195
                  Index           =   0
                  Left            =   75
                  TabIndex        =   21
                  Top             =   240
                  Width           =   1590
               End
               Begin VB.OptionButton optOrientacaoHoriz 
                  Caption         =   "Direita/Esquerda"
                  Height          =   195
                  Index           =   1
                  Left            =   60
                  TabIndex        =   17
                  Top             =   510
                  Width           =   1590
               End
            End
            Begin VB.TextBox txtNumPoltronas 
               Height          =   315
               Left            =   1950
               TabIndex        =   14
               Top             =   675
               Width           =   795
            End
            Begin VB.ComboBox cboNumeracao 
               Height          =   315
               Left            =   210
               TabIndex        =   13
               Text            =   "Combo1"
               Top             =   270
               Width           =   3660
            End
            Begin VB.Label lblNumPriCol 
               AutoSize        =   -1  'True
               Caption         =   "Número 1a. Poltrona:"
               Height          =   195
               Left            =   240
               TabIndex        =   15
               Top             =   720
               Width           =   1500
            End
         End
         Begin VB.Frame fraIdentificacao 
            Caption         =   "Identificação"
            Height          =   1995
            Left            =   120
            TabIndex        =   6
            Top             =   150
            Width           =   2625
            Begin Spin.SpinNumber spnColunas 
               Height          =   315
               Left            =   1095
               TabIndex        =   10
               Top             =   1170
               Width           =   1080
               _ExtentX        =   1905
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
               Max             =   52
               Value           =   "0"
            End
            Begin Spin.SpinNumber spnFilas 
               Height          =   315
               Left            =   1110
               TabIndex        =   9
               Top             =   615
               Width           =   1080
               _ExtentX        =   1905
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
               Max             =   52
               Value           =   "0"
            End
            Begin VB.Label lblColunas 
               AutoSize        =   -1  'True
               Caption         =   "Colunas:"
               Height          =   195
               Left            =   360
               TabIndex        =   12
               Top             =   1170
               Width           =   615
            End
            Begin VB.Label lblFilas 
               AutoSize        =   -1  'True
               Caption         =   "Filas:"
               Height          =   195
               Left            =   375
               TabIndex        =   11
               Top             =   615
               Width           =   360
            End
         End
         Begin FPSpread.vaSpread gridPoltronas 
            Height          =   4515
            Left            =   135
            TabIndex        =   5
            Top             =   2250
            Width           =   9600
            _Version        =   196608
            _ExtentX        =   16933
            _ExtentY        =   7964
            _StockProps     =   64
            DisplayColHeaders=   0   'False
            DisplayRowHeaders=   0   'False
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            MaxCols         =   20
            MaxRows         =   20
            SpreadDesigner  =   "frmPoltronas.frx":0000
            UnitType        =   2
            UserResize      =   0
         End
      End
      Begin Combo.cboCodDesc ccd_sal_cd 
         Height          =   315
         Left            =   555
         TabIndex        =   2
         Top             =   210
         Width           =   4575
         _ExtentX        =   8070
         _ExtentY        =   556
         NomeTabela      =   "tb_sala"
         NomeCampoCodigo =   "sal_cd"
         NomeCampoDescricao=   "sal_desc"
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
         Filtro          =   "sal_dt_des is null"
         MostraBotaoAtualiza=   0   'False
      End
      Begin VB.Label lblLotacao 
         Caption         =   "Lotação:"
         Height          =   255
         Left            =   5640
         TabIndex        =   28
         Top             =   255
         Width           =   630
      End
      Begin VB.Label lblTotPoltronas 
         Caption         =   "Total de Poltronas:"
         Height          =   210
         Left            =   7620
         TabIndex        =   27
         Top             =   255
         Width           =   1350
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Sala:"
         Height          =   195
         Left            =   135
         TabIndex        =   3
         Top             =   270
         Width           =   360
      End
   End
   Begin Comandos.cmdComandos cmdComandos 
      Align           =   2  'Align Bottom
      Height          =   765
      Left            =   0
      TabIndex        =   1
      Top             =   7755
      Width           =   10260
      _ExtentX        =   18098
      _ExtentY        =   1349
      EnabledNovo     =   0   'False
      VisibleNovo     =   0   'False
      VisibleExclui   =   0   'False
   End
End
Attribute VB_Name = "frmPoltronas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Const BLACK = &H80000008
Const DESELECT_BLOCK = 14

Dim dummy

Private Sub cboNumeracao_Click()
    If cboNumeracao.ListIndex = 0 Then
       txtNumPoltronas.Text = "0001"
    ElseIf cboNumeracao.ListIndex = 1 Then
       txtNumPoltronas.Text = "A01"
    ElseIf cboNumeracao.ListIndex = 2 Then
       txtNumPoltronas.Text = "01A"
    End If
       
    Gera_Numeracao

End Sub

Private Sub ccd_sal_cd_AfterProcuraClick()
    If ccd_sal_cd.Descricao <> "" Then
        cmdComandos.EnabledAltera = True
    End If
End Sub

Private Sub ccd_sal_cd_Change()
    CarregaDados
End Sub

Private Sub cmdComandos_Click(ByVal iButtonClicked As Comandos.eButtonClicked, Cancel As Boolean)
    Select Case iButtonClicked
        Case ButtonAltera
            Call HabilitaManut(True)
            
        Case ButtonGrava
            If Grava() Then
                Call HabilitaManut(False)
            Else
                Cancel = True
            End If
    
        Case ButtonFecha
            Unload Me
            
        Case ButtonCancela
            Call HabilitaManut(False)
            CarregaDados
    End Select
End Sub

Private Sub Form_Load()
    Set ccd_sal_cd.ConexaoADO = dbConnect
    
    Call HabilitaManut(False)
    Call inicboNumeracao

    Dim CineAx As New CineAux.CineAx
    Call CineAx.CentralizaTela(MDIFrmAdmCine, Me)
    Set CineAx = Nothing

End Sub

Private Sub HabilitaManut(ByVal bHabilita As Boolean)
    fraManutencao.Enabled = bHabilita
    ccd_sal_cd.Enabled = Not bHabilita
End Sub

Private Sub inicboNumeracao()
    cboNumeracao.Clear
    cboNumeracao.AddItem "Sequencial"
    cboNumeracao.AddItem "Letra x Número"
    cboNumeracao.AddItem "Número x Letra"
    cboNumeracao.ListIndex = 0
End Sub

Private Function verificaTela()
    verificaTela = False
    If Val(txtLotacao.Text) < Val(txtTotPoltronas.Text) Then
        MsgBox "Número de Poltronas Ativas/Inativas/Especiais maior que Lotação da sala"
        Exit Function
    End If
    verificaTela = True
End Function

Private Function Grava() As Boolean
    Grava = False
    'If !verificaTela() Then
    '    Exit Function
    'End If

    Dim clsTB_POLTRONAS As New Cine2005.clsTB_POLTRONAS
    
    On Error GoTo Grava_Erro
    
    Set clsTB_POLTRONAS.ConexaoADO = dbConnect
    
    clsTB_POLTRONAS.sal_cd = ccd_sal_cd.codigo
    clsTB_POLTRONAS.pol_num_filas = spnFilas.Value
    clsTB_POLTRONAS.pol_num_colunas = spnColunas.Value
    clsTB_POLTRONAS.pol_tp_numeracao = cboNumeracao.ListIndex
    clsTB_POLTRONAS.pol_num_pri_col = txtNumPoltronas.Text
    clsTB_POLTRONAS.pol_num_horiz = optOrientacaoHoriz(0).Value
    clsTB_POLTRONAS.pol_num_vert = optOrientacaoVert(0).Value
    clsTB_POLTRONAS.pol_poltronas = Val(txtTotPoltronas.Text)
    clsTB_POLTRONAS.pol_mat_poltr = Monta_MatrPoltr()

    If Not clsTB_POLTRONAS.Alterar() Then
        MsgBox "Não foi possível salvar poltronas!" & vbCrLf & clsTB_POLTRONAS.MensagemErro, vbInformation, App.ProductName
        GoTo Grava_Fim
    End If

    Grava = True
    GoTo Grava_Fim

Grava_Erro:
    MsgBox "Erro de execução! 'Grava/frmPoltronas'", vbCritical, App.ProductName
    
Grava_Fim:
    Set clsTB_POLTRONAS = Nothing
End Function

Private Sub gridPoltronas_EditMode(ByVal Col As Long, ByVal Row As Long, ByVal mode As Integer, ByVal ChangeMade As Boolean)
    gridPoltronas.EditMode = False
End Sub

Private Sub gridPoltronas_KeyDown(KeyCode As Integer, Shift As Integer)
    
    'Mouse = Ampulheta
    Screen.MousePointer = Hourglass
    
    If KeyCode = vbKeyF2 Then
       Pinta_Celulas vbKeyF2
    ElseIf KeyCode = vbKeyF3 Then
       Pinta_Celulas vbKeyF3
    ElseIf KeyCode = vbKeyF4 Then
       Pinta_Celulas vbKeyF4
    ElseIf KeyCode = vbKeyF5 Then
       Pinta_Celulas vbKeyF5
    ElseIf KeyCode <> vbKeyLeft And KeyCode <> vbKeyUp And KeyCode <> vbKeyRight And KeyCode <> vbKeyDown Then
       KeyCode = 0
    End If
      
    gridPoltronas.EditMode = False

    'Mouse = SETA
    Screen.MousePointer = Arrow

End Sub

Private Sub gridPoltronas_KeyPress(KeyAscii As Integer)
    
    KeyAscii = 0
    gridPoltronas.EditMode = False

End Sub

Private Sub Pinta_Celulas(nKey As Integer)
    Dim nContL   As Integer
    Dim nContC   As Integer
    Dim nLimMinL As Integer
    Dim nLimMaxL As Integer
    Dim nLimMinC As Integer
    Dim nLimMaxC As Integer

    If gridPoltronas.IsBlockSelected Then
       nLimMinL = gridPoltronas.SelBlockRow
       nLimMaxL = gridPoltronas.SelBlockRow2
       nLimMinC = gridPoltronas.SelBlockCol
       nLimMaxC = gridPoltronas.SelBlockCol2
       gridPoltronas.Action = DESELECT_BLOCK
    Else
       nLimMinL = gridPoltronas.ActiveRow
       nLimMaxL = gridPoltronas.ActiveRow
       nLimMinC = gridPoltronas.ActiveCol
       nLimMaxC = gridPoltronas.ActiveCol
    End If

    For nContL = nLimMinL To nLimMaxL
        For nContC = nLimMinC To nLimMaxC
            gridPoltronas.Row = nContL
            gridPoltronas.Col = nContC
            gridPoltronas.Row2 = nContL
            gridPoltronas.Col2 = nContC
            gridPoltronas.BlockMode = True
            
            If nKey = vbKeyF2 Then
               gridPoltronas.BackColor = shp_F2.FillColor
            ElseIf nKey = vbKeyF3 Then
               gridPoltronas.BackColor = shp_F3.FillColor
            ElseIf nKey = vbKeyF4 Then
               gridPoltronas.BackColor = shp_F4.FillColor
            ElseIf nKey = vbKeyF5 Then
               gridPoltronas.BackColor = shp_F5.FillColor
            End If
    
            gridPoltronas.Text = ""
            gridPoltronas.BlockMode = False
        Next
    Next
    
    Gera_Numeracao

    dummy = Conta_Poltr()

    If dummy > Val(txtLotacao.Text) Then
       MsgBox "Número de Poltronas Ativas/Inativas/Especiais maior que Lotação da sala"
    End If
End Sub

Private Sub Gera_Numeracao()
    If cboNumeracao.ListIndex = 0 Then
       Gera_NumCadSeq Val(txtNumPoltronas.Text)
    ElseIf cboNumeracao.ListIndex = 1 Then
       Gera_NumcadLN Asc(Mid$(txtNumPoltronas.Text, 1, 1)), Val(Mid$(txtNumPoltronas.Text, 2, 2))
    ElseIf cboNumeracao.ListIndex = 2 Then
       Gera_NumCadNL Val(Mid$(txtNumPoltronas.Text, 1, 2)), Asc(Mid$(txtNumPoltronas.Text, 3, 1))
    End If
End Sub

Private Sub Gera_NumCadSeq(nIndN As Integer)
    
    'nIndN Controla a parte número das poltronas

    Dim nContL   As Integer
    Dim nContC   As Integer
    Dim nLimMinC As Integer
    Dim nLimMaxC As Integer
    Dim nStepC   As Integer
    Dim nLimMinL As Integer
    Dim nLimMaxL As Integer
    Dim nStepL   As Integer

    'Direção da numeração

    'De cima para baixo
    If optOrientacaoVert(0).Value Then
       nLimMinL = 1
       nLimMaxL = gridPoltronas.MaxRows
       nStepL = 1
    'De baixo para cima
    Else
       nLimMinL = gridPoltronas.MaxRows
       nLimMaxL = 1
       nStepL = -1
    End If

    'Da esquerda para direita
    If optOrientacaoHoriz(0).Value Then
       nLimMinC = 1
       nLimMaxC = gridPoltronas.MaxCols
       nStepC = 1
    'Da direita para esquerda
    Else
       nLimMinC = gridPoltronas.MaxCols
       nLimMaxC = 1
       nStepC = -1
    End If

    For nContL = nLimMinL To nLimMaxL Step nStepL
         For nContC = nLimMinC To nLimMaxC Step nStepC
            gridPoltronas.Row = nContL
            gridPoltronas.Col = nContC
            gridPoltronas.Row2 = nContL
            gridPoltronas.Col2 = nContC
            gridPoltronas.BlockMode = True
            
            'If fTela Then
            '   gridPoltronas.BackColor = GRAY
            '   gridPoltronas.ForeColor = BLACK
            '   gridPoltronas.Text = ""
            'Else
            If gridPoltronas.BackColor <> shp_F5.FillColor Then
               gridPoltronas.Text = Format(nIndN, "0000")
               nIndN = nIndN + 1
            Else
               gridPoltronas.Text = ""
            End If
            
            gridPoltronas.BlockMode = False
         Next
    Next

End Sub

Private Sub Gera_NumcadLN(nChar As Integer, nIndN As Integer)
    
    'nChar Controla a parte Letra das poltronas
    'nIndN Controla a parte número das poltronas

    Dim nContL   As Integer
    Dim nContC   As Integer
    Dim nIndC    As Integer
    Dim nLimMinC As Integer
    Dim nLimMaxC As Integer
    Dim nLimMinL As Integer
    Dim nLimMaxL As Integer
    Dim nStepC   As Integer
    Dim nStepL   As Integer

    nIndC = nIndN

    'Direção da numeração

    'De cima para baixo
    If optOrientacaoVert(0).Value Then
       nLimMinL = 1
       nLimMaxL = gridPoltronas.MaxRows
       nStepL = 1
    'De baixo para cima
    Else
       nLimMinL = gridPoltronas.MaxRows
       nLimMaxL = 1
       nStepL = -1
    End If
    
    'Da esquerda para direita
    If optOrientacaoHoriz(0).Value Then
       nLimMinC = 1
       nLimMaxC = gridPoltronas.MaxCols
       nStepC = 1
    'Da direita para esquerda
    Else
       nLimMinC = gridPoltronas.MaxCols
       nLimMaxC = 1
       nStepC = -1
    End If

    For nContL = nLimMinL To nLimMaxL Step nStepL
        For nContC = nLimMinC To nLimMaxC Step nStepC
            gridPoltronas.Row = nContL
            gridPoltronas.Col = nContC
            gridPoltronas.Row2 = nContL
            gridPoltronas.Col2 = nContC
            gridPoltronas.BlockMode = True

            'If fTela Then
            '   gridPoltronas.BackColor = GRAY
            '   gridPoltronas.ForeColor = BLACK
            '   gridPoltronas.Text = ""
            'Else
            If gridPoltronas.BackColor <> shp_F5.FillColor Then
               gridPoltronas.Text = Chr(nChar) & Format(nIndC, "00")
               nIndC = nIndC + 1
            Else
               gridPoltronas.Text = ""
            End If

            gridPoltronas.BlockMode = False
        Next

        If nIndC <> nIndN Then
           nIndC = 1
           nIndN = nIndC
           nChar = nChar + 1
           If nChar = 91 Then nChar = 97 'Continua nas letras minúsculas
        End If
    Next

End Sub

Private Sub Gera_NumCadNL(nIndL As Integer, nIndN As Integer)
 
    'nIndL Controla a parte número das poltronas
    'nIndN Controla a parte letra das poltronas
    
    Dim nContL   As Integer
    Dim nContC   As Integer
    Dim nChar    As Integer
    Dim nLimMinC As Integer
    Dim nLimMaxC As Integer
    Dim nStepC   As Integer
    Dim nLimMinL As Integer
    Dim nLimMaxL As Integer
    Dim nStepL   As Integer

    nChar = nIndN

    'Direção da numeração

    'De cima para baixo
    If optOrientacaoVert(0).Value Then
       nLimMinL = 1
       nLimMaxL = gridPoltronas.MaxRows
       nStepL = 1
    'De baixo para cima
    Else
       nLimMinL = gridPoltronas.MaxRows
       nLimMaxL = 1
       nStepL = -1
    End If
    
    'Da esquerda para direita
    If optOrientacaoHoriz(0).Value Then
       nLimMinC = 1
       nLimMaxC = gridPoltronas.MaxCols
       nStepC = 1
    'Da direita para esquerda
    Else
       nLimMinC = gridPoltronas.MaxCols
       nLimMaxC = 1
       nStepC = -1
    End If
        
    For nContL = nLimMinL To nLimMaxL Step nStepL
        For nContC = nLimMinC To nLimMaxC Step nStepC
            gridPoltronas.Row = nContL
            gridPoltronas.Col = nContC
            gridPoltronas.Row2 = nContL
            gridPoltronas.Col2 = nContC
            gridPoltronas.BlockMode = True

            'If fTela Then
            '   gridPoltronas.BackColor = GRAY
            '   gridPoltronas.ForeColor = BLACK
            '   gridPoltronas.Text = ""
            'Else
            If gridPoltronas.BackColor <> shp_F5.FillColor Then
               gridPoltronas.Text = Format(nIndL, "00") & Chr(nChar)
               nChar = nChar + 1
               If nChar > 90 Then nChar = nChar + 6
            Else
               gridPoltronas.Text = ""
            End If
        
            gridPoltronas.BlockMode = False
        Next

        If nIndN <> nChar Then
           nChar = 65
           nIndN = nChar
           nIndL = nIndL + 1
        End If
    Next

End Sub

Private Sub optOrientacaoHoriz_Click(Index As Integer)
    Gera_Numeracao
End Sub

Private Sub optOrientacaoVert_Click(Index As Integer)
    Gera_Numeracao
End Sub

Private Sub spnColunas_Change()
    Redim_Cols
End Sub

Private Sub spnColunas_LostFocus()
    Redim_Cols
End Sub

Private Sub Redim_Rows()

    Dim nRowsAtu, nRowsAnt, nScrollB, nContL, nLimMinL, nLimMaxL

    nRowsAtu = IIf(IsNumeric(spnFilas.Value), spnFilas.Value, 0)
    nRowsAnt = gridPoltronas.MaxRows

    If nRowsAtu = gridPoltronas.MaxRows Then
       Exit Sub
    End If

    gridPoltronas.MaxRows = nRowsAtu

    If nRowsAtu > nRowsAnt Then

       nLimMinL = nRowsAnt + 1
       nLimMaxL = nRowsAtu

       For nContL = nLimMinL To nLimMaxL

           gridPoltronas.RowHeight(nContL) = 210

           gridPoltronas.Row = nContL
           gridPoltronas.Col = 1
           gridPoltronas.Row2 = nContL
           gridPoltronas.Col2 = gridPoltronas.MaxCols
           gridPoltronas.BlockMode = True
           gridPoltronas.BackColor = shp_F5.FillColor
           gridPoltronas.ForeColor = BLACK
           gridPoltronas.Text = ""
           gridPoltronas.BlockMode = False
       Next

    End If

    nScrollB = 0
    If nRowsAtu > 20 Then
       nScrollB = nScrollB + 2
    End If

   If gridPoltronas.MaxCols > 20 Then
       nScrollB = nScrollB + 1
    End If
    gridPoltronas.ScrollBars = nScrollB

    Gera_Numeracao
       
End Sub

Private Sub Redim_Cols()
    
    Dim nColsAtu, nColsAnt, nScrollB, nContC, nLimMinC, nLimMaxC

    nColsAtu = IIf(IsNumeric(spnColunas.Value), spnColunas.Value, 0)
    nColsAnt = gridPoltronas.MaxCols

    If nColsAtu = gridPoltronas.MaxCols Then
       Exit Sub
    End If

    gridPoltronas.MaxCols = nColsAtu

    If nColsAtu > nColsAnt Then

       nLimMinC = nColsAnt + 1
       nLimMaxC = nColsAtu

       For nContC = nLimMinC To nLimMaxC

           gridPoltronas.ColWidth(nContC) = 465

           gridPoltronas.Row = 1
           gridPoltronas.Col = nContC
           gridPoltronas.Row2 = gridPoltronas.MaxRows
           gridPoltronas.Col2 = nContC
           gridPoltronas.BlockMode = True

           gridPoltronas.BackColor = shp_F5.FillColor
           gridPoltronas.ForeColor = BLACK
           gridPoltronas.Text = ""
           gridPoltronas.BlockMode = False
       Next

    End If

    nScrollB = 0
    If nColsAtu > 20 Then
       nScrollB = nScrollB + 1
    End If

    If gridPoltronas.MaxRows > 20 Then
       nScrollB = nScrollB + 2
    End If
    gridPoltronas.ScrollBars = nScrollB
    
    Gera_Numeracao

End Sub

Private Sub CarregaDados()
    Dim clsTB_SALA      As New Cine2005.clsTB_SALA
    Dim clsTB_POLTRONAS As New Cine2005.clsTB_POLTRONAS
    Dim oRs             As ADODB.Recordset
    
    On Error GoTo CarregaDados_Erro
    
    Set clsTB_SALA.ConexaoADO = dbConnect
    clsTB_SALA.sal_cd = ccd_sal_cd.codigo
    
    Call clsTB_SALA.Selecionar(oRs)
    
    txtLotacao.Text = oRs.Fields.Item("sal_lugares").Value
    
    Set clsTB_POLTRONAS.ConexaoADO = dbConnect
    clsTB_POLTRONAS.sal_cd = ccd_sal_cd.codigo
    
    If (clsTB_POLTRONAS.Selecionar(oRs)) Then
        If Not (oRs.EOF And oRs.BOF) Then
            spnFilas.Value = oRs.Fields.Item("pol_num_filas").Value
            spnColunas.Value = oRs.Fields.Item("pol_num_colunas").Value
            cboNumeracao.ListIndex = oRs.Fields.Item("pol_tp_numeracao").Value
            txtNumPoltronas.Text = oRs.Fields.Item("pol_num_pri_col").Value
            optOrientacaoHoriz(0).Value = oRs.Fields.Item("pol_num_horiz").Value
            optOrientacaoHoriz(1).Value = Not oRs.Fields.Item("pol_num_horiz").Value
            optOrientacaoVert(0).Value = oRs.Fields.Item("pol_num_vert").Value
            optOrientacaoVert(1).Value = Not oRs.Fields.Item("pol_num_vert").Value
            txtTotPoltronas.Text = oRs.Fields.Item("pol_poltronas").Value
            
            gridPoltronas.MaxCols = 0
            gridPoltronas.MaxRows = 0
            Redim_Cols
            Redim_Rows
            
            Call Monta_Poltronas(oRs.Fields.Item("pol_mat_poltr").Value)
            Gera_Numeracao
        Else
            spnFilas.Value = 0
            spnColunas.Value = 0
            cboNumeracao.ListIndex = 0
            txtNumPoltronas.Text = "0001"
            optOrientacaoHoriz(0).Value = True
            optOrientacaoVert(0).Value = True
            txtTotPoltronas.Text = 0
            
            gridPoltronas.MaxCols = 0
            gridPoltronas.MaxRows = 0
            Redim_Cols
            Redim_Rows
            
            Call Monta_Poltronas(oRs.Fields.Item("").Value)
            Gera_Numeracao
        End If
    Else
        MsgBox "Não foi possível carregar dados!" & vbCrLf & clsTB_POLTRONAS.MensagemErro, vbInformation, App.ProductName
        GoTo CarregaDados_Erro
    End If
    
CarregaDados_Erro:
    If oRs.State = 1 Then oRs.Close
    
    Set oRs = Nothing
    Set clsTB_POLTRONAS = Nothing
    Set clsTB_SALA = Nothing
End Sub

Private Function Conta_Poltr()
    Dim nContL   As Integer
    Dim nContC   As Integer
    Dim nRowOld  As Integer
    Dim nColOld  As Integer
    Dim nTotPolt As Integer

    'Acumula o número de poltronas ativas/inativas no Grid
    nRowOld = gridPoltronas.Row
    nColOld = gridPoltronas.Col
    For nContL = 1 To gridPoltronas.MaxRows     'Loop das filas
        For nContC = 1 To gridPoltronas.MaxCols 'Loop das colunas
            gridPoltronas.Row = nContL
            gridPoltronas.Col = nContC
            
            'Acumula 1 em nTotPolt para poltronas Ativas/Inativas
            '                                     Azul-Claro e Amarelo
            If gridPoltronas.BackColor = shp_F2.FillColor Or gridPoltronas.BackColor = shp_F3.FillColor Or gridPoltronas.BackColor = shp_F4.FillColor Then
               nTotPolt = nTotPolt + 1
            End If
        Next
    Next
    gridPoltronas.Row = nRowOld
    gridPoltronas.Col = nColOld
    
    'Atualiza Caixa "Poltronas restantes"
    txtTotPoltronas.Text = Format(nTotPolt, "#,##0")
    
    'Retorna a somatória de poltronas ativas/inativas
    Conta_Poltr = nTotPolt
    
End Function

Private Sub spnFilas_Change()
    Redim_Rows
End Sub

Private Sub spnFilas_LostFocus()
    Redim_Rows
End Sub

Private Sub txtNumPoltronas_LostFocus()
    If cboNumeracao.ListIndex = 0 Then
       If Not IsNumeric(Trim$(txtNumPoltronas.Text)) Then
          MsgBox "Número incompatível com tipo de numeração!" & Chr$(10) & "(O formato correto é 9999)"
          txtNumPoltronas.SelStart = 0
          txtNumPoltronas.SelLength = Len(Trim(txtNumPoltronas.Text))
          txtNumPoltronas.SetFocus
          Exit Sub
       End If
    ElseIf cboNumeracao.ListIndex = 1 Then
       If (Asc(Mid$(txtNumPoltronas.Text, 1, 1)) < 65 Or Asc(Mid$(txtNumPoltronas.Text, 1, 1)) > 122) Or Not IsNumeric(Trim$(Mid$(txtNumPoltronas.Text, 2, 3))) Then
          MsgBox "Número incompatível com tipo de numeração!" & Chr$(10) & "(O formato correto é X99)"
          txtNumPoltronas.SelStart = 0
          txtNumPoltronas.SelLength = Len(Trim(txtNumPoltronas.Text))
          txtNumPoltronas.SetFocus
          Exit Sub
       End If
    ElseIf cboNumeracao.ListIndex = 2 Then
       If Not IsNumeric(Trim$(Mid$(txtNumPoltronas.Text, 1, 2))) Or (Asc(Mid$(txtNumPoltronas.Text, 3, 1)) < 65 Or Asc(Mid$(txtNumPoltronas.Text, 3, 1)) > 122) Then
          MsgBox "Número incompatível com tipo de numeração!" & Chr$(10) & "(O formato correto é 99X)"
          txtNumPoltronas.SelStart = 0
          txtNumPoltronas.SelLength = Len(Trim(txtNumPoltronas.Text))
          txtNumPoltronas.SetFocus
          Exit Sub
       End If
    End If

    Gera_Numeracao
End Sub

Private Function Monta_MatrPoltr()
    Dim nContL     As Integer
    Dim nContC     As Integer
    Dim nRowOld    As Integer
    Dim nColOld    As Integer
    Dim cMatrPoltr As String

    nRowOld = gridPoltronas.Row
    nColOld = gridPoltronas.Col

    cMatrPoltr = ""

    For nContL = 1 To gridPoltronas.MaxRows     'Loop das filas
        For nContC = 1 To gridPoltronas.MaxCols 'Loop das colunas
            gridPoltronas.Row = nContL
            gridPoltronas.Col = nContC
            
            'Monta string cMatrPoltr
            If gridPoltronas.BackColor = shp_F2.FillColor Then
               cMatrPoltr = cMatrPoltr + "1"
            ElseIf gridPoltronas.BackColor = shp_F3.FillColor Then
               cMatrPoltr = cMatrPoltr + "2"
            ElseIf gridPoltronas.BackColor = shp_F4.FillColor Then
               cMatrPoltr = cMatrPoltr + "3"
            ElseIf gridPoltronas.BackColor = shp_F5.FillColor Then
               cMatrPoltr = cMatrPoltr + "4"
            End If
        
        Next

    Next

    gridPoltronas.Row = nRowOld
    gridPoltronas.Col = nColOld
    
    Monta_MatrPoltr = cMatrPoltr
End Function

Private Sub Monta_Poltronas(matrPoltr As String)
    Dim nContL    As Integer
    Dim nContC    As Integer
    Dim nPoltrona As Integer
    Dim cPoltrona As String

    nPoltrona = 0  'Numerador das Poltronas dentro da matriz MatrPoltr
                   'Cada caracter representa uma cor como:
                   '1= BLUE LIGHT Poltrona Ativa
                   '2= YELLOW     Poltrona Inativa
                   '3= GRAY       Sem Poltrona
                   'A ordem das poltronas no campo memorando MatrPoltr
                   'será sempre Cima/Baixo, Esquerda/Direita

    For nContL = 1 To gridPoltronas.MaxRows     'Loop das filas
        For nContC = 1 To gridPoltronas.MaxCols 'Loop das colunas
            gridPoltronas.Row = nContL
            gridPoltronas.Col = nContC
            gridPoltronas.Row2 = nContL
            gridPoltronas.Col2 = nContC
            gridPoltronas.BlockMode = True

            'Pega poltrona no string MatrPoltr
            nPoltrona = nPoltrona + 1
            cPoltrona = Mid$(matrPoltr, nPoltrona, 1)
            If cPoltrona = "1" Then
               gridPoltronas.BackColor = shp_F2.FillColor
            ElseIf cPoltrona = "2" Then
               gridPoltronas.BackColor = shp_F3.FillColor
            ElseIf cPoltrona = "3" Then
               gridPoltronas.BackColor = shp_F4.FillColor
            ElseIf cPoltrona = "4" Then
               gridPoltronas.BackColor = shp_F5.FillColor
            End If

            gridPoltronas.Text = ""

            gridPoltronas.BlockMode = False
        Next
    Next
End Sub

