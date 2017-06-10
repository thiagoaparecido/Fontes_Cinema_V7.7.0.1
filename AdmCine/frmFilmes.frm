VERSION 5.00
Object = "{C0A63B80-4B21-11D3-BD95-D426EF2C7949}#1.0#0"; "Vsflex7L.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{234FCAFF-8D53-4DC2-9CD1-F90F4F6CB524}#30.0#0"; "Comandos.ocx"
Object = "{1753B334-0119-4B34-9134-D2B3CD181550}#34.0#0"; "Combo.ocx"
Object = "{91C13016-45DE-491B-BAC0-37F755626532}#29.0#0"; "Spin.ocx"
Begin VB.Form frmFilmes 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Cadastro - Filmes"
   ClientHeight    =   7065
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   11010
   Icon            =   "frmFilmes.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7065
   ScaleWidth      =   11010
   Begin Comandos.cmdComandos cmdComandos 
      Align           =   2  'Align Bottom
      Height          =   765
      Left            =   0
      TabIndex        =   14
      Top             =   6300
      Width           =   11010
      _ExtentX        =   19420
      _ExtentY        =   1349
      EnabledAltera   =   -1  'True
      EnabledExclui   =   -1  'True
   End
   Begin VB.Frame fraGrid 
      Caption         =   "Filmes Cadastrados"
      Height          =   3000
      Left            =   60
      TabIndex        =   18
      Top             =   60
      Width           =   10890
      Begin VSFlex7LCtl.VSFlexGrid VSFlexGrid 
         Height          =   2610
         Left            =   180
         TabIndex        =   13
         Top             =   240
         Width           =   10575
         _cx             =   18653
         _cy             =   4604
         _ConvInfo       =   1
         Appearance      =   1
         BorderStyle     =   1
         Enabled         =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MousePointer    =   0
         BackColor       =   -2147483643
         ForeColor       =   -2147483640
         BackColorFixed  =   -2147483633
         ForeColorFixed  =   -2147483630
         BackColorSel    =   -2147483635
         ForeColorSel    =   -2147483634
         BackColorBkg    =   -2147483636
         BackColorAlternate=   -2147483643
         GridColor       =   -2147483633
         GridColorFixed  =   -2147483632
         TreeColor       =   -2147483632
         FloodColor      =   192
         SheetBorder     =   -2147483642
         FocusRect       =   1
         HighLight       =   1
         AllowSelection  =   -1  'True
         AllowBigSelection=   -1  'True
         AllowUserResizing=   0
         SelectionMode   =   1
         GridLines       =   1
         GridLinesFixed  =   2
         GridLineWidth   =   1
         Rows            =   1
         Cols            =   12
         FixedRows       =   1
         FixedCols       =   0
         RowHeightMin    =   0
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   -1  'True
         FormatString    =   ""
         ScrollTrack     =   0   'False
         ScrollBars      =   3
         ScrollTips      =   0   'False
         MergeCells      =   0
         MergeCompare    =   0
         AutoResize      =   -1  'True
         AutoSizeMode    =   0
         AutoSearch      =   0
         AutoSearchDelay =   2
         MultiTotals     =   -1  'True
         SubtotalPosition=   1
         OutlineBar      =   0
         OutlineCol      =   0
         Ellipsis        =   0
         ExplorerBar     =   0
         PicturesOver    =   0   'False
         FillStyle       =   0
         RightToLeft     =   0   'False
         PictureType     =   0
         TabBehavior     =   0
         OwnerDraw       =   0
         Editable        =   0
         ShowComboButton =   -1  'True
         WordWrap        =   0   'False
         TextStyle       =   0
         TextStyleFixed  =   0
         OleDragMode     =   0
         OleDropMode     =   0
         ComboSearch     =   3
         AutoSizeMouse   =   -1  'True
         FrozenRows      =   0
         FrozenCols      =   0
         AllowUserFreezing=   0
         BackColorFrozen =   0
         ForeColorFrozen =   0
         WallPaperAlignment=   9
      End
   End
   Begin VB.Frame fraManut 
      Caption         =   "Manutenção"
      Enabled         =   0   'False
      Height          =   3195
      Left            =   75
      TabIndex        =   15
      Top             =   3060
      Width           =   10860
      Begin VB.TextBox txt_ano_filme 
         Height          =   315
         Left            =   1635
         MaxLength       =   4
         TabIndex        =   0
         Top             =   225
         Width           =   780
      End
      Begin VB.TextBox txt_cd_filme 
         Height          =   315
         Left            =   3435
         MaxLength       =   4
         TabIndex        =   1
         Top             =   225
         Width           =   780
      End
      Begin VB.Frame fraCopias 
         Caption         =   "Copias"
         Height          =   1110
         Left            =   2220
         TabIndex        =   25
         Top             =   1965
         Width           =   6435
         Begin VB.CommandButton cmdExcNumCopia 
            Caption         =   "Exclui"
            Height          =   375
            Left            =   1455
            TabIndex        =   10
            Top             =   645
            Width           =   1080
         End
         Begin VB.CommandButton cmdInsNumCopia 
            Caption         =   "Insere"
            Height          =   375
            Left            =   255
            TabIndex        =   9
            Top             =   645
            Width           =   1080
         End
         Begin VB.TextBox txt_num_copia 
            Height          =   315
            Left            =   1725
            MaxLength       =   4
            TabIndex        =   8
            Top             =   270
            Width           =   720
         End
         Begin MSHierarchicalFlexGridLib.MSHFlexGrid hfgCopias 
            Height          =   780
            Left            =   3000
            TabIndex        =   29
            Top             =   210
            Width           =   3270
            _ExtentX        =   5768
            _ExtentY        =   1376
            _Version        =   393216
            Rows            =   7
            Cols            =   1
            FixedRows       =   0
            FixedCols       =   0
            ScrollBars      =   2
            _NumberOfBands  =   1
            _Band(0).Cols   =   1
         End
         Begin VB.Label lbl_num_copia 
            AutoSize        =   -1  'True
            Caption         =   "Número da Cópia:"
            Height          =   195
            Left            =   375
            TabIndex        =   26
            Top             =   270
            Width           =   1275
         End
      End
      Begin VB.Frame fraNacio 
         Caption         =   "Nacionalidade"
         Height          =   1110
         Left            =   8730
         TabIndex        =   24
         Top             =   495
         Width           =   1530
         Begin VB.OptionButton optNacional 
            Caption         =   "Estrangeiro"
            Height          =   225
            Index           =   1
            Left            =   165
            TabIndex        =   12
            Top             =   645
            Width           =   1230
         End
         Begin VB.OptionButton optNacional 
            Caption         =   "Nacional"
            Height          =   225
            Index           =   0
            Left            =   165
            TabIndex        =   11
            Top             =   315
            Width           =   1230
         End
      End
      Begin VB.TextBox txt_fil_nm 
         Height          =   315
         Left            =   1635
         MaxLength       =   50
         TabIndex        =   2
         Top             =   570
         Width           =   6780
      End
      Begin Spin.SpinNumber spn_fil_censura 
         Height          =   315
         Left            =   7365
         TabIndex        =   6
         Top             =   1275
         Width           =   1035
         _ExtentX        =   1826
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
         Max             =   9999
         Value           =   "0"
      End
      Begin Spin.SpinNumber spn_fil_duracao 
         Height          =   315
         Left            =   4905
         TabIndex        =   7
         Top             =   1275
         Width           =   1035
         _ExtentX        =   1826
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
         Max             =   9999
         Value           =   "0"
      End
      Begin Spin.SpinNumber spn_trai_durac 
         Height          =   315
         Left            =   4905
         TabIndex        =   5
         Top             =   1620
         Width           =   1035
         _ExtentX        =   1826
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
         Max             =   9999
         Value           =   "0"
      End
      Begin MSComCtl2.DTPicker dtp_fil_dt_ini 
         Height          =   315
         Left            =   1635
         TabIndex        =   3
         Top             =   1275
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   556
         _Version        =   393216
         Format          =   61538305
         CurrentDate     =   38483
      End
      Begin MSComCtl2.DTPicker dtp_fil_dt_fim 
         Height          =   315
         Left            =   1635
         TabIndex        =   4
         Top             =   1620
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   556
         _Version        =   393216
         Format          =   61538305
         CurrentDate     =   38483
      End
      Begin Combo.cboCodDesc dis_cd 
         Height          =   315
         Left            =   1635
         TabIndex        =   30
         Top             =   900
         Width           =   6750
         _ExtentX        =   11906
         _ExtentY        =   556
         NomeTabela      =   "tb_distribuidora"
         NomeCampoCodigo =   "dis_cd"
         NomeCampoDescricao=   "dis_nm"
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
         TamCampoCodigo  =   5
         MostraBotaoNovo =   0   'False
         CodigoVisible   =   0   'False
         MostraBotaoAtualiza=   0   'False
      End
      Begin VB.Label lbl_ano_filme 
         AutoSize        =   -1  'True
         Caption         =   "Ano:"
         Height          =   195
         Left            =   615
         TabIndex        =   28
         Top             =   225
         Width           =   330
      End
      Begin VB.Label lbl_cd_filme 
         AutoSize        =   -1  'True
         Caption         =   "Código:"
         Height          =   195
         Left            =   2700
         TabIndex        =   27
         Top             =   225
         Width           =   540
      End
      Begin VB.Label lbl_trai_durac 
         AutoSize        =   -1  'True
         Caption         =   "Trailer (em minutos):"
         Height          =   195
         Left            =   3225
         TabIndex        =   23
         Top             =   1620
         Width           =   1410
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "Distribuidora:"
         Height          =   195
         Left            =   615
         TabIndex        =   22
         Top             =   915
         Width           =   915
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Término:"
         Height          =   195
         Left            =   615
         TabIndex        =   21
         Top             =   1620
         Width           =   615
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Início:"
         Height          =   195
         Left            =   615
         TabIndex        =   20
         Top             =   1275
         Width           =   450
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Duração (em minutos):"
         Height          =   195
         Left            =   3225
         TabIndex        =   19
         Top             =   1275
         Width           =   1590
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Classificação:"
         Height          =   195
         Left            =   6255
         TabIndex        =   17
         Top             =   1275
         Width           =   975
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Nome:"
         Height          =   195
         Left            =   615
         TabIndex        =   16
         Top             =   570
         Width           =   465
      End
   End
End
Attribute VB_Name = "frmFilmes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private sOperacao As String
Private filCdGrv  As Long

Private Sub cmdComandos_Click(ByVal iButtonClicked As Comandos.eButtonClicked, Cancel As Boolean)
    
    Select Case iButtonClicked
    
        Case ButtonAltera
            sOperacao = "A"
        
            If VSFlexGrid.RowSel > 0 And VSFlexGrid.Rows > 1 Then
                Call CarregaControles
                Call HabilitaManut(True)
                
                'dis_cd.codigo = ""
                
            Else
                Cancel = True
            End If
            
        Case ButtonNovo
            sOperacao = "I"
            Call LimpaControles
            Call HabilitaManut(True)

        Case ButtonExclui
            If VSFlexGrid.RowSel > 0 And VSFlexGrid.Rows > 1 Then
                If MsgBox("Confirma exclusão do Filme selecionado?", vbYesNo + vbQuestion + vbDefaultButton2, App.ProductName) = vbYes Then
                    If Exclui() Then
                        Call VSFlexGrid.RemoveItem(VSFlexGrid.RowSel)
                        Call CarregaControles
                    End If
                End If
            End If
    
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
            Call CarregaControles
            
    End Select

End Sub

Private Sub cmdExcNumCopia_Click()
    If hfgCopias.Row >= 0 And hfgCopias.Rows > 0 Then
        If hfgCopias.Rows > 1 Then
            hfgCopias.RemoveItem hfgCopias.Row
        Else
'            hfgCopias.TextMatrix(hfgCopias.Row, 0) = ""
'            'hfgCopias.TextMatrix(hfgCopias.row, 1) = ""
             hfgCopias.Rows = 0
        End If
    End If
End Sub

Private Sub cmdInsNumCopia_Click()
    Dim i As Integer
    
    If txt_num_copia.Text <> "" And IsNumeric(txt_num_copia.Text) Then
        For i = 0 To hfgCopias.Rows - 1
            If hfgCopias.TextMatrix(i, 0) = CInt(txt_num_copia.Text) Then
                MsgBox "Número de copias já existe!", vbInformation, App.ProductName
                Exit Sub
            End If
        Next i
            
        hfgCopias.AddItem CInt(txt_num_copia.Text)
        
        txt_num_copia.Text = ""
    Else
        MsgBox "Número de cópia invalido!", vbCritical, App.ProductName
    End If

End Sub

Private Sub Form_Load()
    filCdGrv = 0
    
    Set dis_cd.ConexaoADO = dbConnect
    
    Call HabilitaManut(False)
    Call PreencheGrid
    
    Dim CineAx As New CineAux.CineAx
    Call CineAx.CentralizaTela(MDIFrmAdmCine, Me)
    
    Set CineAx = Nothing
    
    Dim log As New clsLog
    
    Set log.ConexaoADO = dbConnect
    log.usu_nm = strUsuario
    log.slg_descricao = "Entrou na tela de " & Me.Caption & " do módulo " & App.ProductName
    
    log.insereLog
    
End Sub

Private Sub HabilitaManut(ByVal bHabilita As Boolean)
    fraGrid.Enabled = Not bHabilita
    fraManut.Enabled = bHabilita
    
    If bHabilita Then
        If sOperacao = "A" Then
            txt_ano_filme.Enabled = False
            txt_cd_filme.Enabled = False
        End If
    Else
        txt_ano_filme.Enabled = True
        txt_cd_filme.Enabled = True
    End If
End Sub

Private Sub CarregaControles()
    Dim filme As New clsTB_Filme
    Dim i       As Integer
    
    Call LimpaControles
    If VSFlexGrid.RowSel > 0 And VSFlexGrid.Rows > 1 Then
        txt_fil_nm.Text = VSFlexGrid.TextMatrix(VSFlexGrid.RowSel, VSFlexGrid.ColIndex("Filme"))
        spn_fil_censura.Value = VSFlexGrid.TextMatrix(VSFlexGrid.RowSel, VSFlexGrid.ColIndex("Censura"))
        spn_fil_duracao.Value = VSFlexGrid.TextMatrix(VSFlexGrid.RowSel, VSFlexGrid.ColIndex("Duração"))
        dtp_fil_dt_ini.Value = VSFlexGrid.TextMatrix(VSFlexGrid.RowSel, VSFlexGrid.ColIndex("Início"))
        dtp_fil_dt_fim.Value = VSFlexGrid.TextMatrix(VSFlexGrid.RowSel, VSFlexGrid.ColIndex("Término"))
        txt_ano_filme.Text = VSFlexGrid.TextMatrix(VSFlexGrid.RowSel, VSFlexGrid.ColIndex("Ano"))
        txt_cd_filme.Text = VSFlexGrid.TextMatrix(VSFlexGrid.RowSel, VSFlexGrid.ColIndex("Código"))
        dis_cd.codigo = CInt(VSFlexGrid.TextMatrix(VSFlexGrid.RowSel, VSFlexGrid.ColIndex("dis_cd")))
        spn_trai_durac.Value = VSFlexGrid.TextMatrix(VSFlexGrid.RowSel, VSFlexGrid.ColIndex("fil_durac_trai"))
        
        If VSFlexGrid.TextMatrix(VSFlexGrid.RowSel, VSFlexGrid.ColIndex("fil_id_nacio")) = "N" Then
            optNacional(0).Value = True
        Else
            optNacional(1).Value = True
        End If
        
        filme.fil_cd = VSFlexGrid.TextMatrix(VSFlexGrid.RowSel, VSFlexGrid.ColIndex("fil_cd"))
        Set filme.ConexaoADO = dbConnect
        If (filme.CarregaCopias) Then
            For i = 1 To filme.copias.Count
                hfgCopias.AddItem filme.copias.Item(i).NumCopia
            Next i
        End If
    End If
End Sub

Private Sub LimpaControles()
    txt_fil_nm.Text = ""
    spn_fil_censura.Value = 0
    spn_fil_duracao.Value = 0
    dtp_fil_dt_ini.Value = Date
    dtp_fil_dt_fim.Value = Date
    txt_ano_filme.Text = ""
    txt_cd_filme.Text = ""
    dis_cd.codigo = ""
    optNacional(0).Value = True
    spn_trai_durac.Value = 0
    
    hfgCopias.Rows = 0
End Sub

Private Function Grava() As Boolean
    Dim i      As Integer
    Dim copias As New clsCopias

    On Error GoTo Grava_Erro
    
    If Not Consiste() Then
        Exit Function
    End If
    
    Dim clsTB_Filme As New Cine2005.clsTB_Filme
    
    Set clsTB_Filme.ConexaoADO = dbConnect
    
    clsTB_Filme.fil_cd = CLng(txt_ano_filme.Text & Format(CInt(txt_cd_filme.Text), "0000"))
    clsTB_Filme.fil_nm = eliminaCaracProblemas(txt_fil_nm.Text)
    clsTB_Filme.fil_censura = spn_fil_censura.Value
    clsTB_Filme.fil_duracao = spn_fil_duracao.Value
    clsTB_Filme.fil_dt_ini = dtp_fil_dt_ini.Value
    clsTB_Filme.fil_dt_fim = dtp_fil_dt_fim.Value
    clsTB_Filme.dis_cd = dis_cd.codigo
    clsTB_Filme.fil_durac_trai = spn_trai_durac.Value
    clsTB_Filme.fil_id_nacio = IIf(optNacional(0).Value, "N", "E")
    
    filCdGrv = clsTB_Filme.fil_cd
    
    For i = 0 To hfgCopias.Rows - 1
        If IsNumeric(hfgCopias.TextMatrix(i, 0)) Then
            clsTB_Filme.copias.Add hfgCopias.TextMatrix(i, 0)
        End If
    Next i
    
    If sOperacao = "I" Then
        If Not clsTB_Filme.Incluir() Then
            MsgBox "Não foi possível incluir o Filme!" & vbCrLf & clsTB_Filme.MensagemErro, vbInformation, App.ProductName
            GoTo Grava_Fim
        End If
    Else
        clsTB_Filme.fil_cd = VSFlexGrid.TextMatrix(VSFlexGrid.RowSel, 0)
        If Not clsTB_Filme.Alterar() Then
            MsgBox "Não foi possível alterar o Filme Selecionado!" & vbCrLf & clsTB_Filme.MensagemErro, vbInformation, App.ProductName
            GoTo Grava_Fim
        End If
    End If
    
    Call PreencheGrid

    Grava = True
    GoTo Grava_Fim

Grava_Erro:
    MsgBox "Erro de execução! 'Grava/frmFilmes'", vbCritical, App.ProductName
    
Grava_Fim:
    Set clsTB_Filme = Nothing
    
End Function

Private Function Exclui() As Boolean

    On Error GoTo Exclui_Erro
    
    Dim clsTB_Filme As New Cine2005.clsTB_Filme

    Set clsTB_Filme.ConexaoADO = dbConnect
    
    clsTB_Filme.fil_cd = VSFlexGrid.TextMatrix(VSFlexGrid.RowSel, 0)
    
    If Not clsTB_Filme.Excluir() Then
        MsgBox "Não foi possível excluir o Filme Selecionado!" & vbCrLf & clsTB_Filme.MensagemErro, vbInformation, App.ProductName
        GoTo Exclui_Fim
    End If
            
    Exclui = True
    GoTo Exclui_Fim
    
Exclui_Erro:
    MsgBox "Erro de execução! 'Exclui/frmFilmes'", vbCritical, App.ProductName
    
Exclui_Fim:
    Set clsTB_Filme = Nothing
    
End Function

Private Sub PreencheGrid()
    Dim oRs         As New ADODB.Recordset
    Dim clsTB_Filme As New Cine2005.clsTB_Filme
    Dim i           As Integer
    
    On Error GoTo PreencheGrid_Erro
    
    Set clsTB_Filme.ConexaoADO = dbConnect
    
    If Not clsTB_Filme.PreencheGrid(oRs) Then
        MsgBox clsTB_Filme.MensagemErro, vbCritical, App.ProductName
        GoTo PreencheGrid_Fim
    End If

    Call CarregaGridAutomatico(VSFlexGrid, oRs, VSFlexGrid.Row)
    VSFlexGrid.ColHidden(0) = True
    VSFlexGrid.ColHidden(1) = True
    VSFlexGrid.ColHidden(2) = True
    VSFlexGrid.ColHidden(3) = True
            
    VSFlexGrid.MergeCells = flexMergeRestrictColumns
    'VSFlexGrid.MergeCol(3) = True
    'VSFlexGrid.MergeCol(4) = True
    'VSFlexGrid.MergeCol(5) = True

    Call CarregaControles
    
    For i = 1 To VSFlexGrid.Rows - 1
        If IsNumeric(VSFlexGrid.TextMatrix(i, 0)) Then
            If filCdGrv = CLng(VSFlexGrid.TextMatrix(i, 0)) Then
                VSFlexGrid.Row = i
                
                If VSFlexGrid.Row < 10 Then
                    VSFlexGrid.TopRow = 1
                Else
                    VSFlexGrid.TopRow = i - 5
                End If
            End If
        End If
    Next i

    GoTo PreencheGrid_Fim
    
PreencheGrid_Erro:
    MsgBox "Erro de execução! 'PreencheGrid/frmFilmes'", vbCritical, App.ProductName
    
PreencheGrid_Fim:
    If oRs.State = 1 Then oRs.Close
    Set oRs = Nothing
    Set clsTB_Filme = Nothing
    
End Sub

Private Function Consiste() As Boolean
    
    Dim sMens As String
    
    If Trim(txt_ano_filme.Text) = "" Then
        sMens = sMens & "Ano do Filme deve ser informado!" & vbCrLf
    End If
    
    If Not IsNumeric(txt_ano_filme.Text) Then
        sMens = sMens & "Ano do Filme invalido!" & vbCrLf
    End If
    
    If Len(Trim(txt_ano_filme.Text)) <> 4 Then
        sMens = sMens & "Ano do Filme deve ter quatro digitos!" & vbCrLf
    End If
    
    
    If Trim(txt_cd_filme.Text) = "" Then
        sMens = sMens & "Código do Filme deve ser informado!" & vbCrLf
    End If
    
    If Not IsNumeric(txt_cd_filme.Text) Then
        sMens = sMens & "Código do Filme invalido!" & vbCrLf
    End If
    
    If Trim(txt_fil_nm.Text) = "" Then
        sMens = sMens & "Nome do Filme deve ser informado!" & vbCrLf
    End If
    
    If dis_cd.codigo = "" Then
        sMens = sMens & "Distribuidora do Filme deve ser informada!" & vbCrLf
    End If
    
    If spn_fil_duracao.Value = 0 Then
        sMens = sMens & "Duração do Filme ( em minutos ) deve ser informado!" & vbCrLf
    End If

    If dtp_fil_dt_ini.Value > dtp_fil_dt_fim.Value Then
        sMens = sMens & "Data Término deve ser superior a data de Início!" & vbCrLf
    End If
    
    If sMens <> "" Then
        MsgBox sMens, vbInformation, App.ProductName
        Exit Function
    End If
    
    Consiste = True
    
End Function

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    Dim log As New clsLog
    
    Set log.ConexaoADO = dbConnect
    log.usu_nm = strUsuario
    log.slg_descricao = "Saiu da tela de " & Me.Caption & " do módulo " & App.ProductName
    
    log.insereLog
End Sub

Private Sub VSFlexGrid_RowColChange()
    Call CarregaControles
End Sub

Private Function eliminaCaracProblemas(texto As String)
    eliminaCaracProblemas = Replace(texto, ",", "")
    eliminaCaracProblemas = Replace(eliminaCaracProblemas, "'", "")
    eliminaCaracProblemas = Replace(eliminaCaracProblemas, """", "")
    
End Function
