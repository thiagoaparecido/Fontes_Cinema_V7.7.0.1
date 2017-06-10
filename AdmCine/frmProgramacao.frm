VERSION 5.00
Object = "{C0A63B80-4B21-11D3-BD95-D426EF2C7949}#1.0#0"; "Vsflex7L.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{234FCAFF-8D53-4DC2-9CD1-F90F4F6CB524}#32.0#0"; "Comandos.ocx"
Begin VB.Form frmProgramacao 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Cadastro - Período Sessões"
   ClientHeight    =   4920
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4635
   Icon            =   "frmProgramacao.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4920
   ScaleWidth      =   4635
   Begin Comandos.cmdComandos cmdComandos 
      Align           =   2  'Align Bottom
      Height          =   765
      Left            =   0
      TabIndex        =   7
      Top             =   4155
      Width           =   4635
      _ExtentX        =   8176
      _ExtentY        =   1349
      EnabledAltera   =   -1  'True
      EnabledExclui   =   -1  'True
   End
   Begin VB.Frame fraManut 
      Caption         =   "Manutenção"
      Enabled         =   0   'False
      Height          =   720
      Left            =   60
      TabIndex        =   2
      Top             =   3420
      Width           =   4515
      Begin VB.CheckBox chkCopia 
         Caption         =   "Copia sessões do periodo Anterior"
         Height          =   285
         Left            =   135
         TabIndex        =   8
         Top             =   630
         Visible         =   0   'False
         Width           =   3645
      End
      Begin MSComCtl2.DTPicker dtp_prg_dt_ini 
         Height          =   315
         Left            =   660
         TabIndex        =   3
         Top             =   240
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   556
         _Version        =   393216
         Format          =   16777217
         CurrentDate     =   38483
      End
      Begin MSComCtl2.DTPicker dtp_prg_dt_fim 
         Height          =   315
         Left            =   2880
         TabIndex        =   4
         Top             =   240
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   556
         _Version        =   393216
         Format          =   16777217
         CurrentDate     =   38483
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Início:"
         Height          =   195
         Left            =   120
         TabIndex        =   6
         Top             =   300
         Width           =   450
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Término:"
         Height          =   195
         Left            =   2160
         TabIndex        =   5
         Top             =   300
         Width           =   615
      End
   End
   Begin VB.Frame fraGrid 
      Caption         =   "Programações"
      Height          =   3315
      Left            =   60
      TabIndex        =   0
      Top             =   60
      Width           =   4515
      Begin VSFlex7LCtl.VSFlexGrid VSFlexGrid 
         Height          =   2775
         Left            =   180
         TabIndex        =   1
         Top             =   360
         Width           =   4155
         _cx             =   7329
         _cy             =   4895
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
         Cols            =   2
         FixedRows       =   1
         FixedCols       =   0
         RowHeightMin    =   0
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   -1  'True
         FormatString    =   $"frmProgramacao.frx":000C
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
End
Attribute VB_Name = "frmProgramacao"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private sOperacao As String

Private Sub cmdComandos_Click(ByVal iButtonClicked As Comandos.eButtonClicked, Cancel As Boolean)
    
    Select Case iButtonClicked
    
        Case ButtonAltera
            If permiteAltExc() Then
                sOperacao = "A"
            
                If VSFlexGrid.RowSel > 0 And VSFlexGrid.Rows > 1 Then
                    Call CarregaControles
                    Call HabilitaManut(True)
                Else
                    Cancel = True
                End If
            Else
                Cancel = True
                MsgBox "Não é possível alterar programação. Período anterior a data atual", vbCritical, App.ProductName
            End If
            
        Case ButtonNovo
            sOperacao = "I"
            Call LimpaControles
            Call HabilitaManut(True)

        Case ButtonExclui
            If permiteAltExc() Then
                If VSFlexGrid.RowSel > 0 And VSFlexGrid.Rows > 1 Then
                    If MsgBox("Confirma exclusão da Programação selecionada?", vbYesNo + vbQuestion + vbDefaultButton2, App.ProductName) = vbYes Then
                        If Exclui() Then
                            Call VSFlexGrid.RemoveItem(VSFlexGrid.RowSel)
                            Call CarregaControles
                        End If
                    End If
                End If
            Else
                Cancel = True
                MsgBox "Não é possível excluir programação. Período anterior a data atual", vbCritical, App.ProductName
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

Private Sub Form_Activate()
    Call CarregaControles
End Sub

Private Sub Form_Load()

    Call HabilitaManut(False)
    Call PreencheGrid
    
    Dim CineAx As New CineAux.CineAx
    Call CineAx.CentralizaTela(MDIFrmAdmCine, Me)
    
    Set CineAx = Nothing
    Call LimpaControles
    
    Dim log As New clsLog
    
    Set log.ConexaoADO = dbConnect
    log.usu_nm = strUsuario
    log.slg_descricao = "Entrou na tela de " & Me.Caption & " do módulo " & App.ProductName
    
    log.insereLog
    
End Sub

Private Sub HabilitaManut(ByVal bHabilita As Boolean)
    fraGrid.Enabled = Not bHabilita
    fraManut.Enabled = bHabilita
    If sOperacao = "I" And bHabilita Then
        chkCopia.Visible = True
        fraManut.Height = 1035
        Me.Height = 5730
    Else
        chkCopia.Visible = False
        fraManut.Height = 720
        Me.Height = 5400
    End If
End Sub

Private Sub CarregaControles()
    Call LimpaControles
    If VSFlexGrid.RowSel > 0 And VSFlexGrid.Rows > 1 Then
        dtp_prg_dt_ini.Value = VSFlexGrid.TextMatrix(VSFlexGrid.RowSel, VSFlexGrid.ColIndex("Início"))
        dtp_prg_dt_fim.Value = VSFlexGrid.TextMatrix(VSFlexGrid.RowSel, VSFlexGrid.ColIndex("Término"))
    End If
End Sub

Private Sub LimpaControles()
    dtp_prg_dt_ini.Value = Date
    dtp_prg_dt_fim.Value = Date
    chkCopia.Value = vbUnchecked
End Sub

Private Function Grava() As Boolean
    On Error GoTo Grava_Erro
    
    Dim copia As Boolean
    Dim clsTB_PROGRAMACAO As New Cine2005.clsTB_PROGRAMACAO
    
    If Not Consiste() Then
        Exit Function
    End If
    
    Grava = False
    
    Set clsTB_PROGRAMACAO.ConexaoADO = dbConnect
    
    clsTB_PROGRAMACAO.prg_dt_ini = dtp_prg_dt_ini.Value
    clsTB_PROGRAMACAO.prg_dt_fim = dtp_prg_dt_fim.Value
    
    If sOperacao = "I" Then
        If chkCopia.Value = vbChecked Then
            copia = True
        Else
            copia = False
        End If
        
        If copia Then
            If Not clsTB_PROGRAMACAO.VerificaCopia() Then
                If MsgBox("Não existem filmes cadastrados para este período! Deseja continuar?", vbYesNo, App.ProductName) = vbNo Then
                    Exit Function
                End If
            End If
        End If
        
        If Not clsTB_PROGRAMACAO.Incluir(copia) Then
            MsgBox "Não foi possível incluir a Programação!" & vbCrLf & clsTB_PROGRAMACAO.MensagemErro, vbInformation, App.ProductName
            GoTo Grava_Fim
        End If
    Else
        clsTB_PROGRAMACAO.prg_cd = VSFlexGrid.TextMatrix(VSFlexGrid.RowSel, 0)
        If Not clsTB_PROGRAMACAO.Alterar() Then
            MsgBox "Não foi possível alterar a Programação Selecionada!" & vbCrLf & clsTB_PROGRAMACAO.MensagemErro, vbInformation, App.ProductName
            GoTo Grava_Fim
        End If
    End If
    
    Call PreencheGrid

    Grava = True
    GoTo Grava_Fim

Grava_Erro:
    MsgBox "Erro de execução! 'Grava/frmProgramacao'", vbCritical, App.ProductName
    
Grava_Fim:
    Set clsTB_PROGRAMACAO = Nothing
    
End Function

Private Function Exclui() As Boolean

    On Error GoTo Exclui_Erro
    
    Dim clsTB_PROGRAMACAO As New Cine2005.clsTB_PROGRAMACAO

    Set clsTB_PROGRAMACAO.ConexaoADO = dbConnect
    
    clsTB_PROGRAMACAO.prg_cd = VSFlexGrid.TextMatrix(VSFlexGrid.RowSel, 0)
    
    If Not clsTB_PROGRAMACAO.Excluir() Then
        MsgBox "Não foi possível excluir a Programação Selecionada!" & vbCrLf & "Possivelmente existem sessões associadas com este período!!", vbInformation, App.ProductName
        GoTo Exclui_Fim
    End If
            
    Exclui = True
    GoTo Exclui_Fim
    
Exclui_Erro:
    MsgBox "Erro de execução! 'Exclui/frmProgramacao'", vbCritical, App.ProductName
    
Exclui_Fim:
    Set clsTB_PROGRAMACAO = Nothing
    
End Function

Private Sub PreencheGrid()

    On Error GoTo PreencheGrid_Erro
    
    Dim oRs As New ADODB.Recordset
    Dim clsTB_PROGRAMACAO As New Cine2005.clsTB_PROGRAMACAO
    
    Set clsTB_PROGRAMACAO.ConexaoADO = dbConnect
    
    If Not clsTB_PROGRAMACAO.PreencheGrid(oRs) Then
        MsgBox clsTB_PROGRAMACAO.MensagemErro, vbCritical, App.ProductName
        GoTo PreencheGrid_Fim
    End If

    Call CarregaGridAutomatico(VSFlexGrid, oRs, VSFlexGrid.Row)
    VSFlexGrid.ColHidden(0) = True
            
    Call CarregaControles

    GoTo PreencheGrid_Fim
    
PreencheGrid_Erro:
    MsgBox "Erro de execução! 'PreencheGrid/frmProgramacao'", vbCritical, App.ProductName
    
PreencheGrid_Fim:
    If oRs.State = 1 Then oRs.Close
    Set oRs = Nothing
    Set clsTB_PROGRAMACAO = Nothing
    
End Sub

Private Function Consiste() As Boolean
    
    Dim sMens As String
    
    If dtp_prg_dt_ini.Value > dtp_prg_dt_fim.Value Then
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

Private Function permiteAltExc() As Boolean
    Dim dtIni  As Date
    Dim dtFim  As Date
    Dim dtAtu  As Date
    
    permiteAltExc = False
    
    dtIni = dtp_prg_dt_ini.Value
    
    dtFim = dtp_prg_dt_fim.Value
    
    dtAtu = CDate(Format(Date, "Short Date"))
    
    If verificaPeriodo(dtIni, dtFim, dtAtu) > 0 Then
        permiteAltExc = True
    End If
End Function

