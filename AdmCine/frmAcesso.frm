VERSION 5.00
Object = "{234FCAFF-8D53-4DC2-9CD1-F90F4F6CB524}#32.0#0"; "Comandos.ocx"
Begin VB.Form frmAcesso 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Cadastro - Acesso"
   ClientHeight    =   6705
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6555
   Icon            =   "frmAcesso.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6705
   ScaleWidth      =   6555
   Begin VB.Frame fraGrid 
      Caption         =   "Perfis de Acesso"
      Height          =   5805
      Left            =   60
      TabIndex        =   1
      Top             =   60
      Width           =   6435
      Begin VB.Frame fraCaixa 
         Caption         =   "Caixa"
         Height          =   3525
         Left            =   150
         TabIndex        =   3
         Top             =   2025
         Width           =   6135
         Begin VB.CheckBox chkAdm 
            Caption         =   "Check1"
            Height          =   195
            Index           =   0
            Left            =   2760
            TabIndex        =   46
            Top             =   600
            Width           =   225
         End
         Begin VB.CheckBox chkCax 
            Caption         =   "Check1"
            Height          =   195
            Index           =   9
            Left            =   5295
            TabIndex        =   44
            Top             =   3090
            Width           =   225
         End
         Begin VB.CheckBox ckkGer 
            Caption         =   "Check1"
            Height          =   195
            Index           =   9
            Left            =   4027
            TabIndex        =   43
            Top             =   3090
            Width           =   225
         End
         Begin VB.CheckBox chkAdm 
            Caption         =   "Check1"
            Height          =   195
            Index           =   9
            Left            =   2760
            TabIndex        =   42
            Top             =   3090
            Width           =   225
         End
         Begin VB.CheckBox chkCax 
            Caption         =   "Check1"
            Height          =   195
            Index           =   8
            Left            =   5295
            TabIndex        =   41
            Top             =   2820
            Width           =   225
         End
         Begin VB.CheckBox ckkGer 
            Caption         =   "Check1"
            Height          =   195
            Index           =   8
            Left            =   4027
            TabIndex        =   40
            Top             =   2820
            Width           =   225
         End
         Begin VB.CheckBox chkAdm 
            Caption         =   "Check1"
            Height          =   195
            Index           =   8
            Left            =   2760
            TabIndex        =   39
            Top             =   2820
            Width           =   225
         End
         Begin VB.CheckBox chkCax 
            Caption         =   "Check1"
            Height          =   195
            Index           =   7
            Left            =   5295
            TabIndex        =   38
            Top             =   2550
            Width           =   225
         End
         Begin VB.CheckBox ckkGer 
            Caption         =   "Check1"
            Height          =   195
            Index           =   7
            Left            =   4027
            TabIndex        =   37
            Top             =   2550
            Width           =   225
         End
         Begin VB.CheckBox chkAdm 
            Caption         =   "Check1"
            Height          =   195
            Index           =   7
            Left            =   2760
            TabIndex        =   36
            Top             =   2550
            Width           =   225
         End
         Begin VB.CheckBox chkCax 
            Caption         =   "Check1"
            Height          =   195
            Index           =   6
            Left            =   5295
            TabIndex        =   35
            Top             =   2250
            Width           =   225
         End
         Begin VB.CheckBox ckkGer 
            Caption         =   "Check1"
            Height          =   195
            Index           =   6
            Left            =   4027
            TabIndex        =   34
            Top             =   2250
            Width           =   225
         End
         Begin VB.CheckBox chkAdm 
            Caption         =   "Check1"
            Height          =   195
            Index           =   6
            Left            =   2760
            TabIndex        =   33
            Top             =   2250
            Width           =   225
         End
         Begin VB.CheckBox chkCax 
            Caption         =   "Check1"
            Height          =   195
            Index           =   5
            Left            =   5295
            TabIndex        =   32
            Top             =   1980
            Width           =   225
         End
         Begin VB.CheckBox ckkGer 
            Caption         =   "Check1"
            Height          =   195
            Index           =   5
            Left            =   4027
            TabIndex        =   31
            Top             =   1980
            Width           =   225
         End
         Begin VB.CheckBox chkAdm 
            Caption         =   "Check1"
            Height          =   195
            Index           =   5
            Left            =   2760
            TabIndex        =   30
            Top             =   1980
            Width           =   225
         End
         Begin VB.CheckBox chkCax 
            Caption         =   "Check1"
            Height          =   195
            Index           =   4
            Left            =   5295
            TabIndex        =   29
            Top             =   1710
            Width           =   225
         End
         Begin VB.CheckBox ckkGer 
            Caption         =   "Check1"
            Height          =   195
            Index           =   4
            Left            =   4027
            TabIndex        =   28
            Top             =   1710
            Width           =   225
         End
         Begin VB.CheckBox chkAdm 
            Caption         =   "Check1"
            Height          =   195
            Index           =   4
            Left            =   2760
            TabIndex        =   27
            Top             =   1710
            Width           =   225
         End
         Begin VB.CheckBox chkCax 
            Caption         =   "Check1"
            Height          =   195
            Index           =   3
            Left            =   5295
            TabIndex        =   26
            Top             =   1440
            Width           =   225
         End
         Begin VB.CheckBox ckkGer 
            Caption         =   "Check1"
            Height          =   195
            Index           =   3
            Left            =   4027
            TabIndex        =   25
            Top             =   1440
            Width           =   225
         End
         Begin VB.CheckBox chkAdm 
            Caption         =   "Check1"
            Height          =   195
            Index           =   3
            Left            =   2760
            TabIndex        =   24
            Top             =   1440
            Width           =   225
         End
         Begin VB.CheckBox chkCax 
            Caption         =   "Check1"
            Height          =   195
            Index           =   2
            Left            =   5295
            TabIndex        =   23
            Top             =   1170
            Width           =   225
         End
         Begin VB.CheckBox ckkGer 
            Caption         =   "Check1"
            Height          =   195
            Index           =   2
            Left            =   4027
            TabIndex        =   22
            Top             =   1170
            Width           =   225
         End
         Begin VB.CheckBox chkAdm 
            Caption         =   "Check1"
            Height          =   195
            Index           =   2
            Left            =   2760
            TabIndex        =   21
            Top             =   1170
            Width           =   225
         End
         Begin VB.CheckBox chkCax 
            Caption         =   "Check1"
            Height          =   195
            Index           =   1
            Left            =   5295
            TabIndex        =   20
            Top             =   870
            Width           =   225
         End
         Begin VB.CheckBox ckkGer 
            Caption         =   "Check1"
            Height          =   195
            Index           =   1
            Left            =   4027
            TabIndex        =   19
            Top             =   870
            Width           =   225
         End
         Begin VB.CheckBox chkAdm 
            Caption         =   "Check1"
            Height          =   195
            Index           =   1
            Left            =   2760
            TabIndex        =   18
            Top             =   870
            Width           =   225
         End
         Begin VB.CheckBox chkCax 
            Caption         =   "Check1"
            Height          =   195
            Index           =   0
            Left            =   5295
            TabIndex        =   17
            Top             =   600
            Width           =   225
         End
         Begin VB.CheckBox ckkGer 
            Caption         =   "Check1"
            Height          =   195
            Index           =   0
            Left            =   4027
            TabIndex        =   16
            Top             =   600
            Width           =   225
         End
         Begin VB.Label lblCaxAdm 
            AutoSize        =   -1  'True
            Caption         =   "Administrador"
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
            Left            =   2340
            TabIndex        =   45
            Top             =   255
            Width           =   1155
         End
         Begin VB.Label labCaxCax 
            AutoSize        =   -1  'True
            Caption         =   "Caixa"
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
            Left            =   5167
            TabIndex        =   15
            Top             =   255
            Width           =   480
         End
         Begin VB.Label lblCaxGer 
            AutoSize        =   -1  'True
            Caption         =   "Gerente"
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
            Left            =   3794
            TabIndex        =   14
            Top             =   255
            Width           =   690
         End
         Begin VB.Label lblCaixa 
            AutoSize        =   -1  'True
            Caption         =   "Reimpressão"
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
            Index           =   9
            Left            =   480
            TabIndex        =   13
            Top             =   3090
            Width           =   1095
         End
         Begin VB.Label lblCaixa 
            AutoSize        =   -1  'True
            Caption         =   "Posição Caixa"
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
            Index           =   8
            Left            =   480
            TabIndex        =   12
            Top             =   2820
            Width           =   1215
         End
         Begin VB.Label lblCaixa 
            AutoSize        =   -1  'True
            Caption         =   "Cacelamento Operação"
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
            Index           =   7
            Left            =   480
            TabIndex        =   11
            Top             =   2550
            Width           =   1995
         End
         Begin VB.Label lblCaixa 
            AutoSize        =   -1  'True
            Caption         =   "Cancelamento Ingresso"
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
            Index           =   6
            Left            =   480
            TabIndex        =   10
            Top             =   2250
            Width           =   1995
         End
         Begin VB.Label lblCaixa 
            AutoSize        =   -1  'True
            Caption         =   "Cancelamento Combo"
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
            Index           =   5
            Left            =   480
            TabIndex        =   9
            Top             =   1980
            Width           =   1845
         End
         Begin VB.Label lblCaixa 
            AutoSize        =   -1  'True
            Caption         =   "Modo Talão"
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
            Index           =   4
            Left            =   480
            TabIndex        =   8
            Top             =   1710
            Width           =   1020
         End
         Begin VB.Label lblCaixa 
            AutoSize        =   -1  'True
            Caption         =   "Libera Caixa"
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
            Index           =   3
            Left            =   480
            TabIndex        =   7
            Top             =   1440
            Width           =   1065
         End
         Begin VB.Label lblCaixa 
            AutoSize        =   -1  'True
            Caption         =   "Fechamento Caixa"
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
            Index           =   2
            Left            =   480
            TabIndex        =   6
            Top             =   1170
            Width           =   1575
         End
         Begin VB.Label lblCaixa 
            AutoSize        =   -1  'True
            Caption         =   "Sangria"
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
            Index           =   1
            Left            =   480
            TabIndex        =   5
            Top             =   870
            Width           =   660
         End
         Begin VB.Label lblCaixa 
            AutoSize        =   -1  'True
            Caption         =   "Abre Caixa"
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
            Index           =   0
            Left            =   480
            TabIndex        =   4
            Top             =   600
            Width           =   930
         End
      End
      Begin VB.Frame fraAdmi 
         Caption         =   "Administração"
         Height          =   1485
         Left            =   150
         TabIndex        =   2
         Top             =   315
         Width           =   6135
         Begin VB.CheckBox chkAdm 
            Caption         =   "Check1"
            Enabled         =   0   'False
            Height          =   240
            Index           =   10
            Left            =   2775
            TabIndex        =   56
            Top             =   975
            Value           =   1  'Checked
            Width           =   240
         End
         Begin VB.CheckBox ckkGer 
            Caption         =   "Check1"
            Height          =   195
            Index           =   10
            Left            =   4050
            TabIndex        =   55
            Top             =   975
            Width           =   225
         End
         Begin VB.CheckBox chkCax 
            Caption         =   "Check1"
            Height          =   195
            Index           =   10
            Left            =   5325
            TabIndex        =   54
            Top             =   975
            Width           =   225
         End
         Begin VB.CheckBox chkAdmAdm 
            Caption         =   "Check1"
            Enabled         =   0   'False
            Height          =   240
            Left            =   2775
            TabIndex        =   53
            Top             =   525
            Value           =   1  'Checked
            Width           =   240
         End
         Begin VB.CheckBox chkAdmCax 
            Caption         =   "Check1"
            Height          =   195
            Left            =   5325
            TabIndex        =   51
            Top             =   555
            Width           =   225
         End
         Begin VB.CheckBox ckdAdmGer 
            Caption         =   "Check1"
            Height          =   195
            Left            =   4050
            TabIndex        =   50
            Top             =   555
            Width           =   225
         End
         Begin VB.Label lblCaixa 
            AutoSize        =   -1  'True
            Caption         =   "Perfis de Acesso"
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
            Index           =   10
            Left            =   525
            TabIndex        =   57
            Top             =   1050
            Width           =   1440
         End
         Begin VB.Label lblAdmAdm 
            AutoSize        =   -1  'True
            Caption         =   "Administrador"
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
            Left            =   2385
            TabIndex        =   52
            Top             =   210
            Width           =   1155
         End
         Begin VB.Label labAdmCax 
            AutoSize        =   -1  'True
            Caption         =   "Caixa"
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
            Left            =   5205
            TabIndex        =   49
            Top             =   210
            Width           =   480
         End
         Begin VB.Label lblAdmGer 
            AutoSize        =   -1  'True
            Caption         =   "Gerente"
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
            Left            =   3840
            TabIndex        =   48
            Top             =   210
            Width           =   690
         End
         Begin VB.Label lbAdm 
            AutoSize        =   -1  'True
            Caption         =   "Abre Administração"
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
            Left            =   525
            TabIndex        =   47
            Top             =   555
            Width           =   1650
         End
      End
   End
   Begin Comandos.cmdComandos cmdComandos 
      Align           =   2  'Align Bottom
      Height          =   765
      Left            =   0
      TabIndex        =   0
      Top             =   5940
      Width           =   6555
      _ExtentX        =   11562
      _ExtentY        =   1349
      EnabledNovo     =   0   'False
      EnabledAltera   =   -1  'True
   End
End
Attribute VB_Name = "frmAcesso"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private sOperacao As String

Private Sub Form_Load()
    
    
    Call HabilitaManut(False)
    Call CarregaControles
    
    Dim CineAx As New CineAux.CineAx
    Call CineAx.CentralizaTela(MDIFrmAdmCine, Me)
    
    Set CineAx = Nothing
    
End Sub

Private Sub cmdComandos_Click(ByVal iButtonClicked As Comandos.eButtonClicked, Cancel As Boolean)
    
    Select Case iButtonClicked
    
        Case ButtonAltera
            sOperacao = "A"
        
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
    End Select
End Sub

Private Sub HabilitaManut(ByVal bHabilita As Boolean)
    fraGrid.Enabled = bHabilita
End Sub

Private Sub CarregaControles()
    Dim acesso As New clsAcesso
    Dim oRs    As ADODB.Recordset
    
    Call LimpaControles
    
    Set acesso.ConexaoADO = dbConnect
    Call acesso.CarregaAcessos(oRs)
    
    Do While Not oRs.EOF
        'If oRs.Fields("mod_cd") = 1 Then
            Select Case oRs.Fields("per_cd")
                Case 9
                    chkAdmAdm.Value = vbChecked
                Case 8
                    ckdAdmGer.Value = vbChecked
                Case 1
                    chkAdmCax.Value = vbChecked
            End Select
        'Else
            Select Case oRs.Fields("per_cd")
                Case 9
                    chkAdm(oRs.Fields("fun_cd").Value - 1).Value = vbChecked
                Case 8
                    ckkGer(oRs.Fields("fun_cd").Value - 1).Value = vbChecked
                Case 1
                    chkCax(oRs.Fields("fun_cd").Value - 1).Value = vbChecked
            End Select
        'End If
        oRs.MoveNext
    Loop
    chkAdmAdm.Value = 1
    chkAdmAdm.Enabled = False
    chkAdm(10).Value = 1
    chkAdm(10).Enabled = False
    Set acesso = Nothing
End Sub

Private Sub LimpaControles()
    Dim i As Integer
    
    chkAdmAdm.Value = vbUnchecked
    ckdAdmGer.Value = vbUnchecked
    chkAdmCax.Value = vbUnchecked
    
    For i = 0 To 10
        chkAdm(i).Value = vbUnchecked
        ckkGer(i).Value = vbUnchecked
        chkCax(i).Value = vbUnchecked
    Next i
End Sub

Private Function Grava() As Boolean

    On Error GoTo Grava_Erro
    
    If Not Consiste() Then
        Exit Function
    End If
    
    Dim acesso  As New clsAcesso
    Dim acessos As New clcPerfisFuncao
    Dim i       As Integer
    
    Set acesso.ConexaoADO = dbConnect
    
    If chkAdmAdm.Value = vbChecked Then
        Call acessos.Add(1, 1, 9)
    End If
    If ckdAdmGer.Value = vbChecked Then
        Call acessos.Add(1, 1, 8)
    End If
    If chkAdmCax.Value = vbChecked Then
        Call acessos.Add(1, 1, 1)
    End If
    
    For i = 0 To 10
        If chkAdm(i).Value = vbChecked Then
            If i = 10 Then
                Call acessos.Add(i + 1, 1, 9)
            Else
                Call acessos.Add(i + 1, 2, 9)
            End If
        End If
        If ckkGer(i).Value = vbChecked Then
            If i = 10 Then
                Call acessos.Add(i + 1, 1, 8)
            Else
                Call acessos.Add(i + 1, 2, 8)
            End If
        End If
        If chkCax(i).Value = vbChecked Then
            If i = 10 Then
                Call acessos.Add(i + 1, 1, 1)
            Else
                Call acessos.Add(i + 1, 2, 1)
            End If
        End If
    Next i

    Call acesso.AtualizaAcessos(acessos)

    Grava = True
    GoTo Grava_Fim

Grava_Erro:
    MsgBox "Erro de execução! 'Grava/frmAcesso'", vbCritical, App.ProductName
    
Grava_Fim:
    Set acesso = Nothing
    
End Function

Private Function Consiste() As Boolean
    
    Dim sMens As String
    
    If chkAdmAdm.Value = vbUnchecked And ckdAdmGer.Value = vbUnchecked And chkAdmCax.Value = vbUnchecked Then
        sMens = "Não é possivel excluir todos os acessos da Administração"
    End If
    
    If sMens <> "" Then
        MsgBox sMens, vbInformation, App.ProductName
        Exit Function
    End If
    
    Consiste = True
    If ckkGer(10).Value = 1 Then ckdAdmGer.Value = 1
    If chkCax(10).Value = 1 Then chkAdmCax.Value = 1
    
End Function

