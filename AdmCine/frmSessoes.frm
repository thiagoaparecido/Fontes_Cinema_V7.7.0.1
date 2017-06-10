VERSION 5.00
Object = "{C0A63B80-4B21-11D3-BD95-D426EF2C7949}#1.0#0"; "Vsflex7L.ocx"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{234FCAFF-8D53-4DC2-9CD1-F90F4F6CB524}#32.0#0"; "Comandos.ocx"
Object = "{1753B334-0119-4B34-9134-D2B3CD181550}#36.0#0"; "Combo.ocx"
Object = "{91C13016-45DE-491B-BAC0-37F755626532}#31.0#0"; "Spin.ocx"
Begin VB.Form frmSessoes 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Cadastro - Sessões"
   ClientHeight    =   6945
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   10785
   Icon            =   "frmSessoes.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6945
   ScaleWidth      =   10785
   Begin Comandos.cmdComandos cmdComandos 
      Align           =   2  'Align Bottom
      Height          =   765
      Left            =   0
      TabIndex        =   15
      Top             =   6180
      Width           =   10785
      _ExtentX        =   19024
      _ExtentY        =   1349
      EnabledNovo     =   0   'False
   End
   Begin VB.Frame fraGrid 
      Caption         =   "Sessões Cadastradas"
      Height          =   3315
      Left            =   60
      TabIndex        =   19
      Top             =   60
      Width           =   10635
      Begin VSFlex7LCtl.VSFlexGrid VSFlexGrid 
         Height          =   2415
         Left            =   180
         TabIndex        =   1
         Top             =   720
         Width           =   10275
         _cx             =   18124
         _cy             =   4260
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
         Cols            =   25
         FixedRows       =   1
         FixedCols       =   0
         RowHeightMin    =   0
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   -1  'True
         FormatString    =   $"frmSessoes.frx":000C
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
      Begin Combo.cboCodDesc ccd_prg_cd 
         Height          =   315
         Left            =   1320
         TabIndex        =   0
         Top             =   300
         Width           =   5415
         _ExtentX        =   9551
         _ExtentY        =   556
         NomeTabela      =   "tb_programacao"
         NomeCampoCodigo =   "prg_cd"
         NomeCampoDescricao=   "convert(char(10),prg_dt_ini,111) + ' - ' + convert(char(10),prg_dt_fim,111)"
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
         CodigoVisible   =   0   'False
         Filtro          =   "prg_dt_des is null"
         MostraBotaoAtualiza=   0   'False
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Programação:"
         Height          =   195
         Left            =   240
         TabIndex        =   20
         Top             =   360
         Width           =   990
      End
   End
   Begin VB.Frame fraManut 
      Caption         =   "Manutenção"
      Enabled         =   0   'False
      Height          =   2745
      Left            =   60
      TabIndex        =   16
      Top             =   3390
      Width           =   10680
      Begin VB.CheckBox chkPreEstreia 
         Caption         =   "Pré-estréia"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   8460
         TabIndex        =   34
         Top             =   2130
         Width           =   1695
      End
      Begin Spin.SpinNumber spnPeriodo 
         Height          =   315
         Left            =   9960
         TabIndex        =   4
         Top             =   300
         Width           =   495
         _ExtentX        =   873
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
         Max             =   8
         Min             =   1
         Value           =   "1"
      End
      Begin VB.Frame fraSessoes 
         Caption         =   "Sessões"
         Height          =   885
         Left            =   180
         TabIndex        =   22
         Top             =   810
         Width           =   10305
         Begin MSMask.MaskEdBox msk_ses_cd 
            Height          =   315
            Index           =   2
            Left            =   1590
            TabIndex        =   6
            Tag             =   "2"
            Top             =   450
            Width           =   750
            _ExtentX        =   1323
            _ExtentY        =   556
            _Version        =   393216
            MaxLength       =   5
            Format          =   "hh:mm"
            Mask            =   "##:##"
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox msk_ses_cd 
            Height          =   315
            Index           =   3
            Left            =   2520
            TabIndex        =   7
            Tag             =   "3"
            Top             =   450
            Width           =   750
            _ExtentX        =   1323
            _ExtentY        =   556
            _Version        =   393216
            MaxLength       =   5
            Format          =   "hh:mm"
            Mask            =   "##:##"
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox msk_ses_cd 
            Height          =   315
            Index           =   4
            Left            =   3450
            TabIndex        =   8
            Tag             =   "4"
            Top             =   450
            Width           =   750
            _ExtentX        =   1323
            _ExtentY        =   556
            _Version        =   393216
            MaxLength       =   5
            Format          =   "hh:mm"
            Mask            =   "##:##"
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox msk_ses_cd 
            Height          =   315
            Index           =   5
            Left            =   4380
            TabIndex        =   9
            Tag             =   "5"
            Top             =   450
            Width           =   750
            _ExtentX        =   1323
            _ExtentY        =   556
            _Version        =   393216
            MaxLength       =   5
            Format          =   "hh:mm"
            Mask            =   "##:##"
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox msk_ses_cd 
            Height          =   315
            Index           =   10
            Left            =   9030
            TabIndex        =   14
            Tag             =   "10"
            Top             =   450
            Width           =   750
            _ExtentX        =   1323
            _ExtentY        =   556
            _Version        =   393216
            MaxLength       =   5
            Format          =   "hh:mm"
            Mask            =   "##:##"
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox msk_ses_cd 
            Height          =   315
            Index           =   1
            Left            =   660
            TabIndex        =   5
            Tag             =   "1"
            Top             =   450
            Width           =   750
            _ExtentX        =   1323
            _ExtentY        =   556
            _Version        =   393216
            MaxLength       =   5
            Format          =   "hh:mm"
            Mask            =   "##:##"
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox msk_ses_cd 
            Height          =   315
            Index           =   7
            Left            =   6240
            TabIndex        =   11
            Tag             =   "7"
            Top             =   450
            Width           =   750
            _ExtentX        =   1323
            _ExtentY        =   556
            _Version        =   393216
            MaxLength       =   5
            Format          =   "hh:mm"
            Mask            =   "##:##"
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox msk_ses_cd 
            Height          =   315
            Index           =   8
            Left            =   7170
            TabIndex        =   12
            Tag             =   "8"
            Top             =   450
            Width           =   750
            _ExtentX        =   1323
            _ExtentY        =   556
            _Version        =   393216
            MaxLength       =   5
            Format          =   "hh:mm"
            Mask            =   "##:##"
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox msk_ses_cd 
            Height          =   315
            Index           =   9
            Left            =   8100
            TabIndex        =   13
            Tag             =   "9"
            Top             =   450
            Width           =   750
            _ExtentX        =   1323
            _ExtentY        =   556
            _Version        =   393216
            MaxLength       =   5
            Format          =   "hh:mm"
            Mask            =   "##:##"
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox msk_ses_cd 
            Height          =   315
            Index           =   6
            Left            =   5310
            TabIndex        =   10
            Tag             =   "6"
            Top             =   450
            Width           =   750
            _ExtentX        =   1323
            _ExtentY        =   556
            _Version        =   393216
            MaxLength       =   5
            Format          =   "hh:mm"
            Mask            =   "##:##"
            PromptChar      =   "_"
         End
         Begin VB.Label lblsessao 
            AutoSize        =   -1  'True
            Caption         =   "01"
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
            Left            =   870
            TabIndex        =   32
            Top             =   225
            Width           =   225
         End
         Begin VB.Label lblsessao 
            AutoSize        =   -1  'True
            Caption         =   "10"
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
            Left            =   9300
            TabIndex        =   31
            Top             =   225
            Width           =   225
         End
         Begin VB.Label lblsessao 
            AutoSize        =   -1  'True
            Caption         =   "09"
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
            Left            =   8355
            TabIndex        =   30
            Top             =   225
            Width           =   225
         End
         Begin VB.Label lblsessao 
            AutoSize        =   -1  'True
            Caption         =   "08"
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
            Left            =   7410
            TabIndex        =   29
            Top             =   225
            Width           =   225
         End
         Begin VB.Label lblsessao 
            AutoSize        =   -1  'True
            Caption         =   "07"
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
            Left            =   6480
            TabIndex        =   28
            Top             =   210
            Width           =   225
         End
         Begin VB.Label lblsessao 
            AutoSize        =   -1  'True
            Caption         =   "06"
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
            Left            =   5550
            TabIndex        =   27
            Top             =   225
            Width           =   225
         End
         Begin VB.Label lblsessao 
            AutoSize        =   -1  'True
            Caption         =   "05"
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
            Left            =   4605
            TabIndex        =   26
            Top             =   225
            Width           =   225
         End
         Begin VB.Label lblsessao 
            AutoSize        =   -1  'True
            Caption         =   "04"
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
            Left            =   3675
            TabIndex        =   25
            Top             =   225
            Width           =   225
         End
         Begin VB.Label lblsessao 
            AutoSize        =   -1  'True
            Caption         =   "03"
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
            Left            =   2745
            TabIndex        =   24
            Top             =   225
            Width           =   225
         End
         Begin VB.Label lblsessao 
            AutoSize        =   -1  'True
            Caption         =   "02"
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
            Left            =   1800
            TabIndex        =   23
            Top             =   225
            Width           =   225
         End
      End
      Begin VB.Frame fraDiasSemanaSessao 
         Caption         =   "Dias da Semana"
         Height          =   825
         Left            =   180
         TabIndex        =   21
         Top             =   1815
         Width           =   7815
         Begin VB.CheckBox chk_ses_dia_semana 
            Caption         =   "Segunda"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Index           =   2
            Left            =   750
            TabIndex        =   42
            Tag             =   "1"
            Top             =   255
            Width           =   1095
         End
         Begin VB.CheckBox chk_ses_dia_semana 
            Caption         =   "Terça"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Index           =   3
            Left            =   750
            TabIndex        =   41
            Tag             =   "2"
            Top             =   480
            Width           =   1095
         End
         Begin VB.CheckBox chk_ses_dia_semana 
            Caption         =   "Quarta"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Index           =   4
            Left            =   2490
            TabIndex        =   40
            Tag             =   "3"
            Top             =   255
            Width           =   1095
         End
         Begin VB.CheckBox chk_ses_dia_semana 
            Caption         =   "Quinta"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Index           =   5
            Left            =   2490
            TabIndex        =   39
            Tag             =   "4"
            Top             =   480
            Width           =   1095
         End
         Begin VB.CheckBox chk_ses_dia_semana 
            Caption         =   "Sexta"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Index           =   6
            Left            =   4230
            TabIndex        =   38
            Tag             =   "5"
            Top             =   255
            Width           =   1095
         End
         Begin VB.CheckBox chk_ses_dia_semana 
            Caption         =   "Sábado"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Index           =   7
            Left            =   4230
            TabIndex        =   37
            Tag             =   "6"
            Top             =   480
            Width           =   1095
         End
         Begin VB.CheckBox chk_ses_dia_semana 
            Caption         =   "Domingo"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Index           =   1
            Left            =   5970
            TabIndex        =   36
            Tag             =   "7"
            Top             =   255
            Width           =   1095
         End
         Begin VB.CheckBox chk_ses_dia_semana 
            Caption         =   "Feriados"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Index           =   8
            Left            =   5970
            TabIndex        =   35
            Tag             =   "8"
            Top             =   480
            Width           =   1095
         End
      End
      Begin Combo.cboCodDesc ccd_fil_cd 
         Height          =   315
         Left            =   5220
         TabIndex        =   3
         Top             =   300
         Width           =   4005
         _ExtentX        =   7064
         _ExtentY        =   556
         NomeTabela      =   "TB_FILME"
         NomeCampoCodigo =   "fil_cd"
         NomeCampoDescricao=   "fil_nm"
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
         MostraBotaoAtualiza=   0   'False
      End
      Begin Combo.cboCodDesc ccd_sal_cd 
         Height          =   315
         Left            =   600
         TabIndex        =   2
         Top             =   300
         Width           =   4000
         _ExtentX        =   7064
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
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Período:"
         Height          =   195
         Left            =   9300
         TabIndex        =   33
         Top             =   360
         Width           =   615
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Filme:"
         Height          =   195
         Left            =   4740
         TabIndex        =   18
         Top             =   360
         Width           =   405
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Sala:"
         Height          =   195
         Left            =   180
         TabIndex        =   17
         Top             =   360
         Width           =   360
      End
   End
End
Attribute VB_Name = "frmSessoes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private sOperacao As String

Private Sub ccd_fil_cd_BeforeProcuraClick(Cancel As Boolean)

    Dim sFiltro As String
    
    sFiltro = " fil_dt_des is null " & _
              "  and ( (   fil_dt_ini <= ( select prg_dt_ini from tb_programacao where prg_cd = " & ccd_prg_cd.codigo & " ) " & _
              "        and fil_dt_fim > ( select prg_dt_ini from tb_programacao where prg_cd = " & ccd_prg_cd.codigo & " ) " & _
              "      ) " & _
              "   or (     fil_dt_ini >= ( select prg_dt_ini from tb_programacao where prg_cd = " & ccd_prg_cd.codigo & " ) " & _
              "        and fil_dt_ini <  ( select prg_dt_fim from tb_programacao where prg_cd = " & ccd_prg_cd.codigo & " ) " & _
              "      ) )"

    ccd_fil_cd.Filtro = sFiltro

End Sub

Private Sub ccd_prg_cd_AfterProcuraClick()
    
    If ccd_prg_cd.codigo <> "" Then
        Call PreencheGrid
    End If
    
    cmdComandos.EnabledAltera = (ccd_prg_cd.codigo <> "")
    cmdComandos.EnabledExclui = (ccd_prg_cd.codigo <> "")
    cmdComandos.EnabledNovo = (ccd_prg_cd.codigo <> "")
    
End Sub

Private Sub ccd_prg_cd_Change()
    Call ccd_prg_cd_AfterProcuraClick
End Sub

Private Sub Form_Load()

    Set ccd_prg_cd.ConexaoADO = dbConnect
    Set ccd_sal_cd.ConexaoADO = dbConnect
    Set ccd_fil_cd.ConexaoADO = dbConnect
    
    Call HabilitaManut(False)

    Dim CineAx As New CineAux.CineAx
    Call CineAx.CentralizaTela(MDIFrmAdmCine, Me)
    
    Set CineAx = Nothing
    
    VSFlexGrid.ColHidden(VSFlexGrid.ColIndex("KEY_PRG_CD")) = True
    VSFlexGrid.ColHidden(VSFlexGrid.ColIndex("KEY_SAL_CD")) = True
    VSFlexGrid.ColHidden(VSFlexGrid.ColIndex("KEY_FIL_CD")) = True

    Set ccd_prg_cd.NomeForm = frmProgramacao
    
    Dim log As New clsLog
    
    Set log.ConexaoADO = dbConnect
    log.usu_nm = strUsuario
    log.slg_descricao = "Entrou na tela de " & Me.Caption & " do módulo " & App.ProductName
    
    log.insereLog
    
End Sub

Private Sub cmdComandos_Click(ByVal iButtonClicked As Comandos.eButtonClicked, Cancel As Boolean)
    
    Select Case iButtonClicked
    
        Case ButtonAltera
            If permiteAltExc() Then
                sOperacao = "A"
            
                ccd_sal_cd.Enabled = False
                ccd_fil_cd.Enabled = False
                spnPeriodo.Enabled = False
                
                If VSFlexGrid.RowSel > 0 And VSFlexGrid.Rows > 1 Then
                    Call CarregaControles
                    Call HabilitaManut(True)
                Else
                    Cancel = True
                End If
            Else
                Cancel = True
                MsgBox "Não é possível alterar sessão. Período anterior a data atual", vbCritical, App.ProductName
            End If
            
        Case ButtonNovo
            sOperacao = "I"
            
            ccd_sal_cd.Enabled = True
            ccd_fil_cd.Enabled = True
            spnPeriodo.Enabled = True
            
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
                MsgBox "Não é possível excluir sessão. Período anterior a data atual", vbCritical, App.ProductName
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

Private Sub HabilitaManut(ByVal bHabilita As Boolean)
    fraGrid.Enabled = Not bHabilita
    fraManut.Enabled = bHabilita
End Sub

Private Sub CarregaControles()

    Call LimpaControles

    If VSFlexGrid.RowSel > 0 And VSFlexGrid.Rows > 1 Then
    
        ccd_sal_cd.codigo = VSFlexGrid.TextMatrix(VSFlexGrid.RowSel, VSFlexGrid.ColIndex("KEY_SAL_CD"))
        ccd_sal_cd.Refresh
        ccd_fil_cd.codigo = VSFlexGrid.TextMatrix(VSFlexGrid.RowSel, VSFlexGrid.ColIndex("KEY_FIL_CD"))
        ccd_fil_cd.Refresh
        
        spnPeriodo.Value = VSFlexGrid.TextMatrix(VSFlexGrid.RowSel, VSFlexGrid.ColIndex("KEY_SES_PERIODO"))
        
        chkPreEstreia.Value = IIf(VSFlexGrid.TextMatrix(VSFlexGrid.RowSel, VSFlexGrid.ColIndex("KEY_PREESTREIA")) = "S", vbChecked, vbUnchecked)
        
        Dim i As Integer
    
        For i = 1 To 10
            If VSFlexGrid.TextMatrix(VSFlexGrid.RowSel, VSFlexGrid.ColIndex("KEY_SESSAO" & Format(i, "00"))) <> "" Then
                msk_ses_cd(i).Text = VSFlexGrid.TextMatrix(VSFlexGrid.RowSel, VSFlexGrid.ColIndex("KEY_SESSAO" & Format(i, "00")))
            Else
                msk_ses_cd(i).Text = "__:__"
            End If
        Next
        
        For i = 1 To 8
            If VSFlexGrid.TextMatrix(VSFlexGrid.RowSel, VSFlexGrid.ColIndex("KEY_DIA" & i)) = "" Then
                chk_ses_dia_semana(i).Value = vbUnchecked
            Else
                chk_ses_dia_semana(i).Value = vbChecked
            End If
        Next
        
    End If
    
End Sub

Private Sub LimpaControles()

    Dim i As Integer
    
    ccd_sal_cd.codigo = ""
    ccd_fil_cd.codigo = ""
    spnPeriodo.Value = 0
    
    For i = 1 To 10
        msk_ses_cd(i).Text = "__:__"
    Next
    
    For i = 1 To 8
        chk_ses_dia_semana(i).Value = vbUnchecked
    Next
    
End Sub

Private Function Grava() As Boolean

    On Error GoTo Grava_Erro
    
    If Not Consiste() Then
        Exit Function
    End If
    
    Dim clsTB_SESSAO As New Cine2005.clsTB_SESSAO
    
    Set clsTB_SESSAO.ConexaoADO = dbConnect
    
    clsTB_SESSAO.prg_cd = ccd_prg_cd.codigo
    clsTB_SESSAO.sal_cd = ccd_sal_cd.codigo
    clsTB_SESSAO.fil_cd = ccd_fil_cd.codigo
    clsTB_SESSAO.ses_periodo = spnPeriodo.Value
    If chkPreEstreia.Value = vbChecked Then
        clsTB_SESSAO.ses_pre_estreia = "S"
    Else
        clsTB_SESSAO.ses_pre_estreia = "N"
    End If
    
    If sOperacao = "A" Then
        If Not clsTB_SESSAO.Excluir() Then
            MsgBox clsTB_SESSAO.MensagemErro, vbCritical, App.ProductName
            GoTo Grava_Fim
        End If
    End If
    
    Dim iSessoes As Integer
    Dim iDias As Integer
    
    For iSessoes = 1 To 10
    
        If msk_ses_cd(iSessoes).Text <> "__:__" Then
        
            For iDias = 1 To 8
            
                If chk_ses_dia_semana(iDias).Value = vbChecked Then
                
                    clsTB_SESSAO.ses_cd = iSessoes
                    clsTB_SESSAO.ses_dia_semana = iDias
                    
                    If CVDate(msk_ses_cd(iSessoes).Text) < dtHoraMaxSessao Then
                        clsTB_SESSAO.ses_horario = strDataRef2 & " " & msk_ses_cd(iSessoes).Text
                    Else
                        clsTB_SESSAO.ses_horario = strDataRef1 & " " & msk_ses_cd(iSessoes).Text
                    End If
                    
                    If Not clsTB_SESSAO.Incluir() Then
                        MsgBox "Não foi possível incluir a Pogramação!" & vbCrLf & clsTB_SESSAO.MensagemErro, vbInformation, App.ProductName
                        GoTo Grava_Fim
                    End If
                    
                End If
            Next
            
        End If
        
    Next
    
    Call PreencheGrid

    Grava = True
    GoTo Grava_Fim

Grava_Erro:
    MsgBox "Erro de execução! 'Grava/frmSessoes'", vbCritical, App.ProductName
    
Grava_Fim:
    Set clsTB_SESSAO = Nothing
    
End Function
Private Function Exclui() As Boolean

    On Error GoTo Exclui_Erro
    
    Dim clsTB_SESSAO As New Cine2005.clsTB_SESSAO

    Set clsTB_SESSAO.ConexaoADO = dbConnect
    
    clsTB_SESSAO.prg_cd = ccd_prg_cd.codigo
    clsTB_SESSAO.sal_cd = ccd_sal_cd.codigo
    clsTB_SESSAO.fil_cd = ccd_fil_cd.codigo
    clsTB_SESSAO.ses_periodo = spnPeriodo.Value
    
    If Not clsTB_SESSAO.Excluir() Then
        MsgBox "Não foi possível excluir a Programação Selecionada!" & vbCrLf & clsTB_SESSAO.MensagemErro, vbInformation, App.ProductName
        GoTo Exclui_Fim
    End If
            
    Exclui = True
    GoTo Exclui_Fim
    
Exclui_Erro:
    MsgBox "Erro de execução! 'Exclui/frmSessoes'", vbCritical, App.ProductName
    
Exclui_Fim:
    Set clsTB_SESSAO = Nothing
    
End Function
Private Sub PreencheGrid()
    Dim oRs          As New ADODB.Recordset
    Dim clsTB_SESSAO As New Cine2005.clsTB_SESSAO
    Dim iProgramacao As Integer
    Dim iSala        As Integer
    Dim iFilme       As Long
    Dim iPeriodo     As Integer

    On Error GoTo PreencheGrid_Erro
    
    Call LimpaControles
    
    Set clsTB_SESSAO.ConexaoADO = dbConnect
    
    clsTB_SESSAO.prg_cd = ccd_prg_cd.codigo
    
    If Not clsTB_SESSAO.PreencheGrid(oRs) Then
        MsgBox clsTB_SESSAO.MensagemErro, vbCritical, App.ProductName
        GoTo PreencheGrid_Fim
    End If

    VSFlexGrid.Rows = 1

    Do While Not oRs.EOF()
    
        If iProgramacao <> oRs.Fields("prg_cd") Or _
            iSala <> oRs.Fields("sal_cd") Or _
            iFilme <> oRs.Fields("fil_cd") Or _
            iPeriodo <> oRs.Fields("ses_periodo") Then
        
            iProgramacao = oRs.Fields("prg_cd")
            iSala = oRs.Fields("sal_cd")
            iFilme = oRs.Fields("fil_cd")
            iPeriodo = oRs.Fields("ses_periodo")
        
            VSFlexGrid.Rows = VSFlexGrid.Rows + 1
            
            VSFlexGrid.TextMatrix(VSFlexGrid.Rows - 1, VSFlexGrid.ColIndex("KEY_PRG_CD")) = oRs.Fields("prg_cd")
            VSFlexGrid.TextMatrix(VSFlexGrid.Rows - 1, VSFlexGrid.ColIndex("KEY_SAL_CD")) = oRs.Fields("sal_cd")
            VSFlexGrid.TextMatrix(VSFlexGrid.Rows - 1, VSFlexGrid.ColIndex("KEY_FIL_CD")) = oRs.Fields("fil_cd")
            VSFlexGrid.TextMatrix(VSFlexGrid.Rows - 1, VSFlexGrid.ColIndex("KEY_SES_PERIODO")) = oRs.Fields("ses_periodo")
            VSFlexGrid.TextMatrix(VSFlexGrid.Rows - 1, VSFlexGrid.ColIndex("KEY_SALA")) = oRs.Fields("sal_desc")
            VSFlexGrid.TextMatrix(VSFlexGrid.Rows - 1, VSFlexGrid.ColIndex("KEY_FILME")) = oRs.Fields("fil_nm")
            VSFlexGrid.TextMatrix(VSFlexGrid.Rows - 1, VSFlexGrid.ColIndex("KEY_PREESTREIA")) = oRs.Fields("ses_pre_estreia")
        End If
        
        VSFlexGrid.TextMatrix(VSFlexGrid.Rows - 1, VSFlexGrid.ColIndex("KEY_SESSAO" & Format(oRs.Fields("ses_cd"), "00"))) = Format(oRs.Fields("ses_horario"), "hh:mm")
        VSFlexGrid.TextMatrix(VSFlexGrid.Rows - 1, VSFlexGrid.ColIndex("KEY_DIA" & oRs.Fields("ses_dia_semana"))) = True
        
        oRs.MoveNext
        
    Loop
    
    VSFlexGrid.MergeCells = flexMergeRestrictColumns
    VSFlexGrid.MergeCol(VSFlexGrid.ColIndex("KEY_SALA")) = True

    VSFlexGrid.AutoSizeMode = flexAutoSizeColWidth
    Call VSFlexGrid.AutoSize(0, VSFlexGrid.Cols - 1)
    
    VSFlexGrid.ColWidth(VSFlexGrid.ColIndex("KEY_DIA8")) = 1000

    VSFlexGrid.FrozenCols = 6

    VSFlexGrid.Row = IIf(VSFlexGrid.Rows > 1, 1, 0)
    Call CarregaControles

    GoTo PreencheGrid_Fim
    
PreencheGrid_Erro:
    MsgBox "Erro de execução! 'PreencheGrid/frmSessoes'", vbCritical, App.ProductName
    
PreencheGrid_Fim:
    If oRs.State = 1 Then oRs.Close
    Set oRs = Nothing
    Set clsTB_SESSAO = Nothing
    
End Sub

Private Function Consiste() As Boolean
    
    Dim sMens   As String
    Dim i       As Integer
    Dim bMarcou As Boolean
    Dim n       As Integer
    
    Dim oRs          As New ADODB.Recordset
    Dim clsTB_SESSAO As New Cine2005.clsTB_SESSAO
    
    Set clsTB_SESSAO.ConexaoADO = dbConnect
    
    clsTB_SESSAO.prg_cd = ccd_prg_cd.codigo
    clsTB_SESSAO.sal_cd = ccd_sal_cd.codigo
    clsTB_SESSAO.fil_cd = ccd_fil_cd.codigo
    
    If sOperacao = "I" Then
    
        clsTB_SESSAO.ses_periodo = spnPeriodo.Value
        
        ' Verifica se já não existe sessões cadastradas para o filme
        
        If Not clsTB_SESSAO.Selecionar(oRs) Then
            MsgBox clsTB_SESSAO.MensagemErro, vbCritical, App.ProductName
            Exit Function
        End If
        
        If Not oRs.EOF() Then
            MsgBox "Este filme já foi cadastrado para esta sala nesse período!", vbCritical, App.ProductName
            GoTo Consiste_fim
        End If
        
        If oRs.State = 1 Then oRs.Close
    
    End If
    
    ' Verifica se não está duplicando para dias iguais
    
    clsTB_SESSAO.ses_periodo = Empty
    
    If Not clsTB_SESSAO.Selecionar(oRs) Then
        MsgBox clsTB_SESSAO.MensagemErro, vbCritical, App.ProductName
        Exit Function
    End If
    
    Do While Not oRs.EOF()
        For i = 1 To 8
            If chk_ses_dia_semana(i).Value = vbChecked Then
                If oRs.Fields("ses_periodo") <> spnPeriodo.Value Then
                    If oRs.Fields("ses_dia_semana") = i Then
                        MsgBox "O(s) dia(s) da semana selecionado(s) já pertence(m) a outro período!", vbCritical, App.ProductName
                        GoTo Consiste_fim
                    End If
                End If
            End If
        Next
        oRs.MoveNext
    Loop
    
    If oRs.State = 1 Then oRs.Close
    
    If ccd_sal_cd.codigo = "" Then
        sMens = sMens & "Sala deve ser informada!" & vbCrLf
    End If

    If ccd_fil_cd.codigo = "" Then
        sMens = sMens & "Filme deve ser informado!" & vbCrLf
    End If
    
    bMarcou = False
    
    For i = 1 To 10
        If msk_ses_cd(i).Text <> "__:__" Then
            bMarcou = True
            Exit For
        End If
    Next
    
    If Not bMarcou Then
        sMens = sMens & "Deve-se informar pelo menos uma sessão!" & vbCrLf
    Else
        For i = 1 To 10
            If msk_ses_cd(i).Text <> "__:__" Then
                If Not IsDate(msk_ses_cd(i).Text) Then
                    sMens = sMens & "Horário inválido!" & vbCrLf
                    msk_ses_cd(i).SetFocus
                    Exit For
                ElseIf chkPreEstreia.Value = vbChecked Then
                    If Not (CVDate(msk_ses_cd(i).Text) >= CVDate("00:00") And CVDate(msk_ses_cd(i).Text) <= dtHoraMaxSessao) Then
                        sMens = sMens & "Horário inválido! Para Pré-estréia as sessões devem ser de 00:00 as " & Format(dtHoraMaxSessao, "Hh:Nn") & vbCrLf
                        msk_ses_cd(i).SetFocus
                        Exit For
                    End If
                End If
            End If
        Next
    End If

    bMarcou = False
    n = 0
    For i = 1 To 8
        If chk_ses_dia_semana(i).Value = vbChecked Then
            n = n + 1
            bMarcou = True
            If chkPreEstreia.Value <> vbChecked Then
                Exit For
            End If
        End If
    Next
    
    If Not bMarcou Then
        sMens = sMens & "Deve-se informar pelo menos um dia da semana!" & vbCrLf
    ElseIf chkPreEstreia.Value = vbChecked Then
        If n > 1 Then
            sMens = sMens & "Pré-estréia deve ser apenas um dia da semana!" & vbCrLf
        End If
    End If

    If sMens <> "" Then
        MsgBox sMens, vbInformation, App.ProductName
        Exit Function
    End If
    
    ' Pega tempo do filme
    
    Dim iTempoFilme As Integer
    Dim clsTB_Filme As New Cine2005.clsTB_Filme
    
    Set clsTB_Filme.ConexaoADO = dbConnect
    
    clsTB_Filme.fil_cd = ccd_fil_cd.codigo
    
    If Not clsTB_Filme.Selecionar(oRs) Then
        MsgBox clsTB_Filme.MensagemErro, vbCritical, App.ProductName
        Exit Function
    End If
    
    If Not oRs.EOF() Then
        iTempoFilme = oRs.Fields("fil_duracao")
    End If
    
    oRs.Close
    Set clsTB_Filme = Nothing
    
    ' Verifica se existe colisão nas sessões
    
    Dim dtSessaoAnt As Date
    Dim dtSessaoAux As Date
    Dim dtSessaoAtu As Date
    
    dtSessaoAnt = Empty
    
    For i = 1 To 10
        If msk_ses_cd(i).Text <> "__:__" Then
            If dtSessaoAnt <> Empty Then
            
                ' Soma o tempo entre as sessões
                dtSessaoAux = DateAdd("n", intTempoEntreSessoes, dtSessaoAnt)
                
                ' Soma o tempo do filme
                dtSessaoAux = DateAdd("n", iTempoFilme, dtSessaoAux)
                
                If CDate(msk_ses_cd(i).Text) > CDate(dtHoraMaxSessao) Then
                    dtSessaoAtu = CDate(strDataRef1 & " " & msk_ses_cd(i).Text)
                Else
                    dtSessaoAtu = CDate(strDataRef2 & " " & msk_ses_cd(i).Text)
                End If
                
                dtSessaoAtu = DateAdd("s", 1, dtSessaoAtu)
                
                If dtSessaoAux > dtSessaoAtu Then
                    sMens = "Ocorreu uma colisão de horário!" & vbCrLf
                    sMens = sMens & "O horário " & msk_ses_cd(i).Text & " deve ser maior ou igual a " & Format(dtSessaoAux, "Hh:Nn") & "!"
                    MsgBox sMens, vbCritical, App.ProductName
                    Exit Function
                End If
            
            End If
            If CDate(msk_ses_cd(i).Text) > CDate(dtHoraMaxSessao) Then
                dtSessaoAnt = CDate(strDataRef1 & " " & msk_ses_cd(i).Text)
            Else
                dtSessaoAnt = CDate(strDataRef2 & " " & msk_ses_cd(i).Text)
            End If
        End If
    Next
    
    ' Verifica se para a mesma programação não existe outro filme na mesma sala e mesmo horário para o mesmo dia
    
    Dim iDia As Integer
    
    For i = 1 To 10
    
        If msk_ses_cd(i).Text <> "__:__" Then
            
            If CDate(msk_ses_cd(i).Text) > CDate(dtHoraMaxSessao) Then
                clsTB_SESSAO.ses_horario = CDate(strDataRef1 & " " & msk_ses_cd(i).Text)
            Else
                clsTB_SESSAO.ses_horario = CDate(strDataRef2 & " " & msk_ses_cd(i).Text)
            End If
            
            clsTB_SESSAO.prg_cd = ccd_prg_cd.codigo
            clsTB_SESSAO.sal_cd = ccd_sal_cd.codigo
            clsTB_SESSAO.fil_cd = ccd_fil_cd.codigo
            
            For iDia = 1 To 8
                If chk_ses_dia_semana(iDia).Value = vbChecked Then
                    clsTB_SESSAO.ses_dia_semana = iDia
                    If clsTB_SESSAO.TemColisao Then
                        MsgBox clsTB_SESSAO.MensagemErro, vbCritical, App.ProductName
                        GoTo Consiste_fim
                    End If
                End If
            Next
        End If
    Next
        
    Consiste = True
    
Consiste_fim:
    If oRs.State = 1 Then oRs.Close
    Set clsTB_SESSAO = Nothing
    
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
    Dim dia    As Integer
    Dim mes    As Integer
    Dim ano    As Integer
    Dim strAux As String
    
    permiteAltExc = False
    
    strAux = ccd_prg_cd.Descricao
    
    If IsNumeric(Mid(strAux, 9, 2)) Then
        dia = CInt(Mid(strAux, 9, 2))
    Else
        Exit Function
    End If
    
    If IsNumeric(Mid(strAux, 6, 2)) Then
        mes = CInt(Mid(strAux, 6, 2))
    Else
        Exit Function
    End If
    
    If IsNumeric(Mid(strAux, 1, 4)) Then
        ano = CInt(Mid(strAux, 1, 4))
    Else
        Exit Function
    End If
    
    dtIni = DateSerial(ano, mes, dia)
    
    If IsNumeric(Mid(strAux, 22, 2)) Then
        dia = CInt(Mid(strAux, 22, 2))
    Else
        Exit Function
    End If
    
    If IsNumeric(Mid(strAux, 19, 2)) Then
        mes = CInt(Mid(strAux, 19, 2))
    Else
        Exit Function
    End If
    
    If IsNumeric(Mid(strAux, 14, 4)) Then
        ano = CInt(Mid(strAux, 14, 4))
    Else
        Exit Function
    End If
    
    dtFim = DateSerial(ano, mes, dia)
    
    dtAtu = CDate(Format(Date, "Short Date"))
    
    If verificaPeriodo(dtIni, dtFim, dtAtu) > 0 Then
        permiteAltExc = True
    End If
End Function

