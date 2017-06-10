VERSION 5.00
Object = "{C0A63B80-4B21-11D3-BD95-D426EF2C7949}#1.0#0"; "Vsflex7L.ocx"
Object = "{234FCAFF-8D53-4DC2-9CD1-F90F4F6CB524}#32.0#0"; "Comandos.ocx"
Object = "{1753B334-0119-4B34-9134-D2B3CD181550}#36.0#0"; "Combo.ocx"
Object = "{91C13016-45DE-491B-BAC0-37F755626532}#31.0#0"; "Spin.ocx"
Begin VB.Form frmPreco1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Cadastro - Preços"
   ClientHeight    =   7170
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   10635
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7170
   ScaleWidth      =   10635
   Begin VB.Frame fraManut 
      Caption         =   "Manutenção"
      Enabled         =   0   'False
      Height          =   3435
      Left            =   45
      TabIndex        =   27
      Top             =   2910
      Width           =   10470
      Begin VB.Frame fraDiasSemanaSessao 
         Caption         =   "Dias da Semana"
         Height          =   765
         Left            =   75
         TabIndex        =   42
         Top             =   2535
         Width           =   10230
         Begin VB.Frame Frame1 
            BorderStyle     =   0  'None
            Caption         =   "Frame1"
            Height          =   540
            Left            =   8025
            TabIndex        =   51
            Top             =   150
            Width           =   2265
            Begin VB.CheckBox optPromocao 
               Caption         =   "Mostrar Promoção no Videl Hall"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   390
               Left            =   75
               TabIndex        =   52
               Top             =   75
               Width           =   1890
            End
         End
         Begin VB.CheckBox chk_pre_dia_semana 
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
            Left            =   6615
            TabIndex        =   23
            Tag             =   "8"
            Top             =   420
            Width           =   1095
         End
         Begin VB.CheckBox chk_pre_dia_semana 
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
            Left            =   6615
            TabIndex        =   22
            Tag             =   "7"
            Top             =   195
            Width           =   1095
         End
         Begin VB.CheckBox chk_pre_dia_semana 
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
            Left            =   4875
            TabIndex        =   21
            Tag             =   "6"
            Top             =   420
            Width           =   1095
         End
         Begin VB.CheckBox chk_pre_dia_semana 
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
            Left            =   4875
            TabIndex        =   20
            Tag             =   "5"
            Top             =   195
            Width           =   1095
         End
         Begin VB.CheckBox chk_pre_dia_semana 
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
            Left            =   3135
            TabIndex        =   19
            Tag             =   "4"
            Top             =   420
            Width           =   1095
         End
         Begin VB.CheckBox chk_pre_dia_semana 
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
            Left            =   3135
            TabIndex        =   18
            Tag             =   "3"
            Top             =   195
            Width           =   1095
         End
         Begin VB.CheckBox chk_pre_dia_semana 
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
            Left            =   1350
            TabIndex        =   17
            Tag             =   "2"
            Top             =   420
            Width           =   1095
         End
         Begin VB.CheckBox chk_pre_dia_semana 
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
            Left            =   1350
            TabIndex        =   16
            Tag             =   "1"
            Top             =   195
            Width           =   1095
         End
      End
      Begin VB.Frame fra_inteira 
         Caption         =   "Inteira"
         Height          =   780
         Left            =   120
         TabIndex        =   35
         Top             =   690
         Width           =   10230
         Begin VB.TextBox flt_pre_vl_inteira_6 
            Alignment       =   1  'Right Justify
            Height          =   315
            Left            =   9240
            TabIndex        =   9
            Text            =   "0,00"
            Top             =   285
            Width           =   735
         End
         Begin VB.TextBox flt_pre_vl_inteira_5 
            Alignment       =   1  'Right Justify
            Height          =   315
            Left            =   7575
            TabIndex        =   8
            Text            =   "0,00"
            Top             =   285
            Width           =   735
         End
         Begin VB.TextBox flt_pre_vl_inteira_4 
            Alignment       =   1  'Right Justify
            Height          =   315
            Left            =   5910
            TabIndex        =   7
            Text            =   "0,00"
            Top             =   285
            Width           =   735
         End
         Begin VB.TextBox flt_pre_vl_inteira_3 
            Alignment       =   1  'Right Justify
            Height          =   315
            Left            =   4260
            TabIndex        =   6
            Text            =   "0,00"
            Top             =   285
            Width           =   735
         End
         Begin VB.TextBox flt_pre_vl_inteira_2 
            Alignment       =   1  'Right Justify
            Height          =   315
            Left            =   2595
            TabIndex        =   5
            Text            =   "0,00"
            Top             =   285
            Width           =   735
         End
         Begin VB.TextBox flt_pre_vl_inteira_1 
            Alignment       =   1  'Right Justify
            Height          =   315
            Left            =   945
            TabIndex        =   4
            Text            =   "0,00"
            Top             =   285
            Width           =   735
         End
         Begin VB.Label lbl_pre_vl_inteira_1 
            AutoSize        =   -1  'True
            Caption         =   "Preço 1:"
            Height          =   195
            Left            =   255
            TabIndex        =   41
            Top             =   330
            Width           =   600
         End
         Begin VB.Label lbl_pre_vl_inteira_2 
            AutoSize        =   -1  'True
            Caption         =   "Preço 2:"
            Height          =   195
            Left            =   1920
            TabIndex        =   40
            Top             =   330
            Width           =   600
         End
         Begin VB.Label lbl_pre_vl_inteira_3 
            AutoSize        =   -1  'True
            Caption         =   "Preço 3:"
            Height          =   195
            Left            =   3585
            TabIndex        =   39
            Top             =   330
            Width           =   600
         End
         Begin VB.Label lbl_pre_vl_inteira_4 
            AutoSize        =   -1  'True
            Caption         =   "Preço 4:"
            Height          =   195
            Left            =   5235
            TabIndex        =   37
            Top             =   330
            Width           =   600
         End
         Begin VB.Label lbl_pre_vl_inteira_5 
            AutoSize        =   -1  'True
            Caption         =   "Preço 5:"
            Height          =   195
            Left            =   6900
            TabIndex        =   38
            Top             =   330
            Width           =   600
         End
         Begin VB.Label lbl_pre_vl_inteira_6 
            AutoSize        =   -1  'True
            Caption         =   "Preço 6:"
            Height          =   195
            Left            =   8565
            TabIndex        =   36
            Top             =   330
            Width           =   600
         End
      End
      Begin VB.Frame fra_meia 
         Caption         =   "Meia"
         Height          =   780
         Left            =   120
         TabIndex        =   28
         Top             =   1710
         Width           =   10230
         Begin VB.TextBox flt_pre_vl_meia_6 
            Alignment       =   1  'Right Justify
            Height          =   315
            Left            =   9240
            TabIndex        =   15
            Text            =   "0,00"
            Top             =   285
            Width           =   735
         End
         Begin VB.TextBox flt_pre_vl_meia_5 
            Alignment       =   1  'Right Justify
            Height          =   315
            Left            =   7575
            TabIndex        =   14
            Text            =   "0,00"
            Top             =   285
            Width           =   735
         End
         Begin VB.TextBox flt_pre_vl_meia_4 
            Alignment       =   1  'Right Justify
            Height          =   315
            Left            =   5910
            TabIndex        =   13
            Text            =   "0,00"
            Top             =   285
            Width           =   735
         End
         Begin VB.TextBox flt_pre_vl_meia_3 
            Alignment       =   1  'Right Justify
            Height          =   315
            Left            =   4260
            TabIndex        =   12
            Text            =   "0,00"
            Top             =   285
            Width           =   735
         End
         Begin VB.TextBox flt_pre_vl_meia_2 
            Alignment       =   1  'Right Justify
            Height          =   315
            Left            =   2595
            TabIndex        =   11
            Text            =   "0,00"
            Top             =   285
            Width           =   735
         End
         Begin VB.TextBox flt_pre_vl_meia_1 
            Alignment       =   1  'Right Justify
            Height          =   315
            Left            =   945
            TabIndex        =   10
            Text            =   "0,00"
            Top             =   285
            Width           =   735
         End
         Begin VB.Label lbl_pre_vl_meia_1 
            AutoSize        =   -1  'True
            Caption         =   "Preço 1:"
            Height          =   195
            Left            =   255
            TabIndex        =   33
            Top             =   330
            Width           =   600
         End
         Begin VB.Label lbl_pre_vl_meia_2 
            AutoSize        =   -1  'True
            Caption         =   "Preço 2:"
            Height          =   195
            Left            =   1917
            TabIndex        =   34
            Top             =   330
            Width           =   600
         End
         Begin VB.Label lbl_pre_vl_meia_3 
            AutoSize        =   -1  'True
            Caption         =   "Preço 3:"
            Height          =   195
            Left            =   3579
            TabIndex        =   32
            Top             =   330
            Width           =   600
         End
         Begin VB.Label lbl_pre_vl_meia_4 
            AutoSize        =   -1  'True
            Caption         =   "Preço 4:"
            Height          =   195
            Left            =   5241
            TabIndex        =   30
            Top             =   330
            Width           =   600
         End
         Begin VB.Label lbl_pre_vl_meia_5 
            AutoSize        =   -1  'True
            Caption         =   "Preço 5:"
            Height          =   195
            Left            =   6903
            TabIndex        =   31
            Top             =   330
            Width           =   600
         End
         Begin VB.Label lbl_pre_vl_meia_6 
            AutoSize        =   -1  'True
            Caption         =   "Preço 6:"
            Height          =   195
            Left            =   8565
            TabIndex        =   29
            Top             =   330
            Width           =   600
         End
      End
      Begin Spin.SpinNumber spnPeriodo 
         Height          =   315
         Left            =   9795
         TabIndex        =   3
         Top             =   285
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
      Begin Combo.cboCodDesc ccd_fil_cd 
         Height          =   315
         Left            =   720
         TabIndex        =   2
         Top             =   285
         Width           =   8295
         _ExtentX        =   14631
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
         ItemFixo        =   "<< Todos os Filmes >>"
      End
      Begin VB.Label lblHoraLimite5 
         AutoSize        =   -1  'True
         Caption         =   "00:00"
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
         Left            =   8700
         TabIndex        =   50
         Top             =   1515
         Width           =   555
      End
      Begin VB.Label lblHoraLimite4 
         AutoSize        =   -1  'True
         Caption         =   "00:00"
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
         Left            =   7050
         TabIndex        =   49
         Top             =   1515
         Width           =   555
      End
      Begin VB.Label lblHoraLimite3 
         AutoSize        =   -1  'True
         Caption         =   "00:00"
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
         Left            =   5355
         TabIndex        =   48
         Top             =   1515
         Width           =   555
      End
      Begin VB.Label lblHoraLimite2 
         AutoSize        =   -1  'True
         Caption         =   "00:00"
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
         Left            =   3705
         TabIndex        =   47
         Top             =   1515
         Width           =   555
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Horários:"
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
         Left            =   780
         TabIndex        =   46
         Top             =   1515
         Width           =   975
      End
      Begin VB.Label lblHoraLimite1 
         AutoSize        =   -1  'True
         Caption         =   "00:00"
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
         Left            =   2040
         TabIndex        =   45
         Top             =   1515
         Width           =   555
      End
      Begin VB.Label lbl_periodo 
         AutoSize        =   -1  'True
         Caption         =   "Período:"
         Height          =   195
         Left            =   9135
         TabIndex        =   44
         Top             =   345
         Width           =   615
      End
      Begin VB.Label lbl_filme 
         AutoSize        =   -1  'True
         Caption         =   "Filme:"
         Height          =   195
         Left            =   240
         TabIndex        =   43
         Top             =   345
         Width           =   405
      End
   End
   Begin VB.Frame fraGrid 
      Caption         =   "Preços Cadastrados"
      Height          =   2910
      Left            =   75
      TabIndex        =   25
      Top             =   0
      Width           =   10470
      Begin VSFlex7LCtl.VSFlexGrid VSFlexGrid 
         Height          =   2085
         Left            =   75
         TabIndex        =   1
         Top             =   750
         Width           =   10275
         _cx             =   18124
         _cy             =   3678
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
         FormatString    =   $"frmPreco1.frx":0000
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
         WordWrap        =   -1  'True
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
      Begin Combo.cboCodDesc ccd_ppr_cd 
         Height          =   315
         Left            =   1245
         TabIndex        =   0
         Top             =   300
         Width           =   8955
         _ExtentX        =   15796
         _ExtentY        =   556
         NomeTabela      =   "tb_prog_preco"
         NomeCampoCodigo =   "ppr_cd"
         NomeCampoDescricao=   "convert(char(10),ppr_dt_ini,111) + ' - ' + convert(char(10),ppr_dt_fim,111)"
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
         Filtro          =   "ppr_dt_des is null"
         MostraBotaoAtualiza=   0   'False
      End
      Begin VB.Label lbl_prog 
         AutoSize        =   -1  'True
         Caption         =   "Programação:"
         Height          =   195
         Left            =   165
         TabIndex        =   26
         Top             =   360
         Width           =   990
      End
   End
   Begin Comandos.cmdComandos cmdComandos 
      Align           =   2  'Align Bottom
      Height          =   765
      Left            =   0
      TabIndex        =   24
      Top             =   6405
      Width           =   10635
      _ExtentX        =   18759
      _ExtentY        =   1349
      EnabledNovo     =   0   'False
   End
End
Attribute VB_Name = "frmPreco1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private sOperacao As String

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

Const m_def_QtdeDecimais = 2
Const m_def_QtdeInteiros = 9

Private Sub ccd_fil_cd_BeforeProcuraClick(Cancel As Boolean)

    Dim sFiltro As String
    
    sFiltro = " fil_dt_des is null " & _
              "  and ( (   fil_dt_ini <= ( select ppr_dt_ini from tb_prog_preco where ppr_cd = " & ccd_ppr_cd.codigo & " ) " & _
              "        and fil_dt_fim > ( select ppr_dt_ini from tb_prog_preco where ppr_cd = " & ccd_ppr_cd.codigo & " ) " & _
              "      ) " & _
              "   or (     fil_dt_ini >= ( select ppr_dt_ini from tb_prog_preco where ppr_cd = " & ccd_ppr_cd.codigo & " ) " & _
              "        and fil_dt_ini <  ( select ppr_dt_fim from tb_prog_preco where ppr_cd = " & ccd_ppr_cd.codigo & " ) " & _
              "      ) )"

    ccd_fil_cd.Filtro = sFiltro

End Sub

Private Sub ccd_ppr_cd_AfterProcuraClick()
    
    If ccd_ppr_cd.codigo <> "" Then
        Call PreencheGrid
    End If
    
    cmdComandos.EnabledAltera = (ccd_ppr_cd.codigo <> "")
    cmdComandos.EnabledExclui = (ccd_ppr_cd.codigo <> "")
    cmdComandos.EnabledNovo = (ccd_ppr_cd.codigo <> "")
    
End Sub

Private Sub ccd_ppr_cd_Change()
    Call ccd_ppr_cd_AfterProcuraClick
End Sub

Private Sub flt_pre_vl_inteira_1_Change()
    Call f_Change(flt_pre_vl_inteira_1)
End Sub

Private Sub flt_pre_vl_inteira_1_GotFocus()
    Call f_GotFocus(flt_pre_vl_inteira_1)
End Sub

Private Sub flt_pre_vl_inteira_1_KeyDown(KeyCode As Integer, Shift As Integer)
    Call f_KeyDown(flt_pre_vl_inteira_1, KeyCode, Shift)
End Sub

Private Sub flt_pre_vl_inteira_1_KeyPress(KeyAscii As Integer)
    Call f_KeyPress(flt_pre_vl_inteira_1, KeyAscii)
End Sub

Private Sub flt_pre_vl_inteira_1_LostFocus()
    Call f_LostFocus(flt_pre_vl_inteira_1)
End Sub

Private Sub flt_pre_vl_inteira_2_Change()
    Call f_Change(flt_pre_vl_inteira_2)
End Sub

Private Sub flt_pre_vl_inteira_2_GotFocus()
    Call f_GotFocus(flt_pre_vl_inteira_2)
End Sub

Private Sub flt_pre_vl_inteira_2_KeyDown(KeyCode As Integer, Shift As Integer)
    Call f_KeyDown(flt_pre_vl_inteira_2, KeyCode, Shift)
End Sub

Private Sub flt_pre_vl_inteira_2_KeyPress(KeyAscii As Integer)
    Call f_KeyPress(flt_pre_vl_inteira_2, KeyAscii)
End Sub

Private Sub flt_pre_vl_inteira_2_LostFocus()
    Call f_LostFocus(flt_pre_vl_inteira_2)
End Sub

Private Sub flt_pre_vl_inteira_3_Change()
    Call f_Change(flt_pre_vl_inteira_3)
End Sub

Private Sub flt_pre_vl_inteira_3_GotFocus()
    Call f_GotFocus(flt_pre_vl_inteira_3)
End Sub

Private Sub flt_pre_vl_inteira_3_KeyDown(KeyCode As Integer, Shift As Integer)
    Call f_KeyDown(flt_pre_vl_inteira_3, KeyCode, Shift)
End Sub

Private Sub flt_pre_vl_inteira_3_KeyPress(KeyAscii As Integer)
    Call f_KeyPress(flt_pre_vl_inteira_3, KeyAscii)
End Sub

Private Sub flt_pre_vl_inteira_3_LostFocus()
    Call f_LostFocus(flt_pre_vl_inteira_3)
End Sub

Private Sub flt_pre_vl_inteira_4_Change()
    Call f_Change(flt_pre_vl_inteira_4)
End Sub

Private Sub flt_pre_vl_inteira_4_GotFocus()
    Call f_GotFocus(flt_pre_vl_inteira_4)
End Sub

Private Sub flt_pre_vl_inteira_4_KeyDown(KeyCode As Integer, Shift As Integer)
    Call f_KeyDown(flt_pre_vl_inteira_4, KeyCode, Shift)
End Sub

Private Sub flt_pre_vl_inteira_4_KeyPress(KeyAscii As Integer)
    Call f_KeyPress(flt_pre_vl_inteira_4, KeyAscii)
End Sub

Private Sub flt_pre_vl_inteira_4_LostFocus()
    Call f_LostFocus(flt_pre_vl_inteira_4)
End Sub

Private Sub flt_pre_vl_inteira_5_Change()
    Call f_Change(flt_pre_vl_inteira_5)
End Sub

Private Sub flt_pre_vl_inteira_5_GotFocus()
    Call f_GotFocus(flt_pre_vl_inteira_5)
End Sub

Private Sub flt_pre_vl_inteira_5_KeyDown(KeyCode As Integer, Shift As Integer)
    Call f_KeyDown(flt_pre_vl_inteira_5, KeyCode, Shift)
End Sub

Private Sub flt_pre_vl_inteira_5_KeyPress(KeyAscii As Integer)
    Call f_KeyPress(flt_pre_vl_inteira_5, KeyAscii)
End Sub

Private Sub flt_pre_vl_inteira_5_LostFocus()
    Call f_LostFocus(flt_pre_vl_inteira_5)
End Sub

Private Sub flt_pre_vl_inteira_6_Change()
    Call f_Change(flt_pre_vl_inteira_6)
End Sub

Private Sub flt_pre_vl_inteira_6_GotFocus()
    Call f_GotFocus(flt_pre_vl_inteira_6)
End Sub

Private Sub flt_pre_vl_inteira_6_KeyDown(KeyCode As Integer, Shift As Integer)
    Call f_KeyDown(flt_pre_vl_inteira_6, KeyCode, Shift)
End Sub

Private Sub flt_pre_vl_inteira_6_KeyPress(KeyAscii As Integer)
    Call f_KeyPress(flt_pre_vl_inteira_6, KeyAscii)
End Sub

Private Sub flt_pre_vl_inteira_6_LostFocus()
    Call f_LostFocus(flt_pre_vl_inteira_6)
End Sub

Private Sub flt_pre_vl_meia_1_Change()
    Call f_Change(flt_pre_vl_meia_1)
End Sub

Private Sub flt_pre_vl_meia_1_GotFocus()
    Call f_GotFocus(flt_pre_vl_meia_1)
End Sub

Private Sub flt_pre_vl_meia_1_KeyDown(KeyCode As Integer, Shift As Integer)
    Call f_KeyDown(flt_pre_vl_meia_1, KeyCode, Shift)
End Sub

Private Sub flt_pre_vl_meia_1_KeyPress(KeyAscii As Integer)
    Call f_KeyPress(flt_pre_vl_meia_1, KeyAscii)
End Sub

Private Sub flt_pre_vl_meia_1_LostFocus()
    Call f_LostFocus(flt_pre_vl_meia_1)
End Sub

Private Sub flt_pre_vl_meia_2_Change()
    Call f_Change(flt_pre_vl_meia_2)
End Sub

Private Sub flt_pre_vl_meia_2_GotFocus()
    Call f_GotFocus(flt_pre_vl_meia_2)
End Sub

Private Sub flt_pre_vl_meia_2_KeyDown(KeyCode As Integer, Shift As Integer)
    Call f_KeyDown(flt_pre_vl_meia_2, KeyCode, Shift)
End Sub

Private Sub flt_pre_vl_meia_2_KeyPress(KeyAscii As Integer)
    Call f_KeyPress(flt_pre_vl_meia_2, KeyAscii)
End Sub

Private Sub flt_pre_vl_meia_2_LostFocus()
    Call f_LostFocus(flt_pre_vl_meia_2)
End Sub

Private Sub flt_pre_vl_meia_3_Change()
    Call f_Change(flt_pre_vl_meia_3)
End Sub

Private Sub flt_pre_vl_meia_3_GotFocus()
    Call f_GotFocus(flt_pre_vl_meia_3)
End Sub

Private Sub flt_pre_vl_meia_3_KeyDown(KeyCode As Integer, Shift As Integer)
    Call f_KeyDown(flt_pre_vl_meia_3, KeyCode, Shift)
End Sub

Private Sub flt_pre_vl_meia_3_KeyPress(KeyAscii As Integer)
    Call f_KeyPress(flt_pre_vl_meia_3, KeyAscii)
End Sub

Private Sub flt_pre_vl_meia_3_LostFocus()
    Call f_LostFocus(flt_pre_vl_meia_3)
End Sub

Private Sub flt_pre_vl_meia_4_Change()
    Call f_Change(flt_pre_vl_meia_4)
End Sub

Private Sub flt_pre_vl_meia_4_GotFocus()
    Call f_GotFocus(flt_pre_vl_meia_4)
End Sub

Private Sub flt_pre_vl_meia_4_KeyDown(KeyCode As Integer, Shift As Integer)
    Call f_KeyDown(flt_pre_vl_meia_4, KeyCode, Shift)
End Sub

Private Sub flt_pre_vl_meia_4_KeyPress(KeyAscii As Integer)
    Call f_KeyPress(flt_pre_vl_meia_4, KeyAscii)
End Sub

Private Sub flt_pre_vl_meia_4_LostFocus()
    Call f_LostFocus(flt_pre_vl_meia_4)
End Sub

Private Sub flt_pre_vl_meia_5_Change()
    Call f_Change(flt_pre_vl_meia_5)
End Sub

Private Sub flt_pre_vl_meia_5_GotFocus()
    Call f_GotFocus(flt_pre_vl_meia_5)
End Sub

Private Sub flt_pre_vl_meia_5_KeyDown(KeyCode As Integer, Shift As Integer)
    Call f_KeyDown(flt_pre_vl_meia_5, KeyCode, Shift)
End Sub

Private Sub flt_pre_vl_meia_5_KeyPress(KeyAscii As Integer)
    Call f_KeyPress(flt_pre_vl_meia_5, KeyAscii)
End Sub

Private Sub flt_pre_vl_meia_5_LostFocus()
    Call f_LostFocus(flt_pre_vl_meia_5)
End Sub

Private Sub flt_pre_vl_meia_6_Change()
    Call f_Change(flt_pre_vl_meia_6)
End Sub

Private Sub flt_pre_vl_meia_6_GotFocus()
    Call f_GotFocus(flt_pre_vl_meia_6)
End Sub

Private Sub flt_pre_vl_meia_6_KeyDown(KeyCode As Integer, Shift As Integer)
    Call f_KeyDown(flt_pre_vl_meia_6, KeyCode, Shift)
End Sub

Private Sub flt_pre_vl_meia_6_KeyPress(KeyAscii As Integer)
    Call f_KeyPress(flt_pre_vl_meia_6, KeyAscii)
End Sub

Private Sub flt_pre_vl_meia_6_LostFocus()
    Call f_LostFocus(flt_pre_vl_meia_6)
End Sub

Private Sub VSFlexGrid_RowColChange()
    Call CarregaControles
End Sub

Private Sub Form_Load()

    Set ccd_ppr_cd.ConexaoADO = dbConnect
    Set ccd_fil_cd.ConexaoADO = dbConnect
    
    lblHoraLimite1.Caption = Format(dtHoraLimite12, "Hh:Nn")
    lblHoraLimite2.Caption = Format(dtHoraLimite23, "Hh:Nn")
    lblHoraLimite3.Caption = Format(dtHoraLimite34, "Hh:Nn")
    lblHoraLimite4.Caption = Format(dtHoraLimite45, "Hh:Nn")
    lblHoraLimite5.Caption = Format(dtHoraLimite56, "Hh:Nn")
    
    Call HabilitaManut(False)

    Dim CineAx As New CineAux.CineAx
    Call CineAx.CentralizaTela(MDIFrmAdmCine, Me)
    
    Set CineAx = Nothing
    
    VSFlexGrid.ColHidden(VSFlexGrid.ColIndex("KEY_PPR_CD")) = True
    VSFlexGrid.ColHidden(VSFlexGrid.ColIndex("KEY_FIL_CD")) = True
    VSFlexGrid.ColHidden(VSFlexGrid.ColIndex("")) = True
    VSFlexGrid.RowHeight(0) = 500

    Set ccd_ppr_cd.NomeForm = frmProgPreco
    ccd_ppr_cd.NomeCampoDescricao = "convert(char(10),convert(Datetime,ppr_dt_ini,103),111) + ' - ' + convert(char(10),convert(datetime,ppr_dt_fim,103),111) + case when ppr_flg_promocao = 1 then ' - ' + 'PROMOÇÃO' + ' - ' + ppr_patrocinador else '' end"
    
    Dim log As New clsLog
    
    Set log.ConexaoADO = dbConnect
    log.usu_nm = strUsuario
    log.slg_descricao = "Entrou na tela de " & Me.Caption & " do módulo " & App.ProductName
    
    log.insereLog
    
    m_QtdeDecimais = m_def_QtdeDecimais
    m_QtdeInteiros = m_def_QtdeInteiros
    msSeparadorDecimal = ","
    msMilhar = "."
    msLostFormat = fusLostFormat()
    
End Sub

Private Sub cmdComandos_Click(ByVal iButtonClicked As Comandos.eButtonClicked, Cancel As Boolean)
    
    Select Case iButtonClicked
    
        Case ButtonAltera
            If permiteAltExc() Then
                sOperacao = "A"
            
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
                MsgBox "Não é possível alterar preço. Período anterior a data atual", vbCritical, App.ProductName
            End If
            
        Case ButtonNovo
            sOperacao = "I"
            
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
                MsgBox "Não é possível excluir preço. Período anterior a data atual", vbCritical, App.ProductName
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
    
        ccd_fil_cd.codigo = VSFlexGrid.TextMatrix(VSFlexGrid.RowSel, VSFlexGrid.ColIndex("KEY_FIL_CD"))
        If ccd_fil_cd.codigo <> "" Then ccd_fil_cd.Refresh
        
        spnPeriodo.Value = VSFlexGrid.TextMatrix(VSFlexGrid.RowSel, VSFlexGrid.ColIndex("KEY_PRE_PERIODO"))
        flt_pre_vl_inteira_1.Text = VSFlexGrid.TextMatrix(VSFlexGrid.RowSel, VSFlexGrid.ColIndex("KEY_PRE_VL_INTEIRA1"))
        flt_pre_vl_inteira_2.Text = VSFlexGrid.TextMatrix(VSFlexGrid.RowSel, VSFlexGrid.ColIndex("KEY_PRE_VL_INTEIRA2"))
        flt_pre_vl_inteira_3.Text = VSFlexGrid.TextMatrix(VSFlexGrid.RowSel, VSFlexGrid.ColIndex("KEY_PRE_VL_INTEIRA3"))
        flt_pre_vl_inteira_4.Text = VSFlexGrid.TextMatrix(VSFlexGrid.RowSel, VSFlexGrid.ColIndex("KEY_PRE_VL_INTEIRA4"))
        flt_pre_vl_inteira_5.Text = VSFlexGrid.TextMatrix(VSFlexGrid.RowSel, VSFlexGrid.ColIndex("KEY_PRE_VL_INTEIRA5"))
        flt_pre_vl_inteira_6.Text = VSFlexGrid.TextMatrix(VSFlexGrid.RowSel, VSFlexGrid.ColIndex("KEY_PRE_VL_INTEIRA6"))
        
        flt_pre_vl_meia_1.Text = VSFlexGrid.TextMatrix(VSFlexGrid.RowSel, VSFlexGrid.ColIndex("KEY_PRE_VL_MEIA1"))
        flt_pre_vl_meia_2.Text = VSFlexGrid.TextMatrix(VSFlexGrid.RowSel, VSFlexGrid.ColIndex("KEY_PRE_VL_MEIA2"))
        flt_pre_vl_meia_3.Text = VSFlexGrid.TextMatrix(VSFlexGrid.RowSel, VSFlexGrid.ColIndex("KEY_PRE_VL_MEIA3"))
        flt_pre_vl_meia_4.Text = VSFlexGrid.TextMatrix(VSFlexGrid.RowSel, VSFlexGrid.ColIndex("KEY_PRE_VL_MEIA4"))
        flt_pre_vl_meia_5.Text = VSFlexGrid.TextMatrix(VSFlexGrid.RowSel, VSFlexGrid.ColIndex("KEY_PRE_VL_MEIA5"))
        flt_pre_vl_meia_6.Text = VSFlexGrid.TextMatrix(VSFlexGrid.RowSel, VSFlexGrid.ColIndex("KEY_PRE_VL_MEIA6"))
        optPromocao.Value = VSFlexGrid.TextMatrix(VSFlexGrid.RowSel, 24)
        
        Dim i As Integer
    
        For i = 1 To 8
            If VSFlexGrid.TextMatrix(VSFlexGrid.RowSel, VSFlexGrid.ColIndex("KEY_DIA" & i)) = "" Then
                chk_pre_dia_semana(i).Value = vbUnchecked
            Else
                chk_pre_dia_semana(i).Value = vbChecked
            End If
        Next
        
    End If
    
End Sub

Private Sub LimpaControles()

    Dim i As Integer
    
    ccd_fil_cd.codigo = ""
    spnPeriodo.Value = 0
    
    flt_pre_vl_inteira_1.Text = 0
    flt_pre_vl_inteira_2.Text = 0
    flt_pre_vl_inteira_3.Text = 0
    flt_pre_vl_inteira_4.Text = 0
    flt_pre_vl_inteira_5.Text = 0
    flt_pre_vl_inteira_6.Text = 0
    
    flt_pre_vl_meia_1.Text = 0
    flt_pre_vl_meia_2.Text = 0
    flt_pre_vl_meia_3.Text = 0
    flt_pre_vl_meia_4.Text = 0
    flt_pre_vl_meia_5.Text = 0
    flt_pre_vl_meia_6.Text = 0

    For i = 1 To 8
        chk_pre_dia_semana(i).Value = vbUnchecked
    Next
    optPromocao.Value = False
End Sub

Private Function Grava() As Boolean

    On Error GoTo Grava_Erro
    
    If Not Consiste() Then
        Exit Function
    End If
    
    Dim clsTB_PRECO As New Cine2005.clsTB_PRECO
    
    Set clsTB_PRECO.ConexaoADO = dbConnect
    
    clsTB_PRECO.ppr_cd = ccd_ppr_cd.codigo
    clsTB_PRECO.fil_cd = Val(ccd_fil_cd.codigo)
    clsTB_PRECO.pre_periodo = spnPeriodo.Value
    
    If sOperacao = "A" Then
        If Not clsTB_PRECO.Excluir() Then
            MsgBox clsTB_PRECO.MensagemErro, vbCritical, App.ProductName
            GoTo Grava_Fim
        End If
    End If
    
    Dim iDias As Integer
    
    For iDias = 1 To 8
    
        If chk_pre_dia_semana(iDias).Value = vbChecked Then
        
            clsTB_PRECO.pre_dia_semana = iDias
            
            clsTB_PRECO.pre_vl_inteira_ate = flt_pre_vl_inteira_1.Text
            clsTB_PRECO.pre_vl_inteira_apos = flt_pre_vl_inteira_2.Text
            clsTB_PRECO.pre_vl_inteira3 = flt_pre_vl_inteira_3.Text
            clsTB_PRECO.pre_vl_inteira4 = flt_pre_vl_inteira_4.Text
            clsTB_PRECO.pre_vl_inteira5 = flt_pre_vl_inteira_5.Text
            clsTB_PRECO.pre_vl_inteira6 = flt_pre_vl_inteira_6.Text
            
            clsTB_PRECO.pre_vl_meia_ate = flt_pre_vl_meia_1.Text
            clsTB_PRECO.pre_vl_meia_apos = flt_pre_vl_meia_2.Text
            clsTB_PRECO.pre_vl_meia3 = flt_pre_vl_meia_3.Text
            clsTB_PRECO.pre_vl_meia4 = flt_pre_vl_meia_4.Text
            clsTB_PRECO.pre_vl_meia5 = flt_pre_vl_meia_5.Text
            clsTB_PRECO.pre_vl_meia6 = flt_pre_vl_meia_6.Text
            clsTB_PRECO.pre_promocao = optPromocao.Value
            
            If Not clsTB_PRECO.Incluir() Then
                MsgBox "Não foi possível incluir a Pogramação!" & vbCrLf & clsTB_PRECO.MensagemErro, vbInformation, App.ProductName
                GoTo Grava_Fim
            End If
            
        End If
    Next

    Call PreencheGrid

    Grava = True
    GoTo Grava_Fim

Grava_Erro:
    MsgBox "Erro de execução! 'Grava/frmPreco'", vbCritical, App.ProductName
    
Grava_Fim:
    Set clsTB_PRECO = Nothing
    
End Function

Private Function Exclui() As Boolean

    On Error GoTo Exclui_Erro
    
    Dim clsTB_PRECO As New Cine2005.clsTB_PRECO

    Set clsTB_PRECO.ConexaoADO = dbConnect
    
    clsTB_PRECO.ppr_cd = ccd_ppr_cd.codigo
    clsTB_PRECO.fil_cd = ccd_fil_cd.codigo
    clsTB_PRECO.pre_periodo = spnPeriodo.Value
    
    If Not clsTB_PRECO.Excluir() Then
        MsgBox "Não foi possível excluir a Programação Selecionada!" & vbCrLf & clsTB_PRECO.MensagemErro, vbInformation, App.ProductName
        GoTo Exclui_Fim
    End If
            
    Exclui = True
    GoTo Exclui_Fim
    
Exclui_Erro:
    MsgBox "Erro de execução! 'Exclui/frmPreco'", vbCritical, App.ProductName
    
Exclui_Fim:
    Set clsTB_PRECO = Nothing
    
End Function

Private Sub PreencheGrid()

    On Error GoTo PreencheGrid_Erro
    
    Call LimpaControles
    
    Dim oRs As New ADODB.Recordset
    Dim clsTB_PRECO As New Cine2005.clsTB_PRECO
    Dim iProgramacao As Integer, iSala As Integer, iFilme As Long, iPeriodo As Integer
    
    Set clsTB_PRECO.ConexaoADO = dbConnect
    
    clsTB_PRECO.ppr_cd = ccd_ppr_cd.codigo
    
    If Not clsTB_PRECO.PreencheGrid(oRs) Then
        MsgBox clsTB_PRECO.MensagemErro, vbCritical, App.ProductName
        GoTo PreencheGrid_Fim
    End If

    VSFlexGrid.Rows = 1

    Do While Not oRs.EOF()
    
        If iProgramacao <> oRs.Fields("ppr_cd") Or _
            iFilme <> oRs.Fields("fil_cd") Or _
            iPeriodo <> oRs.Fields("pre_periodo") Then
        
            iProgramacao = oRs.Fields("ppr_cd")
            iFilme = oRs.Fields("fil_cd")
            iPeriodo = oRs.Fields("pre_periodo")
                
            On Error GoTo 0
        
            VSFlexGrid.Rows = VSFlexGrid.Rows + 1
            
            VSFlexGrid.TextMatrix(VSFlexGrid.Rows - 1, VSFlexGrid.ColIndex("KEY_PPR_CD")) = oRs.Fields("ppr_cd")
            VSFlexGrid.TextMatrix(VSFlexGrid.Rows - 1, VSFlexGrid.ColIndex("KEY_FIL_CD")) = oRs.Fields("fil_cd")
            VSFlexGrid.TextMatrix(VSFlexGrid.Rows - 1, VSFlexGrid.ColIndex("KEY_PRE_PERIODO")) = oRs.Fields("pre_periodo")
            VSFlexGrid.TextMatrix(VSFlexGrid.Rows - 1, VSFlexGrid.ColIndex("KEY_FILME")) = IIf(IsNull(oRs.Fields("fil_nm")), "** VALORES PADRÃO **", oRs.Fields("fil_nm"))
            
            VSFlexGrid.TextMatrix(VSFlexGrid.Rows - 1, VSFlexGrid.ColIndex("KEY_PRE_VL_INTEIRA1")) = oRs.Fields("pre_vl_inteira_ate")
            VSFlexGrid.TextMatrix(VSFlexGrid.Rows - 1, VSFlexGrid.ColIndex("KEY_PRE_VL_INTEIRA2")) = oRs.Fields("pre_vl_inteira_apos")
            VSFlexGrid.TextMatrix(VSFlexGrid.Rows - 1, VSFlexGrid.ColIndex("KEY_PRE_VL_INTEIRA3")) = oRs.Fields("pre_vl_inteira3")
            VSFlexGrid.TextMatrix(VSFlexGrid.Rows - 1, VSFlexGrid.ColIndex("KEY_PRE_VL_INTEIRA4")) = oRs.Fields("pre_vl_inteira4")
            VSFlexGrid.TextMatrix(VSFlexGrid.Rows - 1, VSFlexGrid.ColIndex("KEY_PRE_VL_INTEIRA5")) = oRs.Fields("pre_vl_inteira5")
            VSFlexGrid.TextMatrix(VSFlexGrid.Rows - 1, VSFlexGrid.ColIndex("KEY_PRE_VL_INTEIRA6")) = oRs.Fields("pre_vl_inteira6")
            
            VSFlexGrid.TextMatrix(VSFlexGrid.Rows - 1, VSFlexGrid.ColIndex("KEY_PRE_VL_MEIA1")) = oRs.Fields("pre_vl_meia_ate")
            VSFlexGrid.TextMatrix(VSFlexGrid.Rows - 1, VSFlexGrid.ColIndex("KEY_PRE_VL_MEIA2")) = oRs.Fields("pre_vl_meia_apos")
            VSFlexGrid.TextMatrix(VSFlexGrid.Rows - 1, VSFlexGrid.ColIndex("KEY_PRE_VL_MEIA3")) = oRs.Fields("pre_vl_meia3")
            VSFlexGrid.TextMatrix(VSFlexGrid.Rows - 1, VSFlexGrid.ColIndex("KEY_PRE_VL_MEIA4")) = oRs.Fields("pre_vl_meia4")
            VSFlexGrid.TextMatrix(VSFlexGrid.Rows - 1, VSFlexGrid.ColIndex("KEY_PRE_VL_MEIA5")) = oRs.Fields("pre_vl_meia5")
            VSFlexGrid.TextMatrix(VSFlexGrid.Rows - 1, VSFlexGrid.ColIndex("KEY_PRE_VL_MEIA6")) = oRs.Fields("pre_vl_meia6")
            'VSFlexGrid.TextMatrix(VSFlexGrid.Rows - 1, VSFlexGrid.ColIndex("KEY_PRE_PROMOCAO")) = oRs.Fields("pre_promocao")
            VSFlexGrid.TextMatrix(VSFlexGrid.Rows - 1, 24) = IIf(oRs.Fields("pre_promocao"), 1, 0)
        
        End If
        
        VSFlexGrid.TextMatrix(VSFlexGrid.Rows - 1, VSFlexGrid.ColIndex("KEY_DIA" & oRs.Fields("pre_dia_semana"))) = True
        
        oRs.MoveNext
        
    Loop
    
    VSFlexGrid.MergeCells = flexMergeRestrictColumns
    VSFlexGrid.MergeCol(VSFlexGrid.ColIndex("KEY_FILME")) = True

    VSFlexGrid.AutoSizeMode = flexAutoSizeColWidth
    
    Call VSFlexGrid.AutoSize(VSFlexGrid.ColIndex("KEY_FILME"))
    
    VSFlexGrid.RowHeight(0) = 500
    
    VSFlexGrid.ColWidth(VSFlexGrid.ColIndex("KEY_DIA8")) = 1000

    VSFlexGrid.FrozenCols = 4

    VSFlexGrid.Row = IIf(VSFlexGrid.Rows > 1, 1, 0)
    Call CarregaControles

    GoTo PreencheGrid_Fim
    
PreencheGrid_Erro:
    MsgBox "Erro de execução! 'PreencheGrid/frmPreco'", vbCritical, App.ProductName
    
PreencheGrid_Fim:
    If oRs.State = 1 Then oRs.Close
    Set oRs = Nothing
    Set clsTB_PRECO = Nothing
    
End Sub

Private Function Consiste() As Boolean
    
    Dim sMens As String
    Dim i As Integer
    Dim bMarcou As Boolean
    
    Dim oRs As New ADODB.Recordset
    Dim clsTB_PRECO As New Cine2005.clsTB_PRECO
    
    Set clsTB_PRECO.ConexaoADO = dbConnect
    
    clsTB_PRECO.ppr_cd = ccd_ppr_cd.codigo
    clsTB_PRECO.fil_cd = Val(ccd_fil_cd.codigo)
    
    If sOperacao = "I" Then
    
        clsTB_PRECO.pre_periodo = spnPeriodo.Value
        
        ' Verifica se já não existe preço cadastrado para o filme
        
        If Not clsTB_PRECO.Selecionar(oRs) Then
            MsgBox clsTB_PRECO.MensagemErro, vbCritical, App.ProductName
            Exit Function
        End If
        
        If Not oRs.EOF() Then
            MsgBox "Este filme já tem preços cadastrados para este período!", vbCritical, App.ProductName
            GoTo Consiste_fim
        End If
        
        If oRs.State = 1 Then oRs.Close
    
    End If
    
    ' Verifica se não está duplicando para dias iguais
    
    clsTB_PRECO.pre_periodo = Empty
    
    If Not clsTB_PRECO.Selecionar(oRs) Then
        MsgBox clsTB_PRECO.MensagemErro, vbCritical, App.ProductName
        Exit Function
    End If
    
    Do While Not oRs.EOF()
        For i = 1 To 8
            If chk_pre_dia_semana(i).Value = vbChecked Then
                If oRs.Fields("pre_periodo") <> spnPeriodo.Value Then
                    If oRs.Fields("pre_dia_semana") = i Then
                        MsgBox "O(s) dia(s) da semana selecionado(s) já pertence(m) a outro período!", vbCritical, App.ProductName
                        GoTo Consiste_fim
                    End If
                End If
            End If
        Next
        oRs.MoveNext
    Loop
    
    If oRs.State = 1 Then oRs.Close
    
'    If ccd_fil_cd.codigo = "" Then
'        sMens = sMens & "Filme deve ser informado!" & vbCrLf
'    End If
    
    bMarcou = False
    
    For i = 1 To 8
        If chk_pre_dia_semana(i).Value = vbChecked Then
            bMarcou = True
            Exit For
        End If
    Next
    
    If Not bMarcou Then
        sMens = sMens & "Deve-se informar pelo menos um dia da semana!" & vbCrLf
    End If

    If sMens <> "" Then
        MsgBox sMens, vbInformation, App.ProductName
        Exit Function
    End If
    
    Consiste = True
    
Consiste_fim:
    If oRs.State = 1 Then oRs.Close
    Set clsTB_PRECO = Nothing
    
End Function

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    Dim log As New clsLog
    
    Set log.ConexaoADO = dbConnect
    log.usu_nm = strUsuario
    log.slg_descricao = "Saiu da tela de " & Me.Caption & " do módulo " & App.ProductName
    
    log.insereLog
End Sub

Private Sub Frame2_DragDrop(Source As Control, X As Single, Y As Single)

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

Private Function permiteAltExc() As Boolean
    Dim dtIni  As Date
    Dim dtFim  As Date
    Dim dtAtu  As Date
    Dim dia    As Integer
    Dim mes    As Integer
    Dim ano    As Integer
    Dim strAux As String
    
    permiteAltExc = False
    
    strAux = ccd_ppr_cd.Descricao
    
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


Private Sub VSFlexGrid1_Click()

End Sub
