VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Begin VB.Form frmParametrosVideoHall 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Parametros VideoHall"
   ClientHeight    =   6300
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   9435
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6300
   ScaleWidth      =   9435
   Begin VB.Frame fraParametros 
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      Height          =   2100
      Left            =   45
      TabIndex        =   84
      Top             =   0
      Width           =   9300
      Begin VB.TextBox txtVelocMsg 
         Height          =   315
         Left            =   3345
         MaxLength       =   3
         TabIndex        =   96
         Top             =   666
         Width           =   660
      End
      Begin VB.TextBox txtMensagen 
         Height          =   315
         Left            =   1110
         TabIndex        =   95
         Top             =   1665
         Width           =   8175
      End
      Begin VB.Frame fraTelas 
         Caption         =   "Telas apresetadas"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1560
         Left            =   5955
         TabIndex        =   89
         Top             =   15
         Width           =   2295
         Begin VB.CheckBox chkSessoes 
            Caption         =   "Sessões"
            Height          =   240
            Left            =   180
            TabIndex        =   94
            Top             =   285
            Width           =   930
         End
         Begin VB.CheckBox chkFilme 
            Caption         =   "Filme"
            Height          =   240
            Left            =   180
            TabIndex        =   93
            Top             =   600
            Width           =   930
         End
         Begin VB.CheckBox chkPrecos 
            Caption         =   "Preços"
            Height          =   240
            Left            =   180
            TabIndex        =   92
            Top             =   900
            Width           =   930
         End
         Begin VB.CheckBox chkImagem 
            Caption         =   "Imagens"
            Height          =   240
            Left            =   180
            TabIndex        =   91
            Top             =   1200
            Width           =   930
         End
         Begin VB.CheckBox chkTrailer 
            Caption         =   "Trailler"
            Height          =   240
            Left            =   1320
            TabIndex        =   90
            Top             =   300
            Width           =   930
         End
      End
      Begin VB.TextBox txtVendaDepois 
         Height          =   315
         Left            =   4575
         MaxLength       =   2
         TabIndex        =   88
         Top             =   984
         Width           =   660
      End
      Begin VB.TextBox txtVendaAntes 
         Height          =   315
         Left            =   3345
         MaxLength       =   2
         TabIndex        =   87
         Top             =   984
         Width           =   660
      End
      Begin VB.TextBox txtIntermitencia 
         Height          =   315
         Left            =   3345
         MaxLength       =   5
         TabIndex        =   86
         Top             =   348
         Width           =   660
      End
      Begin VB.TextBox txtTransicao 
         Height          =   315
         Left            =   3345
         MaxLength       =   9
         TabIndex        =   85
         Top             =   30
         Width           =   660
      End
      Begin MSMask.MaskEdBox mkeHrLimPeriodo 
         Height          =   315
         Left            =   3345
         TabIndex        =   97
         Top             =   1305
         Visible         =   0   'False
         Width           =   660
         _ExtentX        =   1164
         _ExtentY        =   556
         _Version        =   393216
         MaxLength       =   5
         Format          =   "hh:mm"
         Mask            =   "##:##"
         PromptChar      =   "_"
      End
      Begin MSComDlg.CommonDialog CommonDialog2 
         Left            =   5190
         Top             =   180
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
      Begin VB.Label lblVelocMsg 
         AutoSize        =   -1  'True
         Caption         =   "Velocidade Mensagem:"
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
         Left            =   210
         TabIndex        =   107
         Top             =   720
         Width           =   1980
      End
      Begin VB.Label lblMensagem 
         AutoSize        =   -1  'True
         Caption         =   "Mensagem:"
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
         Left            =   60
         TabIndex        =   106
         Top             =   1665
         Width           =   975
      End
      Begin VB.Label lblHrLimPeriodo 
         AutoSize        =   -1  'True
         Caption         =   "Horário limite entre períodos:"
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
         Left            =   210
         TabIndex        =   105
         Top             =   1365
         Visible         =   0   'False
         Width           =   2490
      End
      Begin VB.Label lblDepois 
         AutoSize        =   -1  'True
         Caption         =   "Depois"
         Height          =   195
         Left            =   5295
         TabIndex        =   104
         Top             =   1050
         Width           =   495
      End
      Begin VB.Label lblAntes 
         AutoSize        =   -1  'True
         Caption         =   "Antes"
         Height          =   195
         Left            =   4050
         TabIndex        =   103
         Top             =   1050
         Width           =   405
      End
      Begin VB.Label lblVenda 
         AutoSize        =   -1  'True
         Caption         =   "Sessão ""À Venda"" (Minutos)"
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
         Left            =   210
         TabIndex        =   102
         Top             =   1050
         Width           =   2430
      End
      Begin VB.Label lblMileSegundos 
         AutoSize        =   -1  'True
         Caption         =   "Mile Segundos"
         Height          =   195
         Left            =   4050
         TabIndex        =   101
         Top             =   405
         Width           =   1050
      End
      Begin VB.Label lblIntermitencia 
         AutoSize        =   -1  'True
         Caption         =   "Intermitência de texto:"
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
         Left            =   210
         TabIndex        =   100
         Top             =   405
         Width           =   1920
      End
      Begin VB.Label lblSegundos 
         AutoSize        =   -1  'True
         Caption         =   "Segundos"
         Height          =   195
         Left            =   4035
         TabIndex        =   99
         Top             =   90
         Width           =   720
      End
      Begin VB.Label lblTransicao 
         AutoSize        =   -1  'True
         Caption         =   "Tempo de Transição entre Telas:"
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
         Left            =   210
         TabIndex        =   98
         Top             =   90
         Width           =   2835
      End
   End
   Begin VB.Frame fraCores 
      Caption         =   "Cores"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3705
      Left            =   0
      TabIndex        =   3
      Top             =   2100
      Width           =   9375
      Begin Threed.SSPanel sspT1Filme 
         Height          =   300
         Left            =   45
         TabIndex        =   4
         Top             =   465
         Width           =   3075
         _Version        =   65536
         _ExtentX        =   5424
         _ExtentY        =   529
         _StockProps     =   15
         Caption         =   "Filme"
         ForeColor       =   16777215
         BackColor       =   16711680
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelWidth      =   2
      End
      Begin Threed.SSPanel sspT2Mensagem 
         Height          =   270
         Left            =   3150
         TabIndex        =   5
         Top             =   1680
         Width           =   3060
         _Version        =   65536
         _ExtentX        =   5397
         _ExtentY        =   476
         _StockProps     =   15
         Caption         =   "Mensagem"
         ForeColor       =   16777215
         BackColor       =   16711680
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelWidth      =   2
      End
      Begin Threed.SSPanel sspT3Mensagem 
         Height          =   300
         Left            =   6255
         TabIndex        =   6
         Top             =   2865
         Width           =   3075
         _Version        =   65536
         _ExtentX        =   5424
         _ExtentY        =   529
         _StockProps     =   15
         Caption         =   "Mensagem"
         ForeColor       =   16777215
         BackColor       =   16711680
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelWidth      =   2
      End
      Begin Threed.SSPanel sspT1Hora 
         Height          =   600
         Left            =   45
         TabIndex        =   7
         Top             =   765
         Width           =   810
         _Version        =   65536
         _ExtentX        =   1429
         _ExtentY        =   1058
         _StockProps     =   15
         Caption         =   "99:99"
         ForeColor       =   16777215
         BackColor       =   16711680
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelWidth      =   2
      End
      Begin Threed.SSPanel sspT1Titulo1 
         Height          =   300
         Index           =   0
         Left            =   840
         TabIndex        =   8
         Top             =   765
         Width           =   1140
         _Version        =   65536
         _ExtentX        =   2011
         _ExtentY        =   529
         _StockProps     =   15
         Caption         =   "Até 17:00"
         ForeColor       =   16777215
         BackColor       =   16711680
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelWidth      =   2
      End
      Begin Threed.SSPanel sspT1Titulo1 
         Height          =   300
         Index           =   1
         Left            =   1965
         TabIndex        =   9
         Top             =   765
         Width           =   1155
         _Version        =   65536
         _ExtentX        =   2037
         _ExtentY        =   529
         _StockProps     =   15
         Caption         =   "Após 17:00"
         ForeColor       =   16777215
         BackColor       =   16711680
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelWidth      =   2
      End
      Begin Threed.SSPanel sspT1Titulo2 
         Height          =   300
         Index           =   0
         Left            =   840
         TabIndex        =   10
         Top             =   1065
         Width           =   570
         _Version        =   65536
         _ExtentX        =   1005
         _ExtentY        =   529
         _StockProps     =   15
         Caption         =   "Inteira"
         ForeColor       =   16777215
         BackColor       =   16711680
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelWidth      =   2
      End
      Begin Threed.SSPanel sspT1Titulo2 
         Height          =   300
         Index           =   1
         Left            =   1395
         TabIndex        =   11
         Top             =   1065
         Width           =   570
         _Version        =   65536
         _ExtentX        =   1005
         _ExtentY        =   529
         _StockProps     =   15
         Caption         =   "Meia"
         ForeColor       =   16777215
         BackColor       =   16711680
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelWidth      =   2
      End
      Begin Threed.SSPanel sspT1Titulo2 
         Height          =   300
         Index           =   2
         Left            =   1965
         TabIndex        =   12
         Top             =   1065
         Width           =   570
         _Version        =   65536
         _ExtentX        =   1005
         _ExtentY        =   529
         _StockProps     =   15
         Caption         =   "Inteira"
         ForeColor       =   16777215
         BackColor       =   16711680
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelWidth      =   2
      End
      Begin Threed.SSPanel sspT1Titulo2 
         Height          =   300
         Index           =   3
         Left            =   2535
         TabIndex        =   13
         Top             =   1065
         Width           =   570
         _Version        =   65536
         _ExtentX        =   1005
         _ExtentY        =   529
         _StockProps     =   15
         Caption         =   "Meia"
         ForeColor       =   16777215
         BackColor       =   16711680
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelWidth      =   2
      End
      Begin Threed.SSPanel sspT1Lin1 
         Height          =   645
         Index           =   0
         Left            =   45
         TabIndex        =   14
         Top             =   1365
         Width           =   810
         _Version        =   65536
         _ExtentX        =   1429
         _ExtentY        =   1138
         _StockProps     =   15
         Caption         =   "Seg,Ter, Qua"
         ForeColor       =   16777215
         BackColor       =   16711680
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelWidth      =   2
         Alignment       =   6
      End
      Begin Threed.SSPanel sspT1Lin1 
         Height          =   645
         Index           =   1
         Left            =   840
         TabIndex        =   15
         Top             =   1365
         Width           =   570
         _Version        =   65536
         _ExtentX        =   1005
         _ExtentY        =   1138
         _StockProps     =   15
         Caption         =   "99,99"
         ForeColor       =   16777215
         BackColor       =   16711680
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelWidth      =   2
      End
      Begin Threed.SSPanel sspT1Lin1 
         Height          =   645
         Index           =   2
         Left            =   1395
         TabIndex        =   16
         Top             =   1365
         Width           =   570
         _Version        =   65536
         _ExtentX        =   1005
         _ExtentY        =   1138
         _StockProps     =   15
         Caption         =   "99,99"
         ForeColor       =   16777215
         BackColor       =   16711680
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelWidth      =   2
      End
      Begin Threed.SSPanel sspT1Lin1 
         Height          =   645
         Index           =   3
         Left            =   1965
         TabIndex        =   17
         Top             =   1365
         Width           =   570
         _Version        =   65536
         _ExtentX        =   1005
         _ExtentY        =   1138
         _StockProps     =   15
         Caption         =   "99,99"
         ForeColor       =   16777215
         BackColor       =   16711680
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelWidth      =   2
      End
      Begin Threed.SSPanel sspT1Lin1 
         Height          =   645
         Index           =   4
         Left            =   2535
         TabIndex        =   18
         Top             =   1365
         Width           =   570
         _Version        =   65536
         _ExtentX        =   1005
         _ExtentY        =   1138
         _StockProps     =   15
         Caption         =   "99,99"
         ForeColor       =   16777215
         BackColor       =   16711680
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelWidth      =   2
      End
      Begin Threed.SSPanel sspT1Lin2 
         Height          =   645
         Index           =   0
         Left            =   45
         TabIndex        =   19
         Top             =   2010
         Width           =   810
         _Version        =   65536
         _ExtentX        =   1429
         _ExtentY        =   1138
         _StockProps     =   15
         Caption         =   "Seg,Ter, Qua"
         ForeColor       =   16777215
         BackColor       =   16711680
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelWidth      =   2
         Alignment       =   6
      End
      Begin Threed.SSPanel sspT1Lin2 
         Height          =   645
         Index           =   1
         Left            =   840
         TabIndex        =   20
         Top             =   2010
         Width           =   570
         _Version        =   65536
         _ExtentX        =   1005
         _ExtentY        =   1138
         _StockProps     =   15
         Caption         =   "99,99"
         ForeColor       =   16777215
         BackColor       =   16711680
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelWidth      =   2
      End
      Begin Threed.SSPanel sspT1Lin2 
         Height          =   645
         Index           =   2
         Left            =   1395
         TabIndex        =   21
         Top             =   2010
         Width           =   570
         _Version        =   65536
         _ExtentX        =   1005
         _ExtentY        =   1138
         _StockProps     =   15
         Caption         =   "99,99"
         ForeColor       =   16777215
         BackColor       =   16711680
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelWidth      =   2
      End
      Begin Threed.SSPanel sspT1Lin2 
         Height          =   645
         Index           =   3
         Left            =   1965
         TabIndex        =   22
         Top             =   2010
         Width           =   570
         _Version        =   65536
         _ExtentX        =   1005
         _ExtentY        =   1138
         _StockProps     =   15
         Caption         =   "99,99"
         ForeColor       =   16777215
         BackColor       =   16711680
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelWidth      =   2
      End
      Begin Threed.SSPanel sspT1Lin2 
         Height          =   645
         Index           =   4
         Left            =   2535
         TabIndex        =   23
         Top             =   2010
         Width           =   570
         _Version        =   65536
         _ExtentX        =   1005
         _ExtentY        =   1138
         _StockProps     =   15
         Caption         =   "99,99"
         ForeColor       =   16777215
         BackColor       =   16711680
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelWidth      =   2
      End
      Begin Threed.SSPanel sspT1Mensagem 
         Height          =   510
         Left            =   45
         TabIndex        =   24
         Top             =   2670
         Width           =   3060
         _Version        =   65536
         _ExtentX        =   5397
         _ExtentY        =   900
         _StockProps     =   15
         Caption         =   "Mensagem"
         ForeColor       =   16777215
         BackColor       =   16711680
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelWidth      =   2
      End
      Begin Threed.SSPanel sspT2Filme1 
         Height          =   615
         Left            =   3150
         TabIndex        =   25
         Top             =   465
         Width           =   1950
         _Version        =   65536
         _ExtentX        =   3440
         _ExtentY        =   1085
         _StockProps     =   15
         Caption         =   "Filme"
         ForeColor       =   16777215
         BackColor       =   16711680
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelWidth      =   2
      End
      Begin Threed.SSPanel sspT2Sala1 
         Height          =   300
         Index           =   0
         Left            =   5100
         TabIndex        =   26
         Top             =   465
         Width           =   1095
         _Version        =   65536
         _ExtentX        =   1931
         _ExtentY        =   529
         _StockProps     =   15
         Caption         =   "Sala-01"
         ForeColor       =   16777215
         BackColor       =   16711680
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelWidth      =   2
      End
      Begin Threed.SSPanel sspT2Sessoes1 
         Height          =   300
         Left            =   5100
         TabIndex        =   27
         Top             =   765
         Width           =   1095
         _Version        =   65536
         _ExtentX        =   1931
         _ExtentY        =   529
         _StockProps     =   15
         Caption         =   "Sessões"
         ForeColor       =   16777215
         BackColor       =   16711680
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelWidth      =   2
      End
      Begin Threed.SSPanel sspT2Sessoes1L1 
         Height          =   300
         Index           =   1
         Left            =   5655
         TabIndex        =   28
         Top             =   1065
         Width           =   555
         _Version        =   65536
         _ExtentX        =   979
         _ExtentY        =   529
         _StockProps     =   15
         Caption         =   "99:99"
         ForeColor       =   16777215
         BackColor       =   16711680
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelWidth      =   2
      End
      Begin Threed.SSPanel sspT2Sessoes1L1 
         Height          =   300
         Index           =   0
         Left            =   5100
         TabIndex        =   29
         Top             =   1065
         Width           =   555
         _Version        =   65536
         _ExtentX        =   979
         _ExtentY        =   529
         _StockProps     =   15
         Caption         =   "99:99"
         ForeColor       =   16777215
         BackColor       =   16711680
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelWidth      =   2
      End
      Begin Threed.SSPanel sspT2Sessoes1L2 
         Height          =   300
         Index           =   1
         Left            =   5655
         TabIndex        =   30
         Top             =   1365
         Width           =   555
         _Version        =   65536
         _ExtentX        =   979
         _ExtentY        =   529
         _StockProps     =   15
         Caption         =   "99:99"
         ForeColor       =   16777215
         BackColor       =   16711680
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelWidth      =   2
      End
      Begin Threed.SSPanel sspT2Sessoes1L2 
         Height          =   300
         Index           =   0
         Left            =   5100
         TabIndex        =   31
         Top             =   1365
         Width           =   555
         _Version        =   65536
         _ExtentX        =   979
         _ExtentY        =   529
         _StockProps     =   15
         Caption         =   "99:99"
         ForeColor       =   16777215
         BackColor       =   16711680
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelWidth      =   2
      End
      Begin Threed.SSPanel sspT2Sessao1 
         Height          =   600
         Left            =   4380
         TabIndex        =   32
         Top             =   1080
         Width           =   705
         _Version        =   65536
         _ExtentX        =   1244
         _ExtentY        =   1058
         _StockProps     =   15
         Caption         =   "99,99"
         ForeColor       =   16777215
         BackColor       =   16711680
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelWidth      =   2
      End
      Begin Threed.SSPanel sspT2Titulo1 
         Height          =   300
         Index           =   0
         Left            =   3165
         TabIndex        =   33
         Top             =   1080
         Width           =   1200
         _Version        =   65536
         _ExtentX        =   2117
         _ExtentY        =   529
         _StockProps     =   15
         Caption         =   "Prox. Sessão"
         ForeColor       =   16777215
         BackColor       =   16711680
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelWidth      =   2
      End
      Begin Threed.SSPanel sspT2Titulo1 
         Height          =   300
         Index           =   1
         Left            =   3165
         TabIndex        =   34
         Top             =   1380
         Width           =   1200
         _Version        =   65536
         _ExtentX        =   2117
         _ExtentY        =   529
         _StockProps     =   15
         Caption         =   "A Venda"
         ForeColor       =   16777215
         BackColor       =   16711680
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelWidth      =   2
      End
      Begin Threed.SSPanel sspT2Filme2 
         Height          =   615
         Left            =   3150
         TabIndex        =   35
         Top             =   1965
         Width           =   1950
         _Version        =   65536
         _ExtentX        =   3440
         _ExtentY        =   1085
         _StockProps     =   15
         Caption         =   "Filme"
         ForeColor       =   16777215
         BackColor       =   16711680
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelWidth      =   2
      End
      Begin Threed.SSPanel sspT2Sala2 
         Height          =   300
         Index           =   1
         Left            =   5100
         TabIndex        =   36
         Top             =   1965
         Width           =   1095
         _Version        =   65536
         _ExtentX        =   1931
         _ExtentY        =   529
         _StockProps     =   15
         Caption         =   "Sala-01"
         ForeColor       =   16777215
         BackColor       =   16711680
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelWidth      =   2
      End
      Begin Threed.SSPanel sspT2Sessoes2 
         Height          =   300
         Left            =   5100
         TabIndex        =   37
         Top             =   2265
         Width           =   1095
         _Version        =   65536
         _ExtentX        =   1931
         _ExtentY        =   529
         _StockProps     =   15
         Caption         =   "Sessões"
         ForeColor       =   16777215
         BackColor       =   16711680
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelWidth      =   2
      End
      Begin Threed.SSPanel sspT2Sessoes2L1 
         Height          =   300
         Index           =   1
         Left            =   5655
         TabIndex        =   38
         Top             =   2565
         Width           =   555
         _Version        =   65536
         _ExtentX        =   979
         _ExtentY        =   529
         _StockProps     =   15
         Caption         =   "99:99"
         ForeColor       =   16777215
         BackColor       =   16711680
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelWidth      =   2
      End
      Begin Threed.SSPanel sspT2Sessoes2L1 
         Height          =   300
         Index           =   0
         Left            =   5100
         TabIndex        =   39
         Top             =   2565
         Width           =   555
         _Version        =   65536
         _ExtentX        =   979
         _ExtentY        =   529
         _StockProps     =   15
         Caption         =   "99:99"
         ForeColor       =   16777215
         BackColor       =   16711680
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelWidth      =   2
      End
      Begin Threed.SSPanel sspT2Sessoes2L2 
         Height          =   300
         Index           =   1
         Left            =   5655
         TabIndex        =   40
         Top             =   2865
         Width           =   555
         _Version        =   65536
         _ExtentX        =   979
         _ExtentY        =   529
         _StockProps     =   15
         Caption         =   "99:99"
         ForeColor       =   16777215
         BackColor       =   16711680
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelWidth      =   2
      End
      Begin Threed.SSPanel sspT2Sessoes2L2 
         Height          =   300
         Index           =   0
         Left            =   5100
         TabIndex        =   41
         Top             =   2865
         Width           =   555
         _Version        =   65536
         _ExtentX        =   979
         _ExtentY        =   529
         _StockProps     =   15
         Caption         =   "99:99"
         ForeColor       =   16777215
         BackColor       =   16711680
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelWidth      =   2
      End
      Begin Threed.SSPanel sspT2Sessao2 
         Height          =   600
         Left            =   4380
         TabIndex        =   42
         Top             =   2580
         Width           =   705
         _Version        =   65536
         _ExtentX        =   1244
         _ExtentY        =   1058
         _StockProps     =   15
         Caption         =   "99,99"
         ForeColor       =   16777215
         BackColor       =   16711680
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelWidth      =   2
      End
      Begin Threed.SSPanel sspT2Titulo2 
         Height          =   300
         Index           =   0
         Left            =   3165
         TabIndex        =   43
         Top             =   2580
         Width           =   1200
         _Version        =   65536
         _ExtentX        =   2117
         _ExtentY        =   529
         _StockProps     =   15
         Caption         =   "Prox. Sessão"
         ForeColor       =   16777215
         BackColor       =   16711680
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelWidth      =   2
      End
      Begin Threed.SSPanel sspT2Titulo2 
         Height          =   300
         Index           =   1
         Left            =   3165
         TabIndex        =   44
         Top             =   2880
         Width           =   1200
         _Version        =   65536
         _ExtentX        =   2117
         _ExtentY        =   529
         _StockProps     =   15
         Caption         =   "A Venda"
         ForeColor       =   16777215
         BackColor       =   16711680
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelWidth      =   2
      End
      Begin Threed.SSPanel sspT3Titulo 
         Height          =   300
         Index           =   0
         Left            =   6255
         TabIndex        =   45
         Top             =   765
         Width           =   660
         _Version        =   65536
         _ExtentX        =   1164
         _ExtentY        =   529
         _StockProps     =   15
         Caption         =   "Sessão"
         ForeColor       =   16777215
         BackColor       =   16711680
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelWidth      =   2
      End
      Begin Threed.SSPanel sspT3Lin1 
         Height          =   300
         Index           =   0
         Left            =   6255
         TabIndex        =   46
         Top             =   1065
         Width           =   660
         _Version        =   65536
         _ExtentX        =   1164
         _ExtentY        =   529
         _StockProps     =   15
         Caption         =   "99:99"
         ForeColor       =   16777215
         BackColor       =   16711680
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelWidth      =   2
      End
      Begin Threed.SSPanel SSPanel44 
         Height          =   300
         Index           =   2
         Left            =   8085
         TabIndex        =   47
         Top             =   765
         Width           =   465
         _Version        =   65536
         _ExtentX        =   820
         _ExtentY        =   529
         _StockProps     =   15
         Caption         =   "Sala"
         ForeColor       =   16777215
         BackColor       =   16711680
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelWidth      =   2
         Font3D          =   2
         Begin Threed.SSPanel SSPanel46 
            Height          =   300
            Index           =   0
            Left            =   0
            TabIndex        =   48
            Top             =   0
            Width           =   1200
            _Version        =   65536
            _ExtentX        =   2117
            _ExtentY        =   529
            _StockProps     =   15
            Caption         =   "Filme"
            ForeColor       =   16777215
            BackColor       =   16711680
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BevelWidth      =   2
            Font3D          =   2
            Begin Threed.SSPanel sspT3Titulo 
               Height          =   300
               Index           =   3
               Left            =   0
               TabIndex        =   49
               Top             =   0
               Width           =   465
               _Version        =   65536
               _ExtentX        =   820
               _ExtentY        =   529
               _StockProps     =   15
               Caption         =   "Sala"
               ForeColor       =   16777215
               BackColor       =   16711680
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               BevelWidth      =   2
            End
         End
      End
      Begin Threed.SSPanel sspT3Lin1 
         Height          =   300
         Index           =   2
         Left            =   8085
         TabIndex        =   50
         Top             =   1065
         Width           =   465
         _Version        =   65536
         _ExtentX        =   820
         _ExtentY        =   529
         _StockProps     =   15
         Caption         =   "99"
         ForeColor       =   16777215
         BackColor       =   16711680
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelWidth      =   2
      End
      Begin Threed.SSPanel sspT3Titulo 
         Height          =   300
         Index           =   1
         Left            =   6900
         TabIndex        =   51
         Top             =   765
         Width           =   1200
         _Version        =   65536
         _ExtentX        =   2117
         _ExtentY        =   529
         _StockProps     =   15
         Caption         =   "Filme"
         ForeColor       =   16777215
         BackColor       =   16711680
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelWidth      =   2
      End
      Begin Threed.SSPanel sspT3Lin1 
         Height          =   300
         Index           =   1
         Left            =   6900
         TabIndex        =   52
         Top             =   1065
         Width           =   1200
         _Version        =   65536
         _ExtentX        =   2117
         _ExtentY        =   529
         _StockProps     =   15
         Caption         =   "XXXXXXXX"
         ForeColor       =   16777215
         BackColor       =   16711680
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelWidth      =   2
      End
      Begin Threed.SSPanel SSPanel48 
         Height          =   300
         Index           =   3
         Left            =   8550
         TabIndex        =   53
         Top             =   765
         Width           =   780
         _Version        =   65536
         _ExtentX        =   1376
         _ExtentY        =   529
         _StockProps     =   15
         ForeColor       =   16777215
         BackColor       =   16711680
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelWidth      =   2
         Font3D          =   2
         Begin Threed.SSPanel sspT3Titulo 
            Height          =   300
            Index           =   2
            Left            =   0
            TabIndex        =   54
            Top             =   0
            Width           =   1200
            _Version        =   65536
            _ExtentX        =   2117
            _ExtentY        =   529
            _StockProps     =   15
            ForeColor       =   16777215
            BackColor       =   16711680
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BevelWidth      =   2
            Font3D          =   2
         End
      End
      Begin Threed.SSPanel sspT3Lin1 
         Height          =   300
         Index           =   3
         Left            =   8550
         TabIndex        =   55
         Top             =   1065
         Width           =   780
         _Version        =   65536
         _ExtentX        =   1376
         _ExtentY        =   529
         _StockProps     =   15
         Caption         =   "A Venda"
         ForeColor       =   16777215
         BackColor       =   16711680
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelWidth      =   2
      End
      Begin Threed.SSPanel sspT3Hora 
         Height          =   300
         Left            =   6255
         TabIndex        =   56
         Top             =   465
         Width           =   660
         _Version        =   65536
         _ExtentX        =   1164
         _ExtentY        =   529
         _StockProps     =   15
         Caption         =   "99:99"
         ForeColor       =   16777215
         BackColor       =   16711680
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelWidth      =   2
      End
      Begin Threed.SSPanel sspT3TituloTel 
         Height          =   300
         Left            =   6915
         TabIndex        =   57
         Top             =   465
         Width           =   1635
         _Version        =   65536
         _ExtentX        =   2884
         _ExtentY        =   529
         _StockProps     =   15
         Caption         =   "Prôximas Sessões"
         ForeColor       =   16777215
         BackColor       =   16711680
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelWidth      =   2
      End
      Begin Threed.SSPanel sspT3Data 
         Height          =   300
         Left            =   8550
         TabIndex        =   58
         Top             =   465
         Width           =   780
         _Version        =   65536
         _ExtentX        =   1376
         _ExtentY        =   529
         _StockProps     =   15
         Caption         =   "99/99/99"
         ForeColor       =   16777215
         BackColor       =   16711680
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelWidth      =   2
      End
      Begin Threed.SSPanel sspT3Lin2 
         Height          =   300
         Index           =   0
         Left            =   6255
         TabIndex        =   59
         Top             =   1365
         Width           =   660
         _Version        =   65536
         _ExtentX        =   1164
         _ExtentY        =   529
         _StockProps     =   15
         Caption         =   "99:99"
         ForeColor       =   16777215
         BackColor       =   16711680
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelWidth      =   2
      End
      Begin Threed.SSPanel sspT3Lin2 
         Height          =   300
         Index           =   2
         Left            =   8085
         TabIndex        =   60
         Top             =   1365
         Width           =   465
         _Version        =   65536
         _ExtentX        =   820
         _ExtentY        =   529
         _StockProps     =   15
         Caption         =   "99"
         ForeColor       =   16777215
         BackColor       =   16711680
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelWidth      =   2
      End
      Begin Threed.SSPanel sspT3Lin2 
         Height          =   300
         Index           =   1
         Left            =   6900
         TabIndex        =   61
         Top             =   1365
         Width           =   1200
         _Version        =   65536
         _ExtentX        =   2117
         _ExtentY        =   529
         _StockProps     =   15
         Caption         =   "XXXXXXXX"
         ForeColor       =   16777215
         BackColor       =   16711680
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelWidth      =   2
      End
      Begin Threed.SSPanel sspT3Lin2 
         Height          =   300
         Index           =   3
         Left            =   8550
         TabIndex        =   62
         Top             =   1365
         Width           =   780
         _Version        =   65536
         _ExtentX        =   1376
         _ExtentY        =   529
         _StockProps     =   15
         Caption         =   "A Venda"
         ForeColor       =   16777215
         BackColor       =   16711680
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelWidth      =   2
      End
      Begin Threed.SSPanel sspT3Lin1 
         Height          =   300
         Index           =   4
         Left            =   6255
         TabIndex        =   63
         Top             =   1665
         Width           =   660
         _Version        =   65536
         _ExtentX        =   1164
         _ExtentY        =   529
         _StockProps     =   15
         Caption         =   "99:99"
         ForeColor       =   16777215
         BackColor       =   16711680
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelWidth      =   2
      End
      Begin Threed.SSPanel sspT3Lin1 
         Height          =   300
         Index           =   6
         Left            =   8085
         TabIndex        =   64
         Top             =   1665
         Width           =   465
         _Version        =   65536
         _ExtentX        =   820
         _ExtentY        =   529
         _StockProps     =   15
         Caption         =   "99"
         ForeColor       =   16777215
         BackColor       =   16711680
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelWidth      =   2
      End
      Begin Threed.SSPanel sspT3Lin1 
         Height          =   300
         Index           =   5
         Left            =   6900
         TabIndex        =   65
         Top             =   1665
         Width           =   1200
         _Version        =   65536
         _ExtentX        =   2117
         _ExtentY        =   529
         _StockProps     =   15
         Caption         =   "XXXXXXXX"
         ForeColor       =   16777215
         BackColor       =   16711680
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelWidth      =   2
      End
      Begin Threed.SSPanel sspT3Lin1 
         Height          =   300
         Index           =   7
         Left            =   8550
         TabIndex        =   66
         Top             =   1665
         Width           =   780
         _Version        =   65536
         _ExtentX        =   1376
         _ExtentY        =   529
         _StockProps     =   15
         Caption         =   "A Venda"
         ForeColor       =   16777215
         BackColor       =   16711680
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelWidth      =   2
      End
      Begin Threed.SSPanel sspT3Lin2 
         Height          =   300
         Index           =   4
         Left            =   6255
         TabIndex        =   67
         Top             =   1965
         Width           =   660
         _Version        =   65536
         _ExtentX        =   1164
         _ExtentY        =   529
         _StockProps     =   15
         Caption         =   "99:99"
         ForeColor       =   16777215
         BackColor       =   16711680
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelWidth      =   2
      End
      Begin Threed.SSPanel sspT3Lin2 
         Height          =   300
         Index           =   6
         Left            =   8085
         TabIndex        =   68
         Top             =   1965
         Width           =   465
         _Version        =   65536
         _ExtentX        =   820
         _ExtentY        =   529
         _StockProps     =   15
         Caption         =   "99"
         ForeColor       =   16777215
         BackColor       =   16711680
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelWidth      =   2
      End
      Begin Threed.SSPanel sspT3Lin2 
         Height          =   300
         Index           =   5
         Left            =   6900
         TabIndex        =   69
         Top             =   1965
         Width           =   1200
         _Version        =   65536
         _ExtentX        =   2117
         _ExtentY        =   529
         _StockProps     =   15
         Caption         =   "XXXXXXXX"
         ForeColor       =   16777215
         BackColor       =   16711680
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelWidth      =   2
      End
      Begin Threed.SSPanel sspT3Lin2 
         Height          =   300
         Index           =   7
         Left            =   8550
         TabIndex        =   70
         Top             =   1965
         Width           =   780
         _Version        =   65536
         _ExtentX        =   1376
         _ExtentY        =   529
         _StockProps     =   15
         Caption         =   "A Venda"
         ForeColor       =   16777215
         BackColor       =   16711680
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelWidth      =   2
      End
      Begin Threed.SSPanel sspT3Lin1 
         Height          =   300
         Index           =   8
         Left            =   6255
         TabIndex        =   71
         Top             =   2265
         Width           =   660
         _Version        =   65536
         _ExtentX        =   1164
         _ExtentY        =   529
         _StockProps     =   15
         Caption         =   "99:99"
         ForeColor       =   16777215
         BackColor       =   16711680
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelWidth      =   2
      End
      Begin Threed.SSPanel sspT3Lin1 
         Height          =   300
         Index           =   10
         Left            =   8085
         TabIndex        =   72
         Top             =   2265
         Width           =   465
         _Version        =   65536
         _ExtentX        =   820
         _ExtentY        =   529
         _StockProps     =   15
         Caption         =   "99"
         ForeColor       =   16777215
         BackColor       =   16711680
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelWidth      =   2
      End
      Begin Threed.SSPanel sspT3Lin1 
         Height          =   300
         Index           =   9
         Left            =   6900
         TabIndex        =   73
         Top             =   2265
         Width           =   1200
         _Version        =   65536
         _ExtentX        =   2117
         _ExtentY        =   529
         _StockProps     =   15
         Caption         =   "XXXXXXXX"
         ForeColor       =   16777215
         BackColor       =   16711680
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelWidth      =   2
      End
      Begin Threed.SSPanel sspT3Lin1 
         Height          =   300
         Index           =   11
         Left            =   8550
         TabIndex        =   74
         Top             =   2265
         Width           =   780
         _Version        =   65536
         _ExtentX        =   1376
         _ExtentY        =   529
         _StockProps     =   15
         Caption         =   "A Venda"
         ForeColor       =   16777215
         BackColor       =   16711680
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelWidth      =   2
      End
      Begin Threed.SSPanel sspT3Lin2 
         Height          =   300
         Index           =   8
         Left            =   6255
         TabIndex        =   75
         Top             =   2565
         Width           =   660
         _Version        =   65536
         _ExtentX        =   1164
         _ExtentY        =   529
         _StockProps     =   15
         Caption         =   "99:99"
         ForeColor       =   16777215
         BackColor       =   16711680
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelWidth      =   2
      End
      Begin Threed.SSPanel sspT3Lin2 
         Height          =   300
         Index           =   10
         Left            =   8085
         TabIndex        =   76
         Top             =   2565
         Width           =   465
         _Version        =   65536
         _ExtentX        =   820
         _ExtentY        =   529
         _StockProps     =   15
         Caption         =   "99"
         ForeColor       =   16777215
         BackColor       =   16711680
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelWidth      =   2
      End
      Begin Threed.SSPanel sspT3Lin2 
         Height          =   300
         Index           =   9
         Left            =   6900
         TabIndex        =   77
         Top             =   2565
         Width           =   1200
         _Version        =   65536
         _ExtentX        =   2117
         _ExtentY        =   529
         _StockProps     =   15
         Caption         =   "XXXXXXXX"
         ForeColor       =   16777215
         BackColor       =   16711680
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelWidth      =   2
      End
      Begin Threed.SSPanel sspT3Lin2 
         Height          =   300
         Index           =   11
         Left            =   8550
         TabIndex        =   78
         Top             =   2565
         Width           =   780
         _Version        =   65536
         _ExtentX        =   1376
         _ExtentY        =   529
         _StockProps     =   15
         Caption         =   "A Venda"
         ForeColor       =   16777215
         BackColor       =   16711680
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelWidth      =   2
      End
      Begin Threed.SSPanel sspLotado 
         Height          =   300
         Left            =   3945
         TabIndex        =   79
         Top             =   3285
         Width           =   1575
         _Version        =   65536
         _ExtentX        =   2778
         _ExtentY        =   529
         _StockProps     =   15
         Caption         =   "Lotado"
         ForeColor       =   16777215
         BackColor       =   16711680
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelWidth      =   2
      End
      Begin VB.Label lblLotado 
         AutoSize        =   -1  'True
         Caption         =   "Mensagem Lotado:"
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
         Left            =   2280
         TabIndex        =   83
         Top             =   3330
         Width           =   1620
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Tela Próximas Sessões"
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
         Left            =   6840
         TabIndex        =   82
         Top             =   225
         Width           =   1965
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Tela Filme"
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
         Left            =   4245
         TabIndex        =   81
         Top             =   240
         Width           =   885
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Tela Preços"
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
         Left            =   975
         TabIndex        =   80
         Top             =   240
         Width           =   1035
      End
   End
   Begin VB.CommandButton cmdAlteraParametros 
      Caption         =   "Alterar"
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
      Left            =   3645
      TabIndex        =   2
      Top             =   5865
      Width           =   1980
   End
   Begin VB.CommandButton cmdCancelaParametros 
      Caption         =   "Cancela"
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
      Left            =   4770
      TabIndex        =   1
      Top             =   5865
      Width           =   1980
   End
   Begin VB.CommandButton cmdOkParametros 
      Caption         =   "Ok"
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
      Left            =   2520
      TabIndex        =   0
      Top             =   5865
      Width           =   1980
   End
End
Attribute VB_Name = "frmParametrosVideoHall"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private gRSParametros  As New ADODB.Recordset
Private modoParametros As String
Private Sub Form_Load()
    
    iTop = ((MDIFrmAdmCine.Height - Me.Height) \ 2)
    iLeft = ((MDIFrmAdmCine.width - Me.width) \ 2)
    Me.Move iLeft, iTop
  
    fraParametros.Enabled = False
   
   cmdAlteraParametros.Visible = True
   cmdOkParametros.Visible = False
   cmdCancelaParametros.Visible = False
   
   cmdAlteraParametros.Enabled = True
   cmdOkParametros.Enabled = False
   cmdCancelaParametros.Enabled = False
   
   cmdAlteraParametros.ZOrder 0
   
   Call CarregaParametros
End Sub
Private Sub CarregaParametros()
   Dim sMsg As String
   Dim i    As Integer
   
   On Error GoTo TrataErro

   gRSParametros.Open "tb_parametros", "FILE NAME=" & App.Path & "\Cinema.udl", adOpenDynamic, adLockOptimistic, adCmdTableDirect
   
   
   If gRSParametros.EOF And gRSParametros.BOF Then
      gRSParametros.Close
      If Not insereParamDefault Then
         Exit Sub
      End If
      gRSParametros.Open "tb_parametros", gConnect, adOpenDynamic, adLockOptimistic, adCmdTableDirect
   End If

   txtTransicao.Text = CStr(gRSParametros.Fields("transicao").Value)
   txtIntermitencia.Text = CStr(gRSParametros.Fields("intermitencia").Value)
   txtVelocMsg.Text = CStr(gRSParametros.Fields("velocMsg").Value)
   txtVendaAntes.Text = CStr(gRSParametros.Fields("vendaAndes").Value)
   txtVendaDepois.Text = CStr(gRSParametros.Fields("vendaDepois").Value)
   mkeHrLimPeriodo.Text = Format(gRSParametros.Fields("hrLimitePeriodo").Value, "Hh:Nn")
   If gRSParametros.Fields("telaSessoes").Value Then
      chkSessoes.Value = vbChecked
   Else
      chkSessoes.Value = vbUnchecked
   End If
   If gRSParametros.Fields("telaFilme").Value Then
      chkFilme.Value = vbChecked
   Else
      chkFilme.Value = vbUnchecked
   End If
   If gRSParametros.Fields("telaPrecos").Value Then
      chkPrecos.Value = vbChecked
   Else
      chkPrecos.Value = vbUnchecked
   End If
   If gRSParametros.Fields("telaTrailer").Value Then
      chkTrailer.Value = vbChecked
   Else
      chkTrailer.Value = vbUnchecked
   End If
   If gRSParametros.Fields("telaImagem").Value Then
      chkImagem.Value = vbChecked
   Else
      chkImagem.Value = vbUnchecked
   End If
   
   If Not IsNull(gRSParametros.Fields("mensagem").Value) Then
      txtMensagen.Text = gRSParametros.Fields("mensagem").Value
   Else
      txtMensagen.Text = ""
   End If
   
   sspT1Filme.BackColor = gRSParametros.Fields("corFundT1Filme").Value
   sspT1Filme.ForeColor = gRSParametros.Fields("corTextT1Filme").Value
   sspT1Hora.BackColor = gRSParametros.Fields("corFundT1Hora").Value
   sspT1Hora.ForeColor = gRSParametros.Fields("corTextT1Hora").Value
   
   For i = sspT1Titulo1.LBound To sspT1Titulo1.UBound
      sspT1Titulo1(i).BackColor = gRSParametros.Fields("corFundT1Tutulo1").Value
      sspT1Titulo1(i).ForeColor = gRSParametros.Fields("corTextT1Tutulo1").Value
   Next i
   
   For i = sspT1Titulo2.LBound To sspT1Titulo2.UBound
      sspT1Titulo2(i).BackColor = gRSParametros.Fields("corFundT1Titulo2").Value
      sspT1Titulo2(i).ForeColor = gRSParametros.Fields("corTextT1Titulo2").Value
   Next i
   
   For i = sspT1Lin1.LBound To sspT1Lin1.UBound
      sspT1Lin1(i).BackColor = gRSParametros.Fields("corFundT1Lin1").Value
      sspT1Lin1(i).ForeColor = gRSParametros.Fields("corTextT1Lin1").Value
   Next i
   
   For i = sspT1Lin2.LBound To sspT1Lin2.UBound
      sspT1Lin2(i).BackColor = gRSParametros.Fields("corFundT1Lin2").Value
      sspT1Lin2(i).ForeColor = gRSParametros.Fields("corTextT1Lin2").Value
   Next i
   
   sspT1Mensagem.BackColor = gRSParametros.Fields("corFundT1Mensagem").Value
   sspT1Mensagem.ForeColor = gRSParametros.Fields("corTextT1Mensagem").Value
   sspT2Filme1.BackColor = gRSParametros.Fields("corFundT2Filme1").Value
   sspT2Filme1.ForeColor = gRSParametros.Fields("corTextT2Filme1").Value
   sspT2Filme2.BackColor = gRSParametros.Fields("corFundT2Filme2").Value
   sspT2Filme2.ForeColor = gRSParametros.Fields("corTextT2Filme2").Value
   
   For i = sspT2Titulo1.LBound To sspT2Titulo1.UBound
      sspT2Titulo1(i).BackColor = gRSParametros.Fields("corFundT2Titulo1").Value
      sspT2Titulo1(i).ForeColor = gRSParametros.Fields("corTextT2Titulo1").Value
   Next i
   
   For i = sspT2Titulo2.LBound To sspT2Titulo2.UBound
      sspT2Titulo2(i).BackColor = gRSParametros.Fields("corFundT2Titulo2").Value
      sspT2Titulo2(i).ForeColor = gRSParametros.Fields("corTextT2Titulo2").Value
   Next i
   
   sspT2Sessao1.BackColor = gRSParametros.Fields("corFundT2Sessao1").Value
   sspT2Sessao1.ForeColor = gRSParametros.Fields("corTextT2Sessao1").Value
   sspT2Sessao2.BackColor = gRSParametros.Fields("corFundT2Sessao2").Value
   sspT2Sessao2.ForeColor = gRSParametros.Fields("corTextT2Sessao2").Value
   
   For i = sspT2Sala1.LBound To sspT2Sala1.UBound
      sspT2Sala1(i).BackColor = gRSParametros.Fields("corFundT2Sala1").Value
      sspT2Sala1(i).ForeColor = gRSParametros.Fields("corTextT2Sala1").Value
   Next i
   
   For i = sspT2Sala2.LBound To sspT2Sala2.UBound
      sspT2Sala2(i).BackColor = gRSParametros.Fields("corFundT2Sala2").Value
      sspT2Sala2(i).ForeColor = gRSParametros.Fields("corTextT2Sala2").Value
   Next i
   
   sspT2Sessoes1.BackColor = gRSParametros.Fields("corFundT2Sessoes1").Value
   sspT2Sessoes1.ForeColor = gRSParametros.Fields("corTextT2Sessoes1").Value
   
   For i = sspT2Sessoes1L1.LBound To sspT2Sessoes1L1.UBound
      sspT2Sessoes1L1(i).BackColor = gRSParametros.Fields("corFundT2Sessoes1L1").Value
      sspT2Sessoes1L1(i).ForeColor = gRSParametros.Fields("corTextT2Sessoes1L1").Value
   Next i
   
   For i = sspT2Sessoes1L2.LBound To sspT2Sessoes1L2.UBound
      sspT2Sessoes1L2(i).BackColor = gRSParametros.Fields("corFundT2Sessoes1L2").Value
      sspT2Sessoes1L2(i).ForeColor = gRSParametros.Fields("corTextT2Sessoes1L2").Value
   Next i
   
   sspT2Sessoes2.BackColor = gRSParametros.Fields("corFundT2Sessoes2").Value
   sspT2Sessoes2.ForeColor = gRSParametros.Fields("corTextT2Sessoes2").Value
   
   For i = sspT2Sessoes2L1.LBound To sspT2Sessoes2L1.UBound
      sspT2Sessoes2L1(i).BackColor = gRSParametros.Fields("corFundT2Sessoes2L1").Value
      sspT2Sessoes2L1(i).ForeColor = gRSParametros.Fields("corTextT2Sessoes2L1").Value
   Next i
   
   For i = sspT2Sessoes2L2.LBound To sspT2Sessoes2L2.UBound
      sspT2Sessoes2L2(i).BackColor = gRSParametros.Fields("corFundT2Sessoes2L2").Value
      sspT2Sessoes2L2(i).ForeColor = gRSParametros.Fields("corTextT2Sessoes2L2").Value
   Next i
   
   sspT2Mensagem.BackColor = gRSParametros.Fields("corFundT2Mensagem").Value
   sspT2Mensagem.ForeColor = gRSParametros.Fields("corTextT2Mensagem").Value
   sspT3Hora.BackColor = gRSParametros.Fields("corFundT3Hora").Value
   sspT3Hora.ForeColor = gRSParametros.Fields("corTextT3Hora").Value
   sspT3Data.BackColor = gRSParametros.Fields("corFundT3Data").Value
   sspT3Data.ForeColor = gRSParametros.Fields("corTextT3Data").Value
   sspT3TituloTel.BackColor = gRSParametros.Fields("corFundT3TituloTela").Value
   sspT3TituloTel.ForeColor = gRSParametros.Fields("corTextT3TituloTela").Value
   
   For i = sspT3Titulo.LBound To sspT3Titulo.UBound
      sspT3Titulo(i).BackColor = gRSParametros.Fields("corFundT3Titulo").Value
      sspT3Titulo(i).ForeColor = gRSParametros.Fields("corTextT3Titulo").Value
   Next i
   
   For i = sspT3Lin1.LBound To sspT3Lin1.UBound
      sspT3Lin1(i).BackColor = gRSParametros.Fields("corFundT3Lin1").Value
      sspT3Lin1(i).ForeColor = gRSParametros.Fields("corTextT3Lin1").Value
   Next i
   
   For i = sspT3Lin2.LBound To sspT3Lin2.UBound
      sspT3Lin2(i).BackColor = gRSParametros.Fields("corFundT3Lin2").Value
      sspT3Lin2(i).ForeColor = gRSParametros.Fields("corTextT3Lin2").Value
   Next i
   
   sspT3Mensagem.BackColor = gRSParametros.Fields("corFundT3Mensagem").Value
   sspT3Mensagem.ForeColor = gRSParametros.Fields("corTextT3Mensagem").Value
   
   sspLotado.BackColor = gRSParametros.Fields("corFundLotado").Value
   sspLotado.ForeColor = gRSParametros.Fields("corTextLotado").Value
   
   gRSParametros.Close
   
   Exit Sub
   
TrataErro:
   If gRSParametros.State = adStateOpen Then
       gRSParametros.Close
   End If
   
   sMsg = "Ocorreu um erro em carregaParametros." & vbCrLf
   sMsg = sMsg & Err.Number & " - " & Err.Description
   MsgBox sMsg, vbCritical, App.ProductName
   Exit Sub
   Resume 0
End Sub

Private Function veririficaParametros() As Boolean
   Dim telas As Integer

   veririficaParametros = False
   
   If Not IsNumeric(txtTransicao.Text) Then
      MsgBox "Tempo de transição invalido.", vbCritical, App.ProductName
      txtTransicao.SetFocus
      Exit Function
   End If
   
   If Not IsNumeric(txtIntermitencia.Text) Then
      MsgBox "Tempo de intermitencia invalido.", vbCritical, App.ProductName
      txtIntermitencia.SetFocus
      Exit Function
   End If
   
   If Not IsNumeric(txtVelocMsg.Text) Then
      MsgBox "Velocidade da Mensagem invalida.", vbCritical, App.ProductName
      txtIntermitencia.SetFocus
      Exit Function
   End If
   
   If Not IsNumeric(txtVendaAntes.Text) Then
      MsgBox "Tempo de à venda antes invalido.", vbCritical, App.ProductName
      txtVendaAntes.SetFocus
      Exit Function
   End If
   
   If Not IsNumeric(txtVendaDepois.Text) Then
      MsgBox "Tempo de à venda depois invalido.", vbCritical, App.ProductName
      txtVendaDepois.SetFocus
      Exit Function
   End If
   
   If Not IsDate(mkeHrLimPeriodo.Text) Then
      MsgBox "Horário limite entre períodos invalido.", vbCritical, App.ProductName
      mkeHrLimPeriodo.SetFocus
      Exit Function
   End If
   
   telas = 0
   
   If chkSessoes.Value = vbChecked Then
      telas = telas + 1
   End If
   If chkFilme.Value = vbChecked Then
      telas = telas + 1
   End If
   If chkPrecos.Value = vbChecked Then
      telas = telas + 1
   End If
   If chkTrailer.Value = vbChecked Then
      telas = telas + 1
   End If
   If chkImagem.Value = vbChecked Then
      telas = telas + 1
   End If
   
   If telas < 2 Then
      MsgBox "É necessário selecionar no mínimo DUAS tela para apresentação.", vbCritical, App.ProductName
      chkSessoes.SetFocus
      Exit Function
   End If
   
   veririficaParametros = True

End Function

Private Function alteraParametros() As Boolean
   Dim sMsg     As String
   
   On Error GoTo TrataErro
   
   alteraParametros = False
   
   gRSParametros.Open "tb_parametros", "FILE NAME=" & App.Path & "\Cinema.udl", adOpenDynamic, adLockOptimistic, adCmdTableDirect
   
   
   
   gRSParametros.Fields("transicao").Value = CLng(txtTransicao.Text)
   gRSParametros.Fields("intermitencia").Value = CLng(txtIntermitencia.Text)
   gRSParametros.Fields("velocMsg").Value = CLng(txtVelocMsg.Text)
   gRSParametros.Fields("vendaAndes").Value = CLng(txtVendaAntes.Text)
   gRSParametros.Fields("vendaDepois").Value = CLng(txtVendaDepois.Text)
   gRSParametros.Fields("hrLimitePeriodo").Value = CDate(mkeHrLimPeriodo.Text)
   
   If chkSessoes.Value = vbChecked Then
      gRSParametros.Fields("telaSessoes").Value = True
   Else
      gRSParametros.Fields("telaSessoes").Value = False
   End If
   If chkFilme.Value = vbChecked Then
      gRSParametros.Fields("telaFilme").Value = True
   Else
      gRSParametros.Fields("telaFilme").Value = False
   End If
   If chkPrecos.Value = vbChecked Then
      gRSParametros.Fields("telaPrecos").Value = True
   Else
      gRSParametros.Fields("telaPrecos").Value = False
   End If
   If chkTrailer.Value = vbChecked Then
      gRSParametros.Fields("telaTrailer").Value = True
   Else
      gRSParametros.Fields("telaTrailer").Value = False
   End If
      
   If chkImagem.Value = vbChecked Then
      gRSParametros.Fields("telaImagem").Value = True
   Else
      gRSParametros.Fields("telaImagem").Value = False
   End If
      
   gRSParametros.Fields("mensagem").Value = txtMensagen.Text
   
   gRSParametros.Fields("corFundT1Filme").Value = sspT1Filme.BackColor
   gRSParametros.Fields("corTextT1Filme").Value = sspT1Filme.ForeColor
   gRSParametros.Fields("corFundT1Hora").Value = sspT1Hora.BackColor
   gRSParametros.Fields("corTextT1Hora").Value = sspT1Hora.ForeColor
   gRSParametros.Fields("corFundT1Tutulo1").Value = sspT1Titulo1(sspT1Titulo1.LBound).BackColor
   gRSParametros.Fields("corTextT1Tutulo1").Value = sspT1Titulo1(sspT1Titulo1.LBound).ForeColor
   gRSParametros.Fields("corFundT1Titulo2").Value = sspT1Titulo2(sspT1Titulo2.LBound).BackColor
   gRSParametros.Fields("corTextT1Titulo2").Value = sspT1Titulo2(sspT1Titulo2.LBound).ForeColor
   gRSParametros.Fields("corFundT1Lin1").Value = sspT1Lin1(sspT1Lin1.LBound).BackColor
   gRSParametros.Fields("corTextT1Lin1").Value = sspT1Lin1(sspT1Lin1.LBound).ForeColor
   gRSParametros.Fields("corFundT1Lin2").Value = sspT1Lin2(sspT1Lin2.LBound).BackColor
   gRSParametros.Fields("corTextT1Lin2").Value = sspT1Lin2(sspT1Lin2.LBound).ForeColor
   gRSParametros.Fields("corFundT1Mensagem").Value = sspT1Mensagem.BackColor
   gRSParametros.Fields("corTextT1Mensagem").Value = sspT1Mensagem.ForeColor
   gRSParametros.Fields("corFundT2Filme1").Value = sspT2Filme1.BackColor
   gRSParametros.Fields("corTextT2Filme1").Value = sspT2Filme1.ForeColor
   gRSParametros.Fields("corFundT2Filme2").Value = sspT2Filme2.BackColor
   gRSParametros.Fields("corTextT2Filme2").Value = sspT2Filme2.ForeColor
   gRSParametros.Fields("corFundT2Titulo1").Value = sspT2Titulo1(sspT2Titulo1.LBound).BackColor
   gRSParametros.Fields("corTextT2Titulo1").Value = sspT2Titulo1(sspT2Titulo1.LBound).ForeColor
   gRSParametros.Fields("corFundT2Titulo2").Value = sspT2Titulo2(sspT2Titulo2.LBound).BackColor
   gRSParametros.Fields("corTextT2Titulo2").Value = sspT2Titulo2(sspT2Titulo2.LBound).ForeColor
   gRSParametros.Fields("corFundT2Sessao1").Value = sspT2Sessao1.BackColor
   gRSParametros.Fields("corTextT2Sessao1").Value = sspT2Sessao1.ForeColor
   gRSParametros.Fields("corFundT2Sessao2").Value = sspT2Sessao2.BackColor
   gRSParametros.Fields("corTextT2Sessao2").Value = sspT2Sessao2.ForeColor
   gRSParametros.Fields("corFundT2Sala1").Value = sspT2Sala1(sspT2Sala1.LBound).BackColor
   gRSParametros.Fields("corTextT2Sala1").Value = sspT2Sala1(sspT2Sala1.LBound).ForeColor
   gRSParametros.Fields("corFundT2Sala2").Value = sspT2Sala2(sspT2Sala2.LBound).BackColor
   gRSParametros.Fields("corTextT2Sala2").Value = sspT2Sala2(sspT2Sala2.LBound).ForeColor
   gRSParametros.Fields("corFundT2Sessoes1").Value = sspT2Sessoes1.BackColor
   gRSParametros.Fields("corTextT2Sessoes1").Value = sspT2Sessoes1.ForeColor
   gRSParametros.Fields("corFundT2Sessoes1L1").Value = sspT2Sessoes1L1(sspT2Sessoes1L1.LBound).BackColor
   gRSParametros.Fields("corTextT2Sessoes1L1").Value = sspT2Sessoes1L1(sspT2Sessoes1L1.LBound).ForeColor
   gRSParametros.Fields("corFundT2Sessoes1L2").Value = sspT2Sessoes1L2(sspT2Sessoes1L2.LBound).BackColor
   gRSParametros.Fields("corTextT2Sessoes1L2").Value = sspT2Sessoes1L2(sspT2Sessoes1L2.LBound).ForeColor
   gRSParametros.Fields("corFundT2Sessoes2").Value = sspT2Sessoes2.BackColor
   gRSParametros.Fields("corTextT2Sessoes2").Value = sspT2Sessoes2.ForeColor
   gRSParametros.Fields("corFundT2Sessoes2L1").Value = sspT2Sessoes2L1(sspT2Sessoes2L1.LBound).BackColor
   gRSParametros.Fields("corTextT2Sessoes2L1").Value = sspT2Sessoes2L1(sspT2Sessoes2L1.LBound).ForeColor
   gRSParametros.Fields("corFundT2Sessoes2L2").Value = sspT2Sessoes2L2(sspT2Sessoes2L2.LBound).BackColor
   gRSParametros.Fields("corTextT2Sessoes2L2").Value = sspT2Sessoes2L2(sspT2Sessoes2L2.LBound).ForeColor
   gRSParametros.Fields("corFundT2Mensagem").Value = sspT2Mensagem.BackColor
   gRSParametros.Fields("corTextT2Mensagem").Value = sspT2Mensagem.ForeColor
   gRSParametros.Fields("corFundT3Hora").Value = sspT3Hora.BackColor
   gRSParametros.Fields("corTextT3Hora").Value = sspT3Hora.ForeColor
   gRSParametros.Fields("corFundT3Data").Value = sspT3Data.BackColor
   gRSParametros.Fields("corTextT3Data").Value = sspT3Data.ForeColor
   gRSParametros.Fields("corFundT3TituloTela").Value = sspT3TituloTel.BackColor
   gRSParametros.Fields("corTextT3TituloTela").Value = sspT3TituloTel.ForeColor
   gRSParametros.Fields("corFundT3Titulo").Value = sspT3Titulo(sspT3Titulo.LBound).BackColor
   gRSParametros.Fields("corTextT3Titulo").Value = sspT3Titulo(sspT3Titulo.LBound).ForeColor
   gRSParametros.Fields("corFundT3Lin1").Value = sspT3Lin1(sspT3Lin1.LBound).BackColor
   gRSParametros.Fields("corTextT3Lin1").Value = sspT3Lin1(sspT3Lin1.LBound).ForeColor
   gRSParametros.Fields("corFundT3Lin2").Value = sspT3Lin2(sspT3Lin2.LBound).BackColor
   gRSParametros.Fields("corTextT3Lin2").Value = sspT3Lin2(sspT3Lin2.LBound).ForeColor
   gRSParametros.Fields("corFundT3Mensagem").Value = sspT3Mensagem.BackColor
   gRSParametros.Fields("corTextT3Mensagem").Value = sspT3Mensagem.ForeColor
   gRSParametros.Fields("corFundLotado").Value = sspLotado.BackColor
   gRSParametros.Fields("corTextLotado").Value = sspLotado.ForeColor
      
   gRSParametros.Update
   
   gRSParametros.Close
   
   alteraParametros = True
   
   Exit Function
   
TrataErro:
   If gRSParametros.State = adStateOpen Then
       gRSParametros.Close
   End If
   
   sMsg = "Ocorreu um erro em alteraParametros." & vbCrLf
   sMsg = sMsg & Err.Number & " - " & Err.Description
   MsgBox sMsg, vbCritical, App.ProductName
   Exit Function
   Resume 0
End Function

Private Sub cmdAlteraParametros_Click()
   Dim i As Integer
   
   fraParametros.Enabled = True
   fraCores.Enabled = True
   
   cmdAlteraParametros.Visible = False
   cmdOkParametros.Visible = True
   cmdCancelaParametros.Visible = True
   
   cmdAlteraParametros.Enabled = False
   cmdOkParametros.Enabled = True
   cmdCancelaParametros.Enabled = True
   
   cmdOkParametros.ZOrder 0
   cmdCancelaParametros.ZOrder 0
   
   modoFeriado = "alterar"
   modoParametros = "alterar"
End Sub

Private Sub cmdOkParametros_Click()
   If Not alteraParametros Then
      Exit Sub
   End If
   
   cmdCancelaParametros_Click
End Sub

Private Sub cmdCancelaParametros_Click()
   fraParametros.Enabled = False
   fraCores.Enabled = False
   
   cmdAlteraParametros.Visible = True
   cmdOkParametros.Visible = False
   cmdCancelaParametros.Visible = False
   
   cmdAlteraParametros.Enabled = True
   cmdOkParametros.Enabled = False
   cmdCancelaParametros.Enabled = False
   
   cmdAlteraParametros.ZOrder 0
   
   Call CarregaParametros
End Sub

Private Sub selCol1(ctrlCor As Control)
   Dim sMsg As String
   Dim cor  As Integer
   
   On Error GoTo TrataErro
   
   If frmTipoCor.loadCor(cor) Then
      CommonDialog2.CancelError = True
      CommonDialog2.DialogTitle = App.ProductName
      CommonDialog2.Flags = cdlCCFullOpen Or cdlCCRGBInit
      If cor = 1 Then
         CommonDialog2.Color = ctrlCor.BackColor
      Else
         CommonDialog2.Color = ctrlCor.ForeColor
      End If
      CommonDialog2.ShowColor
      
      If cor = 1 Then
         ctrlCor.BackColor = CommonDialog2.Color
      Else
         ctrlCor.ForeColor = CommonDialog2.Color
      End If
      
   End If

   Exit Sub

TrataErro:
   If Err.Number <> cdlCancel Then
      sMsg = "Ocorreu um erro em selCol1." & vbCrLf
      sMsg = sMsg & Err.Number & " - " & Err.Description
      MsgBox sMsg, vbCritical, App.ProductName
   End If
End Sub

Private Sub selCol2(ByRef ctrlCor As Object)
   Dim sMsg As String
   Dim cor  As Integer
   Dim i    As Integer
   
   On Error GoTo TrataErro
   
   If frmTipoCor.loadCor(cor) Then
      
      CommonDialog2.CancelError = True
      CommonDialog2.DialogTitle = App.ProductName
      CommonDialog2.Flags = cdlCCFullOpen Or cdlCCRGBInit
      If cor = 1 Then
         CommonDialog2.Color = ctrlCor(ctrlCor.LBound).BackColor
      Else
         CommonDialog2.Color = ctrlCor(ctrlCor.LBound).ForeColor
      End If
      CommonDialog2.ShowColor
      
      If cor = 1 Then
         ctrlCor(ctrlCor.LBound).BackColor = CommonDialog2.Color
         For i = ctrlCor.LBound + 1 To ctrlCor.UBound
            ctrlCor(i).BackColor = ctrlCor(ctrlCor.LBound).BackColor
         Next i
      Else
         ctrlCor(ctrlCor.LBound).ForeColor = CommonDialog2.Color
         For i = ctrlCor.LBound + 1 To ctrlCor.UBound
            ctrlCor(i).ForeColor = ctrlCor(ctrlCor.LBound).ForeColor
         Next i
      End If
      
   End If

   Exit Sub

TrataErro:
   If Err.Number <> cdlCancel Then
      sMsg = "Ocorreu um erro em selCol2." & vbCrLf
      sMsg = sMsg & Err.Number & " - " & Err.Description
      MsgBox sMsg, vbCritical, App.ProductName
   End If
End Sub

'Private Sub mkeData_GotFocus()
'   mkeData.SelStart = 0
'   mkeData.SelLength = Len(mkeData.Text)
'End Sub

'Private Sub mkeDuracao_GotFocus()
'   mkeDuracao.SelStart = 0
'   mkeDuracao.SelLength = Len(mkeDuracao.Text)
'End Sub

'Private Sub mkeSessao_GotFocus(Index As Integer)
   'mkeSessao(Index).SelStart = 0
   'mkeSessao(Index).SelLength = Len(mkeSessao(Index).Text)
'End Sub

Private Sub sspLotado_Click()
   Call selCol1(sspLotado)
End Sub

Private Sub sspT1Filme_Click()
   Call selCol1(sspT1Filme)
End Sub

Private Sub sspT1Hora_Click()
   Call selCol1(sspT1Hora)
End Sub

Private Sub sspT1Lin1_Click(Index As Integer)
   Call selCol2(sspT1Lin1)
End Sub

Private Sub sspT1Lin2_Click(Index As Integer)
   Call selCol2(sspT1Lin2)
End Sub

Private Sub sspT1Mensagem_Click()
   Call selCol1(sspT1Mensagem)
End Sub

Private Sub sspT1Titulo1_Click(Index As Integer)
   Call selCol2(sspT1Titulo1)
End Sub

Private Sub sspT1Titulo2_Click(Index As Integer)
   Call selCol2(sspT1Titulo2)
End Sub

Private Sub sspT2Filme1_Click()
   Call selCol1(sspT2Filme1)
End Sub

Private Sub sspT2Filme2_Click()
   Call selCol1(sspT2Filme2)
End Sub

Private Sub sspT2Mensagem_Click()
   Call selCol1(sspT2Mensagem)
End Sub

Private Sub sspT2Sala1_Click(Index As Integer)
   Call selCol2(sspT2Sala1)
End Sub

Private Sub sspT2Sala2_Click(Index As Integer)
   Call selCol2(sspT2Sala2)
End Sub

Private Sub sspT2Sessao1_Click()
   Call selCol1(sspT2Sessao1)
End Sub

Private Sub sspT2Sessao2_Click()
   Call selCol1(sspT2Sessao2)
End Sub

Private Sub sspT2Sessoes1_Click()
   Call selCol1(sspT2Sessoes1)
End Sub

Private Sub sspT2Sessoes1L1_Click(Index As Integer)
   Call selCol2(sspT2Sessoes1L1)
End Sub

Private Sub sspT2Sessoes1L2_Click(Index As Integer)
   Call selCol2(sspT2Sessoes1L2)
End Sub

Private Sub sspT2Sessoes2_Click()
   Call selCol1(sspT2Sessoes2)
End Sub

Private Sub sspT2Sessoes2L1_Click(Index As Integer)
   Call selCol2(sspT2Sessoes2L1)
End Sub

Private Sub sspT2Sessoes2L2_Click(Index As Integer)
   Call selCol2(sspT2Sessoes2L2)
End Sub

Private Sub sspT2Titulo1_Click(Index As Integer)
   Call selCol2(sspT2Titulo1)
End Sub

Private Sub sspT2Titulo2_Click(Index As Integer)
   Call selCol2(sspT2Titulo2)
End Sub

Private Sub sspT3Data_Click()
   Call selCol1(sspT3Data)
End Sub

Private Sub sspT3Hora_Click()
   Call selCol1(sspT3Hora)
End Sub

Private Sub sspT3Lin1_Click(Index As Integer)
   Call selCol2(sspT3Lin1)
End Sub

Private Sub sspT3Lin2_Click(Index As Integer)
   Call selCol2(sspT3Lin2)
End Sub

Private Sub sspT3Mensagem_Click()
   Call selCol1(sspT3Mensagem)
End Sub

Private Sub sspT3Titulo_Click(Index As Integer)
   Call selCol2(sspT3Titulo)
End Sub

Private Sub sspT3TituloTel_Click()
   Call selCol1(sspT3TituloTel)
End Sub
Private Function insereParamDefault() As Boolean
   Dim sMsg     As String
   
   On Error GoTo TrataErro
   
   insereParamDefault = False
   
   gRSParametros.Open "tb_parametros", gConnect, adOpenDynamic, adLockOptimistic, adCmdTableDirect
   
   gRSParametros.AddNew
   
   gRSParametros.Fields("transicao").Value = 15
   gRSParametros.Fields("intermitencia").Value = 1000
   gRSParametros.Fields("velocMsg").Value = 15
   gRSParametros.Fields("vendaAndes").Value = 20
   gRSParametros.Fields("vendaDepois").Value = 20
   gRSParametros.Fields("hrLimitePeriodo").Value = CDate("17:00")
   gRSParametros.Fields("telaSessoes").Value = True
   gRSParametros.Fields("telaFilme").Value = True
   gRSParametros.Fields("telaPrecos").Value = True
   gRSParametros.Fields("telaTrailer").Value = False
   gRSParametros.Fields("telaImagem").Value = True
   
   gRSParametros.Fields("mensagem").Value = ""
   
   gRSParametros.Fields("corFundT1Filme").Value = 1319162
   gRSParametros.Fields("corTextT1Filme").Value = 16777215
   gRSParametros.Fields("corFundT1Hora").Value = 1319162
   gRSParametros.Fields("corTextT1Hora").Value = 16777215
   gRSParametros.Fields("corFundT1Tutulo1").Value = 1319162
   gRSParametros.Fields("corTextT1Tutulo1").Value = 16777215
   gRSParametros.Fields("corFundT1Titulo2").Value = 4344827
   gRSParametros.Fields("corTextT1Titulo2").Value = 16777215
   gRSParametros.Fields("corFundT1Lin1").Value = 16582188
   gRSParametros.Fields("corTextT1Lin1").Value = 16777215
   gRSParametros.Fields("corFundT1Lin2").Value = 16681289
   gRSParametros.Fields("corTextT1Lin2").Value = 16777215
   gRSParametros.Fields("corFundT1Mensagem").Value = 0
   gRSParametros.Fields("corTextT1Mensagem").Value = 16777215
   gRSParametros.Fields("corFundT2Filme1").Value = 16582188
   gRSParametros.Fields("corTextT2Filme1").Value = 16777215
   gRSParametros.Fields("corFundT2Filme2").Value = 16681289
   gRSParametros.Fields("corTextT2Filme2").Value = 16777215
   gRSParametros.Fields("corFundT2Titulo1").Value = 16582188
   gRSParametros.Fields("corTextT2Titulo1").Value = 16777215
   gRSParametros.Fields("corFundT2Titulo2").Value = 16681289
   gRSParametros.Fields("corTextT2Titulo2").Value = 16777215
   gRSParametros.Fields("corFundT2Sessao1").Value = 1319162
   gRSParametros.Fields("corTextT2Sessao1").Value = 16777215
   gRSParametros.Fields("corFundT2Sessao2").Value = 4344827
   gRSParametros.Fields("corTextT2Sessao2").Value = 16777215
   gRSParametros.Fields("corFundT2Sala1").Value = 16582188
   gRSParametros.Fields("corTextT2Sala1").Value = 16777215
   gRSParametros.Fields("corFundT2Sala2").Value = 16681289
   gRSParametros.Fields("corTextT2Sala2").Value = 16777215
   gRSParametros.Fields("corFundT2Sessoes1").Value = 1319162
   gRSParametros.Fields("corTextT2Sessoes1").Value = 16777215
   gRSParametros.Fields("corFundT2Sessoes1L1").Value = 13072967
   gRSParametros.Fields("corTextT2Sessoes1L1").Value = 16777215
   gRSParametros.Fields("corFundT2Sessoes1L2").Value = 16674851
   gRSParametros.Fields("corTextT2Sessoes1L2").Value = 16777215
   gRSParametros.Fields("corFundT2Sessoes2").Value = 4344827
   gRSParametros.Fields("corTextT2Sessoes2").Value = 16777215
   gRSParametros.Fields("corFundT2Sessoes2L1").Value = 13072967
   gRSParametros.Fields("corTextT2Sessoes2L1").Value = 16777215
   gRSParametros.Fields("corFundT2Sessoes2L2").Value = 16674851
   gRSParametros.Fields("corTextT2Sessoes2L2").Value = 16777215
   gRSParametros.Fields("corFundT2Mensagem").Value = 0
   gRSParametros.Fields("corTextT2Mensagem").Value = 16777215
   gRSParametros.Fields("corFundT3Hora").Value = 1319162
   gRSParametros.Fields("corTextT3Hora").Value = 16777215
   gRSParametros.Fields("corFundT3Data").Value = 1319162
   gRSParametros.Fields("corTextT3Data").Value = 16777215
   gRSParametros.Fields("corFundT3TituloTela").Value = 1319162
   gRSParametros.Fields("corTextT3TituloTela").Value = 16777215
   gRSParametros.Fields("corFundT3Titulo").Value = 4344827
   gRSParametros.Fields("corTextT3Titulo").Value = 16777215
   gRSParametros.Fields("corFundT3Lin1").Value = 16582188
   gRSParametros.Fields("corTextT3Lin1").Value = 16777215
   gRSParametros.Fields("corFundT3Lin2").Value = 16681289
   gRSParametros.Fields("corTextT3Lin2").Value = 16777215
   gRSParametros.Fields("corFundT3Mensagem").Value = 0
   gRSParametros.Fields("corTextT3Mensagem").Value = 16777215
   gRSParametros.Fields("corFundLotado").Value = 255
   gRSParametros.Fields("corTextLotado").Value = 16777215

   gRSParametros.Update
   
   gRSParametros.Close
   
   insereParamDefault = True
   
   Exit Function
   
TrataErro:
   If gRSParametros.State = adStateOpen Then
       gRSParametros.Close
   End If
   
   sMsg = "Ocorreu um erro em insereParamDefault." & vbCrLf
   sMsg = sMsg & Err.Number & " - " & Err.Description
   MsgBox sMsg, vbCritical, App.ProductName
   Exit Function
   Resume 0
End Function

