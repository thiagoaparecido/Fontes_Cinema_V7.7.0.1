VERSION 5.00
Object = "{6BF52A50-394A-11D3-B153-00C04F79FAA6}#1.0#0"; "wmp.dll"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Begin VB.Form frmApresent 
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   ClientHeight    =   7380
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   9825
   ControlBox      =   0   'False
   Icon            =   "frmApresent.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   MouseIcon       =   "frmApresent.frx":000C
   MousePointer    =   99  'Custom
   ScaleHeight     =   492
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   655
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.Frame fravideo 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   7140
      Left            =   0
      TabIndex        =   126
      Top             =   0
      Width           =   9540
      Begin WMPLibCtl.WindowsMediaPlayer WindowsMediaPlayer1 
         Height          =   6900
         Left            =   225
         TabIndex        =   127
         Top             =   150
         Width           =   9225
         URL             =   ""
         rate            =   1
         balance         =   0
         currentPosition =   0
         defaultFrame    =   ""
         playCount       =   1
         autoStart       =   0   'False
         currentMarker   =   0
         invokeURLs      =   -1  'True
         baseURL         =   ""
         volume          =   50
         mute            =   0   'False
         uiMode          =   "none"
         stretchToFit    =   0   'False
         windowlessVideo =   0   'False
         enabled         =   -1  'True
         enableContextMenu=   -1  'True
         fullScreen      =   0   'False
         SAMIStyle       =   ""
         SAMILang        =   ""
         SAMIFilename    =   ""
         captioningID    =   ""
         enableErrorDialogs=   0   'False
         _cx             =   16272
         _cy             =   12171
      End
   End
   Begin VB.Timer TimerVelocMsg 
      Left            =   1245
      Top             =   75
   End
   Begin VB.Timer TimerPisca 
      Left            =   675
      Top             =   75
   End
   Begin VB.Timer TimerTelas 
      Left            =   60
      Top             =   75
   End
   Begin VB.Frame fraImagem 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H00000000&
      Height          =   1.08000e5
      Left            =   0
      MouseIcon       =   "frmApresent.frx":0396
      MousePointer    =   4  'Icon
      TabIndex        =   1
      Top             =   0
      Width           =   1.43880e5
      Begin VB.Image Imagem 
         Height          =   7155
         Left            =   75
         MouseIcon       =   "frmApresent.frx":0720
         MousePointer    =   4  'Icon
         Top             =   75
         Visible         =   0   'False
         Width           =   9555
      End
   End
   Begin VB.Frame fraProxSessoes 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H00000000&
      Height          =   1.08000e5
      Left            =   0
      MouseIcon       =   "frmApresent.frx":0AAA
      MousePointer    =   4  'Icon
      TabIndex        =   3
      Top             =   0
      Width           =   1.44000e5
      Begin Threed.SSPanel sspTituloSessao 
         Height          =   465
         Left            =   15
         TabIndex        =   4
         Top             =   675
         Width           =   1320
         _Version        =   65536
         _ExtentX        =   2328
         _ExtentY        =   820
         _StockProps     =   15
         Caption         =   "SESSÃO"
         ForeColor       =   16777215
         BackColor       =   4344827
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial Black"
            Size            =   12
            Charset         =   0
            Weight          =   900
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelWidth      =   2
         FloodColor      =   16777215
         Font3D          =   2
         Alignment       =   6
      End
      Begin Threed.SSPanel sspTitProxSes 
         Height          =   615
         Left            =   1350
         TabIndex        =   5
         Top             =   30
         Width           =   6585
         _Version        =   65536
         _ExtentX        =   11615
         _ExtentY        =   1085
         _StockProps     =   15
         Caption         =   "PRÓXIMAS SESSÕES"
         ForeColor       =   16777215
         BackColor       =   1319162
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial Black"
            Size            =   20.25
            Charset         =   0
            Weight          =   900
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelWidth      =   2
         FloodColor      =   16777215
         Font3D          =   2
         Alignment       =   6
      End
      Begin Threed.SSPanel sspSessao 
         Height          =   465
         Index           =   0
         Left            =   15
         TabIndex        =   6
         Top             =   1170
         Width           =   1320
         _Version        =   65536
         _ExtentX        =   2328
         _ExtentY        =   820
         _StockProps     =   15
         Caption         =   "99:99"
         ForeColor       =   16777215
         BackColor       =   16582188
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial Black"
            Size            =   14.25
            Charset         =   0
            Weight          =   900
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelWidth      =   2
         FloodColor      =   16777215
         Font3D          =   2
         Alignment       =   6
      End
      Begin Threed.SSPanel sspSessao 
         Height          =   465
         Index           =   1
         Left            =   15
         TabIndex        =   7
         Top             =   1665
         Width           =   1320
         _Version        =   65536
         _ExtentX        =   2328
         _ExtentY        =   820
         _StockProps     =   15
         Caption         =   "99:99"
         ForeColor       =   16777215
         BackColor       =   16681289
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial Black"
            Size            =   14.25
            Charset         =   0
            Weight          =   900
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelWidth      =   2
         FloodColor      =   16777215
         Font3D          =   2
         Alignment       =   6
      End
      Begin Threed.SSPanel sspSessao 
         Height          =   465
         Index           =   2
         Left            =   15
         TabIndex        =   8
         Top             =   2160
         Width           =   1320
         _Version        =   65536
         _ExtentX        =   2328
         _ExtentY        =   820
         _StockProps     =   15
         Caption         =   "99:99"
         ForeColor       =   16777215
         BackColor       =   16582188
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial Black"
            Size            =   14.25
            Charset         =   0
            Weight          =   900
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelWidth      =   2
         FloodColor      =   16777215
         Font3D          =   2
         Alignment       =   6
      End
      Begin Threed.SSPanel sspSessao 
         Height          =   465
         Index           =   3
         Left            =   15
         TabIndex        =   9
         Top             =   2655
         Width           =   1320
         _Version        =   65536
         _ExtentX        =   2328
         _ExtentY        =   820
         _StockProps     =   15
         Caption         =   "99:99"
         ForeColor       =   16777215
         BackColor       =   16681289
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial Black"
            Size            =   14.25
            Charset         =   0
            Weight          =   900
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelWidth      =   2
         FloodColor      =   16777215
         Font3D          =   2
         Alignment       =   6
      End
      Begin Threed.SSPanel sspSessao 
         Height          =   465
         Index           =   4
         Left            =   15
         TabIndex        =   10
         Top             =   3150
         Width           =   1320
         _Version        =   65536
         _ExtentX        =   2328
         _ExtentY        =   820
         _StockProps     =   15
         Caption         =   "99:99"
         ForeColor       =   16777215
         BackColor       =   16582188
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial Black"
            Size            =   14.25
            Charset         =   0
            Weight          =   900
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelWidth      =   2
         FloodColor      =   16777215
         Font3D          =   2
         Alignment       =   6
      End
      Begin Threed.SSPanel sspSessao 
         Height          =   465
         Index           =   5
         Left            =   15
         TabIndex        =   11
         Top             =   3645
         Width           =   1320
         _Version        =   65536
         _ExtentX        =   2328
         _ExtentY        =   820
         _StockProps     =   15
         Caption         =   "99:99"
         ForeColor       =   16777215
         BackColor       =   16681289
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial Black"
            Size            =   14.25
            Charset         =   0
            Weight          =   900
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelWidth      =   2
         FloodColor      =   16777215
         Font3D          =   2
         Alignment       =   6
      End
      Begin Threed.SSPanel sspSessao 
         Height          =   465
         Index           =   6
         Left            =   15
         TabIndex        =   12
         Top             =   4140
         Width           =   1320
         _Version        =   65536
         _ExtentX        =   2328
         _ExtentY        =   820
         _StockProps     =   15
         Caption         =   "99:99"
         ForeColor       =   16777215
         BackColor       =   16582188
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial Black"
            Size            =   14.25
            Charset         =   0
            Weight          =   900
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelWidth      =   2
         FloodColor      =   16777215
         Font3D          =   2
         Alignment       =   6
      End
      Begin Threed.SSPanel sspSessao 
         Height          =   465
         Index           =   7
         Left            =   15
         TabIndex        =   13
         Top             =   4635
         Width           =   1320
         _Version        =   65536
         _ExtentX        =   2328
         _ExtentY        =   820
         _StockProps     =   15
         Caption         =   "99:99"
         ForeColor       =   16777215
         BackColor       =   16681289
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial Black"
            Size            =   14.25
            Charset         =   0
            Weight          =   900
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelWidth      =   2
         FloodColor      =   16777215
         Font3D          =   2
         Alignment       =   6
      End
      Begin Threed.SSPanel sspSessao 
         Height          =   465
         Index           =   8
         Left            =   15
         TabIndex        =   14
         Top             =   5130
         Width           =   1320
         _Version        =   65536
         _ExtentX        =   2328
         _ExtentY        =   820
         _StockProps     =   15
         Caption         =   "99:99"
         ForeColor       =   16777215
         BackColor       =   16582188
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial Black"
            Size            =   14.25
            Charset         =   0
            Weight          =   900
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelWidth      =   2
         FloodColor      =   16777215
         Font3D          =   2
         Alignment       =   6
      End
      Begin Threed.SSPanel sspSessao 
         Height          =   465
         Index           =   9
         Left            =   15
         TabIndex        =   15
         Top             =   5625
         Width           =   1320
         _Version        =   65536
         _ExtentX        =   2328
         _ExtentY        =   820
         _StockProps     =   15
         Caption         =   "99:99"
         ForeColor       =   16777215
         BackColor       =   16681289
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial Black"
            Size            =   14.25
            Charset         =   0
            Weight          =   900
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelWidth      =   2
         FloodColor      =   16777215
         Font3D          =   2
         Alignment       =   6
      End
      Begin Threed.SSPanel sspSessao 
         Height          =   465
         Index           =   10
         Left            =   15
         TabIndex        =   16
         Top             =   6120
         Width           =   1320
         _Version        =   65536
         _ExtentX        =   2328
         _ExtentY        =   820
         _StockProps     =   15
         Caption         =   "99:99"
         ForeColor       =   16777215
         BackColor       =   16582188
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial Black"
            Size            =   14.25
            Charset         =   0
            Weight          =   900
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelWidth      =   2
         FloodColor      =   16777215
         Font3D          =   2
         Alignment       =   6
      End
      Begin Threed.SSPanel sspTituloSala 
         Height          =   465
         Left            =   6900
         TabIndex        =   17
         Top             =   675
         Width           =   1035
         _Version        =   65536
         _ExtentX        =   1826
         _ExtentY        =   820
         _StockProps     =   15
         Caption         =   "SALA"
         ForeColor       =   16777215
         BackColor       =   4344827
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial Black"
            Size            =   12
            Charset         =   0
            Weight          =   900
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelWidth      =   2
         FloodColor      =   16777215
         Font3D          =   2
         Alignment       =   6
      End
      Begin Threed.SSPanel sspSala 
         Height          =   465
         Index           =   0
         Left            =   6900
         TabIndex        =   18
         Top             =   1170
         Width           =   1035
         _Version        =   65536
         _ExtentX        =   1826
         _ExtentY        =   820
         _StockProps     =   15
         Caption         =   "99"
         ForeColor       =   16777215
         BackColor       =   16582188
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial Black"
            Size            =   14.25
            Charset         =   0
            Weight          =   900
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelWidth      =   2
         FloodColor      =   16777215
         Font3D          =   2
         Alignment       =   6
      End
      Begin Threed.SSPanel sspSala 
         Height          =   465
         Index           =   1
         Left            =   6900
         TabIndex        =   19
         Top             =   1665
         Width           =   1035
         _Version        =   65536
         _ExtentX        =   1826
         _ExtentY        =   820
         _StockProps     =   15
         Caption         =   "99"
         ForeColor       =   16777215
         BackColor       =   16681289
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial Black"
            Size            =   14.25
            Charset         =   0
            Weight          =   900
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelWidth      =   2
         FloodColor      =   16777215
         Font3D          =   2
         Alignment       =   6
      End
      Begin Threed.SSPanel sspSala 
         Height          =   465
         Index           =   2
         Left            =   6900
         TabIndex        =   20
         Top             =   2160
         Width           =   1035
         _Version        =   65536
         _ExtentX        =   1826
         _ExtentY        =   820
         _StockProps     =   15
         Caption         =   "99"
         ForeColor       =   16777215
         BackColor       =   16582188
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial Black"
            Size            =   14.25
            Charset         =   0
            Weight          =   900
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelWidth      =   2
         FloodColor      =   16777215
         Font3D          =   2
         Alignment       =   6
      End
      Begin Threed.SSPanel sspSala 
         Height          =   465
         Index           =   3
         Left            =   6900
         TabIndex        =   21
         Top             =   2655
         Width           =   1035
         _Version        =   65536
         _ExtentX        =   1826
         _ExtentY        =   820
         _StockProps     =   15
         Caption         =   "99"
         ForeColor       =   16777215
         BackColor       =   16681289
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial Black"
            Size            =   14.25
            Charset         =   0
            Weight          =   900
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelWidth      =   2
         FloodColor      =   16777215
         Font3D          =   2
         Alignment       =   6
      End
      Begin Threed.SSPanel sspSala 
         Height          =   465
         Index           =   4
         Left            =   6900
         TabIndex        =   22
         Top             =   3150
         Width           =   1035
         _Version        =   65536
         _ExtentX        =   1826
         _ExtentY        =   820
         _StockProps     =   15
         Caption         =   "99"
         ForeColor       =   16777215
         BackColor       =   16582188
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial Black"
            Size            =   14.25
            Charset         =   0
            Weight          =   900
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelWidth      =   2
         FloodColor      =   16777215
         Font3D          =   2
         Alignment       =   6
      End
      Begin Threed.SSPanel sspSala 
         Height          =   465
         Index           =   5
         Left            =   6900
         TabIndex        =   23
         Top             =   3645
         Width           =   1035
         _Version        =   65536
         _ExtentX        =   1826
         _ExtentY        =   820
         _StockProps     =   15
         Caption         =   "99"
         ForeColor       =   16777215
         BackColor       =   16681289
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial Black"
            Size            =   14.25
            Charset         =   0
            Weight          =   900
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelWidth      =   2
         FloodColor      =   16777215
         Font3D          =   2
         Alignment       =   6
      End
      Begin Threed.SSPanel sspSala 
         Height          =   465
         Index           =   6
         Left            =   6900
         TabIndex        =   24
         Top             =   4140
         Width           =   1035
         _Version        =   65536
         _ExtentX        =   1826
         _ExtentY        =   820
         _StockProps     =   15
         Caption         =   "99"
         ForeColor       =   16777215
         BackColor       =   16582188
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial Black"
            Size            =   14.25
            Charset         =   0
            Weight          =   900
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelWidth      =   2
         FloodColor      =   16777215
         Font3D          =   2
         Alignment       =   6
      End
      Begin Threed.SSPanel sspSala 
         Height          =   465
         Index           =   7
         Left            =   6900
         TabIndex        =   25
         Top             =   4635
         Width           =   1035
         _Version        =   65536
         _ExtentX        =   1826
         _ExtentY        =   820
         _StockProps     =   15
         Caption         =   "99"
         ForeColor       =   16777215
         BackColor       =   16681289
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial Black"
            Size            =   14.25
            Charset         =   0
            Weight          =   900
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelWidth      =   2
         FloodColor      =   16777215
         Font3D          =   2
         Alignment       =   6
      End
      Begin Threed.SSPanel sspSala 
         Height          =   465
         Index           =   8
         Left            =   6900
         TabIndex        =   26
         Top             =   5130
         Width           =   1035
         _Version        =   65536
         _ExtentX        =   1826
         _ExtentY        =   820
         _StockProps     =   15
         Caption         =   "99"
         ForeColor       =   16777215
         BackColor       =   16582188
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial Black"
            Size            =   14.25
            Charset         =   0
            Weight          =   900
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelWidth      =   2
         FloodColor      =   16777215
         Font3D          =   2
         Alignment       =   6
      End
      Begin Threed.SSPanel sspSala 
         Height          =   465
         Index           =   9
         Left            =   6900
         TabIndex        =   27
         Top             =   5625
         Width           =   1035
         _Version        =   65536
         _ExtentX        =   1826
         _ExtentY        =   820
         _StockProps     =   15
         Caption         =   "99"
         ForeColor       =   16777215
         BackColor       =   16681289
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial Black"
            Size            =   14.25
            Charset         =   0
            Weight          =   900
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelWidth      =   2
         FloodColor      =   16777215
         Font3D          =   2
         Alignment       =   6
      End
      Begin Threed.SSPanel sspSala 
         Height          =   465
         Index           =   10
         Left            =   6900
         TabIndex        =   28
         Top             =   6120
         Width           =   1035
         _Version        =   65536
         _ExtentX        =   1826
         _ExtentY        =   820
         _StockProps     =   15
         Caption         =   "99"
         ForeColor       =   16777215
         BackColor       =   16582188
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial Black"
            Size            =   14.25
            Charset         =   0
            Weight          =   900
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelWidth      =   2
         FloodColor      =   16777215
         Font3D          =   2
         Alignment       =   6
      End
      Begin Threed.SSPanel sspTituloFilme 
         Height          =   465
         Left            =   1350
         TabIndex        =   29
         Top             =   675
         Width           =   5535
         _Version        =   65536
         _ExtentX        =   9763
         _ExtentY        =   820
         _StockProps     =   15
         Caption         =   "FILME"
         ForeColor       =   16777215
         BackColor       =   4344827
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial Black"
            Size            =   12
            Charset         =   0
            Weight          =   900
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelWidth      =   2
         FloodColor      =   16777215
         Font3D          =   2
         Alignment       =   6
      End
      Begin Threed.SSPanel sspFilme 
         Height          =   465
         Index           =   0
         Left            =   1350
         TabIndex        =   30
         Top             =   1170
         Width           =   5535
         _Version        =   65536
         _ExtentX        =   9763
         _ExtentY        =   820
         _StockProps     =   15
         Caption         =   "XXXXXXXXXXXXXXXXXXXXXXXXXXXXX"
         ForeColor       =   16777215
         BackColor       =   16582188
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial Black"
            Size            =   12
            Charset         =   0
            Weight          =   900
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelWidth      =   2
         FloodColor      =   16777215
         Font3D          =   2
      End
      Begin Threed.SSPanel sspFilme 
         Height          =   465
         Index           =   1
         Left            =   1350
         TabIndex        =   31
         Top             =   1665
         Width           =   5535
         _Version        =   65536
         _ExtentX        =   9763
         _ExtentY        =   820
         _StockProps     =   15
         Caption         =   "XXXXXXXXXXXXXXXXXXXXXXXXXXXXX"
         ForeColor       =   16777215
         BackColor       =   16681289
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial Black"
            Size            =   12
            Charset         =   0
            Weight          =   900
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelWidth      =   2
         FloodColor      =   16777215
         Font3D          =   2
      End
      Begin Threed.SSPanel sspFilme 
         Height          =   465
         Index           =   2
         Left            =   1350
         TabIndex        =   32
         Top             =   2160
         Width           =   5535
         _Version        =   65536
         _ExtentX        =   9763
         _ExtentY        =   820
         _StockProps     =   15
         Caption         =   "XXXXXXXXXXXXXXXXXXXXXXXXXXXXX"
         ForeColor       =   16777215
         BackColor       =   16582188
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial Black"
            Size            =   12
            Charset         =   0
            Weight          =   900
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelWidth      =   2
         FloodColor      =   16777215
         Font3D          =   2
      End
      Begin Threed.SSPanel sspFilme 
         Height          =   465
         Index           =   3
         Left            =   1350
         TabIndex        =   33
         Top             =   2655
         Width           =   5535
         _Version        =   65536
         _ExtentX        =   9763
         _ExtentY        =   820
         _StockProps     =   15
         Caption         =   "XXXXXXXXXXXXXXXXXXXXXXXXXXXXX"
         ForeColor       =   16777215
         BackColor       =   16681289
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial Black"
            Size            =   12
            Charset         =   0
            Weight          =   900
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelWidth      =   2
         FloodColor      =   16777215
         Font3D          =   2
      End
      Begin Threed.SSPanel sspFilme 
         Height          =   465
         Index           =   4
         Left            =   1350
         TabIndex        =   34
         Top             =   3150
         Width           =   5535
         _Version        =   65536
         _ExtentX        =   9763
         _ExtentY        =   820
         _StockProps     =   15
         Caption         =   "XXXXXXXXXXXXXXXXXXXXXXXXXXXXX"
         ForeColor       =   16777215
         BackColor       =   16582188
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial Black"
            Size            =   12
            Charset         =   0
            Weight          =   900
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelWidth      =   2
         FloodColor      =   16777215
         Font3D          =   2
      End
      Begin Threed.SSPanel sspFilme 
         Height          =   465
         Index           =   5
         Left            =   1350
         TabIndex        =   35
         Top             =   3645
         Width           =   5535
         _Version        =   65536
         _ExtentX        =   9763
         _ExtentY        =   820
         _StockProps     =   15
         Caption         =   "XXXXXXXXXXXXXXXXXXXXXXXXXXXXX"
         ForeColor       =   16777215
         BackColor       =   16681289
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial Black"
            Size            =   12
            Charset         =   0
            Weight          =   900
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelWidth      =   2
         FloodColor      =   16777215
         Font3D          =   2
      End
      Begin Threed.SSPanel sspFilme 
         Height          =   465
         Index           =   6
         Left            =   1350
         TabIndex        =   36
         Top             =   4140
         Width           =   5535
         _Version        =   65536
         _ExtentX        =   9763
         _ExtentY        =   820
         _StockProps     =   15
         Caption         =   "XXXXXXXXXXXXXXXXXXXXXXXXXXXXX"
         ForeColor       =   16777215
         BackColor       =   16582188
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial Black"
            Size            =   12
            Charset         =   0
            Weight          =   900
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelWidth      =   2
         FloodColor      =   16777215
         Font3D          =   2
      End
      Begin Threed.SSPanel sspFilme 
         Height          =   465
         Index           =   7
         Left            =   1350
         TabIndex        =   37
         Top             =   4635
         Width           =   5535
         _Version        =   65536
         _ExtentX        =   9763
         _ExtentY        =   820
         _StockProps     =   15
         Caption         =   "XXXXXXXXXXXXXXXXXXXXXXXXXXXXX"
         ForeColor       =   16777215
         BackColor       =   16681289
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial Black"
            Size            =   12
            Charset         =   0
            Weight          =   900
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelWidth      =   2
         FloodColor      =   16777215
         Font3D          =   2
      End
      Begin Threed.SSPanel sspFilme 
         Height          =   465
         Index           =   8
         Left            =   1350
         TabIndex        =   38
         Top             =   5130
         Width           =   5535
         _Version        =   65536
         _ExtentX        =   9763
         _ExtentY        =   820
         _StockProps     =   15
         Caption         =   "XXXXXXXXXXXXXXXXXXXXXXXXXXXXX"
         ForeColor       =   16777215
         BackColor       =   16582188
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial Black"
            Size            =   12
            Charset         =   0
            Weight          =   900
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelWidth      =   2
         FloodColor      =   16777215
         Font3D          =   2
      End
      Begin Threed.SSPanel sspFilme 
         Height          =   465
         Index           =   9
         Left            =   1350
         TabIndex        =   39
         Top             =   5625
         Width           =   5535
         _Version        =   65536
         _ExtentX        =   9763
         _ExtentY        =   820
         _StockProps     =   15
         Caption         =   "XXXXXXXXXXXXXXXXXXXXXXXXXXXXX"
         ForeColor       =   16777215
         BackColor       =   16681289
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial Black"
            Size            =   12
            Charset         =   0
            Weight          =   900
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelWidth      =   2
         FloodColor      =   16777215
         Font3D          =   2
      End
      Begin Threed.SSPanel sspFilme 
         Height          =   465
         Index           =   10
         Left            =   1350
         TabIndex        =   40
         Top             =   6120
         Width           =   5535
         _Version        =   65536
         _ExtentX        =   9763
         _ExtentY        =   820
         _StockProps     =   15
         Caption         =   "XXXXXXXXXXXXXXXXXXXXXXXXXXXXX"
         ForeColor       =   16777215
         BackColor       =   16582188
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial Black"
            Size            =   12
            Charset         =   0
            Weight          =   900
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelWidth      =   2
         FloodColor      =   16777215
         Font3D          =   2
      End
      Begin Threed.SSPanel sspMensagem 
         Height          =   540
         Left            =   30
         TabIndex        =   41
         Top             =   6600
         Width           =   9540
         _Version        =   65536
         _ExtentX        =   16828
         _ExtentY        =   952
         _StockProps     =   15
         ForeColor       =   16777215
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial Black"
            Size            =   12
            Charset         =   0
            Weight          =   900
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelWidth      =   2
         FloodColor      =   16777215
         Font3D          =   2
         Begin VB.Label lblMensagem1 
            AutoSize        =   -1  'True
            Caption         =   "Label1"
            BeginProperty Font 
               Name            =   "Arial Black"
               Size            =   12
               Charset         =   0
               Weight          =   900
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   345
            Left            =   3000
            TabIndex        =   124
            Top             =   90
            Width           =   900
         End
      End
      Begin Threed.SSPanel sspTituloVenda 
         Height          =   465
         Left            =   7950
         TabIndex        =   97
         Top             =   675
         Width           =   1620
         _Version        =   65536
         _ExtentX        =   2857
         _ExtentY        =   820
         _StockProps     =   15
         ForeColor       =   16777215
         BackColor       =   4344827
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial Black"
            Size            =   12
            Charset         =   0
            Weight          =   900
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelWidth      =   2
         FloodColor      =   16777215
         Font3D          =   2
         Alignment       =   6
      End
      Begin Threed.SSPanel sspVenda 
         Height          =   465
         Index           =   0
         Left            =   7950
         TabIndex        =   98
         Top             =   1170
         Width           =   1620
         _Version        =   65536
         _ExtentX        =   2857
         _ExtentY        =   820
         _StockProps     =   15
         Caption         =   "A VENDA"
         ForeColor       =   16777215
         BackColor       =   16582188
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial Black"
            Size            =   14.25
            Charset         =   0
            Weight          =   900
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelWidth      =   2
         FloodColor      =   16777215
         Font3D          =   2
         Alignment       =   6
      End
      Begin Threed.SSPanel sspVenda 
         Height          =   465
         Index           =   1
         Left            =   7950
         TabIndex        =   99
         Top             =   1665
         Width           =   1620
         _Version        =   65536
         _ExtentX        =   2857
         _ExtentY        =   820
         _StockProps     =   15
         Caption         =   "LOTADO"
         ForeColor       =   16777215
         BackColor       =   255
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial Black"
            Size            =   14.25
            Charset         =   0
            Weight          =   900
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelWidth      =   2
         FloodColor      =   16777215
         Font3D          =   2
         Alignment       =   6
      End
      Begin Threed.SSPanel sspVenda 
         Height          =   465
         Index           =   2
         Left            =   7950
         TabIndex        =   100
         Top             =   2160
         Width           =   1620
         _Version        =   65536
         _ExtentX        =   2857
         _ExtentY        =   820
         _StockProps     =   15
         Caption         =   "VENDA"
         ForeColor       =   16777215
         BackColor       =   16582188
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial Black"
            Size            =   14.25
            Charset         =   0
            Weight          =   900
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelWidth      =   2
         FloodColor      =   16777215
         Font3D          =   2
         Alignment       =   6
      End
      Begin Threed.SSPanel sspVenda 
         Height          =   465
         Index           =   3
         Left            =   7950
         TabIndex        =   101
         Top             =   2655
         Width           =   1620
         _Version        =   65536
         _ExtentX        =   2857
         _ExtentY        =   820
         _StockProps     =   15
         Caption         =   "VENDA"
         ForeColor       =   16777215
         BackColor       =   16681289
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial Black"
            Size            =   14.25
            Charset         =   0
            Weight          =   900
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelWidth      =   2
         FloodColor      =   16777215
         Font3D          =   2
         Alignment       =   6
      End
      Begin Threed.SSPanel sspVenda 
         Height          =   465
         Index           =   4
         Left            =   7950
         TabIndex        =   102
         Top             =   3150
         Width           =   1620
         _Version        =   65536
         _ExtentX        =   2857
         _ExtentY        =   820
         _StockProps     =   15
         Caption         =   "VENDA"
         ForeColor       =   16777215
         BackColor       =   16582188
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial Black"
            Size            =   14.25
            Charset         =   0
            Weight          =   900
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelWidth      =   2
         FloodColor      =   16777215
         Font3D          =   2
         Alignment       =   6
      End
      Begin Threed.SSPanel sspVenda 
         Height          =   465
         Index           =   5
         Left            =   7950
         TabIndex        =   103
         Top             =   3645
         Width           =   1620
         _Version        =   65536
         _ExtentX        =   2857
         _ExtentY        =   820
         _StockProps     =   15
         Caption         =   "VENDA"
         ForeColor       =   16777215
         BackColor       =   16681289
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial Black"
            Size            =   14.25
            Charset         =   0
            Weight          =   900
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelWidth      =   2
         FloodColor      =   16777215
         Font3D          =   2
         Alignment       =   6
      End
      Begin Threed.SSPanel sspVenda 
         Height          =   465
         Index           =   6
         Left            =   7950
         TabIndex        =   104
         Top             =   4140
         Width           =   1620
         _Version        =   65536
         _ExtentX        =   2857
         _ExtentY        =   820
         _StockProps     =   15
         Caption         =   "VENDA"
         ForeColor       =   16777215
         BackColor       =   13453827
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial Black"
            Size            =   14.25
            Charset         =   0
            Weight          =   900
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelWidth      =   2
         FloodColor      =   16777215
         Font3D          =   2
         Alignment       =   6
      End
      Begin Threed.SSPanel sspVenda 
         Height          =   465
         Index           =   7
         Left            =   7950
         TabIndex        =   105
         Top             =   4635
         Width           =   1620
         _Version        =   65536
         _ExtentX        =   2857
         _ExtentY        =   820
         _StockProps     =   15
         Caption         =   "VENDA"
         ForeColor       =   16777215
         BackColor       =   16681289
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial Black"
            Size            =   14.25
            Charset         =   0
            Weight          =   900
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelWidth      =   2
         FloodColor      =   16777215
         Font3D          =   2
         Alignment       =   6
      End
      Begin Threed.SSPanel sspVenda 
         Height          =   465
         Index           =   8
         Left            =   7950
         TabIndex        =   106
         Top             =   5130
         Width           =   1620
         _Version        =   65536
         _ExtentX        =   2857
         _ExtentY        =   820
         _StockProps     =   15
         Caption         =   "VENDA"
         ForeColor       =   16777215
         BackColor       =   16582188
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial Black"
            Size            =   14.25
            Charset         =   0
            Weight          =   900
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelWidth      =   2
         FloodColor      =   16777215
         Font3D          =   2
         Alignment       =   6
      End
      Begin Threed.SSPanel sspVenda 
         Height          =   465
         Index           =   9
         Left            =   7950
         TabIndex        =   107
         Top             =   5625
         Width           =   1620
         _Version        =   65536
         _ExtentX        =   2857
         _ExtentY        =   820
         _StockProps     =   15
         Caption         =   "VENDA"
         ForeColor       =   16777215
         BackColor       =   16681289
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial Black"
            Size            =   14.25
            Charset         =   0
            Weight          =   900
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelWidth      =   2
         FloodColor      =   16777215
         Font3D          =   2
         Alignment       =   6
      End
      Begin Threed.SSPanel sspVenda 
         Height          =   465
         Index           =   10
         Left            =   7950
         TabIndex        =   108
         Top             =   6120
         Width           =   1620
         _Version        =   65536
         _ExtentX        =   2857
         _ExtentY        =   820
         _StockProps     =   15
         Caption         =   "VENDA"
         ForeColor       =   16777215
         BackColor       =   16582188
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial Black"
            Size            =   14.25
            Charset         =   0
            Weight          =   900
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelWidth      =   2
         FloodColor      =   16777215
         Font3D          =   2
         Alignment       =   6
      End
      Begin Threed.SSPanel sspHora 
         Height          =   615
         Left            =   15
         TabIndex        =   109
         Top             =   30
         Width           =   1320
         _Version        =   65536
         _ExtentX        =   2328
         _ExtentY        =   1085
         _StockProps     =   15
         Caption         =   "99:99"
         ForeColor       =   16777215
         BackColor       =   1319162
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial Black"
            Size            =   14.25
            Charset         =   0
            Weight          =   900
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelWidth      =   2
         FloodColor      =   16777215
         Font3D          =   2
      End
      Begin Threed.SSPanel sspData 
         Height          =   615
         Left            =   7950
         TabIndex        =   110
         Top             =   30
         Width           =   1620
         _Version        =   65536
         _ExtentX        =   2857
         _ExtentY        =   1085
         _StockProps     =   15
         Caption         =   "99/99/9999"
         ForeColor       =   16777215
         BackColor       =   1319162
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial Black"
            Size            =   12
            Charset         =   0
            Weight          =   900
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelWidth      =   2
         FloodColor      =   16777215
         Font3D          =   2
      End
   End
   Begin VB.Frame fraFilme 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H00000000&
      Height          =   1.08000e5
      Left            =   0
      MouseIcon       =   "frmApresent.frx":0E34
      MousePointer    =   4  'Icon
      TabIndex        =   2
      Top             =   0
      Width           =   1.43880e5
      Begin Threed.SSPanel sspSalaFil 
         Height          =   600
         Index           =   1
         Left            =   7275
         TabIndex        =   43
         Top             =   3960
         Width           =   2300
         _Version        =   65536
         _ExtentX        =   4057
         _ExtentY        =   1058
         _StockProps     =   15
         Caption         =   "SALA - 01"
         ForeColor       =   16777215
         BackColor       =   16681289
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial Black"
            Size            =   15.75
            Charset         =   0
            Weight          =   900
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelWidth      =   2
         FloodColor      =   16777215
         Font3D          =   2
      End
      Begin Threed.SSPanel sspHoraFilme 
         Height          =   675
         Left            =   30
         TabIndex        =   44
         Top             =   3255
         Width           =   9545
         _Version        =   65536
         _ExtentX        =   16836
         _ExtentY        =   1191
         _StockProps     =   15
         Caption         =   "99:99"
         ForeColor       =   16777215
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial Black"
            Size            =   20.25
            Charset         =   0
            Weight          =   900
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelWidth      =   2
         FloodColor      =   16777215
         Font3D          =   2
      End
      Begin Threed.SSPanel sspHrProxSessaoFilme 
         Height          =   1275
         Index           =   0
         Left            =   3690
         TabIndex        =   45
         Top             =   1950
         Width           =   3555
         _Version        =   65536
         _ExtentX        =   6271
         _ExtentY        =   2249
         _StockProps     =   15
         Caption         =   "99:99"
         ForeColor       =   16777215
         BackColor       =   16582188
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial Black"
            Size            =   36
            Charset         =   0
            Weight          =   900
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelWidth      =   2
         FloodColor      =   16777215
         Font3D          =   2
      End
      Begin Threed.SSPanel sspProxSessaoFilme 
         Height          =   420
         Index           =   0
         Left            =   3690
         TabIndex        =   46
         Top             =   1500
         Width           =   3555
         _Version        =   65536
         _ExtentX        =   6271
         _ExtentY        =   741
         _StockProps     =   15
         Caption         =   "PRÓXIMA SESSÃO"
         ForeColor       =   16777215
         BackColor       =   1319162
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial Black"
            Size            =   12
            Charset         =   0
            Weight          =   900
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelWidth      =   2
         FloodColor      =   16777215
         Font3D          =   2
      End
      Begin Threed.SSPanel sspAVenda 
         Height          =   840
         Index           =   0
         Left            =   30
         TabIndex        =   47
         Top             =   2385
         Width           =   3630
         _Version        =   65536
         _ExtentX        =   6403
         _ExtentY        =   1482
         _StockProps     =   15
         Caption         =   "A VENDA"
         ForeColor       =   16777215
         BackColor       =   1319162
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial Black"
            Size            =   14.25
            Charset         =   0
            Weight          =   900
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelWidth      =   2
         FloodColor      =   16777215
         Font3D          =   2
      End
      Begin Threed.SSPanel sspHrSessaoFilme 
         Height          =   405
         Index           =   5
         Left            =   8430
         TabIndex        =   50
         Top             =   1089
         Width           =   1140
         _Version        =   65536
         _ExtentX        =   2011
         _ExtentY        =   714
         _StockProps     =   15
         Caption         =   "99:99"
         ForeColor       =   16777215
         BackColor       =   16674851
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial Black"
            Size            =   12
            Charset         =   0
            Weight          =   900
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelWidth      =   2
         FloodColor      =   16777215
         Font3D          =   2
         Alignment       =   6
      End
      Begin Threed.SSPanel sspHrSessaoFilme 
         Height          =   405
         Index           =   6
         Left            =   8430
         TabIndex        =   51
         Top             =   1521
         Width           =   1140
         _Version        =   65536
         _ExtentX        =   2011
         _ExtentY        =   714
         _StockProps     =   15
         Caption         =   "99:99"
         ForeColor       =   16777215
         BackColor       =   13072967
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial Black"
            Size            =   12
            Charset         =   0
            Weight          =   900
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelWidth      =   2
         FloodColor      =   16777215
         Font3D          =   2
         Alignment       =   6
      End
      Begin Threed.SSPanel sspHrSessaoFilme 
         Height          =   405
         Index           =   7
         Left            =   8430
         TabIndex        =   52
         Top             =   1953
         Width           =   1140
         _Version        =   65536
         _ExtentX        =   2011
         _ExtentY        =   714
         _StockProps     =   15
         Caption         =   "99:99"
         ForeColor       =   16777215
         BackColor       =   16674851
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial Black"
            Size            =   12
            Charset         =   0
            Weight          =   900
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelWidth      =   2
         FloodColor      =   16777215
         Font3D          =   2
         Alignment       =   6
      End
      Begin Threed.SSPanel sspHrSessaoFilme 
         Height          =   405
         Index           =   8
         Left            =   8430
         TabIndex        =   53
         Top             =   2385
         Width           =   1140
         _Version        =   65536
         _ExtentX        =   2011
         _ExtentY        =   714
         _StockProps     =   15
         Caption         =   "99:99"
         ForeColor       =   16777215
         BackColor       =   13072967
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial Black"
            Size            =   12
            Charset         =   0
            Weight          =   900
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelWidth      =   2
         FloodColor      =   16777215
         Font3D          =   2
         Alignment       =   6
      End
      Begin Threed.SSPanel sspHrSessaoFilme 
         Height          =   405
         Index           =   9
         Left            =   8430
         TabIndex        =   54
         Top             =   2820
         Width           =   1140
         _Version        =   65536
         _ExtentX        =   2011
         _ExtentY        =   714
         _StockProps     =   15
         Caption         =   "99:99"
         ForeColor       =   16777215
         BackColor       =   16674851
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial Black"
            Size            =   12
            Charset         =   0
            Weight          =   900
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelWidth      =   2
         FloodColor      =   16777215
         Font3D          =   2
         Alignment       =   6
      End
      Begin Threed.SSPanel sspTitSessoes 
         Height          =   405
         Index           =   1
         Left            =   7275
         TabIndex        =   82
         Top             =   4587
         Width           =   2300
         _Version        =   65536
         _ExtentX        =   4057
         _ExtentY        =   714
         _StockProps     =   15
         Caption         =   "SESSÕES"
         ForeColor       =   16777215
         BackColor       =   4344827
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial Black"
            Size            =   12
            Charset         =   0
            Weight          =   900
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelWidth      =   2
         FloodColor      =   16777215
         Font3D          =   2
         Alignment       =   6
      End
      Begin Threed.SSPanel sspHrSessaoFilme 
         Height          =   405
         Index           =   10
         Left            =   7275
         TabIndex        =   83
         Top             =   5019
         Width           =   1140
         _Version        =   65536
         _ExtentX        =   2011
         _ExtentY        =   714
         _StockProps     =   15
         Caption         =   "99:99"
         ForeColor       =   16777215
         BackColor       =   16674851
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial Black"
            Size            =   12
            Charset         =   0
            Weight          =   900
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelWidth      =   2
         FloodColor      =   16777215
         Font3D          =   2
         Alignment       =   6
      End
      Begin Threed.SSPanel sspHrSessaoFilme 
         Height          =   405
         Index           =   11
         Left            =   7275
         TabIndex        =   84
         Top             =   5451
         Width           =   1140
         _Version        =   65536
         _ExtentX        =   2011
         _ExtentY        =   714
         _StockProps     =   15
         Caption         =   "99:99"
         ForeColor       =   16777215
         BackColor       =   13072967
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial Black"
            Size            =   12
            Charset         =   0
            Weight          =   900
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelWidth      =   2
         FloodColor      =   16777215
         Font3D          =   2
         Alignment       =   6
      End
      Begin Threed.SSPanel sspHrSessaoFilme 
         Height          =   405
         Index           =   12
         Left            =   7275
         TabIndex        =   85
         Top             =   5883
         Width           =   1140
         _Version        =   65536
         _ExtentX        =   2011
         _ExtentY        =   714
         _StockProps     =   15
         Caption         =   "99:99"
         ForeColor       =   16777215
         BackColor       =   16674851
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial Black"
            Size            =   12
            Charset         =   0
            Weight          =   900
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelWidth      =   2
         FloodColor      =   16777215
         Font3D          =   2
         Alignment       =   6
      End
      Begin Threed.SSPanel sspHrSessaoFilme 
         Height          =   405
         Index           =   13
         Left            =   7275
         TabIndex        =   86
         Top             =   6315
         Width           =   1140
         _Version        =   65536
         _ExtentX        =   2011
         _ExtentY        =   714
         _StockProps     =   15
         Caption         =   "99:99"
         ForeColor       =   16777215
         BackColor       =   13072967
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial Black"
            Size            =   12
            Charset         =   0
            Weight          =   900
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelWidth      =   2
         FloodColor      =   16777215
         Font3D          =   2
         Alignment       =   6
      End
      Begin Threed.SSPanel sspHrSessaoFilme 
         Height          =   405
         Index           =   14
         Left            =   7275
         TabIndex        =   87
         Top             =   6750
         Width           =   1140
         _Version        =   65536
         _ExtentX        =   2011
         _ExtentY        =   714
         _StockProps     =   15
         Caption         =   "99:99"
         ForeColor       =   16777215
         BackColor       =   16674851
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial Black"
            Size            =   12
            Charset         =   0
            Weight          =   900
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelWidth      =   2
         FloodColor      =   16777215
         Font3D          =   2
         Alignment       =   6
      End
      Begin Threed.SSPanel sspHrSessaoFilme 
         Height          =   405
         Index           =   15
         Left            =   8430
         TabIndex        =   88
         Top             =   5019
         Width           =   1140
         _Version        =   65536
         _ExtentX        =   2011
         _ExtentY        =   714
         _StockProps     =   15
         Caption         =   "99:99"
         ForeColor       =   16777215
         BackColor       =   16674851
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial Black"
            Size            =   12
            Charset         =   0
            Weight          =   900
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelWidth      =   2
         FloodColor      =   16777215
         Font3D          =   2
         Alignment       =   6
      End
      Begin Threed.SSPanel sspHrSessaoFilme 
         Height          =   405
         Index           =   16
         Left            =   8430
         TabIndex        =   89
         Top             =   5451
         Width           =   1140
         _Version        =   65536
         _ExtentX        =   2011
         _ExtentY        =   714
         _StockProps     =   15
         Caption         =   "99:99"
         ForeColor       =   16777215
         BackColor       =   13072967
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial Black"
            Size            =   12
            Charset         =   0
            Weight          =   900
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelWidth      =   2
         FloodColor      =   16777215
         Font3D          =   2
         Alignment       =   6
      End
      Begin Threed.SSPanel sspHrSessaoFilme 
         Height          =   405
         Index           =   17
         Left            =   8430
         TabIndex        =   90
         Top             =   5883
         Width           =   1140
         _Version        =   65536
         _ExtentX        =   2011
         _ExtentY        =   714
         _StockProps     =   15
         Caption         =   "99:99"
         ForeColor       =   16777215
         BackColor       =   16674851
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial Black"
            Size            =   12
            Charset         =   0
            Weight          =   900
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelWidth      =   2
         FloodColor      =   16777215
         Font3D          =   2
         Alignment       =   6
      End
      Begin Threed.SSPanel sspHrSessaoFilme 
         Height          =   405
         Index           =   18
         Left            =   8430
         TabIndex        =   91
         Top             =   6315
         Width           =   1140
         _Version        =   65536
         _ExtentX        =   2011
         _ExtentY        =   714
         _StockProps     =   15
         Caption         =   "99:99"
         ForeColor       =   16777215
         BackColor       =   13072967
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial Black"
            Size            =   12
            Charset         =   0
            Weight          =   900
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelWidth      =   2
         FloodColor      =   16777215
         Font3D          =   2
         Alignment       =   6
      End
      Begin Threed.SSPanel sspHrSessaoFilme 
         Height          =   405
         Index           =   19
         Left            =   8430
         TabIndex        =   92
         Top             =   6750
         Width           =   1140
         _Version        =   65536
         _ExtentX        =   2011
         _ExtentY        =   714
         _StockProps     =   15
         Caption         =   "99:99"
         ForeColor       =   16777215
         BackColor       =   16674851
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial Black"
            Size            =   12
            Charset         =   0
            Weight          =   900
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelWidth      =   2
         FloodColor      =   16777215
         Font3D          =   2
         Alignment       =   6
      End
      Begin Threed.SSPanel sspDescFilme 
         Height          =   1440
         Index           =   1
         Left            =   30
         TabIndex        =   93
         Top             =   3960
         Width           =   7215
         _Version        =   65536
         _ExtentX        =   12726
         _ExtentY        =   2540
         _StockProps     =   15
         Caption         =   "Rei Leão"
         ForeColor       =   16777215
         BackColor       =   16681289
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial Black"
            Size            =   20.25
            Charset         =   0
            Weight          =   900
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelWidth      =   2
         FloodColor      =   16777215
         Font3D          =   2
      End
      Begin Threed.SSPanel sspHrProxSessaoFilme 
         Height          =   1275
         Index           =   1
         Left            =   3690
         TabIndex        =   94
         Top             =   5880
         Width           =   3555
         _Version        =   65536
         _ExtentX        =   6271
         _ExtentY        =   2249
         _StockProps     =   15
         Caption         =   "99:99"
         ForeColor       =   16777215
         BackColor       =   16681289
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial Black"
            Size            =   36
            Charset         =   0
            Weight          =   900
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelWidth      =   2
         FloodColor      =   16777215
         Font3D          =   2
      End
      Begin Threed.SSPanel sspProxSessaoFilme 
         Height          =   420
         Index           =   1
         Left            =   3690
         TabIndex        =   95
         Top             =   5430
         Width           =   3555
         _Version        =   65536
         _ExtentX        =   6271
         _ExtentY        =   741
         _StockProps     =   15
         Caption         =   "PRÓXIMA SESSÃO"
         ForeColor       =   16777215
         BackColor       =   4344827
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial Black"
            Size            =   12
            Charset         =   0
            Weight          =   900
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelWidth      =   2
         FloodColor      =   16777215
         Font3D          =   2
      End
      Begin Threed.SSPanel sspAVenda 
         Height          =   840
         Index           =   1
         Left            =   30
         TabIndex        =   96
         Top             =   6315
         Width           =   3630
         _Version        =   65536
         _ExtentX        =   6403
         _ExtentY        =   1482
         _StockProps     =   15
         Caption         =   "A VENDA"
         ForeColor       =   16777215
         BackColor       =   4344827
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial Black"
            Size            =   14.25
            Charset         =   0
            Weight          =   900
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelWidth      =   2
         FloodColor      =   16777215
         Font3D          =   2
      End
      Begin Threed.SSPanel sspCensura 
         Height          =   855
         Index           =   0
         Left            =   30
         TabIndex        =   111
         Top             =   1500
         Width           =   3630
         _Version        =   65536
         _ExtentX        =   6403
         _ExtentY        =   1508
         _StockProps     =   15
         Caption         =   "Censura Livre"
         ForeColor       =   16777215
         BackColor       =   16582188
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial Black"
            Size            =   14.25
            Charset         =   0
            Weight          =   900
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelWidth      =   2
         FloodColor      =   16777215
         Font3D          =   2
      End
      Begin Threed.SSPanel sspCensura 
         Height          =   855
         Index           =   1
         Left            =   30
         TabIndex        =   112
         Top             =   5430
         Width           =   3630
         _Version        =   65536
         _ExtentX        =   6403
         _ExtentY        =   1508
         _StockProps     =   15
         Caption         =   "Censura Livre"
         ForeColor       =   16777215
         BackColor       =   16681289
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial Black"
            Size            =   14.25
            Charset         =   0
            Weight          =   900
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelWidth      =   2
         FloodColor      =   16777215
         Font3D          =   2
      End
      Begin Threed.SSPanel sspTitSessoes 
         Height          =   405
         Index           =   0
         Left            =   7275
         TabIndex        =   117
         Top             =   657
         Width           =   2300
         _Version        =   65536
         _ExtentX        =   4057
         _ExtentY        =   714
         _StockProps     =   15
         Caption         =   "SESSÕES"
         ForeColor       =   16777215
         BackColor       =   1319162
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial Black"
            Size            =   12
            Charset         =   0
            Weight          =   900
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelWidth      =   2
         FloodColor      =   16777215
         Font3D          =   2
         Alignment       =   6
      End
      Begin Threed.SSPanel sspHrSessaoFilme 
         Height          =   405
         Index           =   0
         Left            =   7275
         TabIndex        =   118
         Top             =   1089
         Width           =   1140
         _Version        =   65536
         _ExtentX        =   2011
         _ExtentY        =   714
         _StockProps     =   15
         Caption         =   "99:99"
         ForeColor       =   16777215
         BackColor       =   16674851
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial Black"
            Size            =   12
            Charset         =   0
            Weight          =   900
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelWidth      =   2
         FloodColor      =   16777215
         Font3D          =   2
         Alignment       =   6
      End
      Begin Threed.SSPanel sspHrSessaoFilme 
         Height          =   405
         Index           =   1
         Left            =   7275
         TabIndex        =   119
         Top             =   1521
         Width           =   1140
         _Version        =   65536
         _ExtentX        =   2011
         _ExtentY        =   714
         _StockProps     =   15
         Caption         =   "99:99"
         ForeColor       =   16777215
         BackColor       =   13072967
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial Black"
            Size            =   12
            Charset         =   0
            Weight          =   900
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelWidth      =   2
         FloodColor      =   16777215
         Font3D          =   2
         Alignment       =   6
      End
      Begin Threed.SSPanel sspHrSessaoFilme 
         Height          =   405
         Index           =   2
         Left            =   7275
         TabIndex        =   120
         Top             =   1953
         Width           =   1140
         _Version        =   65536
         _ExtentX        =   2011
         _ExtentY        =   714
         _StockProps     =   15
         Caption         =   "99:99"
         ForeColor       =   16777215
         BackColor       =   16674851
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial Black"
            Size            =   12
            Charset         =   0
            Weight          =   900
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelWidth      =   2
         FloodColor      =   16777215
         Font3D          =   2
         Alignment       =   6
      End
      Begin Threed.SSPanel sspHrSessaoFilme 
         Height          =   405
         Index           =   3
         Left            =   7275
         TabIndex        =   121
         Top             =   2385
         Width           =   1140
         _Version        =   65536
         _ExtentX        =   2011
         _ExtentY        =   714
         _StockProps     =   15
         Caption         =   "99:99"
         ForeColor       =   16777215
         BackColor       =   13072967
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial Black"
            Size            =   12
            Charset         =   0
            Weight          =   900
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelWidth      =   2
         FloodColor      =   16777215
         Font3D          =   2
         Alignment       =   6
      End
      Begin Threed.SSPanel sspHrSessaoFilme 
         Height          =   405
         Index           =   4
         Left            =   7275
         TabIndex        =   122
         Top             =   2820
         Width           =   1140
         _Version        =   65536
         _ExtentX        =   2011
         _ExtentY        =   714
         _StockProps     =   15
         Caption         =   "99:99"
         ForeColor       =   16777215
         BackColor       =   16674851
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial Black"
            Size            =   12
            Charset         =   0
            Weight          =   900
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelWidth      =   2
         FloodColor      =   16777215
         Font3D          =   2
         Alignment       =   6
      End
      Begin Threed.SSPanel sspSalaFil 
         Height          =   600
         Index           =   0
         Left            =   7275
         TabIndex        =   123
         Top             =   30
         Width           =   2300
         _Version        =   65536
         _ExtentX        =   4057
         _ExtentY        =   1058
         _StockProps     =   15
         Caption         =   "SALA - 01"
         ForeColor       =   16777215
         BackColor       =   16582188
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial Black"
            Size            =   15.75
            Charset         =   0
            Weight          =   900
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelWidth      =   2
         FloodColor      =   16777215
         Font3D          =   2
      End
      Begin Threed.SSPanel sspDescFilme 
         Height          =   1440
         Index           =   0
         Left            =   30
         TabIndex        =   42
         Top             =   30
         Width           =   7215
         _Version        =   65536
         _ExtentX        =   12726
         _ExtentY        =   2540
         _StockProps     =   15
         Caption         =   "Rei Leão"
         ForeColor       =   16777215
         BackColor       =   16582188
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial Black"
            Size            =   20.25
            Charset         =   0
            Weight          =   900
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelWidth      =   2
         FloodColor      =   16777215
         Font3D          =   2
      End
   End
   Begin VB.Frame fraPrecos 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H00000000&
      Height          =   1.08000e5
      Left            =   0
      MouseIcon       =   "frmApresent.frx":11BE
      MousePointer    =   4  'Icon
      TabIndex        =   0
      Top             =   0
      Width           =   1.44000e5
      Begin Threed.SSPanel sspFilmePrecos 
         Height          =   480
         Left            =   30
         TabIndex        =   48
         Top             =   30
         Width           =   9540
         _Version        =   65536
         _ExtentX        =   16828
         _ExtentY        =   847
         _StockProps     =   15
         Caption         =   "Filme"
         ForeColor       =   16777215
         BackColor       =   1319162
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial Black"
            Size            =   14.25
            Charset         =   0
            Weight          =   900
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelWidth      =   2
         FloodColor      =   16777215
         Font3D          =   2
         Alignment       =   6
      End
      Begin Threed.SSPanel sspDiasSemana 
         Height          =   1140
         Index           =   0
         Left            =   30
         TabIndex        =   49
         Top             =   1470
         Width           =   2655
         _Version        =   65536
         _ExtentX        =   4683
         _ExtentY        =   2011
         _StockProps     =   15
         Caption         =   "Seg, Ter, Qua, Qui, Sex, Sab, Dom, Fer"
         ForeColor       =   16777215
         BackColor       =   13453827
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial Black"
            Size            =   12
            Charset         =   0
            Weight          =   900
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelWidth      =   2
         FloodColor      =   16777215
         Font3D          =   2
         Alignment       =   6
         Begin Threed.SSPanel sspProcional 
            Height          =   375
            Index           =   0
            Left            =   45
            TabIndex        =   113
            Top             =   720
            Width           =   2565
            _Version        =   65536
            _ExtentX        =   4524
            _ExtentY        =   661
            _StockProps     =   15
            Caption         =   "Promocional"
            ForeColor       =   16777215
            BackColor       =   16744576
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial Black"
               Size            =   14.25
               Charset         =   0
               Weight          =   900
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BevelWidth      =   2
            BevelOuter      =   0
            FloodColor      =   16777215
            Font3D          =   2
            Alignment       =   6
         End
      End
      Begin Threed.SSPanel sspInteira 
         Height          =   435
         Index           =   0
         Left            =   2700
         TabIndex        =   55
         Top             =   1005
         Width           =   1695
         _Version        =   65536
         _ExtentX        =   2990
         _ExtentY        =   767
         _StockProps     =   15
         Caption         =   "INTEIRA"
         ForeColor       =   16777215
         BackColor       =   4344827
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial Black"
            Size            =   12
            Charset         =   0
            Weight          =   900
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelWidth      =   2
         FloodColor      =   16777215
         Font3D          =   2
         Alignment       =   6
      End
      Begin Threed.SSPanel sspMeia 
         Height          =   435
         Index           =   0
         Left            =   4410
         TabIndex        =   56
         Top             =   1005
         Width           =   1710
         _Version        =   65536
         _ExtentX        =   3016
         _ExtentY        =   767
         _StockProps     =   15
         Caption         =   "MEIA"
         ForeColor       =   16777215
         BackColor       =   4344827
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial Black"
            Size            =   12
            Charset         =   0
            Weight          =   900
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelWidth      =   2
         FloodColor      =   16777215
         Font3D          =   2
         Alignment       =   6
      End
      Begin Threed.SSPanel sspInteira 
         Height          =   435
         Index           =   1
         Left            =   6135
         TabIndex        =   57
         Top             =   1005
         Width           =   1710
         _Version        =   65536
         _ExtentX        =   3016
         _ExtentY        =   767
         _StockProps     =   15
         Caption         =   "INTEIRA"
         ForeColor       =   16777215
         BackColor       =   4344827
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial Black"
            Size            =   12
            Charset         =   0
            Weight          =   900
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelWidth      =   2
         FloodColor      =   16777215
         Font3D          =   2
         Alignment       =   6
      End
      Begin Threed.SSPanel sspMeia 
         Height          =   435
         Index           =   1
         Left            =   7860
         TabIndex        =   58
         Top             =   1005
         Width           =   1710
         _Version        =   65536
         _ExtentX        =   3016
         _ExtentY        =   767
         _StockProps     =   15
         Caption         =   "MEIA"
         ForeColor       =   16777215
         BackColor       =   4344827
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial Black"
            Size            =   12
            Charset         =   0
            Weight          =   900
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelWidth      =   2
         FloodColor      =   16777215
         Font3D          =   2
         Alignment       =   6
      End
      Begin Threed.SSPanel sspHoraPreco 
         Height          =   900
         Left            =   30
         TabIndex        =   59
         Top             =   540
         Width           =   2655
         _Version        =   65536
         _ExtentX        =   4683
         _ExtentY        =   1587
         _StockProps     =   15
         Caption         =   "99:00"
         ForeColor       =   16777215
         BackColor       =   1319162
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial Black"
            Size            =   24
            Charset         =   0
            Weight          =   900
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelWidth      =   2
         FloodColor      =   16777215
         Font3D          =   2
      End
      Begin Threed.SSPanel sspManha 
         Height          =   435
         Left            =   2700
         TabIndex        =   60
         Top             =   540
         Width           =   3420
         _Version        =   65536
         _ExtentX        =   6032
         _ExtentY        =   767
         _StockProps     =   15
         Caption         =   "Até as 17:00"
         ForeColor       =   16777215
         BackColor       =   1319162
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial Black"
            Size            =   14.25
            Charset         =   0
            Weight          =   900
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelWidth      =   2
         FloodColor      =   16777215
         Font3D          =   2
      End
      Begin Threed.SSPanel sspTarde 
         Height          =   435
         Left            =   6135
         TabIndex        =   61
         Top             =   540
         Width           =   3435
         _Version        =   65536
         _ExtentX        =   6059
         _ExtentY        =   767
         _StockProps     =   15
         Caption         =   "Após as 17:00"
         ForeColor       =   16777215
         BackColor       =   1319162
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial Black"
            Size            =   14.25
            Charset         =   0
            Weight          =   900
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelWidth      =   2
         FloodColor      =   16777215
         Font3D          =   2
      End
      Begin Threed.SSPanel sspIntManha 
         Height          =   1140
         Index           =   0
         Left            =   2700
         TabIndex        =   62
         Top             =   1470
         Width           =   1695
         _Version        =   65536
         _ExtentX        =   2990
         _ExtentY        =   2011
         _StockProps     =   15
         Caption         =   "R$ 999,00"
         ForeColor       =   16777215
         BackColor       =   13453827
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial Black"
            Size            =   14.25
            Charset         =   0
            Weight          =   900
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelWidth      =   2
         FloodColor      =   16777215
         Font3D          =   2
      End
      Begin Threed.SSPanel sspMeiaManha 
         Height          =   1140
         Index           =   0
         Left            =   4410
         TabIndex        =   63
         Top             =   1470
         Width           =   1710
         _Version        =   65536
         _ExtentX        =   3016
         _ExtentY        =   2011
         _StockProps     =   15
         Caption         =   "R$ 999,00"
         ForeColor       =   16777215
         BackColor       =   13453827
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial Black"
            Size            =   14.25
            Charset         =   0
            Weight          =   900
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelWidth      =   2
         FloodColor      =   16777215
         Font3D          =   2
      End
      Begin Threed.SSPanel sspIntTarde 
         Height          =   1140
         Index           =   0
         Left            =   6135
         TabIndex        =   64
         Top             =   1470
         Width           =   1710
         _Version        =   65536
         _ExtentX        =   3016
         _ExtentY        =   2011
         _StockProps     =   15
         Caption         =   "R$ 999,00"
         ForeColor       =   16777215
         BackColor       =   13453827
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial Black"
            Size            =   14.25
            Charset         =   0
            Weight          =   900
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelWidth      =   2
         FloodColor      =   16777215
         Font3D          =   2
      End
      Begin Threed.SSPanel sspMeiaTarde 
         Height          =   1140
         Index           =   0
         Left            =   7860
         TabIndex        =   65
         Top             =   1470
         Width           =   1710
         _Version        =   65536
         _ExtentX        =   3016
         _ExtentY        =   2011
         _StockProps     =   15
         Caption         =   "R$ 999,00"
         ForeColor       =   16777215
         BackColor       =   13453827
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial Black"
            Size            =   14.25
            Charset         =   0
            Weight          =   900
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelWidth      =   2
         FloodColor      =   16777215
         Font3D          =   2
      End
      Begin Threed.SSPanel sspDiasSemana 
         Height          =   1140
         Index           =   1
         Left            =   30
         TabIndex        =   66
         Top             =   2640
         Width           =   2655
         _Version        =   65536
         _ExtentX        =   4683
         _ExtentY        =   2011
         _StockProps     =   15
         Caption         =   "Seg, Ter, Qua"
         ForeColor       =   16777215
         BackColor       =   13453827
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial Black"
            Size            =   12
            Charset         =   0
            Weight          =   900
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelWidth      =   2
         FloodColor      =   16777215
         Font3D          =   2
         Alignment       =   6
         Begin Threed.SSPanel sspProcional 
            Height          =   375
            Index           =   1
            Left            =   45
            TabIndex        =   114
            Top             =   720
            Width           =   2565
            _Version        =   65536
            _ExtentX        =   4524
            _ExtentY        =   661
            _StockProps     =   15
            Caption         =   "Promocional"
            ForeColor       =   16777215
            BackColor       =   16744576
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial Black"
               Size            =   14.25
               Charset         =   0
               Weight          =   900
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BevelWidth      =   2
            BevelOuter      =   0
            FloodColor      =   16777215
            Font3D          =   2
            Alignment       =   6
         End
      End
      Begin Threed.SSPanel sspIntManha 
         Height          =   1140
         Index           =   1
         Left            =   2700
         TabIndex        =   67
         Top             =   2640
         Width           =   1695
         _Version        =   65536
         _ExtentX        =   2990
         _ExtentY        =   2011
         _StockProps     =   15
         Caption         =   "R$ 999,00"
         ForeColor       =   16777215
         BackColor       =   13453827
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial Black"
            Size            =   14.25
            Charset         =   0
            Weight          =   900
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelWidth      =   2
         FloodColor      =   16777215
         Font3D          =   2
      End
      Begin Threed.SSPanel sspMeiaManha 
         Height          =   1140
         Index           =   1
         Left            =   4410
         TabIndex        =   68
         Top             =   2640
         Width           =   1710
         _Version        =   65536
         _ExtentX        =   3016
         _ExtentY        =   2011
         _StockProps     =   15
         Caption         =   "R$ 999,00"
         ForeColor       =   16777215
         BackColor       =   13453827
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial Black"
            Size            =   14.25
            Charset         =   0
            Weight          =   900
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelWidth      =   2
         FloodColor      =   16777215
         Font3D          =   2
      End
      Begin Threed.SSPanel sspIntTarde 
         Height          =   1140
         Index           =   1
         Left            =   6135
         TabIndex        =   69
         Top             =   2640
         Width           =   1710
         _Version        =   65536
         _ExtentX        =   3016
         _ExtentY        =   2011
         _StockProps     =   15
         Caption         =   "R$ 999,00"
         ForeColor       =   16777215
         BackColor       =   13453827
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial Black"
            Size            =   14.25
            Charset         =   0
            Weight          =   900
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelWidth      =   2
         FloodColor      =   16777215
         Font3D          =   2
      End
      Begin Threed.SSPanel sspMeiaTarde 
         Height          =   1140
         Index           =   1
         Left            =   7860
         TabIndex        =   70
         Top             =   2640
         Width           =   1710
         _Version        =   65536
         _ExtentX        =   3016
         _ExtentY        =   2011
         _StockProps     =   15
         Caption         =   "R$ 999,00"
         ForeColor       =   16777215
         BackColor       =   13453827
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial Black"
            Size            =   14.25
            Charset         =   0
            Weight          =   900
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelWidth      =   2
         FloodColor      =   16777215
         Font3D          =   2
      End
      Begin Threed.SSPanel sspDiasSemana 
         Height          =   1140
         Index           =   2
         Left            =   30
         TabIndex        =   71
         Top             =   3810
         Width           =   2655
         _Version        =   65536
         _ExtentX        =   4683
         _ExtentY        =   2011
         _StockProps     =   15
         Caption         =   "Seg, Ter, Qua, Qui, Sex, Sab, Dom, Fer"
         ForeColor       =   16777215
         BackColor       =   13453827
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial Black"
            Size            =   12
            Charset         =   0
            Weight          =   900
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelWidth      =   2
         FloodColor      =   16777215
         Font3D          =   2
         Alignment       =   6
         Begin Threed.SSPanel sspProcional 
            Height          =   375
            Index           =   2
            Left            =   45
            TabIndex        =   115
            Top             =   720
            Width           =   2565
            _Version        =   65536
            _ExtentX        =   4524
            _ExtentY        =   661
            _StockProps     =   15
            Caption         =   "Promocional"
            ForeColor       =   16777215
            BackColor       =   16744576
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial Black"
               Size            =   14.25
               Charset         =   0
               Weight          =   900
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BevelWidth      =   2
            BevelOuter      =   0
            FloodColor      =   16777215
            Font3D          =   2
            Alignment       =   6
         End
      End
      Begin Threed.SSPanel sspIntManha 
         Height          =   1140
         Index           =   2
         Left            =   2700
         TabIndex        =   72
         Top             =   3810
         Width           =   1695
         _Version        =   65536
         _ExtentX        =   2990
         _ExtentY        =   2011
         _StockProps     =   15
         Caption         =   "R$ 999,00"
         ForeColor       =   16777215
         BackColor       =   13453827
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial Black"
            Size            =   14.25
            Charset         =   0
            Weight          =   900
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelWidth      =   2
         FloodColor      =   16777215
         Font3D          =   2
      End
      Begin Threed.SSPanel sspMeiaManha 
         Height          =   1140
         Index           =   2
         Left            =   4410
         TabIndex        =   73
         Top             =   3810
         Width           =   1710
         _Version        =   65536
         _ExtentX        =   3016
         _ExtentY        =   2011
         _StockProps     =   15
         Caption         =   "R$ 999,00"
         ForeColor       =   16777215
         BackColor       =   13453827
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial Black"
            Size            =   14.25
            Charset         =   0
            Weight          =   900
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelWidth      =   2
         FloodColor      =   16777215
         Font3D          =   2
      End
      Begin Threed.SSPanel sspIntTarde 
         Height          =   1140
         Index           =   2
         Left            =   6135
         TabIndex        =   74
         Top             =   3810
         Width           =   1710
         _Version        =   65536
         _ExtentX        =   3016
         _ExtentY        =   2011
         _StockProps     =   15
         Caption         =   "R$ 999,00"
         ForeColor       =   16777215
         BackColor       =   13453827
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial Black"
            Size            =   14.25
            Charset         =   0
            Weight          =   900
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelWidth      =   2
         FloodColor      =   16777215
         Font3D          =   2
      End
      Begin Threed.SSPanel sspMeiaTarde 
         Height          =   1140
         Index           =   2
         Left            =   7860
         TabIndex        =   75
         Top             =   3810
         Width           =   1710
         _Version        =   65536
         _ExtentX        =   3016
         _ExtentY        =   2011
         _StockProps     =   15
         Caption         =   "R$ 999,00"
         ForeColor       =   16777215
         BackColor       =   13453827
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial Black"
            Size            =   14.25
            Charset         =   0
            Weight          =   900
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelWidth      =   2
         FloodColor      =   16777215
         Font3D          =   2
      End
      Begin Threed.SSPanel sspDiasSemana 
         Height          =   1140
         Index           =   3
         Left            =   30
         TabIndex        =   76
         Top             =   4980
         Width           =   2655
         _Version        =   65536
         _ExtentX        =   4683
         _ExtentY        =   2011
         _StockProps     =   15
         Caption         =   "Seg, Ter, Qua, Qui, Sex, Sab, Dom, Fer"
         ForeColor       =   16777215
         BackColor       =   13453827
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial Black"
            Size            =   12
            Charset         =   0
            Weight          =   900
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelWidth      =   2
         FloodColor      =   16777215
         Font3D          =   2
         Alignment       =   6
         Begin Threed.SSPanel sspProcional 
            Height          =   375
            Index           =   3
            Left            =   45
            TabIndex        =   116
            Top             =   720
            Width           =   2565
            _Version        =   65536
            _ExtentX        =   4524
            _ExtentY        =   661
            _StockProps     =   15
            Caption         =   "Promocional"
            ForeColor       =   16777215
            BackColor       =   16744576
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial Black"
               Size            =   14.25
               Charset         =   0
               Weight          =   900
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BevelWidth      =   2
            BevelOuter      =   0
            FloodColor      =   16777215
            Font3D          =   2
            Alignment       =   6
         End
      End
      Begin Threed.SSPanel sspIntManha 
         Height          =   1140
         Index           =   3
         Left            =   2700
         TabIndex        =   77
         Top             =   4980
         Width           =   1695
         _Version        =   65536
         _ExtentX        =   2990
         _ExtentY        =   2011
         _StockProps     =   15
         Caption         =   "R$ 999,00"
         ForeColor       =   16777215
         BackColor       =   13453827
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial Black"
            Size            =   14.25
            Charset         =   0
            Weight          =   900
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelWidth      =   2
         FloodColor      =   16777215
         Font3D          =   2
      End
      Begin Threed.SSPanel sspMeiaManha 
         Height          =   1140
         Index           =   3
         Left            =   4410
         TabIndex        =   78
         Top             =   4980
         Width           =   1710
         _Version        =   65536
         _ExtentX        =   3016
         _ExtentY        =   2011
         _StockProps     =   15
         Caption         =   "R$ 999,00"
         ForeColor       =   16777215
         BackColor       =   13453827
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial Black"
            Size            =   14.25
            Charset         =   0
            Weight          =   900
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelWidth      =   2
         FloodColor      =   16777215
         Font3D          =   2
      End
      Begin Threed.SSPanel sspIntTarde 
         Height          =   1140
         Index           =   3
         Left            =   6135
         TabIndex        =   79
         Top             =   4980
         Width           =   1710
         _Version        =   65536
         _ExtentX        =   3016
         _ExtentY        =   2011
         _StockProps     =   15
         Caption         =   "R$ 999,00"
         ForeColor       =   16777215
         BackColor       =   13453827
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial Black"
            Size            =   14.25
            Charset         =   0
            Weight          =   900
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelWidth      =   2
         FloodColor      =   16777215
         Font3D          =   2
      End
      Begin Threed.SSPanel sspMeiaTarde 
         Height          =   1140
         Index           =   3
         Left            =   7860
         TabIndex        =   80
         Top             =   4980
         Width           =   1710
         _Version        =   65536
         _ExtentX        =   3016
         _ExtentY        =   2011
         _StockProps     =   15
         Caption         =   "R$ 999,00"
         ForeColor       =   16777215
         BackColor       =   13453827
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial Black"
            Size            =   14.25
            Charset         =   0
            Weight          =   900
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelWidth      =   2
         FloodColor      =   16777215
         Font3D          =   2
      End
      Begin Threed.SSPanel sspMensagemPreco 
         Height          =   1005
         Left            =   30
         TabIndex        =   81
         Top             =   6150
         Width           =   9540
         _Version        =   65536
         _ExtentX        =   16828
         _ExtentY        =   1773
         _StockProps     =   15
         ForeColor       =   16777215
         BackColor       =   13453827
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial Black"
            Size            =   12
            Charset         =   0
            Weight          =   900
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelWidth      =   2
         FloodColor      =   16777215
         Font3D          =   2
         Begin VB.Label lblMensagemPreco1 
            AutoSize        =   -1  'True
            Caption         =   "Label1"
            BeginProperty Font 
               Name            =   "Arial Black"
               Size            =   12
               Charset         =   0
               Weight          =   900
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   345
            Left            =   3000
            TabIndex        =   125
            Top             =   330
            Width           =   900
         End
      End
   End
End
Attribute VB_Name = "frmApresent"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'SLS Resize do form 16/09/2011

Dim MyForm As FRMSIZE
Dim DesignX As Integer
Dim DesignY As Integer

Dim proxFilme   As Integer
Dim proxPreco   As Integer
Dim proxTrailer As Integer
Dim proxImagem  As Integer
Dim proxTela    As Integer
Dim telaAtu     As Integer
Dim abriuArq    As Boolean

Dim ctrlsPisca1()   As Control
Dim strCtrlPisca1() As String
Dim bBranco1        As Boolean

Dim ctrlsPisca2()   As Control
Dim strCtrlPisca2() As String
Dim bBranco2        As Boolean

Dim ctrlsMsg() As Control
Dim wAtualiza As Boolean

Public iqtdAtualiza As Integer


Private Sub preparaProxSessoes()
   Dim I         As Integer
   Dim j         As Integer
   Dim ini       As Integer
   Dim horaAgora As Date
   Dim corFundo  As Long
   Dim corTexto  As Long
   
'   horaAgora = CDate(Format(Now, "Hh:Nn"))
   
   If CDate(Format(Now, "Hh:Nn")) > CDate("03:00") Then
      horaAgora = CDate(dtStrRef1 & " " & Format(Now, "Hh:Nn"))
   Else
      horaAgora = CDate(dtStrRef2 & " " & Format(Now, "Hh:Nn"))
   End If
   
   ReDim ctrlsPisca2(0 To 0) As Control
   ReDim strCtrlPisca2(0 To 0) As String
   bBranco2 = False
   
   'Tela 3
   sspHora.Caption = Format(Now, "Hh:Nn")
   sspData.Caption = Format(Now, "DD/MM/YYYY")
   lblMensagem1.Caption = mensagem
   lblMensagem1.Left = -1 * lblMensagem1.Width + -2
   
   '0 -> 10
   For I = 0 To 10
      sspSessao(I).Caption = ""
      sspFilme(I).Caption = ""
      sspSala(I).Caption = ""
      sspVenda(I).Caption = ""
   
      sspSessao(I).Visible = False
      sspFilme(I).Visible = False
      sspSala(I).Visible = False
      sspVenda(I).Visible = False
   
      If I Mod 2 = 0 Then
         corFundo = corFundT3Lin1
         corTexto = corTextT3Lin1
      Else
         corFundo = corFundT3Lin2
         corTexto = corTextT3Lin2
      End If
      
      sspVenda(I).BackColor = corFundo
      sspVenda(I).ForeColor = corTexto
   Next I

   ini = -1
   For I = LBound(proxSessoes) To UBound(proxSessoes)
      If proxSessoes(I).horario >= DateAdd("n", -1 * vendaDepois, horaAgora) Then
         ini = I
         Exit For
      End If
   Next I
   
   If ini <> -1 Then
      j = 0
      For I = ini To UBound(proxSessoes)
         sspSessao(j).Caption = Format(proxSessoes(I).horario, "Hh:Nn")
         sspFilme(j).Caption = proxSessoes(I).filme
         sspSala(j).Caption = proxSessoes(I).sala
         If verificaLOTADO(proxSessoes(I).codFilme, proxSessoes(I).codSala, proxSessoes(I).horario) Then
            sspVenda(j).Caption = "LOTADO"
            
            sspVenda(j).BackColor = corFundLOTADO
            sspVenda(j).ForeColor = corTextLOTADO
            
            If LBound(ctrlsPisca2) = 0 Then
               ReDim ctrlsPisca2(1 To 1) As Control
               ReDim strCtrlPisca2(1 To 1) As String
            Else
               ReDim Preserve ctrlsPisca2(1 To UBound(ctrlsPisca2) + 1) As Control
               ReDim Preserve strCtrlPisca2(1 To UBound(ctrlsPisca2) + 1) As String
            End If
            
            Set ctrlsPisca2(UBound(ctrlsPisca2)) = sspVenda(j)
         Else
            sspVenda(j).Caption = "A VENDA"
         End If
         
         sspSessao(j).Visible = True
         sspFilme(j).Visible = True
         sspSala(j).Visible = True
         sspVenda(j).Visible = True
         
         j = j + 1
         
         If j >= 11 Then
            Exit For
         End If
      Next I
   End If
End Sub

Private Sub preparaFilme()
   Dim I            As Integer
   Dim j            As Integer
   Dim hrProxSessao As Date
   Dim cdProxSessao As Long
   Dim hrAgoraMin   As Date
   Dim hrAgoraMax   As Date
   Dim proxFilme1   As Integer
   Dim proxFilme2   As Integer
   Dim brancoAux    As Boolean

   For I = 0 To 1
      sspDescFilme(I).Caption = ""
      sspSalaFil(I).Caption = ""
      sspCensura(I).Caption = ""
      sspAVenda(I).Caption = ""
      sspHrProxSessaoFilme(I).Caption = ""
   
      sspDescFilme(I).Visible = False
      sspSalaFil(I).Visible = False
      sspCensura(I).Visible = False
      sspAVenda(I).Visible = False
      sspHrProxSessaoFilme(I).Visible = False
      sspProxSessaoFilme(I).Visible = False
      sspTitSessoes(I).Visible = False
   Next I
   
   sspAVenda(0).BackColor = corFundT2Titulo1
   sspAVenda(0).ForeColor = corTextT2Titulo1
   
   sspAVenda(1).BackColor = corFundT2Titulo2
   sspAVenda(1).ForeColor = corTextT2Titulo2

   sspHoraFilme.Caption = Format(Now, "Hh:Nn")
   
   For I = 0 To 19
      sspHrSessaoFilme(I).Caption = ""
      
      sspHrSessaoFilme(I).Visible = False
   Next I
   
   proxFilme1 = proxFilme
   proxFilme2 = proxFilme1 + 1
   If proxFilme2 < LBound(filmes) Or proxFilme2 > UBound(filmes) Then
      proxFilme2 = LBound(filmes)
   End If
   
'   hrAgoraMin = DateAdd("n", -1 * vendaDepois, CDate(Format(Now, "Hh:Nn")))
   If CDate(Format(Now, "Hh:Nn")) > CDate("03:00") Then
      hrAgoraMin = CDate(dtStrRef1 & " " & Format(Now, "Hh:Nn"))
   Else
      hrAgoraMin = CDate(dtStrRef2 & " " & Format(Now, "Hh:Nn"))
   End If
   hrAgoraMin = DateAdd("n", -1 * vendaDepois, hrAgoraMin)
   
   sspDescFilme(0).Caption = filmes(proxFilme1).descFilme
   sspSalaFil(0).Caption = Trim(filmes(proxFilme1).descSala)
   
   'sspSalaFil(0).Caption = "Sala - " & Trim(filmes(proxFilme1).descSala)
   sspCensura(0).Caption = filmes(proxFilme1).censura
   
   j = 0
   hrProxSessao = Empty
   cdProxSessao = 0
      
   For I = LBound(filmes(proxFilme1).horarios) To UBound(filmes(proxFilme1).horarios)
      sspHrSessaoFilme(j).Caption = Format(filmes(proxFilme1).horarios(I).horario, "Hh:Nn")
   
      If hrAgoraMin <= filmes(proxFilme1).horarios(I).horario And _
         hrProxSessao = Empty Then
         hrProxSessao = filmes(proxFilme1).horarios(I).horario
         cdProxSessao = filmes(proxFilme1).horarios(I).sessao
      End If
      j = j + 1
      If j > 9 Then
         Exit For
      End If
   Next I
   
   If hrProxSessao <> Empty Then
      sspHrProxSessaoFilme(0).Caption = Format(hrProxSessao, "Hh:Nn")
      If verificaLOTADO(filmes(proxFilme1).codFilme, filmes(proxFilme1).codSala, hrProxSessao) Then
      'If verificaLOTADO(filmes(proxFilme1).codFilme, filmes(proxFilme1).codSala, cdProxSessao) Then
         sspAVenda(0).Caption = "LOTADO"
         sspAVenda(0).BackColor = corFundLOTADO
         sspAVenda(0).ForeColor = corTextLOTADO
      Else
         sspAVenda(0).Caption = "A VENDA"
      End If
   Else
         sspAVenda(0).Caption = "ENCERRADO"
   End If
   
   brancoAux = bBranco1
   Do While brancoAux = bBranco1
      DoEvents
   Loop
   
   sspDescFilme(0).Visible = True
   sspSalaFil(0).Visible = True
   sspCensura(0).Visible = True
   sspAVenda(0).Visible = True
   sspHrProxSessaoFilme(0).Visible = True
   sspProxSessaoFilme(0).Visible = True
   sspTitSessoes(0).Visible = True
   
   For I = 0 To 9
      sspHrSessaoFilme(I).Visible = True
   Next I
   
   If proxFilme2 <> proxFilme1 Then
      sspDescFilme(1).Caption = filmes(proxFilme2).descFilme
      sspSalaFil(1).Caption = Trim(filmes(proxFilme2).descSala)
      'sspSalaFil(1).Caption = "Sala - " & Trim(filmes(proxFilme2).descSala)
      sspCensura(1).Caption = filmes(proxFilme2).censura
      
      j = 10
      hrProxSessao = Empty
         
      For I = LBound(filmes(proxFilme2).horarios) To UBound(filmes(proxFilme2).horarios)
         sspHrSessaoFilme(j).Caption = Format(filmes(proxFilme2).horarios(I).horario, "Hh:Nn")
      
         If hrAgoraMin <= filmes(proxFilme2).horarios(I).horario And _
            hrProxSessao = Empty Then
            hrProxSessao = filmes(proxFilme2).horarios(I).horario
            cdProxSessao = filmes(proxFilme2).horarios(I).sessao
         End If
         j = j + 1
         If j > 19 Then
            Exit For
         End If
      Next I
      
      If hrProxSessao <> Empty Then
         sspHrProxSessaoFilme(1).Caption = Format(hrProxSessao, "Hh:Nn")
         If verificaLOTADO(filmes(proxFilme2).codFilme, filmes(proxFilme2).codSala, hrProxSessao) Then
         'If verificaLOTADO(filmes(proxFilme2).codFilme, filmes(proxFilme2).codSala, cdProxSessao) Then
            sspAVenda(1).Caption = "LOTADO"
            sspAVenda(1).BackColor = corFundLOTADO
            sspAVenda(1).ForeColor = corTextLOTADO
         Else
            sspAVenda(1).Caption = "A VENDA"
         End If
      Else
            sspAVenda(1).Caption = "ENCERRADO"
      End If
      
      brancoAux = bBranco1
      Do While brancoAux = bBranco1
         DoEvents
      Loop
      
      sspDescFilme(1).Visible = True
      sspSalaFil(1).Visible = True
      sspCensura(1).Visible = True
      sspAVenda(1).Visible = True
      sspHrProxSessaoFilme(1).Visible = True
      sspProxSessaoFilme(1).Visible = True
      sspTitSessoes(1).Visible = True
      
      For I = 10 To 19
         sspHrSessaoFilme(I).Visible = True
      Next I
   End If
End Sub

Private Sub preparaPrecos()
   Dim I As Integer
   Dim j As Integer

    Call carregaPrecos
    
   sspHoraPreco.Caption = Format(Now, "Hh:Nn")
   
   'Tela 1
   sspFilmePrecos.Caption = ""
   sspManha.Caption = ""
   sspTarde.Caption = ""
   lblMensagemPreco1.Caption = mensagem
   lblMensagemPreco1.Left = -1 * lblMensagemPreco1.Width + -2

   '0 -> 3
   For I = 0 To 3
      sspDiasSemana(I).Caption = ""
      sspIntManha(I).Caption = ""
      sspMeiaManha(I).Caption = ""
      sspIntTarde(I).Caption = ""
      sspMeiaTarde(I).Caption = ""
      sspProcional(I).Caption = ""
   
      sspDiasSemana(I).Visible = False
      sspIntManha(I).Visible = False
      sspMeiaManha(I).Visible = False
      sspIntTarde(I).Visible = False
      sspMeiaTarde(I).Visible = False
      sspProcional(I).Visible = False
   Next I
   
   If precos(proxPreco).codFilme = 0 Then
      sspFilmePrecos.Caption = "PREÇOS"
   Else
      sspFilmePrecos.Caption = "PREÇOS - " & precos(proxPreco).descFilme
   End If
   
   'sspManha.Caption = "Até as " & Format(hrLimitePeriodo, "Hh:Nn")
   'sspTarde.Caption = "Após as " & Format(hrLimitePeriodo, "Hh:Nn")
   
   sspManha.Caption = "Antes das " & Format(wHorario2, "Hh:Nn")
   sspTarde.Caption = "Á partir das " & Format(wHorario1, "Hh:Nn")
   
   j = 0
   For I = LBound(precos(proxPreco).precos) To UBound(precos(proxPreco).precos)
   
      sspDiasSemana(j).Caption = precos(proxPreco).precos(I).descricao
      If precos(proxPreco).precos(I).promocional Then
         sspProcional(j).Caption = "Promocional"
      End If
      If precos(proxPreco).precos(I).vlrIntManha > 0 Then
         sspIntManha(j).Caption = "R$ " & Format(precos(proxPreco).precos(I).vlrIntManha, "##0.00")
      Else
         sspIntManha(j).Caption = ""
      End If
      If precos(proxPreco).precos(I).vlrMeiaManha > 0 Then
         sspMeiaManha(j).Caption = "R$ " & Format(precos(proxPreco).precos(I).vlrMeiaManha, "##0.00")
         Else
         sspMeiaManha(j).Caption = ""
      End If
   
      If precos(proxPreco).precos(I).vlrIntTarde > 0 Then
         sspIntTarde(j).Caption = "R$ " & Format(precos(proxPreco).precos(I).vlrIntTarde, "##0.00")
      Else
         sspIntTarde(j).Caption = ""
      End If
      If precos(proxPreco).precos(I).vlrMeiaTarde > 0 Then
         sspMeiaTarde(j).Caption = "R$ " & Format(precos(proxPreco).precos(I).vlrMeiaTarde, "##0.00")
      Else
         sspMeiaTarde(j).Caption = ""
      End If
      
      sspDiasSemana(j).Visible = True
      sspIntManha(j).Visible = True
      sspMeiaManha(j).Visible = True
      sspIntTarde(j).Visible = True
      sspMeiaTarde(j).Visible = True
      sspProcional(j).Visible = True
      
      j = j + 1
      If j > 3 Then
         Exit For
      End If
   Next I
   
End Sub

Private Sub preparaImagem()
        Dim wFator As Double
       Imagem.Picture = LoadPicture(App.Path + "\" + imagens(proxImagem))
       Imagem.Stretch = False
       Imagem.Picture = LoadPicture(App.Path + "\" + imagens(proxImagem))
       
       wFator = Screen.Height / Imagem.Height - 1
       Imagem.Stretch = True
       Imagem.Height = Imagem.Height + (Imagem.Height * wFator)
       Imagem.Width = Imagem.Width + (Imagem.Width * wFator)
       Imagem.Left = (Screen.Width / 2) - (Imagem.Width / 2)
End Sub

Private Sub preparaTrailer()
    'SLS
    WindowsMediaPlayer1.URL = gGetShortPathName(App.Path + "\" + trailers(proxTrailer))
End Sub


Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = vbKeyEscape Then ' And Shift = vbAltMask Then
     
      TimerTelas.Interval = 0
      TimerPisca.Interval = 0
      TimerVelocMsg.Interval = 0
      
      If telaAtu = 4 Then
         WindowsMediaPlayer1.Visible = False
         WindowsMediaPlayer1.Controls.stop
      End If
      
      'gFechaBase
      
      DoEvents
      
      End
   End If
End Sub

Private Sub Form_Load()
    
   dbConnect.Open "FILE NAME=" + App.Path + "\Cinema.udl"

   cmdConLotacao.CommandType = adCmdStoredProc
   cmdConLotacao.CommandText = "upTB_SESSAO_AUX_Lot"
   cmdConLotacao.CommandTimeout = 30
   cmdConLotacao.ActiveConnection = dbConnect
   
   cmdConLotacao.Parameters.Append cmdConLotacao.CreateParameter("@sal_cd", adInteger, adParamInput)
   cmdConLotacao.Parameters.Append cmdConLotacao.CreateParameter("@fil_cd", adInteger, adParamInput)
   cmdConLotacao.Parameters.Append cmdConLotacao.CreateParameter("@sre_data", adDate, adParamInput)
   cmdConLotacao.Parameters.Append cmdConLotacao.CreateParameter("@sre_horario", adDate, adParamInput)
   cmdConLotacao.Parameters.Append cmdConLotacao.CreateParameter("@sre_lugares", adInteger, adParamOutput)
   cmdConLotacao.Parameters.Append cmdConLotacao.CreateParameter("@sea_lugares_sel", adInteger, adParamOutput)
   cmdConLotacao.Parameters.Append cmdConLotacao.CreateParameter("@sea_lugares_ven", adInteger, adParamOutput)
   cmdConLotacao.Parameters.Append cmdConLotacao.CreateParameter("@sea_inteiras", adInteger, adParamOutput)
   cmdConLotacao.Parameters.Append cmdConLotacao.CreateParameter("@sre_inteiras", adInteger, adParamOutput)
   cmdConLotacao.Parameters.Append cmdConLotacao.CreateParameter("@sea_meias", adInteger, adParamOutput)
   cmdConLotacao.Parameters.Append cmdConLotacao.CreateParameter("@sre_meias", adInteger, adParamOutput)
   cmdConLotacao.Parameters.Append cmdConLotacao.CreateParameter("@sea_cortesias", adInteger, adParamOutput)
   cmdConLotacao.Parameters.Append cmdConLotacao.CreateParameter("@sre_cortesias", adInteger, adParamOutput)
   cmdConLotacao.Parameters.Append cmdConLotacao.CreateParameter("@Erro", adInteger, adParamOutput)
   cmdConLotacao.Parameters.Append cmdConLotacao.CreateParameter("@MsgErr", adVarChar, adParamOutput, 255)
    
    Dim ScaleFactorX As Single, ScaleFactorY As Single  ' Scaling factors
      ' Size of Form in Pixels at design resolution
      DesignX = 800
      DesignY = 600
      RePosForm = True   ' Flag for positioning Form
      DoResize = False   ' Flag for Resize Evefnt
      ' Set up the screen values
      Xtwips = Screen.TwipsPerPixelX
      Ytwips = Screen.TwipsPerPixelY
      Ypixels = Screen.Height / Ytwips ' Y Pixel Resolution
      Xpixels = Screen.Width / Xtwips  ' X Pixel Resolution

      ' Determine scaling factors
      ScaleFactorX = (Xpixels / DesignX)
      ScaleFactorY = (Ypixels / DesignY)
      If ScaleFactorY > ScaleFactorX Then
        Me.WindowState = vbNormal
        'Me.Width = Screen.Width
      End If
      
      ScaleMode = 1  ' twips
      'Exit Sub  ' uncomment to see how Form1 looks without resizing
      
      Resize_For_Resolution ScaleFactorX, ScaleFactorY, Me
      MyForm.Height = Me.Height ' Remember the current size
      MyForm.Width = Me.Width
      
    '*************************************************************
   proxFilme = -100
   proxPreco = -100
   proxTrailer = -100
   proxImagem = -100
   
   Call setaCores
   
   If (Not telaSessoes) And _
      (Not telaFilme) And _
      (Not telaPrecos) And _
      (Not telaImagem) And _
      (Not telaTrailer) Then
      End
   End If
   
   ReDim ctrlsPisca1(1 To 4) As Control
   ReDim strCtrlPisca1(1 To 4) As String
   
   Set ctrlsPisca1(1) = sspHrProxSessaoFilme(0)
   Set ctrlsPisca1(2) = sspAVenda(0)
   
   Set ctrlsPisca1(3) = sspHrProxSessaoFilme(1)
   Set ctrlsPisca1(4) = sspAVenda(1)
   
   bBranco1 = False
   
   ReDim ctrlsPisca2(0 To 0) As Control
   ReDim strCtrlPisca2(0 To 0) As String
   bBranco2 = False
   
   TimerPisca.Interval = intermitencia
   
   ReDim ctrlsMsg(1 To 2) As Control
   Set ctrlsMsg(1) = lblMensagem1
   Set ctrlsMsg(2) = lblMensagemPreco1
   TimerVelocMsg.Interval = 5

   If telaFilme Then
      proxFilme = LBound(filmes)
      Call preparaFilme
      telaAtu = 1
      fraFilme.ZOrder 0
      fraFilme.Visible = True
   ElseIf telaSessoes Then
      Call preparaProxSessoes
      telaAtu = 2
      fraProxSessoes.ZOrder 0
      fraProxSessoes.Visible = True
   ElseIf telaPrecos Then
      proxPreco = LBound(precos)
      Call preparaPrecos
      telaAtu = 3
      fraPrecos.ZOrder 0
      fraPrecos.Visible = True
   ElseIf telaTrailer Then
   '   proxTrailer = LBound(trailers)
   '   Call preparaTrailer
   '
      'Calcula o termino do trailer
      'Salva a transicão atual
      'Altera a transição
      'Inicia video
      'Volta transição
   '   telaAtu = 4
      'MMControl1.hWndDisplay = frmApresent.hWnd
      'MMControl1.Command = "Play"
   '     fraImagem.Visible = True
   '     fraImagem.ZOrder 0
   '     WindowsMediaPlayer1.Visible = True
    '    WindowsMediaPlayer1.fullScreen = True
    ''    WindowsMediaPlayer1.Controls.play
    '    TimerTelas.Enabled = False
        'TimerTelas.Interval = 65000 'Format(WindowsMediaPlayer1.currentMedia.duration, "###.000") * 1000
   ElseIf telaImagem Then
      proxImagem = LBound(imagens)
              
      Call preparaImagem
      telaAtu = 5
   End If

   wAtualiza = True

   preparaProxTela

   TimerTelas.Enabled = True
   TimerTelas.Interval = transicao * 1000
   
   Me.Show
   
End Sub

Private Sub preparaProxTela()

    If wAtualiza = False Then
        If iqtdAtualiza < (iqtdTelas + (5 - iqtdTelas)) Then
            iqtdAtualiza = iqtdAtualiza + 1
        Else
            iqtdAtualiza = 0
            wAtualiza = True
        End If
    End If
   
   
   If telaAtu <> 1 And telaFilme Then
      proxFilme = proxFilme + 2
      If proxFilme < LBound(filmes) Or proxFilme > UBound(filmes) Then
         proxFilme = LBound(filmes)
      End If
      Call preparaFilme
   Else
      'proxTela = telaAtu
      
      Do While True
         proxTela = proxTela + 1
         If proxTela > 5 Then
            proxTela = 2
         End If
         
         'If proxTela = 1 And telaSessoes Then
         '   Exit Do
         'ElseIf proxTela = 2 And telaFilme Then
         If proxTela = 2 And telaSessoes Then
            Exit Do
         'ElseIf proxTela = 3 And telaSessoes Then
         '   Exit Do
         ElseIf proxTela = 3 And telaPrecos Then
            Exit Do
         ElseIf proxTela = 4 And telaTrailer Then
            Exit Do
         ElseIf proxTela = 5 And telaImagem Then
            Exit Do
         End If
      Loop
      
      If proxTela = 2 Then
         Call preparaProxSessoes
      ElseIf proxTela = 3 Then
         proxPreco = proxPreco + 1
         If proxPreco < LBound(precos) Or proxPreco > UBound(precos) Then
            proxPreco = LBound(precos)
         End If
         Call preparaPrecos
         
      ElseIf proxTela = 4 Then
         proxTrailer = proxTrailer + 1
         If proxTrailer < LBound(trailers) Or proxTrailer > UBound(trailers) Then
            proxTrailer = LBound(trailers)
         End If
         Call preparaTrailer
      ElseIf proxTela = 5 Then
         proxImagem = proxImagem + 1
         If proxImagem < LBound(imagens) Or proxImagem > UBound(imagens) Then
            proxImagem = LBound(imagens)
         End If
         Call preparaImagem
      End If
   End If
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
   'Call gFechaBase
End Sub

Private Sub Form_Resize()
    'SLS 16/09/2011
    Dim ScaleFactorX As Single, ScaleFactorY As Single '

      If Not DoResize Then  ' To avoid infinite loop
         DoResize = True
         Exit Sub
      End If

      RePosForm = False
      ScaleFactorX = Me.Width / MyForm.Width   ' How much change?
      ScaleFactorY = Me.Height / MyForm.Height
      Resize_For_Resolution ScaleFactorX, ScaleFactorY, Me
      MyForm.Height = Me.Height ' Remember the current size
      MyForm.Width = Me.Width
End Sub



Private Sub TimerPisca_Timer()
   Dim I As Integer
   
   If Not bBranco1 Then
      For I = LBound(ctrlsPisca1) To UBound(ctrlsPisca1)
         strCtrlPisca1(I) = ctrlsPisca1(I).Caption
         If ctrlsPisca1(I).Visible Then
            ctrlsPisca1(I).Caption = ""
         End If
      Next I
      bBranco1 = True
   Else
      For I = LBound(ctrlsPisca1) To UBound(ctrlsPisca1)
         If ctrlsPisca1(I).Visible Then
            ctrlsPisca1(I).Caption = strCtrlPisca1(I)
         End If
      Next I
      
      bBranco1 = False
   End If
   
   If LBound(ctrlsPisca2) <> 0 Then
      If Not bBranco2 Then
         For I = LBound(ctrlsPisca2) To UBound(ctrlsPisca2)
            strCtrlPisca2(I) = ctrlsPisca2(I).Caption
            If ctrlsPisca2(I).Visible Then
               ctrlsPisca2(I).Caption = ""
            End If
         Next I
         bBranco2 = True
      Else
         For I = LBound(ctrlsPisca2) To UBound(ctrlsPisca2)
            If ctrlsPisca2(I).Visible Then
               ctrlsPisca2(I).Caption = strCtrlPisca2(I)
            End If
         Next I
         
         bBranco2 = False
      End If
   End If

End Sub

Private Sub TimerTelas_Timer()
   
   If telaAtu <> 1 And telaFilme Then
      fraFilme.Visible = True
      fraFilme.ZOrder 0
      telaAtu = 1
   Else
      If proxTela = 2 Then
         fraProxSessoes.Visible = True
         fraProxSessoes.ZOrder 0
         telaAtu = 2
      ElseIf proxTela = 3 Then
         fraPrecos.Visible = True
         fraPrecos.ZOrder 0
         telaAtu = 3
      ElseIf proxTela = 4 Then
         telaAtu = 4
         Imagem.Visible = False
         fraPrecos.Visible = False
         fraFilme.Visible = False
         fraProxSessoes.Visible = False
         fraImagem.Visible = False
         fravideo.Visible = True
         WindowsMediaPlayer1.uiMode = "none"
         WindowsMediaPlayer1.Controls.play
         TimerTelas.Enabled = False
      ElseIf proxTela = 5 Then
         fraImagem.Visible = True
         Imagem.Visible = True
         fraImagem.ZOrder 0
         telaAtu = 5
      End If
   End If
   
    If wAtualiza = True Then
        
        Call Main
        wAtualiza = False
    
    End If
   
   Call preparaProxTela

End Sub

Private Sub setaCores()
   Dim I        As Integer
   Dim corFundo As Long
   Dim corTexto As Long
   
   'Tela 1
   sspFilmePrecos.BackColor = corFundT1Filme
   sspFilmePrecos.ForeColor = corTextT1Filme
   sspHoraPreco.BackColor = corFundT1Hora
   sspHoraPreco.ForeColor = corTextT1Hora
   sspManha.BackColor = corFundT1Tutulo1
   sspManha.ForeColor = corTextT1Tutulo1
   sspTarde.BackColor = corFundT1Tutulo1
   sspTarde.ForeColor = corTextT1Tutulo1
   sspMensagemPreco.BackColor = corFundT1Mensagem
   sspMensagemPreco.ForeColor = corTextT1Mensagem
   lblMensagemPreco1.BackColor = corFundT1Mensagem
   lblMensagemPreco1.ForeColor = corTextT1Mensagem

   sspFilmePrecos.Font3D = 0
   sspHoraPreco.Font3D = 0
   sspManha.Font3D = 0
   sspTarde.Font3D = 0
   sspMensagemPreco.Font3D = 0

   '0 -> 1
   For I = 0 To 1
      sspInteira(I).BackColor = corFundT1Titulo2
      sspInteira(I).ForeColor = corTextT1Titulo2
      sspMeia(I).BackColor = corFundT1Titulo2
      sspMeia(I).ForeColor = corTextT1Titulo2
   
      sspInteira(I).Font3D = 0
      sspMeia(I).Font3D = 0
   Next I

   '0 -> 3
   For I = 0 To 3
      If I Mod 2 = 0 Then
         corFundo = corFundT1Lin1
         corTexto = corTextT1Lin1
      Else
         corFundo = corFundT1Lin2
         corTexto = corTextT1Lin2
      End If
      sspDiasSemana(I).BackColor = corFundo
      sspDiasSemana(I).ForeColor = corTexto
      sspIntManha(I).BackColor = corFundo
      sspIntManha(I).ForeColor = corTexto
      sspMeiaManha(I).BackColor = corFundo
      sspMeiaManha(I).ForeColor = corTexto
      sspIntTarde(I).BackColor = corFundo
      sspIntTarde(I).ForeColor = corTexto
      sspMeiaTarde(I).BackColor = corFundo
      sspMeiaTarde(I).ForeColor = corTexto
      sspProcional(I).BackColor = corFundo
      sspProcional(I).ForeColor = corTexto
   
      sspDiasSemana(I).Font3D = 0
      sspIntManha(I).Font3D = 0
      sspMeiaManha(I).Font3D = 0
      sspIntTarde(I).Font3D = 0
      sspMeiaTarde(I).Font3D = 0
      sspProcional(I).Font3D = 0
   Next I

   'Tela 2
   '0 -> 1
   sspDescFilme(0).BackColor = corFundT2Filme1
   sspDescFilme(0).ForeColor = corTextT2Filme1
   sspSalaFil(0).BackColor = corFundT2Sala1
   sspSalaFil(0).ForeColor = corTextT2Sala1
   sspTitSessoes(0).BackColor = corFundT2Sessoes1
   sspTitSessoes(0).ForeColor = corTextT2Sessoes1
   sspCensura(0).BackColor = corFundT2Titulo1
   sspCensura(0).ForeColor = corTextT2Titulo1
   sspProxSessaoFilme(0).BackColor = corFundT2Titulo1
   sspProxSessaoFilme(0).ForeColor = corTextT2Titulo1
   sspAVenda(0).BackColor = corFundT2Titulo1
   sspAVenda(0).ForeColor = corTextT2Titulo1
   sspHrProxSessaoFilme(0).BackColor = corFundT2Sessao1
   sspHrProxSessaoFilme(0).ForeColor = corTextT2Sessao1
   
   sspDescFilme(0).Font3D = 0
   sspSalaFil(0).Font3D = 0
   sspTitSessoes(0).Font3D = 0
   sspCensura(0).Font3D = 0
   sspProxSessaoFilme(0).Font3D = 0
   sspAVenda(0).Font3D = 0
   sspHrProxSessaoFilme(0).Font3D = 0
   
   sspDescFilme(1).BackColor = corFundT2Filme2
   sspDescFilme(1).ForeColor = corTextT2Filme2
   sspSalaFil(1).BackColor = corFundT2Sala2
   sspSalaFil(1).ForeColor = corTextT2Sala2
   sspTitSessoes(1).BackColor = corFundT2Sessoes2
   sspTitSessoes(1).ForeColor = corTextT2Sessoes2
   sspCensura(1).BackColor = corFundT2Titulo2
   sspCensura(1).ForeColor = corTextT2Titulo2
   sspProxSessaoFilme(1).BackColor = corFundT2Titulo2
   sspProxSessaoFilme(1).ForeColor = corTextT2Titulo2
   sspAVenda(1).BackColor = corFundT2Titulo2
   sspAVenda(1).ForeColor = corTextT2Titulo2
   sspHrProxSessaoFilme(1).BackColor = corFundT2Sessao2
   sspHrProxSessaoFilme(1).ForeColor = corTextT2Sessao2
   
   sspDescFilme(1).Font3D = 0
   sspSalaFil(1).Font3D = 0
   sspTitSessoes(1).Font3D = 0
   sspCensura(1).Font3D = 0
   sspProxSessaoFilme(1).Font3D = 0
   sspAVenda(1).Font3D = 0
   sspHrProxSessaoFilme(1).Font3D = 0
   
   sspHoraFilme.BackColor = corFundT2Mensagem
   sspHoraFilme.ForeColor = corTextT2Mensagem

   sspHoraFilme.Font3D = 0

   ' 0 -> 9
   '10-> 19
   For I = 0 To 4
      If I Mod 2 = 0 Then
         corFundo = corFundT2Sessoes1L1
         corTexto = corTextT2Sessoes1L1
      Else
         corFundo = corFundT2Sessoes1L2
         corTexto = corTextT2Sessoes1L2
      End If
      
      sspHrSessaoFilme(I).BackColor = corFundo
      sspHrSessaoFilme(I).ForeColor = corTexto
      sspHrSessaoFilme(I + 5).BackColor = corFundo
      sspHrSessaoFilme(I + 5).ForeColor = corTexto
   
      sspHrSessaoFilme(I).Font3D = 0
      sspHrSessaoFilme(I + 5).Font3D = 0
   Next I
   
   For I = 10 To 14
      If I Mod 2 = 0 Then
         corFundo = corFundT2Sessoes2L1
         corTexto = corTextT2Sessoes2L1
      Else
         corFundo = corFundT2Sessoes2L2
         corTexto = corTextT2Sessoes2L2
      End If
      
      sspHrSessaoFilme(I).BackColor = corFundo
      sspHrSessaoFilme(I).ForeColor = corTexto
      sspHrSessaoFilme(I + 5).BackColor = corFundo
      sspHrSessaoFilme(I + 5).ForeColor = corTexto
   
      sspHrSessaoFilme(I).Font3D = 0
      sspHrSessaoFilme(I + 5).Font3D = 0
   Next I

   'Tela 3
   sspHora.BackColor = corFundT3Hora
   sspHora.ForeColor = corTextT3Hora
   sspTitProxSes.BackColor = corFundT3TituloTela
   sspTitProxSes.ForeColor = corTextT3TituloTela
   sspData.BackColor = corFundT3Data
   sspData.ForeColor = corTextT3Data
   sspTituloSessao.BackColor = corFundT3Titulo
   sspTituloSessao.ForeColor = corTextT3Titulo
   sspTituloFilme.BackColor = corFundT3Titulo
   sspTituloFilme.ForeColor = corTextT3Titulo
   sspTituloSala.BackColor = corFundT3Titulo
   sspTituloSala.ForeColor = corTextT3Titulo
   sspTituloVenda.BackColor = corFundT3Titulo
   sspTituloVenda.ForeColor = corTextT3Titulo
   sspMensagem.BackColor = corFundT3Mensagem
   sspMensagem.ForeColor = corTextT3Mensagem
   lblMensagem1.BackColor = corFundT3Mensagem
   lblMensagem1.ForeColor = corTextT3Mensagem
   
   sspHora.Font3D = 0
   sspTitProxSes.Font3D = 0
   sspData.Font3D = 0
   sspTituloSessao.Font3D = 0
   sspTituloFilme.Font3D = 0
   sspTituloSala.Font3D = 0
   sspTituloVenda.Font3D = 0
   sspMensagem.Font3D = 0
   
   '0 -> 10
   For I = 0 To 10
      If I Mod 2 = 0 Then
         corFundo = corFundT3Lin1
         corTexto = corTextT3Lin1
      Else
         corFundo = corFundT3Lin2
         corTexto = corTextT3Lin2
      End If
      
      sspSessao(I).BackColor = corFundo
      sspSessao(I).ForeColor = corTexto
      sspFilme(I).BackColor = corFundo
      sspFilme(I).ForeColor = corTexto
      sspSala(I).BackColor = corFundo
      sspSala(I).ForeColor = corTexto
      sspVenda(I).BackColor = corFundo
      sspVenda(I).ForeColor = corTexto
   
      sspSessao(I).Font3D = 0
      sspFilme(I).Font3D = 0
      sspSala(I).Font3D = 0
      sspVenda(I).Font3D = 0
   Next I

End Sub

Private Sub TimerVelocMsg_Timer()
   Dim ctrlAux As Control
   Dim I       As Integer
   
   For I = LBound(ctrlsMsg) To UBound(ctrlsMsg)
      If ctrlsMsg(I).Left < -1 * ctrlsMsg(I).Width Then
         Set ctrlAux = ctrlsMsg(I).Container
         ctrlsMsg(I).Left = ctrlAux.Width
      Else
         ctrlsMsg(I).Left = ctrlsMsg(I).Left - velocMsg
      End If
   Next I
   DoEvents
End Sub

Private Sub WindowsMediaPlayer1_PlayStateChange(ByVal NewState As Long)
        If NewState = 8 Then
            'fravideo.Visible = False
            TimerTelas.Enabled = True
            Call TimerTelas_Timer
        End If
        If (NewState = wmppsPlaying) Then
            WindowsMediaPlayer1.fullScreen = True
        End If
End Sub



