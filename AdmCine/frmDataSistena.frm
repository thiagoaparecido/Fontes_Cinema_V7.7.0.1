VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Begin VB.Form frmDataSistema 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   2715
   ClientLeft      =   45
   ClientTop       =   45
   ClientWidth     =   6240
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2715
   ScaleWidth      =   6240
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin Threed.SSCommand sscConfirma 
      Height          =   675
      Left            =   2715
      TabIndex        =   2
      Top             =   1920
      Width           =   1500
      _Version        =   65536
      _ExtentX        =   2646
      _ExtentY        =   1191
      _StockProps     =   78
      Caption         =   "Confirmar"
      ForeColor       =   12582912
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Frame fraAviso 
      Caption         =   "Aviso"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   75
      TabIndex        =   0
      Top             =   660
      Width           =   6045
      Begin VB.Label lblAviso 
         Alignment       =   2  'Center
         Caption         =   $"frmDataSistena.frx":0000
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   585
         Left            =   105
         TabIndex        =   1
         Top             =   330
         Width           =   5760
      End
   End
   Begin Threed.SSCommand sscSair 
      Height          =   675
      Left            =   4395
      TabIndex        =   3
      Top             =   1920
      Width           =   1500
      _Version        =   65536
      _ExtentX        =   2646
      _ExtentY        =   1191
      _StockProps     =   78
      Caption         =   "Sair"
      ForeColor       =   255
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSComCtl2.DTPicker dtpDataSistema 
      Height          =   360
      Left            =   2760
      TabIndex        =   5
      Top             =   165
      Width           =   2310
      _ExtentX        =   4075
      _ExtentY        =   635
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      CalendarForeColor=   0
      CalendarTitleForeColor=   0
      CustomFormat    =   "dd/MM/yy HH:mm:ss"
      Format          =   58851331
      CurrentDate     =   38483
   End
   Begin VB.Label lblDataSistema 
      AutoSize        =   -1  'True
      Caption         =   "Data do Sistema"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   240
      Left            =   645
      TabIndex        =   4
      Top             =   210
      Width           =   1755
   End
End
Attribute VB_Name = "frmDataSistema"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private altDtSistema As Boolean

Private Sub dtpDataSistema_Change()
    altDtSistema = True
End Sub

Private Sub Form_Load()
    dtpDataSistema.Value = Now
    altDtSistema = False
End Sub

Private Sub sscConfirma_Click()
    If altDtSistema Then
        If MsgBox("Confirma Alteração da data do Sistema?", vbOKCancel) = vbCancel Then
            Exit Sub
        End If
    
        Date = dtpDataSistema.Value
        Time = dtpDataSistema.Value
    End If
    
    bConfirmaDtSistema = True
    Unload Me
End Sub

Private Sub sscSair_Click()
    Unload Me
End Sub
