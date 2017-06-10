VERSION 5.00
Object = "{C0A63B80-4B21-11D3-BD95-D426EF2C7949}#1.0#0"; "Vsflex7L.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{648A5603-2C6E-101B-82B6-000000000014}#1.1#0"; "MSCOMM32.OCX"
Object = "{C9E1AFB0-1172-11D7-83AD-0050DA238ADA}#1.0#0"; "Coptr19.ocx"
Begin VB.Form frmVendaIngresso 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Relatório de Vendas"
   ClientHeight    =   8160
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   7380
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8160
   ScaleWidth      =   7380
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdImprime 
      Cancel          =   -1  'True
      Caption         =   "&Imprime"
      Height          =   390
      Left            =   5805
      TabIndex        =   10
      Top             =   1110
      Width           =   1438
   End
   Begin VB.Frame Frame1 
      Height          =   975
      Left            =   135
      TabIndex        =   3
      Top             =   135
      Width           =   5535
      Begin VB.OptionButton OptTipo 
         Caption         =   "Ambos (Totalizado)"
         Height          =   195
         Index           =   2
         Left            =   2475
         TabIndex        =   13
         Top             =   720
         Value           =   -1  'True
         Width           =   1905
      End
      Begin VB.OptionButton OptTipo 
         Caption         =   "Combos (Detalhado)"
         Height          =   195
         Index           =   1
         Left            =   2475
         TabIndex        =   12
         Top             =   450
         Width           =   1950
      End
      Begin VB.OptionButton OptTipo 
         Caption         =   "Ingressos (Detalhado)"
         Height          =   195
         Index           =   0
         Left            =   2475
         TabIndex        =   11
         Top             =   180
         Width           =   1950
      End
      Begin MSComCtl2.DTPicker dtpDtIni 
         Height          =   315
         Left            =   1080
         TabIndex        =   4
         Top             =   165
         Width           =   1260
         _ExtentX        =   2223
         _ExtentY        =   556
         _Version        =   393216
         Format          =   62324737
         CurrentDate     =   38606
      End
      Begin MSComCtl2.DTPicker dtpDtFim 
         Height          =   315
         Left            =   1080
         TabIndex        =   5
         Top             =   540
         Width           =   1260
         _ExtentX        =   2223
         _ExtentY        =   556
         _Version        =   393216
         Format          =   62324737
         CurrentDate     =   38606
      End
      Begin VB.Label lblDtIni 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Data Inicio:"
         Height          =   195
         Left            =   90
         TabIndex        =   7
         Top             =   225
         Width           =   810
      End
      Begin VB.Label lblDtFim 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Data Fim:"
         Height          =   195
         Left            =   135
         TabIndex        =   6
         Top             =   585
         Width           =   675
      End
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "&OK"
      Default         =   -1  'True
      Height          =   390
      Left            =   5805
      TabIndex        =   2
      Top             =   150
      Width           =   1438
   End
   Begin VB.CommandButton cmdCancela 
      Caption         =   "&Cancela"
      Height          =   390
      Left            =   5805
      TabIndex        =   1
      Top             =   630
      Width           =   1438
   End
   Begin VSFlex7LCtl.VSFlexGrid vsfBoletim 
      Height          =   6795
      Left            =   135
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   1230
      Width           =   5535
      _cx             =   9763
      _cy             =   11986
      _ConvInfo       =   1
      Appearance      =   1
      BorderStyle     =   1
      Enabled         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9.75
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
      BackColorBkg    =   -2147483643
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
      SelectionMode   =   3
      GridLines       =   0
      GridLinesFixed  =   0
      GridLineWidth   =   1
      Rows            =   0
      Cols            =   3
      FixedRows       =   0
      FixedCols       =   0
      RowHeightMin    =   0
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   -1  'True
      FormatString    =   $"frmVendaIngresso.frx":0000
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
      FrozenCols      =   2
      AllowUserFreezing=   0
      BackColorFrozen =   0
      ForeColorFrozen =   0
      WallPaperAlignment=   9
   End
   Begin MSCommLib.MSComm MSComm 
      Left            =   6345
      Top             =   1875
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      CommPort        =   2
      DTREnable       =   0   'False
      Handshaking     =   1
      OutBufferSize   =   1311
      ParityReplace   =   32
      RThreshold      =   1
      BaudRate        =   19200
   End
   Begin VB.Frame fraAviso 
      Height          =   1890
      Left            =   390
      TabIndex        =   8
      Top             =   2790
      Visible         =   0   'False
      Width           =   6630
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Aguarde Carregando Boletins..."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   360
         Left            =   1095
         TabIndex        =   9
         Top             =   735
         Width           =   4425
      End
   End
   Begin VSFlex7LCtl.VSFlexGrid vsfBoletim1 
      Height          =   6075
      Left            =   1530
      TabIndex        =   14
      TabStop         =   0   'False
      Top             =   1710
      Visible         =   0   'False
      Width           =   5625
      _cx             =   9922
      _cy             =   10716
      _ConvInfo       =   1
      Appearance      =   1
      BorderStyle     =   1
      Enabled         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9.75
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
      BackColorBkg    =   -2147483643
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
      SelectionMode   =   3
      GridLines       =   0
      GridLinesFixed  =   0
      GridLineWidth   =   1
      Rows            =   0
      Cols            =   3
      FixedRows       =   0
      FixedCols       =   0
      RowHeightMin    =   0
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   -1  'True
      FormatString    =   $"frmVendaIngresso.frx":004C
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
      FrozenCols      =   2
      AllowUserFreezing=   0
      BackColorFrozen =   0
      ForeColorFrozen =   0
      WallPaperAlignment=   9
   End
   Begin COPTRLib.OPOSPOSPrinter OPOSPOSPrinter1 
      Left            =   6585
      Top             =   3015
      _Version        =   65536
      _ExtentX        =   609
      _ExtentY        =   582
      _StockProps     =   0
   End
End
Attribute VB_Name = "frmVendaIngresso"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private LinhaNormal
Private LinhaTitulo
Private LinhaDupla

Dim sala    As Collection
Dim filme   As Collection
Dim sesExcl As Collection

'Public linha()   As String
'Dim tpLinha() As String

Dim impriminto As Boolean
Dim dataRef    As Date
Dim inicio     As Boolean

Private Sub cmdImprime_Click()

Dim Row             As Long
Dim fonteNormal     As StdFont
Dim fonteTitulo     As StdFont
Dim fonteSubTtitulo As StdFont
Dim tpImp           As Integer
Dim hf              As Single
Dim Y               As Single
Dim tamPag          As Single
Dim strVendCombo    As String
Dim lInicio         As Long

Imprime_Boletim:

    'SLS 18/01/2012
    Call abreImp(OPOSPOSPrinter1)

    If vsfBoletim.Rows <= 0 Then
        Exit Sub
    End If
    
    'On Error GoTo erro_ImprimeBoletim
    
    tpImp = frmTipoImp.tipoImpo
    
    DoEvents
    
    If tpImp = 1 Then
        Set fonteNormal = New StdFont
        fonteNormal.Name = "Courier New"
        fonteNormal.Size = 10
        fonteNormal.Bold = False
        fonteNormal.Italic = False
          
        Set fonteTitulo = New StdFont
        fonteTitulo.Name = "Courier New"
        fonteTitulo.Size = 14
        fonteTitulo.Bold = True
        fonteTitulo.Italic = False
        
        Set fonteSubTtitulo = New StdFont
        fonteSubTtitulo.Name = "Courier New"
        fonteSubTtitulo.Size = 12
        fonteSubTtitulo.Bold = True
        fonteSubTtitulo.Italic = False
        
        Y = 10
    
        'Printer.PaperSize = vbPRPSA4
        Printer.ScaleMode = vbMillimeters
        Printer.Orientation = vbPRORPortrait
        tamPag = Printer.ScaleHeight
        
        For Row = 1 To UBound(linha)
            If tplinha(Row) = "T" Then
                Set Printer.Font = fonteTitulo
            ElseIf tplinha(Row) = "S" Then
                Set Printer.Font = fonteSubTtitulo
            Else
                Set Printer.Font = fonteNormal
            End If
            
            hf = Printer.TextHeight("X")
            
            Printer.CurrentX = 10
            Printer.CurrentY = Y
            Printer.Print linha(Row)
            
            Y = Y + hf
            If Y + 10 > tamPag Then
                Printer.NewPage
                Y = 10
            End If
        Next
    
        Printer.EndDoc
        
    ElseIf tpImp = 2 Then
    
    
        Dim wLinhaInicial As Double
        Dim wLinhaFinal As Double
        Dim wVezes As Integer
        
        'Inicializa impressora
        If Not iniImpressora(OPOSPOSPrinter1) Then
            GoTo erro_ImprimeBoletim
        End If
        
        wVezes = 1
        If UBound(linha) >= 50 Then
            wLinhaInicial = LBound(linha)
            wLinhaFinal = 50
        Else
            wLinhaInicial = LBound(linha)
            wLinhaFinal = UBound(linha)
        End If
        
        Do While wLinhaFinal <= UBound(linha)
            
            If verificaImp(OPOSPOSPrinter1) Then
                strVendCombo = ""
                If wVezes = 1 Then
                    strVendCombo = strVendCombo & strAlinhaCentralizado
                End If
                For Row = wLinhaInicial To wLinhaFinal
                    Select Case tplinha(Row)
                        Case "N"
                            'If Not imprimeNormal(OPOSPOSPrinter1, linha(Row) & Chr$(10)) Then
                            '    GoTo erro_ImprimeBoletim
                            'End If
                            strVendCombo = strVendCombo & strNormal(linha(Row) & Chr$(10))
                        Case "T"
                            'If Not imprimeTitulo(OPOSPOSPrinter1, linha(Row) & Chr$(10)) Then
                            '    GoTo erro_ImprimeBoletim
                            'End If
                            strVendCombo = strVendCombo & strTitulo(linha(Row) & Chr$(10))
                        Case "S"
                            'If Not imprimeDupla(OPOSPOSPrinter1, linha(Row) & Chr$(10)) Then
                            '    GoTo erro_ImprimeBoletim
                            'End If
                            strVendCombo = strVendCombo & strDupla(linha(Row) & Chr$(10))
                    End Select
                Next
                If wLinhaFinal = UBound(linha) Then
                    strVendCombo = strVendCombo & strNormal(Chr$(10) & Chr$(10) & Chr$(10) & Chr$(10))
                    strVendCombo = strVendCombo & strCortaPapel
                End If
                If Not imprime(OPOSPOSPrinter1, strVendCombo) Then
                    GoTo erro_ImprimeBoletim
                End If
                
                lInicio = timeGetTime
                
                Do While lInicio + CInt(pTempImp2) > timeGetTime
                    DoEvents
                Loop
            Else
                GoTo erro_ImprimeBoletim
            End If
            If wLinhaFinal = UBound(linha) Then
                Exit Do
            End If
            wVezes = wVezes + 1
            wLinhaInicial = wLinhaFinal + 1
            wLinhaFinal = wVezes * 50
            If wLinhaFinal > UBound(linha) Then
                wLinhaFinal = UBound(linha)
            End If
        Loop
    End If

    'sls 18/01/2012
    Call fechaImp(OPOSPOSPrinter1)
    
    Exit Sub

erro_ImprimeBoletim:
    Beep
    
    If MsgBox("Impressora não está pronta! Ajuste-a e clique em SIM para tentar imprimir novamente ou NÃO para desistir.", vbQuestion + vbYesNo, App.ProductName) = vbYes Then
        GoTo Imprime_Boletim
    End If

End Sub

Private Sub Form_Activate()
    If inicio Then
        dtpDtIni.Value = Date
        dtpDtFim.Value = Date
        inicio = False
    End If
End Sub

Private Sub Form_Load()
    Dim gerais As New clsGerais
    
    'Call AjustaCOM(MSComm)
    'SLS 08/01/2012
    'If abreImp(OPOSPOSPrinter1) Then
    'End If
    
    Set gerais.ConexaoADO = dbConnect
    
    dataRef = gerais.DataRefAtu
    
    impriminto = False
    inicio = True
End Sub

Private Sub cmdCancela_Click()
    If Not impriminto Then
        Unload Me
    End If
End Sub

Private Sub cmdOK_Click()
    Dim x As Integer
    
    If OptTipo(0).Value = True Then
        impriminto = True
        Label3.Caption = "Aguarde Carregando Boletins..."
        Call ImprimeVendaIngresso(dtpDtIni, dtpDtFim)
        impriminto = False
    ElseIf OptTipo(1).Value = True Then
        impriminto = True
        Label3.Caption = "Aguarde Carregando Combos..."
        DoEvents
        fraAviso.ZOrder 0
        fraAviso.Visible = True
        DoEvents
        
        Call frmVendaCombo.ImprimeVendaCombo(dtpDtIni, dtpDtFim)
        impriminto = False
        
        vsfBoletim.Clear
        vsfBoletim.Rows = 0
        vsfBoletim.FontName = "Courier"
        vsfBoletim.FontSize = 6
        vsfBoletim.GridLines = flexGridNone
    
        For x = 2 To frmVendaCombo.vsfBoletim.Rows + 1
            vsfBoletim.AddItem Chr(9) & linha(x) & Chr(9)
        Next
        fraAviso.Visible = False
        vsfBoletim.ZOrder 0
    Else
        impriminto = True
        
        Call ImprimeVendaIngressoCombo(dtpDtIni, dtpDtFim)
        impriminto = False
        
    End If
    
End Sub

Private Sub ImprimeVendaIngresso(ByVal dDataIni As Date, ByVal dDataFim As Date)

    On Error GoTo ImprimeVendaIngresso_erro
    
    Dim clsBoletim   As New clsBoletim
    Dim oRsIngressos As New ADODB.Recordset
    Dim dtDia        As Date
    Dim primeiro     As Boolean
    Dim dValorTotal  As Double
    Dim iQtdeTotal   As Integer
    Dim dValorDia    As Double
    Dim iQtdeDia     As Integer
        
    DoEvents
    fraAviso.ZOrder 0
    fraAviso.Visible = True
    DoEvents
        
    LinhaNormal = Chr$(27) & Chr$(33) & Chr$(10)
    LinhaTitulo = Chr$(27) & Chr$(33) & Chr$(60)
    LinhaDupla = Chr$(27) & Chr$(33) & Chr$(40)
    
    vsfBoletim.Clear
    vsfBoletim.Rows = 0
    'vsfBoletim.Cols = 2
    vsfBoletim.FontName = "Courier"
    vsfBoletim.FontSize = 6
    vsfBoletim.GridLines = flexGridNone
    
    Set clsBoletim.ConexaoADO = dbConnect
    
    clsBoletim.dtIni = dDataIni
    clsBoletim.dtFim = dDataFim
    
    If Not clsBoletim.VendaIngresso(oRsIngressos) Then
        MsgBox clsBoletim.MensagemErro, vbCritical, App.ProductName
        GoTo ImprimeVendaIngresso_fim
    End If
    
    ReDim linha(1 To 1) As String
    ReDim tplinha(1 To 1) As String
    linha(UBound(linha)) = ""
    tplinha(UBound(tplinha)) = "N"
    
    If IsNull(oRsIngressos.Fields("ope_dt_operacao")) Then
        MsgBox "Não houve movimento no periodo selecionado!", vbInformation, App.ProductName
        oRsIngressos.Close
        GoTo ImprimeVendaIngresso_fim
    End If

    vsfBoletim.AddItem (LinhaTitulo & Chr(9) & "VENDAS INGRESSO" & Chr(9) & Chr$(10))
    ReDim Preserve linha(1 To UBound(linha) + 1) As String
    ReDim Preserve tplinha(1 To UBound(tplinha) + 1) As String
    linha(UBound(linha)) = "VENDAS INGRESSO"
    tplinha(UBound(tplinha)) = "T"
    
    vsfBoletim.AddItem (LinhaNormal & Chr(9) & "" & Chr(9) & Chr$(10))
    ReDim Preserve linha(1 To UBound(linha) + 1) As String
    ReDim Preserve tplinha(1 To UBound(tplinha) + 1) As String
    linha(UBound(linha)) = ""
    tplinha(UBound(tplinha)) = "N"
    
    dValorTotal = 0
    iQtdeTotal = 0
    dValorDia = 0
    iQtdeDia = 0
    
    dtDia = Empty
    primeiro = True
        
    Do While Not oRsIngressos.EOF()
        If dtDia <> oRsIngressos.Fields("ope_dt_operacao") Then
            If Not primeiro Then
                vsfBoletim.AddItem (LinhaNormal & Chr(9) & Alinha("-", 37, esquerda, "-") & Chr(9) & Chr$(10))
                ReDim Preserve linha(1 To UBound(linha) + 1) As String
                ReDim Preserve tplinha(1 To UBound(tplinha) + 1) As String
                linha(UBound(linha)) = Alinha("-", COLUNAS_IMP, esquerda, "-")
                tplinha(UBound(tplinha)) = "N"
            
                vsfBoletim.AddItem (LinhaNormal & Chr(9) & _
                                    Alinha("Total dia", 20, esquerda, " ") & ": " & _
                                    Alinha(Format(iQtdeDia, "#,##0"), 5, Direita) & _
                                    Alinha(Format(dValorDia, "#,##0.00"), 10, Direita) & Chr(9) & Chr$(10))
                ReDim Preserve linha(1 To UBound(linha) + 1) As String
                ReDim Preserve tplinha(1 To UBound(tplinha) + 1) As String
                linha(UBound(linha)) = Alinha(Alinha("Total dia", 20, esquerda, " ") & ": " & _
                                              Alinha(Format(iQtdeDia, "#,##0"), 5, Direita) & _
                                              Alinha(Format(dValorDia, "#,##0.00"), 10, Direita), COLUNAS_IMP, esquerda)
                tplinha(UBound(tplinha)) = "N"
            
                vsfBoletim.AddItem (LinhaNormal & Chr(9) & "" & Chr(9) & Chr$(10))
                ReDim Preserve linha(1 To UBound(linha) + 1) As String
                ReDim Preserve tplinha(1 To UBound(tplinha) + 1) As String
                linha(UBound(linha)) = ""
                tplinha(UBound(tplinha)) = "N"
            End If
            
            vsfBoletim.AddItem (LinhaNormal & Chr(9) & _
                                "Data: " & _
                                Format(oRsIngressos.Fields("ope_dt_operacao"), "DD/MM/YYYY") & Chr(9) & Chr$(10))
            ReDim Preserve linha(1 To UBound(linha) + 1) As String
            ReDim Preserve tplinha(1 To UBound(tplinha) + 1) As String
            linha(UBound(linha)) = Alinha("Data: " & Format(oRsIngressos.Fields("ope_dt_operacao"), "DD/MM/YYYY"), COLUNAS_IMP, esquerda)
            tplinha(UBound(tplinha)) = "N"
            
            primeiro = False
            dtDia = oRsIngressos.Fields("ope_dt_operacao")
            
            dValorDia = 0
            iQtdeDia = 0

        End If
        
        vsfBoletim.AddItem (LinhaNormal & Chr(9) & _
                            Alinha(EliminaAcentos(oRsIngressos.Fields("igt_desc")), 20, esquerda, ".") & ": " & _
                            Alinha(Format(oRsIngressos.Fields("qtde"), "#,##0"), 5, Direita) & _
                            Alinha(Format(oRsIngressos.Fields("valor"), "#,##0.00"), 10, Direita) & Chr(9) & Chr$(10))
        ReDim Preserve linha(1 To UBound(linha) + 1) As String
        ReDim Preserve tplinha(1 To UBound(tplinha) + 1) As String
        linha(UBound(linha)) = Alinha(Alinha(EliminaAcentos(oRsIngressos.Fields("igt_desc")), 20, esquerda, ".") & ": " & _
                                      Alinha(Format(oRsIngressos.Fields("qtde"), "#,##0"), 5, Direita) & _
                                      Alinha(Format(oRsIngressos.Fields("valor"), "#,##0.00"), 10, Direita), COLUNAS_IMP, esquerda)
        tplinha(UBound(tplinha)) = "N"
        
        dValorTotal = dValorTotal + oRsIngressos.Fields("valor")
        iQtdeTotal = iQtdeTotal + oRsIngressos.Fields("qtde")
        dValorDia = dValorDia + oRsIngressos.Fields("valor")
        iQtdeDia = iQtdeDia + oRsIngressos.Fields("qtde")
        
        oRsIngressos.MoveNext
    Loop
    
    vsfBoletim.AddItem (LinhaNormal & Chr(9) & Alinha("-", 37, esquerda, "-") & Chr(9) & Chr$(10))
    ReDim Preserve linha(1 To UBound(linha) + 1) As String
    ReDim Preserve tplinha(1 To UBound(tplinha) + 1) As String
    linha(UBound(linha)) = Alinha("-", COLUNAS_IMP, esquerda, "-")
    tplinha(UBound(tplinha)) = "N"
    
    vsfBoletim.AddItem (LinhaNormal & Chr(9) & _
                        Alinha("Total dia", 20, esquerda, " ") & ": " & _
                        Alinha(Format(iQtdeDia, "#,##0"), 5, Direita) & _
                        Alinha(Format(dValorDia, "#,##0.00"), 10, Direita) & Chr(9) & Chr$(10))
    ReDim Preserve linha(1 To UBound(linha) + 1) As String
    ReDim Preserve tplinha(1 To UBound(tplinha) + 1) As String
    linha(UBound(linha)) = Alinha(Alinha("Total dia", 20, esquerda, " ") & ": " & _
                                  Alinha(Format(iQtdeDia, "#,##0"), 5, Direita) & _
                                  Alinha(Format(dValorDia, "#,##0.00"), 10, Direita), COLUNAS_IMP, esquerda)
    tplinha(UBound(tplinha)) = "N"
    
    vsfBoletim.AddItem (LinhaNormal & Chr(9) & _
                        Alinha("Total geral", 20, esquerda, " ") & ": " & _
                        Alinha(Format(iQtdeTotal, "#,##0"), 5, Direita) & _
                        Alinha(Format(dValorTotal, "#,##0.00"), 10, Direita) & Chr(9) & Chr$(10))
    ReDim Preserve linha(1 To UBound(linha) + 1) As String
    ReDim Preserve tplinha(1 To UBound(tplinha) + 1) As String
    linha(UBound(linha)) = Alinha(Alinha("Total geral", 20, esquerda, " ") & ": " & _
                                  Alinha(Format(iQtdeTotal, "#,##0"), 5, Direita) & _
                                  Alinha(Format(dValorTotal, "#,##0.00"), 10, Direita), COLUNAS_IMP, esquerda)
    tplinha(UBound(tplinha)) = "N"
     
     
    If Not clsBoletim.VendaIngressoTotal(oRsIngressos) Then
        MsgBox clsBoletim.MensagemErro, vbCritical, App.ProductName
        GoTo ImprimeVendaIngresso_fim
    End If
    
    vsfBoletim.AddItem (LinhaNormal & Chr(9) & "" & Chr(9) & Chr$(10))
    ReDim Preserve linha(1 To UBound(linha) + 1) As String
    ReDim Preserve tplinha(1 To UBound(tplinha) + 1) As String
    linha(UBound(linha)) = ""
    tplinha(UBound(tplinha)) = "N"
    
    vsfBoletim.AddItem (LinhaNormal & Chr(9) & "TOTAL NO PERIODO" & Chr(9) & Chr$(10))
    ReDim Preserve linha(1 To UBound(linha) + 1) As String
    ReDim Preserve tplinha(1 To UBound(tplinha) + 1) As String
    linha(UBound(linha)) = "TOTAL NO PERIODO"
    tplinha(UBound(tplinha)) = "N"
    
    dValorTotal = 0
    iQtdeTotal = 0
    
    dtDia = Empty
    primeiro = True
        
    Do While Not oRsIngressos.EOF()
        vsfBoletim.AddItem (LinhaNormal & Chr(9) & _
                            Alinha(EliminaAcentos(oRsIngressos.Fields("igt_desc")), 20, esquerda, ".") & ": " & _
                            Alinha(Format(oRsIngressos.Fields("qtde"), "#,##0"), 5, Direita) & _
                            Alinha(Format(oRsIngressos.Fields("valor"), "#,##0.00"), 10, Direita) & Chr(9) & Chr$(10))
        ReDim Preserve linha(1 To UBound(linha) + 1) As String
        ReDim Preserve tplinha(1 To UBound(tplinha) + 1) As String
        linha(UBound(linha)) = Alinha(Alinha(EliminaAcentos(oRsIngressos.Fields("igt_desc")), 20, esquerda, ".") & ": " & _
                                      Alinha(Format(oRsIngressos.Fields("qtde"), "#,##0"), 5, Direita) & _
                                      Alinha(Format(oRsIngressos.Fields("valor"), "#,##0.00"), 10, Direita), COLUNAS_IMP, esquerda)
        tplinha(UBound(tplinha)) = "N"
        
        dValorTotal = dValorTotal + oRsIngressos.Fields("valor")
        iQtdeTotal = iQtdeTotal + oRsIngressos.Fields("qtde")
        
        oRsIngressos.MoveNext
    Loop
    vsfBoletim.AddItem (LinhaNormal & Chr(9) & Alinha("-", 37, esquerda, "-") & Chr(9) & Chr$(10))
    ReDim Preserve linha(1 To UBound(linha) + 1) As String
    ReDim Preserve tplinha(1 To UBound(tplinha) + 1) As String
    linha(UBound(linha)) = Alinha("-", COLUNAS_IMP, esquerda, "-")
    tplinha(UBound(tplinha)) = "N"
    
    vsfBoletim.AddItem (LinhaNormal & Chr(9) & _
                        Alinha("Total geral", 20, esquerda, " ") & ": " & _
                        Alinha(Format(iQtdeTotal, "#,##0"), 5, Direita) & _
                        Alinha(Format(dValorTotal, "#,##0.00"), 10, Direita) & Chr(9) & Chr$(10))
    ReDim Preserve linha(1 To UBound(linha) + 1) As String
    ReDim Preserve tplinha(1 To UBound(tplinha) + 1) As String
    linha(UBound(linha)) = Alinha(Alinha("Total geral", 20, esquerda, " ") & ": " & _
                                  Alinha(Format(iQtdeTotal, "#,##0"), 5, Direita) & _
                                  Alinha(Format(dValorTotal, "#,##0.00"), 10, Direita), COLUNAS_IMP, esquerda)
    tplinha(UBound(tplinha)) = "N"
     
    ReDim Preserve linha(1 To UBound(linha) + 1) As String
    ReDim Preserve tplinha(1 To UBound(tplinha) + 1) As String
    linha(UBound(linha)) = ""
    tplinha(UBound(tplinha)) = "N"
    
    ReDim Preserve linha(1 To UBound(linha) + 1) As String
    ReDim Preserve tplinha(1 To UBound(tplinha) + 1) As String
    linha(UBound(linha)) = ""
    tplinha(UBound(tplinha)) = "N"
        
     
ImprimeVendaIngresso_fim:
    If oRsIngressos.State = 1 Then oRsIngressos.Close
    
ImprimeVendaIngresso_erro:
    DoEvents
    fraAviso.Visible = False
    vsfBoletim.ZOrder 0
    DoEvents
    
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    'If fechaImp(OPOSPOSPrinter1) Then
    'End If
End Sub

Private Sub MSComm_OnComm()
    Dim EventComm
    
    EventComm = MSComm.CommEvent
    
    Select Case EventComm
        Case 2
            'Recebeu um caractere, provavelmente o retorno
            'do status da impressora
            pbErroImp = False
        Case Is > 1000
            'Erro
            pbErroImp = True
    End Select
    
End Sub

Private Function TrataTipo(ByVal Tipo As String) As String
    TrataTipo = EliminaAcentos(Replace(Replace(Tipo, "PROMOÇÃO", "PROM"), "INTEIRA", "INT"))
End Function

Private Sub ImprimeVendaIngressoCombo(ByVal dDataIni As Date, ByVal dDataFim As Date)

    On Error GoTo ImprimeVendaIngressoCombo_erro
    
    Dim clsBoletim   As New clsBoletim
    Dim oRsIngressos As New ADODB.Recordset
    Dim dtDia        As Date
    Dim primeiro     As Boolean
    Dim dValorTotal  As Double
    Dim iQtdeTotal   As Integer
    Dim dValorDia    As Double
    Dim iQtdeDia     As Integer
    Dim wValorGeral As Double
    
    DoEvents
    fraAviso.ZOrder 0
    fraAviso.Visible = True
    DoEvents
        
    LinhaNormal = Chr$(27) & Chr$(33) & Chr$(10)
    LinhaTitulo = Chr$(27) & Chr$(33) & Chr$(60)
    LinhaDupla = Chr$(27) & Chr$(33) & Chr$(40)
    
    vsfBoletim.Clear
    vsfBoletim.Rows = 0
    vsfBoletim.FontName = "Courier"
    vsfBoletim.FontSize = 6
    vsfBoletim.GridLines = flexGridNone
    
    Set clsBoletim.ConexaoADO = dbConnect
    
    clsBoletim.dtIni = dDataIni
    clsBoletim.dtFim = dDataFim
    
    If Not clsBoletim.VendaIngresso(oRsIngressos) Then
        MsgBox clsBoletim.MensagemErro, vbCritical, App.ProductName
        GoTo ImprimeVendaIngresso_fim
    End If
    
    ReDim linha(1 To 1) As String
    ReDim tplinha(1 To 1) As String
    linha(UBound(linha)) = ""
    tplinha(UBound(tplinha)) = "N"
    
    If IsNull(oRsIngressos.Fields("ope_dt_operacao")) Then
        MsgBox "Não houve movimento no periodo selecionado!", vbInformation, App.ProductName
        oRsIngressos.Close
        GoTo ImprimeVendaIngresso_fim
    End If

    vsfBoletim.AddItem (LinhaTitulo & Chr(9) & "VENDAS INGRESSO" & Chr(9) & Chr$(10))
    ReDim Preserve linha(1 To UBound(linha) + 1) As String
    ReDim Preserve tplinha(1 To UBound(tplinha) + 1) As String
    linha(UBound(linha)) = "VENDAS INGRESSO"
    tplinha(UBound(tplinha)) = "T"
    
    vsfBoletim.AddItem (LinhaNormal & Chr(9) & "" & Chr(9) & Chr$(10))
    ReDim Preserve linha(1 To UBound(linha) + 1) As String
    ReDim Preserve tplinha(1 To UBound(tplinha) + 1) As String
    linha(UBound(linha)) = ""
    tplinha(UBound(tplinha)) = "N"
    
    
    dValorTotal = 0
    iQtdeTotal = 0
    dValorDia = 0
    iQtdeDia = 0
    
    dtDia = Empty
    primeiro = True
        
    Do While Not oRsIngressos.EOF()
        If dtDia <> oRsIngressos.Fields("ope_dt_operacao") Then
'            If Not primeiro Then
'                vsfBoletim.AddItem (LinhaNormal & Chr(9) & Alinha("-", 37, esquerda, "-") & Chr(9) & Chr$(10))
'                ReDim Preserve linha(1 To UBound(linha) + 1) As String
'                ReDim Preserve tpLinha(1 To UBound(tpLinha) + 1) As String
'                linha(UBound(linha)) = Alinha("-", COLUNAS_IMP, esquerda, "-")
'                tpLinha(UBound(tpLinha)) = "N"
'
'                vsfBoletim.AddItem (LinhaNormal & Chr(9) & _
'                                    Alinha("Total dia", 20, esquerda, " ") & ": " & _
'                                    Alinha(Format(iQtdeDia, "#,##0"), 5, Direita) & _
'                                    Alinha(Format(dValorDia, "#,##0.00"), 10, Direita) & Chr(9) & Chr$(10))
'                ReDim Preserve linha(1 To UBound(linha) + 1) As String
'                ReDim Preserve tpLinha(1 To UBound(tpLinha) + 1) As String
'                linha(UBound(linha)) = Alinha(Alinha("Total dia", 20, esquerda, " ") & ": " & _
'                                              Alinha(Format(iQtdeDia, "#,##0"), 5, Direita) & _
'                                              Alinha(Format(dValorDia, "#,##0.00"), 10, Direita), COLUNAS_IMP, esquerda)
'                tpLinha(UBound(tpLinha)) = "N"
'
'                vsfBoletim.AddItem (LinhaNormal & Chr(9) & "" & Chr(9) & Chr$(10))
'                ReDim Preserve linha(1 To UBound(linha) + 1) As String
'                ReDim Preserve tpLinha(1 To UBound(tpLinha) + 1) As String
'                linha(UBound(linha)) = ""
'                tpLinha(UBound(tpLinha)) = "N"
'            End If
            
'            vsfBoletim.AddItem (LinhaNormal & Chr(9) & _
'                                "Data: " & _
'                                Format(oRsIngressos.Fields("ope_dt_operacao"), "DD/MM/YYYY") & Chr(9) & Chr$(10))
'            ReDim Preserve linha(1 To UBound(linha) + 1) As String
'            ReDim Preserve tpLinha(1 To UBound(tpLinha) + 1) As String
'            linha(UBound(linha)) = Alinha("Data: " & Format(oRsIngressos.Fields("ope_dt_operacao"), "DD/MM/YYYY"), COLUNAS_IMP, esquerda)
'            tpLinha(UBound(tpLinha)) = "N"
'
            primeiro = False
            dtDia = oRsIngressos.Fields("ope_dt_operacao")
            
            dValorDia = 0
            iQtdeDia = 0

        End If
        
'        vsfBoletim.AddItem (LinhaNormal & Chr(9) & _
'                            Alinha(EliminaAcentos(oRsIngressos.Fields("igt_desc")), 20, esquerda, ".") & ": " & _
'                            Alinha(Format(oRsIngressos.Fields("qtde"), "#,##0"), 5, Direita) & _
'                            Alinha(Format(oRsIngressos.Fields("valor"), "#,##0.00"), 10, Direita) & Chr(9) & Chr$(10))
'        ReDim Preserve linha(1 To UBound(linha) + 1) As String
'        ReDim Preserve tpLinha(1 To UBound(tpLinha) + 1) As String
'        linha(UBound(linha)) = Alinha(Alinha(EliminaAcentos(oRsIngressos.Fields("igt_desc")), 20, esquerda, ".") & ": " & _
'                                      Alinha(Format(oRsIngressos.Fields("qtde"), "#,##0"), 5, Direita) & _
'                                      Alinha(Format(oRsIngressos.Fields("valor"), "#,##0.00"), 10, Direita), COLUNAS_IMP, esquerda)
'        tpLinha(UBound(tpLinha)) = "N"
'
        dValorTotal = dValorTotal + oRsIngressos.Fields("valor")
        iQtdeTotal = iQtdeTotal + oRsIngressos.Fields("qtde")
        dValorDia = dValorDia + oRsIngressos.Fields("valor")
        iQtdeDia = iQtdeDia + oRsIngressos.Fields("qtde")
        
        oRsIngressos.MoveNext
    Loop
    
    wValorGeral = dValorTotal
'    vsfBoletim.AddItem (LinhaNormal & Chr(9) & Alinha("-", 37, esquerda, "-") & Chr(9) & Chr$(10))
'    ReDim Preserve linha(1 To UBound(linha) + 1) As String
'    ReDim Preserve tpLinha(1 To UBound(tpLinha) + 1) As String
'    linha(UBound(linha)) = Alinha("-", COLUNAS_IMP, esquerda, "-")
'    tpLinha(UBound(tpLinha)) = "N"
'
'    vsfBoletim.AddItem (LinhaNormal & Chr(9) & _
'                        Alinha("Total dia", 20, esquerda, " ") & ": " & _
'                        Alinha(Format(iQtdeDia, "#,##0"), 5, Direita) & _
'                        Alinha(Format(dValorDia, "#,##0.00"), 10, Direita) & Chr(9) & Chr$(10))
'    ReDim Preserve linha(1 To UBound(linha) + 1) As String
'    ReDim Preserve tpLinha(1 To UBound(tpLinha) + 1) As String
'    linha(UBound(linha)) = Alinha(Alinha("Total dia", 20, esquerda, " ") & ": " & _
'                                  Alinha(Format(iQtdeDia, "#,##0"), 5, Direita) & _
'                                  Alinha(Format(dValorDia, "#,##0.00"), 10, Direita), COLUNAS_IMP, esquerda)
'    tpLinha(UBound(tpLinha)) = "N"
'
'    vsfBoletim.AddItem (LinhaNormal & Chr(9) & _
'                        Alinha("Total geral", 20, esquerda, " ") & ": " & _
'                        Alinha(Format(iQtdeTotal, "#,##0"), 5, Direita) & _
'                        Alinha(Format(dValorTotal, "#,##0.00"), 10, Direita) & Chr(9) & Chr$(10))
'    ReDim Preserve linha(1 To UBound(linha) + 1) As String
'    ReDim Preserve tpLinha(1 To UBound(tpLinha) + 1) As String
'    linha(UBound(linha)) = Alinha(Alinha("Total geral", 20, esquerda, " ") & ": " & _
'                                  Alinha(Format(iQtdeTotal, "#,##0"), 5, Direita) & _
'                                  Alinha(Format(dValorTotal, "#,##0.00"), 10, Direita), COLUNAS_IMP, esquerda)
'    tpLinha(UBound(tpLinha)) = "N"
     
     
    If Not clsBoletim.VendaIngressoTotal(oRsIngressos) Then
        MsgBox clsBoletim.MensagemErro, vbCritical, App.ProductName
        GoTo ImprimeVendaIngresso_fim
    End If
    
    'vsfBoletim.AddItem (LinhaNormal & Chr(9) & "" & Chr(9) & Chr$(10))
    ReDim Preserve linha(1 To UBound(linha) + 1) As String
    ReDim Preserve tplinha(1 To UBound(tplinha) + 1) As String
    linha(UBound(linha)) = ""
    tplinha(UBound(tplinha)) = "N"
    
    vsfBoletim.AddItem (LinhaNormal & Chr(9) & "TOTAL NO PERIODO" & Chr(9) & Chr$(10))
    ReDim Preserve linha(1 To UBound(linha) + 1) As String
    ReDim Preserve tplinha(1 To UBound(tplinha) + 1) As String
    linha(UBound(linha)) = "TOTAL NO PERIODO"
    tplinha(UBound(tplinha)) = "N"
    
    dValorTotal = 0
    iQtdeTotal = 0
    
    dtDia = Empty
    primeiro = True
        
    Do While Not oRsIngressos.EOF()
        vsfBoletim.AddItem (LinhaNormal & Chr(9) & _
                            Alinha(EliminaAcentos(oRsIngressos.Fields("igt_desc")), 20, esquerda, ".") & ": " & _
                            Alinha(Format(oRsIngressos.Fields("qtde"), "#,##0"), 5, Direita) & _
                            Alinha(Format(oRsIngressos.Fields("valor"), "#,##0.00"), 10, Direita) & Chr(9) & Chr$(10))
        ReDim Preserve linha(1 To UBound(linha) + 1) As String
        ReDim Preserve tplinha(1 To UBound(tplinha) + 1) As String
        linha(UBound(linha)) = Alinha(Alinha(EliminaAcentos(oRsIngressos.Fields("igt_desc")), 20, esquerda, ".") & ": " & _
                                      Alinha(Format(oRsIngressos.Fields("qtde"), "#,##0"), 5, Direita) & _
                                      Alinha(Format(oRsIngressos.Fields("valor"), "#,##0.00"), 10, Direita), COLUNAS_IMP, esquerda)
        tplinha(UBound(tplinha)) = "N"
        
        dValorTotal = dValorTotal + oRsIngressos.Fields("valor")
        iQtdeTotal = iQtdeTotal + oRsIngressos.Fields("qtde")
        
        oRsIngressos.MoveNext
    Loop
    vsfBoletim.AddItem (LinhaNormal & Chr(9) & Alinha("-", 37, esquerda, "-") & Chr(9) & Chr$(10))
    ReDim Preserve linha(1 To UBound(linha) + 1) As String
    ReDim Preserve tplinha(1 To UBound(tplinha) + 1) As String
    linha(UBound(linha)) = Alinha("-", COLUNAS_IMP, esquerda, "-")
    tplinha(UBound(tplinha)) = "N"
    
    vsfBoletim.AddItem (LinhaNormal & Chr(9) & _
                        Alinha("Sub-Total Geral", 20, esquerda, " ") & ": " & _
                        Alinha(Format(iQtdeTotal, "#,##0"), 5, Direita) & _
                        Alinha(Format(dValorTotal, "#,##0.00"), 10, Direita) & Chr(9) & Chr$(10))
    ReDim Preserve linha(1 To UBound(linha) + 1) As String
    ReDim Preserve tplinha(1 To UBound(tplinha) + 1) As String
    linha(UBound(linha)) = Alinha(Alinha("Total geral", 20, esquerda, " ") & ": " & _
                                  Alinha(Format(iQtdeTotal, "#,##0"), 5, Direita) & _
                                  Alinha(Format(dValorTotal, "#,##0.00"), 10, Direita), COLUNAS_IMP, esquerda)
    tplinha(UBound(tplinha)) = "N"
     
    ReDim Preserve linha(1 To UBound(linha) + 1) As String
    ReDim Preserve tplinha(1 To UBound(tplinha) + 1) As String
    linha(UBound(linha)) = ""
    tplinha(UBound(tplinha)) = "N"
    
    ReDim Preserve linha(1 To UBound(linha) + 1) As String
    ReDim Preserve tplinha(1 To UBound(tplinha) + 1) As String
    linha(UBound(linha)) = ""
    tplinha(UBound(tplinha)) = "N"
    vsfBoletim.AddItem Chr(9) + "" + Chr(9)
    
    
    ReDim Preserve linha(1 To UBound(linha) + 1) As String
    ReDim Preserve tplinha(1 To UBound(tplinha) + 1) As String
    linha(UBound(linha)) = ""
    tplinha(UBound(tplinha)) = "N"
        
     
    On Error GoTo ImprimeVendaCombo_erro
    
    'Dim clsBoletim  As New clsBoletim
    Dim oRsCombos   As New ADODB.Recordset
    'Dim dtDia       As Date
    'Dim primeiro    As Boolean
    'Dim dValorTotal As Double
    'Dim iQtdeTotal  As Integer
    'Dim dValorDia   As Double
    'Dim iQtdeDia    As Integer
        
    'DoEvents
    'fraAviso.ZOrder 0
    'fraAviso.Visible = True
    'DoEvents
        
    'LinhaNormal = Chr$(27) & Chr$(33) & Chr$(10)
    'LinhaTitulo = Chr$(27) & Chr$(33) & Chr$(60)
    'LinhaDupla = Chr$(27) & Chr$(33) & Chr$(40)
    
    'vsfBoletim.Clear
    'vsfBoletim.Rows = 0
    'vsfBoletim.Cols = 2
    'vsfBoletim.FontName = "Courier"
    'vsfBoletim.FontSize = 6
    'vsfBoletim.GridLines = flexGridNone
    
    'Set clsBoletim.ConexaoADO = dbConnect
    
    'clsBoletim.dtIni = dDataIni
    'clsBoletim.dtFim = dDataFim
    
    If Not clsBoletim.VendaCombo(oRsCombos) Then
        MsgBox clsBoletim.MensagemErro, vbCritical, App.ProductName
        GoTo ImprimeVendaCombo_fim
    End If
    
    'ReDim linha(1 To 1) As String
    'ReDim tpLinha(1 To 1) As String
    ReDim Preserve linha(1 To UBound(linha) + 1) As String
    ReDim Preserve tplinha(1 To UBound(tplinha) + 1) As String
    
    linha(UBound(linha)) = ""
    tplinha(UBound(tplinha)) = "N"
    
    If IsNull(oRsCombos.Fields("ope_dt_operacao")) Then
        MsgBox "Não houve movimento no periodo selecionado!", vbInformation, App.ProductName
        oRsCombos.Close
        GoTo ImprimeVendaCombo_fim
    End If

    vsfBoletim.AddItem (LinhaTitulo & Chr(9) & "VENDAS COMBO" & Chr(9) & Chr$(10))
    ReDim Preserve linha(1 To UBound(linha) + 1) As String
    ReDim Preserve tplinha(1 To UBound(tplinha) + 1) As String
    linha(UBound(linha)) = "VENDAS COMBO"
    tplinha(UBound(tplinha)) = "T"
    
    vsfBoletim.AddItem (LinhaNormal & Chr(9) & "" & Chr(9) & Chr$(10))
    ReDim Preserve linha(1 To UBound(linha) + 1) As String
    ReDim Preserve tplinha(1 To UBound(tplinha) + 1) As String
    linha(UBound(linha)) = ""
    tplinha(UBound(tplinha)) = "N"
    
    dValorTotal = 0
    iQtdeTotal = 0
    dValorDia = 0
    iQtdeDia = 0
    
    dtDia = Empty
    primeiro = True
        
    Do While Not oRsCombos.EOF()
        If dtDia <> oRsCombos.Fields("ope_dt_operacao") Then
'            If Not primeiro Then
'                vsfBoletim.AddItem (LinhaNormal & Chr(9) & Alinha("-", 37, esquerda, "-") & Chr(9) & Chr$(10))
'                ReDim Preserve linha(1 To UBound(linha) + 1) As String
'                ReDim Preserve tpLinha(1 To UBound(tpLinha) + 1) As String
'                linha(UBound(linha)) = Alinha("-", COLUNAS_IMP, esquerda, "-")
'                tpLinha(UBound(tpLinha)) = "N"
'
'                vsfBoletim.AddItem (LinhaNormal & Chr(9) & _
'                                    Alinha("Total dia", 20, esquerda, " ") & ": " & _
'                                    Alinha(Format(iQtdeDia, "#,##0"), 5, Direita) & _
'                                    Alinha(Format(dValorDia, "#,##0.00"), 10, Direita) & Chr(9) & Chr$(10))
'                ReDim Preserve linha(1 To UBound(linha) + 1) As String
'                ReDim Preserve tpLinha(1 To UBound(tpLinha) + 1) As String
'                linha(UBound(linha)) = Alinha(Alinha("Total dia", 20, esquerda, " ") & ": " & _
'                                              Alinha(Format(iQtdeDia, "#,##0"), 5, Direita) & _
'                                              Alinha(Format(dValorDia, "#,##0.00"), 10, Direita), COLUNAS_IMP, esquerda)
'                tpLinha(UBound(tpLinha)) = "N"
'
'                vsfBoletim.AddItem (LinhaNormal & Chr(9) & "" & Chr(9) & Chr$(10))
'                ReDim Preserve linha(1 To UBound(linha) + 1) As String
'                ReDim Preserve tpLinha(1 To UBound(tpLinha) + 1) As String
'                linha(UBound(linha)) = ""
'                tpLinha(UBound(tpLinha)) = "N"
'            End If
'
'            vsfBoletim.AddItem (LinhaNormal & Chr(9) & _
'                                "Data: " & _
'                                Format(oRsCombos.Fields("ope_dt_operacao"), "DD/MM/YYYY") & Chr(9) & Chr$(10))
'            ReDim Preserve linha(1 To UBound(linha) + 1) As String
'            ReDim Preserve tpLinha(1 To UBound(tpLinha) + 1) As String
'            linha(UBound(linha)) = Alinha("Data: " & Format(oRsCombos.Fields("ope_dt_operacao"), "DD/MM/YYYY"), COLUNAS_IMP, esquerda)
'            tpLinha(UBound(tpLinha)) = "N"
'
            primeiro = False
            dtDia = oRsCombos.Fields("ope_dt_operacao")
            
            dValorDia = 0
            iQtdeDia = 0

        End If
        
'        vsfBoletim.AddItem (LinhaNormal & Chr(9) & _
'                            Alinha(EliminaAcentos(oRsCombos.Fields("cbo_nm")), 20, esquerda, ".") & ": " & _
'                            Alinha(Format(oRsCombos.Fields("qtde"), "#,##0"), 5, Direita) & _
'                            Alinha(Format(oRsCombos.Fields("valor"), "#,##0.00"), 10, Direita) & Chr(9) & Chr$(10))
'        ReDim Preserve linha(1 To UBound(linha) + 1) As String
'        ReDim Preserve tpLinha(1 To UBound(tpLinha) + 1) As String
'        linha(UBound(linha)) = Alinha(Alinha(EliminaAcentos(oRsCombos.Fields("cbo_nm")), 20, esquerda, ".") & ": " & _
'                                      Alinha(Format(oRsCombos.Fields("qtde"), "#,##0"), 5, Direita) & _
'                                      Alinha(Format(oRsCombos.Fields("valor"), "#,##0.00"), 10, Direita), COLUNAS_IMP, esquerda)
'        tpLinha(UBound(tpLinha)) = "N"
'
        dValorTotal = dValorTotal + oRsCombos.Fields("valor")
        iQtdeTotal = iQtdeTotal + oRsCombos.Fields("qtde")
        dValorDia = dValorDia + oRsCombos.Fields("valor")
        iQtdeDia = iQtdeDia + oRsCombos.Fields("qtde")
        
        oRsCombos.MoveNext
    Loop
    
    wValorGeral = wValorGeral + dValorTotal
'    vsfBoletim.AddItem (LinhaNormal & Chr(9) & Alinha("-", 37, esquerda, "-") & Chr(9) & Chr$(10))
'    ReDim Preserve linha(1 To UBound(linha) + 1) As String
'    ReDim Preserve tpLinha(1 To UBound(tpLinha) + 1) As String
'    linha(UBound(linha)) = Alinha("-", COLUNAS_IMP, esquerda, "-")
'    tpLinha(UBound(tpLinha)) = "N"
'
'    vsfBoletim.AddItem (LinhaNormal & Chr(9) & _
'                        Alinha("Total dia", 20, esquerda, " ") & ": " & _
'                        Alinha(Format(iQtdeDia, "#,##0"), 5, Direita) & _
'                        Alinha(Format(dValorDia, "#,##0.00"), 10, Direita) & Chr(9) & Chr$(10))
'    ReDim Preserve linha(1 To UBound(linha) + 1) As String
'    ReDim Preserve tpLinha(1 To UBound(tpLinha) + 1) As String
'    linha(UBound(linha)) = Alinha(Alinha("Total dia", 20, esquerda, " ") & ": " & _
'                                  Alinha(Format(iQtdeDia, "#,##0"), 5, Direita) & _
'                                  Alinha(Format(dValorDia, "#,##0.00"), 10, Direita), COLUNAS_IMP, esquerda)
'    tpLinha(UBound(tpLinha)) = "N"
    
'    vsfBoletim.AddItem (LinhaNormal & Chr(9) & _
'                        Alinha("Total Geral", 20, esquerda, " ") & ": " & _
'                        Alinha(Format(iQtdeTotal, "#,##0"), 5, Direita) & _
'                        Alinha(Format(dValorTotal, "#,##0.00"), 10, Direita) & Chr(9) & Chr$(10))
'    ReDim Preserve linha(1 To UBound(linha) + 1) As String
'    ReDim Preserve tpLinha(1 To UBound(tpLinha) + 1) As String
'    linha(UBound(linha)) = Alinha(Alinha("Total geral", 20, esquerda, " ") & ": " & _
'                                  Alinha(Format(iQtdeTotal, "#,##0"), 5, Direita) & _
'                                  Alinha(Format(dValorTotal, "#,##0.00"), 10, Direita), COLUNAS_IMP, esquerda)
'    tpLinha(UBound(tpLinha)) = "N"
'
     
    If Not clsBoletim.VendaComboTotal(oRsCombos) Then
        MsgBox clsBoletim.MensagemErro, vbCritical, App.ProductName
        GoTo ImprimeVendaCombo_fim
    End If
    
'    vsfBoletim.AddItem (LinhaNormal & Chr(9) & "" & Chr(9) & Chr$(10))
'    ReDim Preserve linha(1 To UBound(linha) + 1) As String
'    ReDim Preserve tpLinha(1 To UBound(tpLinha) + 1) As String
'    linha(UBound(linha)) = ""
'    tpLinha(UBound(tpLinha)) = "N"
    
    vsfBoletim.AddItem (LinhaNormal & Chr(9) & "TOTAL NO PERIODO" & Chr(9) & Chr$(10))
    ReDim Preserve linha(1 To UBound(linha) + 1) As String
    ReDim Preserve tplinha(1 To UBound(tplinha) + 1) As String
    linha(UBound(linha)) = "TOTAL NO PERIODO"
    tplinha(UBound(tplinha)) = "N"
    
    dValorTotal = 0
    iQtdeTotal = 0
    
    dtDia = Empty
    primeiro = True
        
    Do While Not oRsCombos.EOF()
        vsfBoletim.AddItem (LinhaNormal & Chr(9) & _
                            Alinha(EliminaAcentos(oRsCombos.Fields("cbo_nm")), 20, esquerda, ".") & ": " & _
                            Alinha(Format(oRsCombos.Fields("qtde"), "#,##0"), 5, Direita) & _
                            Alinha(Format(oRsCombos.Fields("valor"), "#,##0.00"), 10, Direita) & Chr(9) & Chr$(10))
        ReDim Preserve linha(1 To UBound(linha) + 1) As String
        ReDim Preserve tplinha(1 To UBound(tplinha) + 1) As String
        linha(UBound(linha)) = Alinha(Alinha(EliminaAcentos(oRsCombos.Fields("cbo_nm")), 20, esquerda, ".") & ": " & _
                                      Alinha(Format(oRsCombos.Fields("qtde"), "#,##0"), 5, Direita) & _
                                      Alinha(Format(oRsCombos.Fields("valor"), "#,##0.00"), 10, Direita), COLUNAS_IMP, esquerda)
        tplinha(UBound(tplinha)) = "N"
        
        dValorTotal = dValorTotal + oRsCombos.Fields("valor")
        iQtdeTotal = iQtdeTotal + oRsCombos.Fields("qtde")
        
        oRsCombos.MoveNext
    Loop
    vsfBoletim.AddItem (LinhaNormal & Chr(9) & Alinha("-", 37, esquerda, "-") & Chr(9) & Chr$(10))
    ReDim Preserve linha(1 To UBound(linha) + 1) As String
    ReDim Preserve tplinha(1 To UBound(tplinha) + 1) As String
    linha(UBound(linha)) = Alinha("-", COLUNAS_IMP, esquerda, "-")
    tplinha(UBound(tplinha)) = "N"
    
    vsfBoletim.AddItem (LinhaNormal & Chr(9) & _
                        Alinha("Sub-Total Geral", 20, esquerda, " ") & ": " & _
                        Alinha(Format(iQtdeTotal, "#,##0"), 5, Direita) & _
                        Alinha(Format(dValorTotal, "#,##0.00"), 10, Direita) & Chr(9) & Chr$(10))
    ReDim Preserve linha(1 To UBound(linha) + 1) As String
    ReDim Preserve tplinha(1 To UBound(tplinha) + 1) As String
    linha(UBound(linha)) = Alinha(Alinha("Total geral", 20, esquerda, " ") & ": " & _
                                  Alinha(Format(iQtdeTotal, "#,##0"), 5, Direita) & _
                                  Alinha(Format(dValorTotal, "#,##0.00"), 10, Direita), COLUNAS_IMP, esquerda)
    tplinha(UBound(tplinha)) = "N"
     
    ReDim Preserve linha(1 To UBound(linha) + 1) As String
    ReDim Preserve tplinha(1 To UBound(tplinha) + 1) As String
    linha(UBound(linha)) = ""
    tplinha(UBound(tplinha)) = "N"
    
    ReDim Preserve linha(1 To UBound(linha) + 1) As String
    ReDim Preserve tplinha(1 To UBound(tplinha) + 1) As String
    linha(UBound(linha)) = ""
    tplinha(UBound(tplinha)) = "N"
    
    'vsfBoletim.AddItem (LinhaNormal & Chr(9) & "TOTAL NO PERIODO" & Chr(9) & Chr$(10))
    'ReDim Preserve linha(1 To UBound(linha) + 1) As String
    'ReDim Preserve tpLinha(1 To UBound(tpLinha) + 1) As String
    'linha(UBound(linha)) = "TOTAL NO PERIODO"
    'tpLinha(UBound(tpLinha)) = "N"
    
    vsfBoletim.AddItem (LinhaNormal & Chr(9) & "" & Chr(9) & Chr$(10))
    ReDim Preserve linha(1 To UBound(linha) + 1) As String
    ReDim Preserve tplinha(1 To UBound(tplinha) + 1) As String
    linha(UBound(linha)) = ""
    tplinha(UBound(tplinha)) = "N"

vsfBoletim.AddItem (LinhaNormal & Chr(9) & _
                            Alinha("TOTAL NO PERIODO", 20, esquerda, ".") & ": " & _
                            Alinha(Format(wValorGeral, "#,##0.00"), 15, Direita) & Chr(9) & Chr$(10))
        ReDim Preserve linha(1 To UBound(linha) + 1) As String
        ReDim Preserve tplinha(1 To UBound(tplinha) + 1) As String
        linha(UBound(linha)) = Alinha("TOTAL NO PERIODO", 20, esquerda, ".") & ": " & _
                            Alinha(Format(wValorGeral, "#,##0.00"), 15, Direita) & Chr(9) & Chr$(10)
        tplinha(UBound(tplinha)) = "N"


        
     
ImprimeVendaCombo_fim:
    If oRsCombos.State = 1 Then oRsCombos.Close
    
ImprimeVendaCombo_erro:
    DoEvents
    fraAviso.Visible = False
    vsfBoletim.ZOrder 0
    DoEvents
     
     
ImprimeVendaIngresso_fim:
    If oRsIngressos.State = 1 Then oRsIngressos.Close
    
ImprimeVendaIngressoCombo_erro:
    DoEvents
    fraAviso.Visible = False
    vsfBoletim.ZOrder 0
    DoEvents
    
End Sub


