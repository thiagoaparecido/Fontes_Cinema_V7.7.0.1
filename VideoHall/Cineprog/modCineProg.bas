Attribute VB_Name = "modCineProg"
Option Explicit

Global Const senhaProteq = "31Z16XQ03"
Global Const senhaBanco = "31Z16XQ03"

Global Const STATUS_OK = 0 '* API call was succesfull        *

Public wHorario1 As String
Public wHorario2 As String
   
Public gsBandoDados  As String
Public gsArqsVideo   As String
Public gbCheck       As Integer

Public clsPC         As New clsPainelControle

Public gRSDados1     As New ADODB.Recordset
Public gRSDados2     As New ADODB.Recordset

Public dtReferencia1 As Date
Public dtReferencia2 As Date
Public dtStrRef1     As String
Public dtStrRef2     As String

Public transicao       As Long
Public intermitencia   As Long
Public velocMsg        As Long
Public vendaAndes      As Long
Public vendaDepois     As Long
Public hrLimitePeriodo As Date
Public telaSessoes     As Boolean
Public telaFilme       As Boolean
Public telaPrecos      As Boolean
Public telaTrailer     As Boolean
Public telaImagem      As Boolean
Public mensagem        As String

Public iqtdTelas As Integer

Public corFundT1Filme      As Long
Public corTextT1Filme      As Long
Public corFundT1Hora       As Long
Public corTextT1Hora       As Long
Public corFundT1Tutulo1    As Long
Public corTextT1Tutulo1    As Long
Public corFundT1Titulo2    As Long
Public corTextT1Titulo2    As Long
Public corFundT1Lin1       As Long
Public corTextT1Lin1       As Long
Public corFundT1Lin2       As Long
Public corTextT1Lin2       As Long
Public corFundT1Mensagem   As Long
Public corTextT1Mensagem   As Long
Public corFundT2Filme1     As Long
Public corTextT2Filme1     As Long
Public corFundT2Filme2     As Long
Public corTextT2Filme2     As Long
Public corFundT2Titulo1    As Long
Public corTextT2Titulo1    As Long
Public corFundT2Titulo2    As Long
Public corTextT2Titulo2    As Long
Public corFundT2Sessao1    As Long
Public corTextT2Sessao1    As Long
Public corFundT2Sessao2    As Long
Public corTextT2Sessao2    As Long
Public corFundT2Sala1      As Long
Public corTextT2Sala1      As Long
Public corFundT2Sala2      As Long
Public corTextT2Sala2      As Long
Public corFundT2Sessoes1   As Long
Public corTextT2Sessoes1   As Long
Public corFundT2Sessoes1L1 As Long
Public corTextT2Sessoes1L1 As Long
Public corFundT2Sessoes1L2 As Long
Public corTextT2Sessoes1L2 As Long
Public corFundT2Sessoes2   As Long
Public corTextT2Sessoes2   As Long
Public corFundT2Sessoes2L1 As Long
Public corTextT2Sessoes2L1 As Long
Public corFundT2Sessoes2L2 As Long
Public corTextT2Sessoes2L2 As Long
Public corFundT2Mensagem   As Long
Public corTextT2Mensagem   As Long
Public corFundT3Hora       As Long
Public corTextT3Hora       As Long
Public corFundT3Data       As Long
Public corTextT3Data       As Long
Public corFundT3TituloTela As Long
Public corTextT3TituloTela As Long
Public corFundT3Titulo     As Long
Public corTextT3Titulo     As Long
Public corFundT3Lin1       As Long
Public corTextT3Lin1       As Long
Public corFundT3Lin2       As Long
Public corTextT3Lin2       As Long
Public corFundT3Mensagem   As Long
Public corTextT3Mensagem   As Long
Public corFundLOTADO       As Long
Public corTextLOTADO         As Long
Public wHorarioLimite       As String

Public Type tp_proxSessao
   horario  As Date
   filme    As String
   sala     As String
   codFilme As Long
   codSala  As Long
   sessao   As Long
End Type

Public Type tp_horarios
   horario As Date
   sessao  As Long
End Type

Public Type tp_filme
   codFilme   As Long
   descFilme  As String
   codSala    As Long
   descSala   As String
   censura    As String
   horarios() As tp_horarios
End Type

Public Type tp_preco
   descricao    As String
   vlrIntManha  As Double
   vlrIntTarde  As Double
   vlrMeiaManha As Double
   vlrMeiaTarde As Double
   promocional  As Boolean
End Type

Type tp_precos
   descFilme As String
   codFilme  As Long
   precos()  As tp_preco
End Type

Public proxSessoes() As tp_proxSessao
Public filmes()      As tp_filme
Public precos()      As tp_precos
Public trailers()    As String
Public imagens()     As String

Sub Main()
      
   dtReferencia1 = CDate("01/01/1900")
   dtReferencia2 = DateAdd("d", 1, dtReferencia1)
   
   dtStrRef1 = Format(dtReferencia1, "Short Date")
   dtStrRef2 = Format(dtReferencia2, "Short Date")
   If Not carregaParametros() Then
      End
   End If
   If telaSessoes Then
      Call carregaProxSessoes
   End If
   If telaFilme Then
      Call carregaFilmes
   End If
   If telaPrecos Then
      Call carregaPrecos
   End If
      
   If telaTrailer Then
      Call caregaTrailers
   End If
   
   If telaImagem Then
      caregaImagens
   End If
   
   Load frmApresent
    
End Sub

Public Function carregaParametros() As Boolean
   Dim sMsg    As String
   Dim strSql  As String
   
   On Error GoTo TrataErro
   
   carregaParametros = False
   
   gRSDados1.Open "tb_parametros", "File Name=" + App.Path + "\Cinema.Udl", adOpenDynamic, adLockReadOnly, adCmdTable
   
   'gRSDados1.Open strSql, "File Name=" + App.Path + "\Cinema.Udl", adOpenDynamic, adLockReadOnly, adCmdText
   
   If gRSDados1.EOF And gRSDados1.BOF Then
      MsgBox "Parâmetros do sistema não encontrados.", vbCritical, App.ProductName
      gRSDados1.Close
      'Call gFechaBase
      End
   End If
   
   transicao = gRSDados1.Fields("transicao").Value
   intermitencia = gRSDados1.Fields("intermitencia").Value
   velocMsg = gRSDados1.Fields("velocMsg").Value
   vendaAndes = gRSDados1.Fields("vendaAndes").Value
   vendaDepois = gRSDados1.Fields("vendaDepois").Value
   hrLimitePeriodo = gRSDados1.Fields("hrLimitePeriodo").Value
   telaSessoes = gRSDados1.Fields("telaSessoes").Value
   telaFilme = gRSDados1.Fields("telaFilme").Value
   telaPrecos = gRSDados1.Fields("telaPrecos").Value
   telaTrailer = gRSDados1.Fields("telaTrailer").Value
   telaImagem = gRSDados1.Fields("telaImagem").Value
   mensagem = gRSDados1.Fields("mensagem").Value
   
   corFundT1Filme = gRSDados1.Fields("corFundT1Filme").Value
   corTextT1Filme = gRSDados1.Fields("corTextT1Filme").Value
   corFundT1Hora = gRSDados1.Fields("corFundT1Hora").Value
   corTextT1Hora = gRSDados1.Fields("corTextT1Hora").Value
   corFundT1Tutulo1 = gRSDados1.Fields("corFundT1Tutulo1").Value
   corTextT1Tutulo1 = gRSDados1.Fields("corTextT1Tutulo1").Value
   corFundT1Titulo2 = gRSDados1.Fields("corFundT1Titulo2").Value
   corTextT1Titulo2 = gRSDados1.Fields("corTextT1Titulo2").Value
   corFundT1Lin1 = gRSDados1.Fields("corFundT1Lin1").Value
   corTextT1Lin1 = gRSDados1.Fields("corTextT1Lin1").Value
   corFundT1Lin2 = gRSDados1.Fields("corFundT1Lin2").Value
   corTextT1Lin2 = gRSDados1.Fields("corTextT1Lin2").Value
   corFundT1Mensagem = gRSDados1.Fields("corFundT1Mensagem").Value
   corTextT1Mensagem = gRSDados1.Fields("corTextT1Mensagem").Value
   corFundT2Filme1 = gRSDados1.Fields("corFundT2Filme1").Value
   corTextT2Filme1 = gRSDados1.Fields("corTextT2Filme1").Value
   corFundT2Filme2 = gRSDados1.Fields("corFundT2Filme2").Value
   corTextT2Filme2 = gRSDados1.Fields("corTextT2Filme2").Value
   corFundT2Titulo1 = gRSDados1.Fields("corFundT2Titulo1").Value
   corTextT2Titulo1 = gRSDados1.Fields("corTextT2Titulo1").Value
   corFundT2Titulo2 = gRSDados1.Fields("corFundT2Titulo2").Value
   corTextT2Titulo2 = gRSDados1.Fields("corTextT2Titulo2").Value
   corFundT2Sessao1 = gRSDados1.Fields("corFundT2Sessao1").Value
   corTextT2Sessao1 = gRSDados1.Fields("corTextT2Sessao1").Value
   corFundT2Sessao2 = gRSDados1.Fields("corFundT2Sessao2").Value
   corTextT2Sessao2 = gRSDados1.Fields("corTextT2Sessao2").Value
   corFundT2Sala1 = gRSDados1.Fields("corFundT2Sala1").Value
   corTextT2Sala1 = gRSDados1.Fields("corTextT2Sala1").Value
   corFundT2Sala2 = gRSDados1.Fields("corFundT2Sala2").Value
   corTextT2Sala2 = gRSDados1.Fields("corTextT2Sala2").Value
   corFundT2Sessoes1 = gRSDados1.Fields("corFundT2Sessoes1").Value
   corTextT2Sessoes1 = gRSDados1.Fields("corTextT2Sessoes1").Value
   corFundT2Sessoes1L1 = gRSDados1.Fields("corFundT2Sessoes1L1").Value
   corTextT2Sessoes1L1 = gRSDados1.Fields("corTextT2Sessoes1L1").Value
   corFundT2Sessoes1L2 = gRSDados1.Fields("corFundT2Sessoes1L2").Value
   corTextT2Sessoes1L2 = gRSDados1.Fields("corTextT2Sessoes1L2").Value
   corFundT2Sessoes2 = gRSDados1.Fields("corFundT2Sessoes2").Value
   corTextT2Sessoes2 = gRSDados1.Fields("corTextT2Sessoes2").Value
   corFundT2Sessoes2L1 = gRSDados1.Fields("corFundT2Sessoes2L1").Value
   corTextT2Sessoes2L1 = gRSDados1.Fields("corTextT2Sessoes2L1").Value
   corFundT2Sessoes2L2 = gRSDados1.Fields("corFundT2Sessoes2L2").Value
   corTextT2Sessoes2L2 = gRSDados1.Fields("corTextT2Sessoes2L2").Value
   corFundT2Mensagem = gRSDados1.Fields("corFundT2Mensagem").Value
   corTextT2Mensagem = gRSDados1.Fields("corTextT2Mensagem").Value
   corFundT3Hora = gRSDados1.Fields("corFundT3Hora").Value
   corTextT3Hora = gRSDados1.Fields("corTextT3Hora").Value
   corFundT3Data = gRSDados1.Fields("corFundT3Data").Value
   corTextT3Data = gRSDados1.Fields("corTextT3Data").Value
   corFundT3TituloTela = gRSDados1.Fields("corFundT3TituloTela").Value
   corTextT3TituloTela = gRSDados1.Fields("corTextT3TituloTela").Value
   corFundT3Titulo = gRSDados1.Fields("corFundT3Titulo").Value
   corTextT3Titulo = gRSDados1.Fields("corTextT3Titulo").Value
   corFundT3Lin1 = gRSDados1.Fields("corFundT3Lin1").Value
   corTextT3Lin1 = gRSDados1.Fields("corTextT3Lin1").Value
   corFundT3Lin2 = gRSDados1.Fields("corFundT3Lin2").Value
   corTextT3Lin2 = gRSDados1.Fields("corTextT3Lin2").Value
   corFundT3Mensagem = gRSDados1.Fields("corFundT3Mensagem").Value
   corTextT3Mensagem = gRSDados1.Fields("corTextT3Mensagem").Value
   corFundLOTADO = gRSDados1.Fields("corFundLOTADO").Value
   corTextLOTADO = gRSDados1.Fields("corTextLOTADO").Value
   gRSDados1.Close
      
    iqtdTelas = 6

    If telaSessoes = True Then iqtdTelas = iqtdTelas - 1
    If telaFilme = True Then iqtdTelas = iqtdTelas - 1
    If telaPrecos = True Then iqtdTelas = iqtdTelas - 1
    If telaTrailer = True Then iqtdTelas = iqtdTelas - 1
    If telaImagem = True Then iqtdTelas = iqtdTelas - 1
      
   carregaParametros = True
   
   Exit Function
   
TrataErro:
   If gRSDados1.State = adStateOpen Then
       gRSDados1.Close
   End If
   
   sMsg = "Ocorreu um erro em carregaParametros." & vbCrLf
   sMsg = sMsg & Err.Number & " - " & Err.Description
   MsgBox sMsg, vbCritical, App.ProductName
   Exit Function
   Resume 0
End Function

Public Function carregaProxSessoes() As Boolean
   Dim sMsg    As String
   Dim strSql  As String
   Dim ini     As Boolean
   
   On Error GoTo TrataErro
   
   carregaProxSessoes = False
   ini = True
   
   strSql = "SELECT tb_sessao.ses_horario as horario,tb_filme.fil_nm as Filme,tb_sessao.sal_cd as Sala,tb_filme.fil_cd as codFilme,tb_sessao.sal_cd as codSala,tb_sessao.ses_horario as Sessao "
   strSql = strSql + "FROM tb_filme INNER JOIN "
   strSql = strSql + "tb_sessao ON tb_filme.fil_cd = tb_sessao.fil_cd INNER JOIN "
   strSql = strSql + "tb_programacao ON tb_sessao.prg_cd = tb_programacao.prg_cd "
   strSql = strSql + "WHERE (tb_filme.fil_dt_ini <= CONVERT(DATETIME, '" + Format(Date, "yyyy-mm-dd") + "', 102)) and (tb_filme.fil_dt_fim >= CONVERT(DATETIME, '" + Format(Date, "yyyy-mm-dd") + "', 102)) "
   strSql = strSql + " and (tb_programacao.prg_dt_ini <= CONVERT(DATETIME, '" + Format(Date, "yyyy-mm-dd") + "', 102)) and (tb_programacao.prg_dt_fim >= CONVERT(DATETIME, '" + Format(Date, "yyyy-mm-dd") + "', 102)) "
   strSql = strSql + "and tb_sessao.ses_dia_semana=" + str(Weekday(Date)) + " order by horario"

   gRSDados1.Open strSql, "File Name=" + App.Path + "\Cinema.Udl", adOpenDynamic, adLockReadOnly, adCmdText
   
   If Not (gRSDados1.EOF And gRSDados1.BOF) Then
      Do While Not gRSDados1.EOF
         If ini Then
            ReDim proxSessoes(1 To 1) As tp_proxSessao
            ini = False
         Else
            ReDim Preserve proxSessoes(1 To UBound(proxSessoes) + 1) As tp_proxSessao
         End If
         
         proxSessoes(UBound(proxSessoes)).horario = gRSDados1.Fields("horario").Value
         proxSessoes(UBound(proxSessoes)).filme = gRSDados1.Fields("Filme").Value
         proxSessoes(UBound(proxSessoes)).sala = gRSDados1.Fields("Sala").Value
         proxSessoes(UBound(proxSessoes)).codFilme = gRSDados1.Fields("codFilme").Value
         proxSessoes(UBound(proxSessoes)).codSala = gRSDados1.Fields("codSala").Value
         proxSessoes(UBound(proxSessoes)).sessao = gRSDados1.Fields("sessao").Value
         
         gRSDados1.MoveNext
      Loop
   Else
      telaSessoes = False
   End If
   
   gRSDados1.Close
   
   carregaProxSessoes = True
   
   Exit Function
   
TrataErro:
   telaSessoes = False
   
   If gRSDados1.State = adStateOpen Then
       gRSDados1.Close
   End If
   
   sMsg = "Ocorreu um erro em carregaProxSessoes." & vbCrLf
   sMsg = sMsg & Err.Number & " - " & Err.Description
   MsgBox sMsg, vbCritical, App.ProductName
End Function

Public Function carregaFilmes() As Boolean
   Dim sMsg    As String
   Dim strSql  As String
   Dim diaSem  As Integer
   Dim bAux    As Boolean
   Dim ini     As Boolean
   
   On Error GoTo TrataErro
   
   carregaFilmes = False
   diaSem = diaSemana(Now)
   ini = True
 
   strSql = "SELECT distinct tb_filme.fil_cd as codFilme,tb_sessao.sal_cd as CodSala,tb_filme.fil_nm as Filme, tb_sala.sal_desc as Sala,tb_filme.fil_censura as Censura  "
   strSql = strSql + "FROM tb_filme INNER JOIN tb_sessao ON tb_filme.fil_cd = tb_sessao.fil_cd INNER JOIN tb_programacao ON tb_sessao.prg_cd = tb_programacao.prg_cd "
   strSql = strSql + " INNER JOIN tb_sala ON tb_sessao.sal_cd = tb_sala.sal_cd "
   'strSql = strSql + "FROM tb_filme Inner Join tb_sessao ON tb_filme.fil_cd = tb_sessao.fil_cd "
   'strSql = strSql + "WHERE (tb_filme.fil_dt_ini <= CONVERT(DATETIME, '" + Format(Date, "yyyy-mm-dd") + "', 102)) and (tb_filme.fil_dt_fim >= CONVERT(DATETIME, '" + Format(Date, "yyyy-mm-dd") + "', 102)) "
   'strSql = strSql + "and tb_sessao.ses_dia_semana=" & diaSemana(Now)
   strSql = strSql & "WHERE (tb_filme.fil_dt_ini <= CONVERT(DATETIME, '" + Format(Date, "yyyy-mm-dd") + "', 102)) and (tb_filme.fil_dt_fim >= CONVERT(DATETIME, '" + Format(Date, "yyyy-mm-dd") + "', 102))  and (tb_programacao.prg_dt_ini <= CONVERT(DATETIME, '" + Format(Date, "yyyy-mm-dd") + "', 102)) and (tb_programacao.prg_dt_fim >= CONVERT(DATETIME, '" + Format(Date, "yyyy-mm-dd") + "', 102)) "
   strSql = strSql & "and tb_sessao.ses_dia_semana= " & diaSemana(Now)
   
   gRSDados1.Open strSql, "File Name=" + App.Path + "\Cinema.Udl", adOpenDynamic, adLockReadOnly, adCmdText
   
   If Not (gRSDados1.EOF And gRSDados1.BOF) Then
      Do While Not gRSDados1.EOF
         If ini Then
            ReDim filmes(1 To 1) As tp_filme
            ini = False
         Else
            ReDim Preserve filmes(1 To UBound(filmes) + 1) As tp_filme
         End If
         
         filmes(UBound(filmes)).codFilme = gRSDados1.Fields("codFilme").Value
         filmes(UBound(filmes)).descFilme = gRSDados1.Fields("Filme").Value
         filmes(UBound(filmes)).codSala = gRSDados1.Fields("codSala").Value
         filmes(UBound(filmes)).descSala = gRSDados1.Fields("Sala").Value
         If Not IsNull(gRSDados1.Fields("censura").Value) Then
            If gRSDados1.Fields("censura").Value = 0 Then
                filmes(UBound(filmes)).censura = "Livre"
            Else
                filmes(UBound(filmes)).censura = str(gRSDados1.Fields("censura").Value) & " anos"
            End If
         Else
            filmes(UBound(filmes)).censura = ""
         End If
         
         bAux = carregaHorariosFilme(diaSem, UBound(filmes))
         
         gRSDados1.MoveNext
      Loop
   Else
      telaFilme = False
   End If
   
   gRSDados1.Close
   
   carregaFilmes = True
   
   Exit Function
   
TrataErro:
   telaFilme = False
   
   If gRSDados1.State = adStateOpen Then
       gRSDados1.Close
   End If
   
   sMsg = "Ocorreu um erro em carregaFilmes." & vbCrLf
   sMsg = sMsg & Err.Number & " - " & Err.Description
   MsgBox sMsg, vbCritical, App.ProductName
End Function

Public Function carregaHorariosFilme(diaSemana As Integer, I As Integer) As Boolean
   Dim sMsg    As String
   Dim strSql  As String
   Dim ini     As Boolean
   
   On Error GoTo TrataErro
   
   carregaHorariosFilme = False
   ini = True
   
   strSql = "SELECT tb_sessao.ses_horario as horario,tb_filme.fil_nm as Filme,tb_sessao.sal_cd as Sala,tb_filme.fil_cd as codFilme,tb_sessao.sal_cd as codSala,ses_cd  as Sessao "
   strSql = strSql & "FROM tb_filme INNER JOIN tb_sessao ON tb_filme.fil_cd = tb_sessao.fil_cd INNER JOIN tb_programacao ON tb_sessao.prg_cd = tb_programacao.prg_cd "
   strSql = strSql & "WHERE (tb_filme.fil_dt_ini <= CONVERT(DATETIME, '" + Format(Date, "yyyy-mm-dd") + "', 102)) and (tb_filme.fil_dt_fim >= CONVERT(DATETIME, '" + Format(Date, "yyyy-mm-dd") + "', 102))  and (tb_programacao.prg_dt_ini <= CONVERT(DATETIME, '" + Format(Date, "yyyy-mm-dd") + "', 102)) and (tb_programacao.prg_dt_fim >= CONVERT(DATETIME, '" + Format(Date, "yyyy-mm-dd") + "', 102)) "
   strSql = strSql & "and tb_sessao.ses_dia_semana= " & diaSemana & " and tb_sessao.sal_cd =" & filmes(I).codSala & " and tb_filme.fil_cd = " & filmes(I).codFilme

   gRSDados2.Open strSql, "File Name=" + App.Path + "\Cinema.Udl", adOpenDynamic, adLockReadOnly, adCmdText
   
   If Not (gRSDados2.EOF And gRSDados2.BOF) Then
      Do While Not gRSDados2.EOF
         If ini Then
            ReDim filmes(I).horarios(1 To 1) As tp_horarios
            ini = False
         Else
            ReDim Preserve filmes(I).horarios(1 To UBound(filmes(I).horarios) + 1) As tp_horarios
         End If
         
         filmes(I).horarios(UBound(filmes(I).horarios)).horario = gRSDados2.Fields("horario").Value
         filmes(I).horarios(UBound(filmes(I).horarios)).sessao = gRSDados2.Fields("sessao").Value
         
         gRSDados2.MoveNext
      Loop
   End If
   
   gRSDados2.Close
   
   carregaHorariosFilme = True
   
   Exit Function
   
TrataErro:
   If gRSDados2.State = adStateOpen Then
       gRSDados2.Close
   End If
   
   sMsg = "Ocorreu um erro em carregaHorariosFilme." & vbCrLf
   sMsg = sMsg & Err.Number & " - " & Err.Description
   MsgBox sMsg, vbCritical, App.ProductName
End Function

Public Function carregaPrecos() As Boolean
   Dim sMsg         As String
   Dim strSql       As String
   Dim ini          As Boolean
   Dim iniPreco     As Boolean
   Dim codFilme     As Long
   Dim vlrIntManha  As Double
   Dim vlrIntTarde  As Double
   Dim vlrIntNoite  As Double
   Dim vlrMeiaManha As Double
   Dim vlrMeiaTarde As Double
   Dim vlrMeiaNoite As Double
   Dim descricao    As String
   Dim promocional  As Boolean
   
   On Error GoTo TrataErro
   
   carregaPrecos = False
   ini = True
   iniPreco = True
   
   strSql = "SELECT  Convert(Time(0),[par_hora_limite12]) as Limite12 ,Convert(Time(0),[par_hora_limite23])as Limite23 ,Convert(Time(0),[par_hora_limite34])as Limite34,Convert(Time(0),[par_hora_limite45])as Limite45,Convert(Time(0),[par_hora_limite56])as Limite56  From tb_parametro"
   gRSDados1.Open strSql, "File Name=" + App.Path + "\Cinema.Udl", adOpenDynamic, adLockReadOnly, adCmdText
   
   Dim wCampos As String
   
   
   'tb_preco.pre_vl_inteira_ate as VlrIntAteHrLim, tb_preco.pre_vl_inteira_apos as VlrIntAposHrLim, tb_preco.pre_vl_Meia_ate as VlrMeiaAteHrLim, tb_preco.pre_vl_Meia_apos as VlrMeiaAposHrLim,
   If FormatDateTime(gRSDados1!Limite12, vbShortTime) > Time Then
        wCampos = "tb_preco.pre_vl_inteira_ate as VlrIntAteHrLim, tb_preco.pre_vl_inteira_apos as VlrIntAposHrLim, tb_preco.pre_vl_Meia_ate as VlrMeiaAteHrLim, tb_preco.pre_vl_Meia_apos as VlrMeiaAposHrLim,"
        wHorario1 = gRSDados1!Limite12
        wHorario2 = gRSDados1!Limite23
    ElseIf FormatDateTime(gRSDados1!Limite23, vbShortTime) > Time Then
        wCampos = "tb_preco.pre_vl_inteira_apos as VlrIntAteHrLim, tb_preco.pre_vl_inteira3 as VlrIntAposHrLim, tb_preco.pre_vl_Meia_apos as VlrMeiaAteHrLim, tb_preco.pre_vl_Meia3 as VlrMeiaAposHrLim,"
        wHorario1 = gRSDados1!Limite23
        wHorario2 = gRSDados1!Limite34
    ElseIf FormatDateTime(gRSDados1!Limite34, vbShortTime) > Time Then
        wCampos = "tb_preco.pre_vl_inteira3 as VlrIntAteHrLim, tb_preco.pre_vl_inteira4 as VlrIntAposHrLim, tb_preco.pre_vl_Meia3 as VlrMeiaAteHrLim, tb_preco.pre_vl_Meia4 as VlrMeiaAposHrLim,"
        wHorario1 = gRSDados1!Limite34
        wHorario2 = gRSDados1!Limite45
    ElseIf FormatDateTime(gRSDados1!Limite45, vbShortTime) > Time Then
        wCampos = "tb_preco.pre_vl_inteira4 as VlrIntAteHrLim, tb_preco.pre_vl_inteira5 as VlrIntAposHrLim, tb_preco.pre_vl_Meia4 as VlrMeiaAteHrLim, tb_preco.pre_vl_Meia5 as VlrMeiaAposHrLim,"
        wHorario1 = gRSDados1!Limite45
        wHorario2 = gRSDados1!Limite56
    Else
        wCampos = "tb_preco.pre_vl_inteira5 as VlrIntAteHrLim, tb_preco.pre_vl_inteira6 as VlrIntAposHrLim, tb_preco.pre_vl_Meia5 as VlrMeiaAteHrLim, tb_preco.pre_vl_Meia6 as VlrMeiaAposHrLim,"
        wHorario1 = gRSDados1!Limite56
        wHorario2 = gRSDados1!Limite56
   End If
   
   gRSDados1.Close
   
    strSql = "SELECT  Distinct tb_filme.fil_cd as codFilme, tb_filme.fil_nm as Filme,tb_preco.pre_dia_semana as DiaSemana," + wCampos + " pre_promocao as Promocional "
    strSql = strSql & "FROM tb_filme INNER JOIN "
    strSql = strSql & "tb_preco ON tb_filme.fil_cd = tb_preco.fil_cd "
    strSql = strSql & "INNER JOIN  tb_sessao ON tb_filme.fil_cd = tb_sessao.fil_cd "
    strSql = strSql & " Inner Join tb_prog_preco on tb_preco.ppr_cd = tb_prog_preco.ppr_cd "
    strSql = strSql & "WHERE (tb_filme.fil_dt_ini <= CONVERT(DATETIME, '" + Format(Date, "yyyy-mm-dd") + "', 102)) and (tb_filme.fil_dt_fim >= CONVERT(DATETIME, '" + Format(Date, "yyyy-mm-dd") + "', 102)) "
    strSql = strSql & " AND tb_prog_preco.ppr_flg_promocao = 0 "
    strSql = strSql & "and ((tb_preco.pre_dia_semana =" & str(Weekday(Date)) + " and tb_preco.pre_promocao = 0 ) or (tb_preco.pre_dia_semana <>" & str(Weekday(Date)) + " and tb_preco.pre_promocao = 1)) "
    strSql = strSql & "Union "
    strSql = strSql & "SELECT tb_preco.fil_cd as codFilme, 'Vlr.Padrão' as Filme,tb_preco.pre_dia_semana as DiaSemana," + wCampos + " pre_promocao as Promocional "
    strSql = strSql & "FROM   tb_preco Inner Join tb_prog_preco "
    strSql = strSql & "on tb_preco.ppr_cd = tb_prog_preco.ppr_cd "
    strSql = strSql & "Where tb_preco.fil_cd = 0 and tb_prog_preco.ppr_flg_promocao = 0 and tb_preco.pre_dia_semana =" & str(Weekday(Date)) + " or (tb_preco.fil_cd = 0 and tb_preco.pre_dia_semana <>" & str(Weekday(Date)) + " and tb_preco.pre_Promocao = 1)"
   
    gRSDados1.Open strSql, "File Name=" + App.Path + "\Cinema.Udl", adOpenDynamic, adLockReadOnly, adCmdText
   
   
   If Not (gRSDados1.EOF And gRSDados1.BOF) Then
      vlrIntManha = gRSDados1.Fields("vlrIntAteHrLim").Value
      vlrIntTarde = gRSDados1.Fields("vlrIntAposHrLim").Value
      vlrMeiaManha = gRSDados1.Fields("vlrMeiaAteHrLim").Value
      vlrMeiaTarde = gRSDados1.Fields("vlrMeiaAposHrLim").Value

      Do While Not gRSDados1.EOF
         If ini Then
            ReDim precos(1 To 1) As tp_precos
            ini = False
            iniPreco = True

            codFilme = gRSDados1.Fields("codFilme").Value
            vlrIntManha = gRSDados1.Fields("vlrIntAteHrLim").Value
            vlrIntTarde = gRSDados1.Fields("vlrIntAposHrLim").Value
            vlrMeiaManha = gRSDados1.Fields("vlrMeiaAteHrLim").Value
            vlrMeiaTarde = gRSDados1.Fields("vlrMeiaAposHrLim").Value
            descricao = descDiaSemana(gRSDados1.Fields("diaSemana").Value)
            promocional = gRSDados1.Fields("promocional").Value
            
            precos(UBound(precos)).descFilme = gRSDados1.Fields("Filme").Value
            precos(UBound(precos)).codFilme = gRSDados1.Fields("codFilme").Value
         Else
            If codFilme <> gRSDados1.Fields("codFilme").Value Then
               If iniPreco Then
                  ReDim precos(UBound(precos)).precos(1 To 1) As tp_preco
               Else
                  ReDim Preserve precos(UBound(precos)).precos(1 To UBound(precos(UBound(precos)).precos) + 1) As tp_preco
               End If
               
               precos(UBound(precos)).precos(UBound(precos(UBound(precos)).precos)).descricao = descricao
               precos(UBound(precos)).precos(UBound(precos(UBound(precos)).precos)).vlrIntManha = vlrIntManha
               precos(UBound(precos)).precos(UBound(precos(UBound(precos)).precos)).vlrIntTarde = vlrIntTarde
               precos(UBound(precos)).precos(UBound(precos(UBound(precos)).precos)).vlrMeiaManha = vlrMeiaManha
               precos(UBound(precos)).precos(UBound(precos(UBound(precos)).precos)).vlrMeiaTarde = vlrMeiaTarde
               precos(UBound(precos)).precos(UBound(precos(UBound(precos)).precos)).promocional = promocional
               
               ReDim Preserve precos(1 To UBound(precos) + 1) As tp_precos
            
               codFilme = gRSDados1.Fields("codFilme").Value
               vlrIntManha = gRSDados1.Fields("vlrIntAteHrLim").Value
               vlrIntTarde = gRSDados1.Fields("vlrIntAposHrLim").Value
               vlrMeiaManha = gRSDados1.Fields("vlrMeiaAteHrLim").Value
               vlrMeiaTarde = gRSDados1.Fields("vlrMeiaAposHrLim").Value
               descricao = descDiaSemana(gRSDados1.Fields("diaSemana").Value)
               promocional = gRSDados1.Fields("promocional").Value
               
               precos(UBound(precos)).descFilme = gRSDados1.Fields("Filme").Value
               precos(UBound(precos)).codFilme = gRSDados1.Fields("codFilme").Value
               'descricao = ""
               iniPreco = True
            ElseIf vlrIntManha <> gRSDados1.Fields("vlrIntAteHrLim").Value Or _
                   vlrIntTarde <> gRSDados1.Fields("vlrIntAposHrLim").Value Or _
                   vlrMeiaManha <> gRSDados1.Fields("vlrMeiaAteHrLim").Value Or _
                   vlrMeiaTarde <> gRSDados1.Fields("vlrMeiaAposHrLim").Value Or _
                   promocional <> gRSDados1.Fields("promocional").Value Then
               If iniPreco Then
                  ReDim precos(UBound(precos)).precos(1 To 1) As tp_preco
                  iniPreco = False
               Else
                  ReDim Preserve precos(UBound(precos)).precos(1 To UBound(precos(UBound(precos)).precos) + 1) As tp_preco
               End If
               
               precos(UBound(precos)).precos(UBound(precos(UBound(precos)).precos)).descricao = descricao
               precos(UBound(precos)).precos(UBound(precos(UBound(precos)).precos)).vlrIntManha = vlrIntManha
               precos(UBound(precos)).precos(UBound(precos(UBound(precos)).precos)).vlrIntTarde = vlrIntTarde
               precos(UBound(precos)).precos(UBound(precos(UBound(precos)).precos)).vlrMeiaManha = vlrMeiaManha
               precos(UBound(precos)).precos(UBound(precos(UBound(precos)).precos)).vlrMeiaTarde = vlrMeiaTarde
               precos(UBound(precos)).precos(UBound(precos(UBound(precos)).precos)).promocional = promocional
            
               vlrIntManha = gRSDados1.Fields("vlrIntAteHrLim").Value
               vlrIntTarde = gRSDados1.Fields("vlrIntAposHrLim").Value
               vlrMeiaManha = gRSDados1.Fields("vlrMeiaAteHrLim").Value
               vlrMeiaTarde = gRSDados1.Fields("vlrMeiaAposHrLim").Value
               descricao = descDiaSemana(gRSDados1.Fields("diaSemana").Value)
               promocional = gRSDados1.Fields("promocional").Value
            Else
               If descricao = "" Then
                  descricao = descDiaSemana(gRSDados1.Fields("diaSemana").Value)
               Else
                  descricao = descricao & ", " & descDiaSemana(gRSDados1.Fields("diaSemana").Value)
               End If
            End If
         End If
         
         gRSDados1.MoveNext
      Loop
   
      If iniPreco Then
         ReDim precos(UBound(precos)).precos(1 To 1) As tp_preco
         iniPreco = False
      Else
         ReDim Preserve precos(UBound(precos)).precos(1 To UBound(precos(UBound(precos)).precos) + 1) As tp_preco
      End If
      
      precos(UBound(precos)).precos(UBound(precos(UBound(precos)).precos)).descricao = descricao
      precos(UBound(precos)).precos(UBound(precos(UBound(precos)).precos)).vlrIntManha = vlrIntManha
      precos(UBound(precos)).precos(UBound(precos(UBound(precos)).precos)).vlrIntTarde = vlrIntTarde
      precos(UBound(precos)).precos(UBound(precos(UBound(precos)).precos)).vlrMeiaManha = vlrMeiaManha
      precos(UBound(precos)).precos(UBound(precos(UBound(precos)).precos)).vlrMeiaTarde = vlrMeiaTarde
      precos(UBound(precos)).precos(UBound(precos(UBound(precos)).precos)).promocional = promocional
   Else
      telaPrecos = False
   End If
   
   gRSDados1.Close
   
   carregaPrecos = True
   
   Exit Function
   
TrataErro:
   telaPrecos = False
   
   If gRSDados1.State = adStateOpen Then
       gRSDados1.Close
   End If
   
   sMsg = "Ocorreu um erro em carregaPrecos." & vbCrLf
   sMsg = sMsg & Err.Number & " - " & Err.Description
   MsgBox sMsg, vbCritical, App.ProductName
   Exit Function
   Resume 0
End Function

Public Function caregaTrailers() As Boolean
   Dim sMsg    As String
   Dim strSql  As String
   Dim ini     As Boolean
   
   On Error GoTo TrataErro
   
   caregaTrailers = False
   ini = True
   
   strSql = "SELECT arquivo "
   strSql = strSql & "FROM tb_trailer "
   strSql = strSql & "ORDER BY arquivo"
   
   gRSDados1.Open strSql, "File Name=" + App.Path + "\Cinema.Udl", adOpenDynamic, adLockReadOnly, adCmdText
   
   If Not (gRSDados1.EOF And gRSDados1.BOF) Then
      Do While Not gRSDados1.EOF
         If ini Then
            ReDim trailers(1 To 1) As String
            ini = False
         Else
            ReDim Preserve trailers(1 To UBound(trailers) + 1) As String
         End If
         
         trailers(UBound(trailers)) = gRSDados1.Fields("arquivo").Value
         
         gRSDados1.MoveNext
      Loop
   Else
      telaTrailer = False
   End If
   
   gRSDados1.Close
   
   caregaTrailers = True
   
   Exit Function
   
TrataErro:
   telaTrailer = False
   
   If gRSDados1.State = adStateOpen Then
       gRSDados1.Close
   End If
   
   sMsg = "Ocorreu um erro em caregaTrailers." & vbCrLf
   sMsg = sMsg & Err.Number & " - " & Err.Description
   MsgBox sMsg, vbCritical, App.ProductName
End Function

Public Function diaSemana(dtRef As Date) As Integer
   diaSemana = 0
   If verificaFeriado(dtRef) Then
      diaSemana = 8
   Else
      Select Case Weekday(dtRef)
         Case vbSunday
            diaSemana = 7
         Case vbMonday
            diaSemana = 1
         Case vbTuesday
            diaSemana = 2
         Case vbWednesday
            diaSemana = 3
         Case vbThursday
            diaSemana = 4
         Case vbFriday
            diaSemana = 5
         Case vbSaturday
            diaSemana = 6
      End Select
   End If
End Function

Public Function descDiaSemana(diaSemana As Integer) As String
   descDiaSemana = ""
   
   Select Case diaSemana
      Case 1
         descDiaSemana = "Dom"
      Case 2
         descDiaSemana = "Seg"
      Case 3
         descDiaSemana = "Ter"
      Case 4
         descDiaSemana = "Qua"
      Case 5
         descDiaSemana = "Qui"
      Case 6
         descDiaSemana = "Sex"
      Case 7
         descDiaSemana = "Sab"
      Case 8
         descDiaSemana = "Fer"
   End Select
End Function

Public Function verificaFeriado(dtRef As Date) As Boolean
   Dim sMsg    As String
   Dim strSql  As String
   
   On Error GoTo TrataErro
   
   verificaFeriado = False
   
   strSql = "SELECT Fer_Data FROM tb_feriado "
   strSql = strSql & "WHERE (fer_data = CONVERT(DATETIME, '" + Format(dtRef, "YYYY-MM-DD 00:00:00") + "', 102))"
   'strSql = strSql & "WHERE Fer_Data = #" & Format(dtRef, "MM/DD/YYYY") & "#"
   
   
   gRSDados1.Open strSql, "File Name=" + App.Path + "\Cinema.Udl", adOpenDynamic, adLockReadOnly, adCmdText
   
   
   
   If gRSDados1.EOF And gRSDados1.BOF Then
      gRSDados1.Close
      Exit Function
   End If
   
   gRSDados1.Close

   verificaFeriado = True
   
   Exit Function
   
TrataErro:
   If gRSDados1.State = adStateOpen Then
       gRSDados1.Close
   End If
   
   sMsg = "Ocorreu um erro em verificaFeriado." & vbCrLf
   sMsg = sMsg & Err.Number & " - " & Err.Description
   MsgBox sMsg, vbCritical, App.ProductName
End Function

Public Function inicializaTabLotacao() As Boolean
   Dim sMsg     As String
   Dim strSql   As String
   Dim atualiza As Boolean
   Dim regs     As Long
   
   
   
   
   On Error GoTo TrataErro
   
   inicializaTabLotacao = False
   atualiza = True
   
   strSql = "SELECT DataDados, Atualizados "
   strSql = strSql & "FROM tb_Temp_Atu_Lot"
   
   gRSDados1.Open strSql, gConnect, adOpenDynamic, adLockReadOnly, adCmdText
   
   If Not (gRSDados1.EOF And gRSDados1.BOF) Then
      If Date = gRSDados1.Fields("DataDados").Value Then
         atualiza = False
      End If
   End If
   
   gRSDados1.Close
   
   If atualiza Then
      strSql = "DELETE FROM tb_Temp_Lotacao "
      gConnect.Execute strSql, regs, adCmdText
      
      strSql = "INSERT INTO tb_Temp_Lotacao "
      strSql = strSql & "(codFilme, codSala, sessao, horario, lotada)"
      strSql = strSql & "SELECT tb_sessoes.codFilme, "
      strSql = strSql & "tb_sessoes.codSala, "
      strSql = strSql & "tb_sessoes.sessao, "
      strSql = strSql & "tb_sessoes.horario, "
      strSql = strSql & "false "
      strSql = strSql & "FROM tb_sessoes "
      strSql = strSql & "WHERE tb_sessoes.diaSemana = " & diaSemana(Now) & " "
      gConnect.Execute strSql, regs, adCmdText
      
      Call atualizaLotAtu
   End If

   inicializaTabLotacao = True

   Exit Function
   
TrataErro:
   If gRSDados1.State = adStateOpen Then
       gRSDados1.Close
   End If
   
   sMsg = "Ocorreu um erro em inicializaTabLotacao." & vbCrLf
   sMsg = sMsg & Err.Number & " - " & Err.Description
   MsgBox sMsg, vbCritical, App.ProductName

End Function

Public Function verificaLOTADO(codFilme As Long, codSala As Long, sessao As Variant) As Boolean
   'Dim strSql As String
   Dim sMsg   As String
   
   On Error GoTo TrataErro
   
  Dim Lugares As Integer, Meias As Integer, Cortesias As Integer, Lotacao As Integer
        
  Lotacao = consLotacaoSel(CInt(codSala), codFilme, Date, CVDate(sessao), Lugares, Meias, Cortesias)

  verificaLOTADO = IIf(Lotacao > 0, False, True)
  
  If verificaLOTADO = False Then 'Se Não Esta Lotado, verificar lotação manual
        Dim strSql  As String
        
        strSql = "SELECT CodFilme FROM tb_Temp_Lotacao Where CodFilme = '" + Trim(str(codFilme)) + "' and CodSala = '" + Trim(str(codSala)) + "' and Horario = '01/01/1901 '+CONVERT(DATETIME, '" + Format(sessao, "HH:mm") + "', 102) and Data = CONVERT(DATETIME, '" + Format(Date, "yyyy-mm-dd") + "', 102)"
        
        gRSDados1.Open strSql, "File Name=" + App.Path + "\Cinema.Udl", adOpenDynamic, adLockReadOnly, adCmdText
        
        If Not (gRSDados1.EOF And gRSDados1.BOF) Then
            verificaLOTADO = True
        End If
        gRSDados1.Close
  End If
  

   Exit Function
   
TrataErro:
   If gRSDados1.State = adStateOpen Then
       gRSDados1.Close
   End If
   
   sMsg = "Ocorreu um erro em verificaLOTADO." & vbCrLf
   sMsg = sMsg & Err.Number & " - " & Err.Description
   MsgBox sMsg, vbCritical, App.ProductName
End Function

Public Function atualizaLotAtu() As Boolean
   Dim sMsg     As String
   Dim strSql   As String
   Dim atualiza As Boolean
   Dim regs     As Long
   
   On Error GoTo TrataErro
   
   atualizaLotAtu = False
   atualiza = False
   
   strSql = "SELECT DataDados, Atualizados "
   strSql = strSql & "FROM tb_Temp_Atu_Lot"
   
   gRSDados1.Open strSql, gConnect, adOpenDynamic, adLockReadOnly, adCmdText
   
   If Not (gRSDados1.EOF And gRSDados1.BOF) Then
      atualiza = True
   End If
   
   gRSDados1.Close
   
   If atualiza Then
      strSql = "UPDATE tb_Temp_Atu_Lot "
      strSql = strSql & "SET Atualizados = True, "
      strSql = strSql & "DataDados = Date()"
      
      gConnect.Execute strSql, regs, adCmdText
   Else
      strSql = "INSERT INTO tb_Temp_Atu_Lot "
      strSql = strSql & "(DataDados, Atualizados) "
      strSql = strSql & "VALUES(Date(), True)"
      
      gConnect.Execute strSql, regs, adCmdText
   End If

   atualizaLotAtu = True
   
   Exit Function
   
TrataErro:
   If gRSDados1.State = adStateOpen Then
       gRSDados1.Close
   End If
   
   sMsg = "Ocorreu um erro em atualizaLotAtu." & vbCrLf
   sMsg = sMsg & Err.Number & " - " & Err.Description
   MsgBox sMsg, vbCritical, App.ProductName
End Function

Public Function caregaImagens() As Boolean
   Dim sMsg    As String
   Dim strSql  As String
   Dim ini     As Boolean
   
   On Error GoTo TrataErro
   
   caregaImagens = False
   ini = True
   
   strSql = "SELECT arquivo "
   strSql = strSql & "FROM tb_imagens "
   strSql = strSql & "ORDER BY arquivo"
   
   gRSDados1.Open strSql, "File Name=" + App.Path + "\Cinema.Udl", adOpenDynamic, adLockReadOnly, adCmdText
   
   If Not (gRSDados1.EOF And gRSDados1.BOF) Then
      Do While Not gRSDados1.EOF
         If ini Then
            ReDim imagens(1 To 1) As String
            ini = False
         Else
            ReDim Preserve imagens(1 To UBound(imagens) + 1) As String
         End If
         
         imagens(UBound(imagens)) = gRSDados1.Fields("arquivo").Value
         
         gRSDados1.MoveNext
      Loop
   Else
      telaImagem = False
   End If
   
   gRSDados1.Close
   
   caregaImagens = True
   
   Exit Function
   
TrataErro:
   telaTrailer = False
   
   If gRSDados1.State = adStateOpen Then
       gRSDados1.Close
   End If
   
   sMsg = "Ocorreu um erro em caregaImagens." & vbCrLf
   sMsg = sMsg & Err.Number & " - " & Err.Description
   MsgBox sMsg, vbCritical, App.ProductName
End Function



