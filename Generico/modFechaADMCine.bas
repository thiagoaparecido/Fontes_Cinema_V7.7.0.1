Attribute VB_Name = "modFecham"
Option Explicit

Public Function ImprimeFechamentoCaixa(ByVal origem As String, ByRef MSC As OPOSPOSPrinter, ByRef ErroImp As Boolean, ByVal iCaixa As Integer, tpImp As Integer, Optional ByVal dtAbertura As Date) As Boolean
    Dim oRs    As New ADODB.Recordset
    Dim oRsAux As ADODB.Recordset
    
    Dim dTotal        As Double
    Dim dTotal2       As Double
    Dim clsFechamento As New clsFechamento
    Dim clsCaixa      As New clsTB_CAIXA
    Dim caixaDesc     As String

    Dim fonteNormal     As StdFont
    Dim fonteTitulo     As StdFont
    Dim fonteSubTitulo  As StdFont
    Dim hf              As Single
    Dim Y               As Single
    Dim tamPag          As Single
    Dim strFechamento   As String
    Dim lInicio         As Long
    
Imprime_FechamentoCaixa:
    
    On Error GoTo erro_ImprimeFechamentoCaixa
    
    Set clsFechamento.ConexaoADO = dbConnect
    
    Set clsCaixa.ConexaoADO = dbConnect
    clsCaixa.cxa_cd = iCaixa
    Set oRsAux = New ADODB.Recordset
    
    If clsCaixa.Selecionar(oRsAux) Then
        caixaDesc = oRsAux.Fields("cxa_desc")
    Else
        caixaDesc = ""
    End If
    oRsAux.Close
    
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
        
        Set fonteSubTitulo = New StdFont
        fonteSubTitulo.Name = "Courier New"
        fonteSubTitulo.Size = 12
        fonteSubTitulo.Bold = True
        fonteSubTitulo.Italic = False
        
        Y = 10
    
        'Printer.PaperSize = vbPRPSA4
        Printer.ScaleMode = vbMillimeters
        Printer.Orientation = vbPRORPortrait
        tamPag = Printer.ScaleHeight
        
        clsFechamento.cxa_cd = iCaixa
        clsFechamento.cxp_dt_abertura = Format(dtAbertura, "dd/MM/yyyy Hh:Nn:Ss")
        
        If Not clsFechamento.Caixa(oRs) Then
            MsgBox clsFechamento.MensagemErro, vbCritical, App.ProductName
            Exit Function
        End If
        
        If Not oRs.EOF Then
            Y = 10
                
            Set Printer.Font = fonteNormal
            hf = Printer.TextHeight("X")
            
            'risco de separação
            Printer.CurrentX = 10
            Printer.CurrentY = Y
            Printer.Print String(42, "_")
            Y = Y + hf
            
            Set Printer.Font = fonteTitulo
            hf = Printer.TextHeight("X")
            Printer.CurrentX = 10
            Printer.CurrentY = Y
            If origem = "C" Then
                Printer.Print "FECHAMENTO DE CAIXA"
            Else
                Printer.Print "FECHAM. ADM. DE CAIXA"
            End If
            Y = Y + hf
            Set Printer.Font = fonteNormal
            hf = Printer.TextHeight("X")
            Y = Y + hf
            
            'Caixa
            Set Printer.Font = fonteSubTitulo
            hf = Printer.TextHeight("X")
            Printer.CurrentX = 10
            Printer.CurrentY = Y
            Printer.Print caixaDesc
            Y = Y + hf
            Set Printer.Font = fonteNormal
            hf = Printer.TextHeight("X")
            Y = Y + hf
            
            'Data Abertura
            Printer.CurrentX = 10
            Printer.CurrentY = Y
            Printer.Print Alinha("DATA ABERTURA...: " & Format$(dtAbertura, "dd/mm/yyyy HH:MM:SS"), COLUNAS_IMP, esquerda)
            Y = Y + hf
            
            'Data Fechamento
            Printer.CurrentX = 10
            Printer.CurrentY = Y
            Printer.Print Alinha("DATA FECHAMENTO.: " & Format$(oRs.Fields("cxp_dt_fechamento"), "dd/mm/yyyy HH:MM:SS"), COLUNAS_IMP, esquerda)
            Y = Y + hf
            
            'Operador do Caixa
            Printer.CurrentX = 10
            Printer.CurrentY = Y
            Printer.Print Alinha("OPERADOR CAIXA..: " & oRs.Fields("usu_nm_abertura"), COLUNAS_IMP, esquerda)
            Y = Y + hf
            
            'Fechou o Caixa
            If oRs.Fields("usu_nm_abertura") <> oRs.Fields("usu_nm_fechamento") Then
                Printer.CurrentX = 10
                Printer.CurrentY = Y
                Printer.Print Alinha("FECHOU CAIXA....: " & oRs.Fields("usu_nm_fechamento"), COLUNAS_IMP, esquerda)
                Y = Y + hf
            End If
            
            ' VALORES
            
            Set Printer.Font = fonteSubTitulo
            hf = Printer.TextHeight("X")
            Y = Y + hf
            Printer.CurrentX = 10
            Printer.CurrentY = Y
            Printer.Print Alinha("VALORES", 19, esquerda)
            Y = Y + hf
            Set Printer.Font = fonteNormal
            hf = Printer.TextHeight("X")
            Y = Y + hf
            
            Set oRsAux = New ADODB.Recordset
            dTotal = 0
                        
            If Not clsFechamento.CaixaValores(oRsAux) Then
                MsgBox clsFechamento.MensagemErro, vbCritical, App.ProductName
                Exit Function
            End If
        
            Do While Not oRsAux.EOF
                If oRsAux.Fields("pag_valor") > 0 Then
                    dTotal = dTotal + oRsAux.Fields("sinal") * oRsAux.Fields("pag_valor")
                End If
                Printer.CurrentX = 10
                Printer.CurrentY = Y
                Printer.Print Alinha(Alinha(EliminaAcentos(oRsAux.Fields("pgt_desc")), 20, esquerda, ".") & ": " & Alinha(Format(oRsAux.Fields("pag_valor"), "$ #,##0.00"), 15, Direita), COLUNAS_IMP, esquerda)
                Y = Y + hf
                
                oRsAux.MoveNext
            Loop
            
            If dTotal > 0 Then
                Printer.CurrentX = 10
                Printer.CurrentY = Y
                Printer.Print String(COLUNAS_IMP, "_")
                Y = Y + hf
                
                Printer.CurrentX = 10
                Printer.CurrentY = Y
                Printer.Print Alinha(Alinha("TOTAL VALORES", 20, esquerda, ".") & ": " & Alinha(Format(dTotal, "$ #,##0.00"), 15, Direita), COLUNAS_IMP, esquerda)
                Y = Y + hf
            End If
            
            oRsAux.Close
            
            ' BILHETES
            
            Set Printer.Font = fonteSubTitulo
            hf = Printer.TextHeight("X")
            Y = Y + hf
            Printer.CurrentX = 10
            Printer.CurrentY = Y
            Printer.Print Alinha("BILHETES", 19, esquerda)
            Y = Y + hf
            Set Printer.Font = fonteNormal
            hf = Printer.TextHeight("X")
            Y = Y + hf
            
            Set oRsAux = New ADODB.Recordset
            dTotal = 0
                        
            If Not clsFechamento.CaixaBilhetes(oRsAux) Then
                MsgBox clsFechamento.MensagemErro, vbCritical, App.ProductName
                Exit Function
            End If
        
            Do While Not oRsAux.EOF
                dTotal = dTotal + oRsAux.Fields("qtde")
                Printer.CurrentX = 10
                Printer.CurrentY = Y
                Printer.Print Alinha(Alinha(EliminaAcentos(oRsAux.Fields("igt_desc")), 20, esquerda, ".") & ": " & Alinha(Format(oRsAux.Fields("qtde"), "##,##0"), 11, Direita), COLUNAS_IMP, esquerda)
                Y = Y + hf
                
                oRsAux.MoveNext
            Loop
            
            If dTotal > 0 Then
                Printer.CurrentX = 10
                Printer.CurrentY = Y
                Printer.Print String(COLUNAS_IMP, "_")
                Y = Y + hf
                
                Printer.CurrentX = 10
                Printer.CurrentY = Y
                Printer.Print Alinha(Alinha("TOTAL BILHETES", 20, esquerda, ".") & ": " & Alinha(Format(dTotal, "##,##0"), 11, Direita), COLUNAS_IMP, esquerda)
                Y = Y + hf
            End If
            
            oRsAux.Close
            
            ' COMBOS
            
            Set oRsAux = New ADODB.Recordset
            dTotal = 0
            dTotal2 = 0
                        
            If Not clsFechamento.CaixaCombos(oRsAux) Then
                MsgBox clsFechamento.MensagemErro, vbCritical, App.ProductName
                Exit Function
            End If
        
            If Not oRsAux.EOF Then
               Set Printer.Font = fonteSubTitulo
               hf = Printer.TextHeight("X")
               Printer.CurrentX = 10
               Printer.CurrentY = Y
               Printer.Print Alinha("COMBOS", 19, esquerda)
               Y = Y + hf
               Set Printer.Font = fonteNormal
               hf = Printer.TextHeight("X")
               Y = Y + hf
            End If
        
            Do While Not oRsAux.EOF
                dTotal = dTotal + oRsAux.Fields("qtde")
                dTotal2 = dTotal2 + oRsAux.Fields("valor")
               Printer.CurrentX = 10
               Printer.CurrentY = Y
               Printer.Print Alinha(Alinha(EliminaAcentos(oRsAux.Fields("cbo_nm")), 20, esquerda, ".") & ": " & _
                                    Alinha(Format(oRsAux.Fields("qtde"), "#,##0"), 5, Direita) & _
                                    Alinha(Format(oRsAux.Fields("valor"), "#,##0.00"), 10, Direita), COLUNAS_IMP, esquerda)
               Y = Y + hf
               oRsAux.MoveNext
            Loop
            
            If dTotal > 0 Then
                Printer.CurrentX = 10
                Printer.CurrentY = Y
                Printer.Print String(COLUNAS_IMP, "_")
                Y = Y + hf
                
                Printer.CurrentX = 10
                Printer.CurrentY = Y
                Printer.Print Alinha(Alinha("TOTAL COMBOS", 20, esquerda, ".") & ": " & _
                                    Alinha(Format(dTotal, "#,##0"), 5, Direita) & _
                                    Alinha(Format(dTotal2, "#,##0.00"), 10, Direita), COLUNAS_IMP, esquerda)
                Y = Y + hf
            End If
            
            oRsAux.Close
        End If
        
        Printer.EndDoc
    
    Else
        If verificaImp(MSC) Then
            clsFechamento.cxa_cd = iCaixa
            clsFechamento.cxp_dt_abertura = Format(dtAbertura, "dd/MM/yyyy Hh:Nn:Ss")
            
            If Not clsFechamento.Caixa(oRs) Then
                MsgBox clsFechamento.MensagemErro, vbCritical, App.ProductName
                Exit Function
            End If
            
            'Inicializa impressora
            If Not iniImpressora(MSC) Then
                GoTo Imprime_FechamentoCaixa
            End If
            
            strFechamento = ""
            
            If Not oRs.EOF Then
                    
                'risco de separação com dupla altura
                'If Not imprimeTraco(MSC) Then
                '    GoTo Imprime_FechamentoCaixa
                'End If
                strFechamento = strFechamento & strTraco

                'If Not AlinhaCentralizado(MSC) Then
                '    GoTo Imprime_FechamentoCaixa
                'End If
                strFechamento = strFechamento & strAlinhaCentralizado
                
                If origem = "C" Then
                    'If Not imprimeTitulo(MSC, "FECHAMENTO DE CAIXA" & Chr$(10) & Chr$(10)) Then
                    '    GoTo Imprime_FechamentoCaixa
                    'End If
                    strFechamento = strFechamento & strTitulo("FECHAMENTO DE CAIXA" & Chr$(10) & Chr$(10))
                Else
                    'If Not imprimeTitulo(MSC, "FECHAM. ADM. DE CAIXA" & Chr$(10) & Chr$(10)) Then
                    '    GoTo Imprime_FechamentoCaixa
                    'End If
                    strFechamento = strFechamento & strTitulo("FECHAM. ADM. DE CAIXA" & Chr$(10) & Chr$(10))
                End If
                
                'Data/Hora do relatório
                'MSC.Output = LinhaDupla & Format$(Now, "dd/mm/yyyy HH:MM:SS") & Chr$(10) & Chr$(10)
                
                'Caixa
                'If Not imprimeDupla(MSC, caixaDesc & Chr$(10) & Chr$(10)) Then
                '    GoTo Imprime_FechamentoCaixa
                'End If
                strFechamento = strFechamento & strDupla(caixaDesc & Chr$(10) & Chr$(10))
                
                'Data Abertura
                'If Not imprimeNormal(MSC, Alinha("DATA ABERTURA...: " & Format$(dtAbertura, "dd/mm/yyyy HH:MM:SS"), COLUNAS_IMP, esquerda) & Chr$(10)) Then
                '    GoTo Imprime_FechamentoCaixa
                'End If
                strFechamento = strFechamento & strNormal(Alinha("DATA ABERTURA...: " & Format$(dtAbertura, "dd/mm/yyyy HH:MM:SS"), COLUNAS_IMP, esquerda) & Chr$(10))
                
                'Data Fechamento
                'If Not imprimeNormal(MSC, Alinha("DATA FECHAMENTO.: " & Format$(oRs.Fields("cxp_dt_fechamento"), "dd/mm/yyyy HH:MM:SS"), COLUNAS_IMP, esquerda) & Chr$(10)) Then
                '    GoTo Imprime_FechamentoCaixa
                'End If
                strFechamento = strFechamento & strNormal(Alinha("DATA FECHAMENTO.: " & Format$(oRs.Fields("cxp_dt_fechamento"), "dd/mm/yyyy HH:MM:SS"), COLUNAS_IMP, esquerda) & Chr$(10))
                
                'Operador do Caixa
                'If Not imprimeNormal(MSC, Alinha("OPERADOR CAIXA..: " & oRs.Fields("usu_nm_abertura"), COLUNAS_IMP, esquerda) & Chr$(10)) Then
                '    GoTo Imprime_FechamentoCaixa
                'End If
                strFechamento = strFechamento & strNormal(Alinha("OPERADOR CAIXA..: " & oRs.Fields("usu_nm_abertura"), COLUNAS_IMP, esquerda) & Chr$(10))
                
                'Fechou o Caixa
                If oRs.Fields("usu_nm_abertura") <> oRs.Fields("usu_nm_fechamento") Then
                    'If Not imprimeNormal(MSC, Alinha("FECHOU CAIXA....: " & oRs.Fields("usu_nm_fechamento"), COLUNAS_IMP, esquerda) & Chr$(10)) Then
                    '    GoTo Imprime_FechamentoCaixa
                    'End If
                    strFechamento = strFechamento & strNormal(Alinha("FECHOU CAIXA....: " & oRs.Fields("usu_nm_fechamento"), COLUNAS_IMP, esquerda) & Chr$(10))
                End If
                
                ' VALORES
                
                'If Not imprimeNormal(MSC, Chr$(10)) Then
                '    GoTo Imprime_FechamentoCaixa
                'End If
                strFechamento = strFechamento & strNormal(Chr$(10))
                
                'If Not imprimeDupla(MSC, Alinha("VALORES", 19, esquerda) & Chr$(10) & Chr$(10)) Then
                '    GoTo Imprime_FechamentoCaixa
                'End If
                strFechamento = strFechamento & strDupla(Alinha("VALORES", 19, esquerda) & Chr$(10) & Chr$(10))
                
                Set oRsAux = New ADODB.Recordset
                dTotal = 0
                            
                If Not clsFechamento.CaixaValores(oRsAux) Then
                    MsgBox clsFechamento.MensagemErro, vbCritical, App.ProductName
                    Exit Function
                End If
            
                Do While Not oRsAux.EOF
                    If oRsAux.Fields("pag_valor") > 0 Then
                        dTotal = dTotal + oRsAux.Fields("sinal") * oRsAux.Fields("pag_valor")
                    End If
                    'If Not imprimeNormal(MSC, Alinha(Alinha(EliminaAcentos(oRsAux.Fields("pgt_desc")), 20, esquerda, ".") & ": " & Alinha(Format(oRsAux.Fields("pag_valor"), "$ #,##0.00"), 15, Direita), COLUNAS_IMP, esquerda) & Chr$(10)) Then
                    '    GoTo Imprime_FechamentoCaixa
                    'End If
                    strFechamento = strFechamento & strNormal(Alinha(Alinha(EliminaAcentos(oRsAux.Fields("pgt_desc")), 20, esquerda, ".") & ": " & Alinha(Format(oRsAux.Fields("pag_valor"), "$ #,##0.00"), 15, Direita), COLUNAS_IMP, esquerda) & Chr$(10))
                    oRsAux.MoveNext
                Loop
                
                If dTotal > 0 Then
                    'If Not imprimeNormal(MSC, String(COLUNAS_IMP, "_") & Chr$(10)) Then
                    '    GoTo Imprime_FechamentoCaixa
                    'End If
                    strFechamento = strFechamento & strNormal(String(COLUNAS_IMP, "_") & Chr$(10))
                    
                    'If Not imprimeNormal(MSC, Alinha(Alinha("TOTAL VALORES", 20, esquerda, ".") & ": " & Alinha(Format(dTotal, "$ #,##0.00"), 15, Direita), COLUNAS_IMP, esquerda) & Chr$(10)) Then
                    '    GoTo Imprime_FechamentoCaixa
                    'End If
                    strFechamento = strFechamento & strNormal(Alinha(Alinha("TOTAL VALORES", 20, esquerda, ".") & ": " & Alinha(Format(dTotal, "$ #,##0.00"), 15, Direita), COLUNAS_IMP, esquerda) & Chr$(10))
                End If
                
                oRsAux.Close
                
                ' BILHETES
                
                'If Not imprimeNormal(MSC, Chr$(10)) Then
                '    GoTo Imprime_FechamentoCaixa
                'End If
                strFechamento = strFechamento & strNormal(Chr$(10))
                
                'If Not imprimeDupla(MSC, Alinha("BILHETES", 19, esquerda) & Chr$(10) & Chr$(10)) Then
                '    GoTo Imprime_FechamentoCaixa
                'End If
                strFechamento = strFechamento & strDupla(Alinha("BILHETES", 19, esquerda) & Chr$(10) & Chr$(10))
                
                Set oRsAux = New ADODB.Recordset
                dTotal = 0
                            
                If Not clsFechamento.CaixaBilhetes(oRsAux) Then
                    MsgBox clsFechamento.MensagemErro, vbCritical, App.ProductName
                    Exit Function
                End If
            
                Do While Not oRsAux.EOF
                    dTotal = dTotal + oRsAux.Fields("qtde")
                    'If Not imprimeNormal(MSC, Alinha(Alinha(EliminaAcentos(oRsAux.Fields("igt_desc")), 20, esquerda, ".") & ": " & Alinha(Format(oRsAux.Fields("qtde"), "##,##0"), 11, Direita), COLUNAS_IMP, esquerda) & Chr$(10)) Then
                    '    GoTo Imprime_FechamentoCaixa
                    'End If
                    strFechamento = strFechamento & strNormal(Alinha(Alinha(EliminaAcentos(oRsAux.Fields("igt_desc")), 20, esquerda, ".") & ": " & Alinha(Format(oRsAux.Fields("qtde"), "##,##0"), 11, Direita), COLUNAS_IMP, esquerda) & Chr$(10))
                    oRsAux.MoveNext
                Loop
                
                If dTotal > 0 Then
                    'If Not imprimeNormal(MSC, String(COLUNAS_IMP, "_") & Chr$(10)) Then
                    '    GoTo Imprime_FechamentoCaixa
                    'End If
                    strFechamento = strFechamento & strNormal(String(COLUNAS_IMP, "_") & Chr$(10))
                    
                    'If Not imprimeNormal(MSC, Alinha(Alinha("TOTAL BILHETES", 20, esquerda, ".") & ": " & Alinha(Format(dTotal, "##,##0"), 11, Direita), COLUNAS_IMP, esquerda) & Chr$(10)) Then
                    '    GoTo Imprime_FechamentoCaixa
                    'End If
                    strFechamento = strFechamento & strNormal(Alinha(Alinha("TOTAL BILHETES", 20, esquerda, ".") & ": " & Alinha(Format(dTotal, "##,##0"), 11, Direita), COLUNAS_IMP, esquerda) & Chr$(10))
                End If
                
                oRsAux.Close
                
                ' COMBOS
                
                Set oRsAux = New ADODB.Recordset
                dTotal = 0
                dTotal2 = 0
                            
                If Not clsFechamento.CaixaCombos(oRsAux) Then
                    MsgBox clsFechamento.MensagemErro, vbCritical, App.ProductName
                    Exit Function
                End If
            
                If Not oRsAux.EOF Then
                    'If Not imprimeNormal(MSC, Chr$(10)) Then
                    '    GoTo Imprime_FechamentoCaixa
                    'End If
                    strFechamento = strFechamento & strNormal(Chr$(10))
                    
                    'If Not imprimeDupla(MSC, Alinha("COMBOS", 19, esquerda) & Chr$(10) & Chr$(10)) Then
                    '    GoTo Imprime_FechamentoCaixa
                    'End If
                    strFechamento = strFechamento & strNormal(Alinha("COMBOS", 19, esquerda) & Chr$(10) & Chr$(10))
                End If
            
                Do While Not oRsAux.EOF
                    dTotal = dTotal + oRsAux.Fields("qtde")
                    dTotal2 = dTotal2 + oRsAux.Fields("valor")
                    'If Not imprimeNormal(MSC, Alinha(Alinha(EliminaAcentos(oRsAux.Fields("cbo_nm")), 20, esquerda, ".") & ": " & _
                    '                                 Alinha(Format(oRsAux.Fields("qtde"), "#,##0"), 5, Direita) & _
                    '                                 Alinha(Format(oRsAux.Fields("valor"), "#,##0.00"), 10, Direita), COLUNAS_IMP, esquerda) & Chr$(10)) Then
                    '    GoTo Imprime_FechamentoCaixa
                    'End If
                    strFechamento = strFechamento & strNormal(Alinha(Alinha(EliminaAcentos(oRsAux.Fields("cbo_nm")), 20, esquerda, ".") & ": " & _
                                                                     Alinha(Format(oRsAux.Fields("qtde"), "#,##0"), 5, Direita) & _
                                                                     Alinha(Format(oRsAux.Fields("valor"), "#,##0.00"), 10, Direita), COLUNAS_IMP, esquerda) & Chr$(10))
                    oRsAux.MoveNext
                Loop
                
                If dTotal > 0 Then
                    'If Not imprimeNormal(MSC, String(COLUNAS_IMP, "_") & Chr$(10)) Then
                    '    GoTo Imprime_FechamentoCaixa
                    'End If
                    strFechamento = strFechamento & strNormal(String(COLUNAS_IMP, "_") & Chr$(10))
                    
                    'If Not imprimeNormal(MSC, Alinha(Alinha("TOTAL COMBOS", 20, esquerda, ".") & ": " & _
                    '                                 Alinha(Format(dTotal, "#,##0"), 5, Direita) & _
                    '                                 Alinha(Format(dTotal2, "#,##0.00"), 10, Direita), COLUNAS_IMP, esquerda) & Chr$(10)) Then
                    '    GoTo Imprime_FechamentoCaixa
                    'End If
                    strFechamento = strFechamento & strNormal(Alinha(Alinha("TOTAL COMBOS", 20, esquerda, ".") & ": " & _
                                                                     Alinha(Format(dTotal, "#,##0"), 5, Direita) & _
                                                                     Alinha(Format(dTotal2, "#,##0.00"), 10, Direita), COLUNAS_IMP, esquerda) & Chr$(10))
                End If
                
                oRsAux.Close
                
                'If Not imprimeNormal(MSC, Chr$(10) & Chr$(10) & Chr$(10) & Chr$(10)) Then
                '    GoTo Imprime_FechamentoCaixa
                'End If
                strFechamento = strFechamento & strNormal(Chr$(10) & Chr$(10) & Chr$(10) & Chr$(10))
                
                'If Not cortaPapel(MSC) Then
                '    GoTo Imprime_FechamentoCaixa
                'End If
                strFechamento = strFechamento & strCortaPapel
                
                If Not imprime(MSC, strFechamento) Then
                    GoTo erro_ImprimeFechamentoCaixa
                End If
                
                lInicio = timeGetTime
                
                Do While lInicio + CInt(pTempImp2) > timeGetTime
                    DoEvents
                Loop
                
                'Do While MSC.OutBufferCount > 0
                '    If ErroImp Then
                '        GoTo erro_ImprimeFechamentoCaixa
                '    End If
                'Loop
            End If
        Else
            GoTo erro_ImprimeFechamentoCaixa
        End If
    End If
    
    ImprimeFechamentoCaixa = True
    
    Exit Function

erro_ImprimeFechamentoCaixa:
    Beep
    
    If MsgBox("Impressora não está pronta! Ajuste-a e clique em SIM para tentar imprimir novamente ou NÃO para desistir.", vbQuestion + vbYesNo, App.ProductName) = vbYes Then
        GoTo Imprime_FechamentoCaixa
    End If

End Function


