Attribute VB_Name = "CalculosMT"
Sub Calc_Bitola_Corrente()
   
'Calculo da Bitola pelo m�todo de Condu��o de Corrente

    Dim i As Integer
    Dim j As Integer
    Dim conta_linha As Integer
    Dim corrente_corrigida As Double
    Dim num_cond As Integer
    conta_linha = 10
    bitola_limite = Worksheets("C�lculos MT").Cells(7, "I").Value
    bitola_limite_min = Worksheets("C�lculos MT").Cells(8, "I").Value
    
    'Rodar o c�digo abaixo at� o final da planilha C�lculos MT
    
    Do While Worksheets("C�lculos MT").Cells(conta_linha, "S").Value <> ""
    
        If Worksheets("C�lculos MT").Cells(conta_linha, 8).Value = "EPR 90" Then
           
            For i = 1 To 13 'Encontrar na planilha de tabelas a coluna correspondente ao m�todo de instala��o
            
                metodo = Worksheets("C�lculos MT").Cells(conta_linha, "O").Value
                metodo_aux = Worksheets("Tabelas_Norma_CAP_COR_MT").Cells(9, 2 + i).Value
                
                'Descobrindo a coluna correta para fazer a busca
                If metodo_aux = metodo Then
                    'Worksheets("C�lculos MT").Cells(conta_linha, 38).Value = i
                    ref_col = i + 2
                End If
                
            Next i
                               
            num_cond = 1 'Condi��o inicial de n�mero de condutores por fase para capacidade de corrente
            
Calculo_CapCorrente_EPR90:
                               
            For j = 1 To 17 'Encontrar na planilha de tabelas a bitola correspondente ao m�todo de instala��o e corrente calculada
            
                corrente_corrigida = Worksheets("C�lculos MT").Cells(conta_linha, "S").Value / num_cond
            
                'If corrente_corrigida <= Worksheets("Tabelas_Norma_CAP_COR_MT").Cells(10, ref_col).Value Then 'Se a corrente corrigida for menor que a primeira correnta da tabela, bitola � 10mm2
                
                'Worksheets("C�lculos MT").Cells(conta_linha, "V").Value = Worksheets("Tabelas_Norma_CAP_COR_MT").Range("B10").Value
                'Worksheets("C�lculos MT").Cells(conta_linha, "U").Value = num_cond
                
                'End If
                
                'Encontrando a c�lula correspondente a bitola limite
                
                For q = 10 To 26
                
                    If bitola_limite = Worksheets("Tabelas_Norma_CAP_COR_MT").Cells(q, "B").Value Then
                        celulalimite = q
                        Exit For
                    End If
                Next q
                    
                secao_limite = Worksheets("Tabelas_Norma_CAP_COR_MT").Cells(celulalimite, "B").Value
                
                If corrente_corrigida > Worksheets("Tabelas_Norma_CAP_COR_MT").Cells(j + 9, ref_col).Value And corrente_corrigida <= Worksheets("Tabelas_Norma_CAP_COR_MT").Cells(j + 10, ref_col) Then 'Se a corrente corrigida for menor ou igual que a pr�xima bitola, e maior do que a bitola anterior, o valor da bitola escolhido � o pr�ximo
                    secao_encontrada = Worksheets("Tabelas_Norma_CAP_COR_MT").Cells(ref_lin + 9, 2).Value
                    ref_lin = j + 1
                    If secao_encontrada <= bitola_limite_min Then
                        Worksheets("C�lculos MT").Cells(conta_linha, "V").Value = bitola_limite_min
                        Worksheets("C�lculos MT").Cells(conta_linha, "U").Value = num_cond
                    ElseIf secao_encontrada > secao_limite Then
                        num_cond = num_cond + 1
                        GoTo Calculo_CapCorrente_EPR90
                    Else
                        Worksheets("C�lculos MT").Cells(conta_linha, "V").Value = secao_encontrada
                        Worksheets("C�lculos MT").Cells(conta_linha, "U").Value = num_cond
                    End If
                End If
                
                If corrente_corrigida > Worksheets("Tabelas_Norma_CAP_COR_MT").Cells(26, ref_col).Value Then
                    num_cond = num_cond + 1 'Incremento do n�mero de condutores
                    GoTo Calculo_CapCorrente_EPR90
                End If
                   
            Next j
        Else
        
            For i = 1 To 13 'Encontrar na planilha de tabelas a coluna correspondente ao m�todo de instala��o
            
                If Worksheets("Tabelas_Norma_CAP_COR_MT").Cells(9, 15 + i).Value = Worksheets("C�lculos MT").Cells(conta_linha, "O").Value Then
                
                'Worksheets("C�lculos MT").Cells(conta_linha, 38).Value = i + 15
                ref_col = i + 15
                
                End If
                
            Next i
                               
            num_cond = 1 'Condi��o inicial de n�mero de condutores por fase para capacidade de corrente
            
Calculo_CapCorrente_EPR105:
                               
            For j = 1 To 17 'Encontrar na planilha de tabelas a bitola correspondente ao m�todo de instala��o e corrente calculada
            
                corrente_corrigida = Worksheets("C�lculos MT").Cells(conta_linha, "S").Value / num_cond
            
                'If corrente_corrigida <= Worksheets("Tabelas_Norma_CAP_COR_MT").Cells(10, ref_col).Value Then 'Se a corrente corrigida for menor que a primeira correnta da tabela, bitola � 10mm2
                
                'Worksheets("C�lculos MT").Cells(conta_linha, "V").Value = Worksheets("Tabelas_Norma_CAP_COR_MT").Range("B10").Value
                'Worksheets("C�lculos MT").Cells(conta_linha, "U").Value = num_cond
                
                'End If
                
                For q = 10 To 26
                
                    If bitola_limite = Worksheets("Tabelas_Norma_CAP_COR_MT").Cells(q, "B").Value Then
                        celulalimite = q
                        Exit For
                    End If
                Next q
                    
                secao_limite = Worksheets("Tabelas_Norma_CAP_COR_MT").Cells(celulalimite, "B").Value
                
                If corrente_corrigida > Worksheets("Tabelas_Norma_CAP_COR_MT").Cells(j + 9, ref_col).Value And corrente_corrigida <= Worksheets("Tabelas_Norma_CAP_COR_MT").Cells(j + 10, ref_col) Then 'Se a corrente corrigida for menor ou igual que a pr�xima bitola, e maior do que a bitola anterior, o valor da bitola escolhido � o pr�ximo
                    ref_lin = j + 1
                    secao_encontrada = Worksheets("Tabelas_Norma_CAP_COR_MT").Cells(ref_lin + 9, 2).Value
                    If secao_encontrada < bitola_limite_min Then
                        Worksheets("C�lculos MT").Cells(conta_linha, "V").Value = bitola_limite_min
                        Worksheets("C�lculos MT").Cells(conta_linha, "U").Value = num_cond
                    ElseIf secao_encontrada > secao_limite Then
                        num_cond = num_cond + 1
                        GoTo Calculo_CapCorrente_EPR105
                    Else
                        Worksheets("C�lculos MT").Cells(conta_linha, "V").Value = Worksheets("Tabelas_Norma_CAP_COR_MT").Cells(ref_lin + 9, 2).Value
                        Worksheets("C�lculos MT").Cells(conta_linha, "U").Value = num_cond
                    End If
                End If
                
                If corrente_corrigida > Worksheets("Tabelas_Norma_CAP_COR_MT").Cells(26, ref_col).Value Then
                    num_cond = num_cond + 1 'Incremento do n�mero de condutores
                    GoTo Calculo_CapCorrente_EPR105
                End If
                   
            Next j
        
        End If
        
    conta_linha = conta_linha + 1
     
    Loop
                                                     
End Sub

Sub Calc_Bitola_Curto_Circuito()

'Calculo da Bitola pelo m�todo de Curto Circuito

    Dim conta_linha As Integer
    Dim corrente_curto As Double
    Dim tempo_op_prote As Double
    Dim bitola_curto_aux As Double
    Dim bitola_curto As Double
    Dim i As Integer
    
    conta_linha = 10 'Vari�vel auxiliar de contagem. Representa primeira linha de dados da aba C�lculos MT
    num_cond = 1
    bitola_limite_min = Worksheets("C�lculos MT").Cells(8, "I").Value
    
    Do While Worksheets("C�lculos MT").Cells(conta_linha, 10).Value <> "" 'Rotina para rodar o c�digo abaixo at� a �ltima c�lula com dados da aba C�lculos MT
    
        corrente_curto = Worksheets("C�lculos MT").Cells(conta_linha, "J").Value 'Corrente de curto circuito conforme aba conforme aba C�lculos MT
        tempo_op_prot = Worksheets("C�lculos MT").Cells(conta_linha, "K").Value 'Tempo de opera��o da prote��o conforme aba conforme aba C�lculos MT
        bitola_curto_aux = (1000 * corrente_curto * Sqr(tempo_op_prot)) / 142 'C�lculo te�rico de bitola de curto circuito para cabo EPR de cobre conex�es aparafusadas
        
        For i = 10 To 26 'Percorrendo bitolas de cabos desde 10mm2 a 630mm2
        
            'If bitola_curto_aux <= Worksheets("Tabelas_Norma_CAP_COR_MT").Cells(10, 2).Value Then
            
                'bitola_curto = Worksheets("Tabelas_Norma_CAP_COR_MT").Cells(10, 2).Value 'Se o c�lculo te�rico de bitola for menor ou igual a 10mm2, a bitola final ser� 10mm2
                
            'End If
            
            If bitola_curto_aux > Worksheets("Tabelas_Norma_CAP_COR_MT").Cells(i, 2).Value And bitola_curto_aux <= Worksheets("Tabelas_Norma_CAP_COR_MT").Cells(i + 1, 2).Value Then
            
                bitola_curto = Worksheets("Tabelas_Norma_CAP_COR_MT").Cells(i + 1, 2).Value 'Rotina para calcular a bitola final de curto circuito, com base na te�rica. Pega o valor de bitola tabelado logo acima do calculado
            
            End If
        If bitola_curto <= bitola_limite_min Then
            Worksheets("C�lculos MT").Cells(conta_linha, "Y").Value = bitola_limite_min 'Registro da bitola calculada de curto circuito na aba C�lculos MT
            Worksheets("C�lculos MT").Cells(conta_linha, "X").Value = num_cond
        Else
            Worksheets("C�lculos MT").Cells(conta_linha, "Y").Value = bitola_curto 'Registro da bitola calculada de curto circuito na aba C�lculos MT
            Worksheets("C�lculos MT").Cells(conta_linha, "X").Value = num_cond
        End If
        
        Next i
           
    conta_linha = conta_linha + 1 'Passa para a pr�xima linha com dados da aba C�lculos MT
    
    Loop

End Sub
Sub Calc_Bitola_Queda_de_Tensao()

    Dim Rca As Double
    Dim Xl As Double
    Dim i As Integer
    Dim conta_linha As Integer
    Dim comprimento As Double
    Dim corrente_corrigida As Double
    Dim tensao As Double
    Dim classe_tensao As String
    Dim Tipo_Cabo As String
    Dim tipo_inst As String
    Dim concat As String
    Dim fp As Double
    Dim queda_tensao As Double
    Dim lim_queda_tensao_regime As Double
    Dim bitola_queda_tensao As Double
    Dim bitola_teste_queda_tensao As Double
    Dim num_cond As Integer
       
    lim_queda_tensao_regime = Worksheets("C�lculos MT").Cells(4, "I").Value 'Limite de queda de tens�o conforme aba C�lculos MT
    bitola_limite = Worksheets("C�lculos MT").Cells(7, "I").Value
    bitola_limite_min = Worksheets("C�lculos MT").Cells(8, "I").Value
    conta_linha = 10 'Vari�vel auxiliar de contagem. Representa primeira linha de dados da aba C�lculos MT
    
    Do While Worksheets("C�lculos MT").Cells(conta_linha, 18).Value <> "" 'Rotina para rodar o c�digo abaixo at� a �ltima c�lula com dados da aba C�lculos MT
                             
        classe_tensao = Worksheets("C�lculos MT").Cells(conta_linha, "I").Value 'Classe de tens�o conforme aba C�lculos MT
        Tipo_Cabo = Worksheets("C�lculos MT").Cells(conta_linha, "H").Value 'Tipo de cabo conforme aba conforme aba C�lculos MT
        tipo_inst = Worksheets("C�lculos MT").Cells(conta_linha, "M").Value 'Tipo de instala��o  conforme aba C�lculos MT
        comprimento = Worksheets("C�lculos MT").Cells(conta_linha, "L").Value 'Comprimento do circuito conforme aba C�lculos MT
        corrente_corrigida = Worksheets("C�lculos MT").Cells(conta_linha, "S").Value 'Corrente corrigida por fatores conforme aba C�lculos MT
        tensao = Worksheets("C�lculos MT").Cells(conta_linha, "C").Value 'N�vel de tens�o conforme aba C�lculos MT
        fp = Worksheets("C�lculos MT").Cells(conta_linha, "E").Value 'Fator de Pot�ncia conforme aba C�lculos MT
        num_cond = 1 'Condi��o inicial de n�mero de condutores por fase para queda de tens�o
        sen = Sqr(1 - fp ^ 2)
        
Calculo_Queda: 'Rotina para calcular a queda de tens�o
        
        For i = 10 To 24 'Percorrendo bitolas de cabos desde 10mm2 a 630mm2
            
            
            bitola_teste_queda_tensao = Worksheets("Tabelas_Norma_CAP_COR_MT").Cells(i, 2).Value
            concat = bitola_teste_queda_tensao & " - " & classe_tensao & " - " & Tipo_Cabo & " - " & tipo_inst 'Concatenar vari�veis para fazer procv na aba Tabelas Norma Queda Tens�o
            
            Rca = Application.WorksheetFunction.VLookup(concat, Worksheets("Tabelas_Norma_Queda_Tensao_MT").Range("E:G"), 2, False) 'Fixando o valor de Rca para bitola
            Xl = Application.WorksheetFunction.VLookup(concat, Worksheets("Tabelas_Norma_Queda_Tensao_MT").Range("E:G"), 3, False) 'Fixando o valor de Xl para bitola

            queda_tensao = (Sqr(3) * (corrente_corrigida / num_cond) * (comprimento / 1000) * ((Rca * fp) + (Xl * sen) / num_cond)) / (tensao * 1000) 'C�lculo queda de tens�o com base na bitola do cabo
                           
        If queda_tensao <= lim_queda_tensao_regime Then 'Se a queda de tens�o calculada for menor do que o limite normativo, o cabo � escolhido
            If bitola_teste_queda_tensao > bitola_limite Then
                num_cond = num_cond + 1
                GoTo Calculo_Queda
            ElseIf bitola_teste_queda_tensao <= bitola_limite_min Then
                Worksheets("C�lculos MT").Cells(conta_linha, "AB").Value = bitola_limite_min
                Worksheets("C�lculos MT").Cells(conta_linha, "AA").Value = num_cond
            Else
                bitola_queda_tensao = bitola_teste_queda_tensao
                Worksheets("C�lculos MT").Cells(conta_linha, "AB").Value = bitola_queda_tensao
                Worksheets("C�lculos MT").Cells(conta_linha, "AA").Value = num_cond
                Exit For
            End If
        If bitola_teste_queda_tensao >= bitola_limite Then
                num_cond = num_cond + 1
                GoTo Calculo_Queda
                Exit For
        End If
        'Else
            'Worksheets("C�lculos MT").Cells(conta_linha, 25).Value = queda_tensao
            'Worksheets("C�lculos MT").Cells(conta_linha, 27).Value = concat
            'Worksheets("C�lculos MT").Cells(conta_linha, 28).Value = Rca
            'Worksheets("C�lculos MT").Cells(conta_linha, 29).Value = Xl
            'Worksheets("C�lculos MT").Cells(conta_linha, "AA").Value = bitola_teste_queda_tensao
                                                                          
        End If
        
        If bitola_teste_queda_tensao >= 630 Then 'Se bitola for maior ou igual a 630mm2, o n�mero de condutores tem que ser aumentado
            num_cond = num_cond + 1 'Incremento do n�mero de condutores
            GoTo Calculo_Queda 'Envia o c�digo para a rotina Calculo_Queda, para ser realizada com o n�mero de condutores incrementado
        End If
        Next i
    conta_linha = conta_linha + 1 'Passa para a pr�xima linha com dados da aba C�lculos MT
    Loop
End Sub
Sub mainMT()
   
    Calc_Bitola_Corrente
    Calc_Bitola_Curto_Circuito
    Calc_Bitola_Queda_de_Tensao
    MelhorEscolha
    NeutroeTerra

End Sub
Sub NeutroeTerra()
    Dim wsCalculosMT As Worksheet
    Dim wsTerraeNeutro As Worksheet
    Dim i As Integer
    Set wsCalculosMT = ThisWorkbook.Sheets("C�lculos MT")
    Set wsTerraeNeutro = ThisWorkbook.Sheets("Terra e Neutro")
    
    LastRowCalcMT = wsCalculosMT.Cells(wsCalculosMT.Rows.Count, "F").End(xlUp).Row
    
    For i = 10 To LastRowCalcMT
    
    sistema = wsCalculosMT.Cells(i, "F").Value
    secao = wsCalculosMT.Cells(i, "AE").Value
        If sistema = "3F" Then
            wsCalculosMT.Cells(i, "AG").Value = "-"
            wsCalculosMT.Cells(i, "AH").Value = "-"
        End If
        
        If sistema = "3F + N" Then
            For j = 2 To 14
                If wsTerraeNeutro.Cells(j, "A").Value = secao Then
                    wsCalculosMT.Cells(i, "AG").Value = "-"
                    wsCalculosMT.Cells(i, "AH").Value = wsTerraeNeutro.Cells(j, "B").Value
                End If
            Next j
        End If
                
        If sistema = "3F + T" Then
            For q = 2 To 14
                If wsTerraeNeutro.Cells(q, "A").Value = secao Then
                    secao_terra = wsTerraeNeutro.Cells(q, "C").Value
                    wsCalculosMT.Cells(i, "AG").Value = secao_terra
                    wsCalculosMT.Cells(i, "AH").Value = "-"
                End If
            Next q
        End If
    Next i
End Sub
Sub MelhorEscolha()

    Dim wsCalculosMT As Worksheet
    Set wsCalculosMT = ThisWorkbook.Sheets("C�lculos MT")
    Dim i As Integer
    LastRowCalcMT = wsCalculosMT.Cells(wsCalculosMT.Rows.Count, "C").End(xlUp).Row
    
    For i = 10 To LastRowCalcMT
    
        'quantidade
        quantidade_cap_corrente = wsCalculosMT.Cells(i, "U").Value
        'secao
        cap_corrente = wsCalculosMT.Cells(i, "V").Value
        'secao
        curto = wsCalculosMT.Cells(i, "Y").Value
        'quantidade
        quantidade_curto = wsCalculosMT.Cells(i, "X").Value
        'quantidade
        quantidade_queda_tensao = wsCalculosMT.Cells(i, "AA").Value
        'secao
        queda_tensao = wsCalculosMT.Cells(i, "AB").Value
        
    '    If (quantidade_queda_tensao) > (quantidade_cap_corrente) Then
    '        wsCalculosMT.Cells(i, "AD").Value = quantidade_queda_tensao
    '    Else
    '        wsCalculosMT.Cells(i, "AD").Value = quantidade_cap_corrente
    '    End If
    '
    '    If queda_tensao > cap_corrente Then
    '        wsCalculosMT.Cells(i, "AE").Value = queda_tensao
    '    Else
    '        wsCalculosMT.Cells(i, "AE").Value = cap_corrente
    '    End If
        
        mult_cap_corrente = quantidade_cap_corrente * cap_corrente
        mult_curto = quantidade_curto * curto
        mult_queda_tensao = quantidade_queda_tensao * queda_tensao
        
        If mult_cap_corrente >= mult_curto And mult_cap_corrente >= mult_queda_tensao Then
        
            wsCalculosMT.Cells(i, "AD").Value = quantidade_cap_corrente
            wsCalculosMT.Cells(i, "AE").Value = cap_corrente
            
        ElseIf mult_curto >= mult_cap_corrente And mult_curto >= mult_queda_tensao Then
       
            wsCalculosMT.Cells(i, "AD").Value = quantidade_curto
            wsCalculosMT.Cells(i, "AE").Value = curto
            
        ElseIf mult_queda_tensao >= mult_cap_corrente And mult_queda_tensao >= mult_curto Then
       
            wsCalculosMT.Cells(i, "AD").Value = quantidade_queda_tensao
            wsCalculosMT.Cells(i, "AE").Value = queda_tensao
            
        End If
        
    Next i
    
End Sub
