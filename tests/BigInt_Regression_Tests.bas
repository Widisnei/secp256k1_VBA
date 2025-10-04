Attribute VB_Name = "BigInt_Regression_Tests"
Option Explicit
Option Compare Binary
Option Base 0

'==============================================================================
' MÓDULO: BigInt_Regression_Tests
' Descrição: Testes de Regressão para Bugs Críticos Já Corrigidos
' Autor: Desenvolvido para compatibilidade com Bitcoin Core
' Data: 2024
'
' CARACTERÍSTICAS TÉCNICAS:
' • Detecção precoce de regressões em bugs críticos
' • Validação de correções de identidade de divisão
' • Testes de propagação de carry em adições
' • Verificação de manipulação de sinais
' • Validação de casos extremos e limites
'
' BUGS TESTADOS (JÁ CORRIGIDOS):
' • Division Identity Bug      - Identidade q*d + r = a falhava
' • Carry Propagation Bug      - Carries não propagavam corretamente
' • Sign Handling Bug          - Sinais incorretos em BN_add
' • Reserved Word Bugs         - Uso de palavras reservadas VBA
' • Boundary Edge Cases        - Casos limites em módulo p
'
' CASOS CRÍTICOS TESTADOS:
' • Divisão com números negativos
' • Adição com carry máximo (0xFFFF...)
' • Manipulação de sinais opostos
' • Parâmetros com nomes seguros (modulus vs mod)
' • Operações no limite do módulo p
'
' ALGORITMOS VALIDADOS:
' • BN_div()                   - Divisão com resto
' • BN_add()                   - Adição com propagação de carry
' • BN_mod_add()               - Adição modular
' • BN_mod_inverse()           - Inversão modular
' • BN_MONT_CTX_set()          - Contexto Montgomery
'
' IMPORTÂNCIA DOS TESTES:
' • Previne reintroducção de bugs críticos
' • Garante estabilidade após modificações
' • Valida correções de segurança
' • Detecta efeitos colaterais de mudanças
'
' COMPATIBILIDADE:
' • Bitcoin Core - Casos de uso reais testados
' • OpenSSL BN_* - Comportamento idêntico validado
' • VBA - Evita palavras reservadas e limitações
' • Produção - Casos que falharam em ambiente real
'==============================================================================

'==============================================================================
' EXECUÇÃO DOS TESTES DE REGRESSÃO
'==============================================================================

' Propósito: Detecta regressões em bugs críticos já corrigidos
' Algoritmo: Execução de 5 suítes de teste para bugs conhecidos
' Retorno: Relatório de regressão via Debug.Print
' Crítico: Falhas indicam reintroducção de bugs graves

Public Sub Run_Regression_Tests()
    Debug.Print "=== TESTES DE REGRESSÃO ==="

    Dim passed As Long, total As Long

    ' Teste 1: Bug identidade de divisão (já corrigido)
    Call Test_Division_Identity_Bug(passed, total)

    ' Teste 2: Bug propagação de carry
    Call Test_Carry_Propagation_Bug(passed, total)

    ' Teste 3: Bug manipulação de sinais em BN_add
    Call Test_Sign_Handling_Bug(passed, total)

    ' Teste 4: Bug palavras reservadas Montgomery
    Call Test_Reserved_Word_Bugs(passed, total)

    ' Teste 5: Casos extremos de condições limite
    Call Test_Boundary_Edge_Cases(passed, total)

    ' Teste 6: Regressões BN_mul/BN_mod_mul 256-bit
    Call Test_BN_Mul_ModMul_Regression(passed, total)

    Debug.Print "=== TESTES DE REGRESSÃO: ", passed, "/", total, " APROVADOS ==="
    If passed = total Then
        Debug.Print "*** NENHUMA REGRESSÃO DETECTADA ***"
    Else
        Debug.Print "*** CRÍTICO: REGRESSÃO DETECTADA ***"
    End If
End Sub

' Testa regressões relacionadas à multiplicação 256-bit e redução modular
Private Sub Test_BN_Mul_ModMul_Regression(ByRef passed As Long, ByRef total As Long)
    Debug.Print "Testando regressões BN_mul/BN_mod_mul 256-bit..."

    Dim a As BIGNUM_TYPE, b As BIGNUM_TYPE
    Dim result_mul As BIGNUM_TYPE, expected_mul As BIGNUM_TYPE
    Dim modulus As BIGNUM_TYPE, result_mod As BIGNUM_TYPE, expected_mod As BIGNUM_TYPE

    a = BN_hex2bn("FFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFE")
    b = BN_hex2bn("AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAB")
    result_mul = BN_new()
    expected_mul = BN_hex2bn("AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA9AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA")

    If BN_mul(result_mul, a, b) And BN_cmp(result_mul, expected_mul) = 0 Then
        passed = passed + 1
        Debug.Print "APROVADO: BN_mul 256-bit combina com resultado esperado"
    Else
        Debug.Print "FALHOU: BN_mul 256-bit incorreto"
        Debug.Print "  Obtido: ", BN_bn2hex(result_mul)
        Debug.Print "  Esperado: ", BN_bn2hex(expected_mul)
    End If
    total = total + 1

    modulus = BN_hex2bn("FFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFEFFFFFC2F")
    result_mod = BN_new()
    expected_mod = BN_hex2bn("AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAB55555A6A555F04B9")

    If BN_mod_mul(result_mod, a, b, modulus) And BN_cmp(result_mod, expected_mod) = 0 Then
        passed = passed + 1
        Debug.Print "APROVADO: BN_mod_mul 256-bit reduz corretamente"
    Else
        Debug.Print "FALHOU: BN_mod_mul 256-bit incorreto"
        Debug.Print "  Obtido: ", BN_bn2hex(result_mod)
        Debug.Print "  Esperado: ", BN_bn2hex(expected_mod)
    End If
    total = total + 1
End Sub

' Testa regressão do bug de identidade de divisão
Private Sub Test_Division_Identity_Bug(ByRef passed As Long, ByRef total As Long)
    Debug.Print "Testando regressão identidade de divisão..."

    ' Este era o caso original que falhava
    Dim a As BIGNUM_TYPE, d As BIGNUM_TYPE, q As BIGNUM_TYPE, r As BIGNUM_TYPE, check As BIGNUM_TYPE
    a = BN_hex2bn("-123456789ABCDEF0123456789ABCDEF")
    d = BN_hex2bn("FEDCBA987654321")
    q = BN_new() : r = BN_new() : check = BN_new()

    Call BN_div(q, r, a, d)
    Call BN_mul(check, q, d)
    Call BN_add(check, check, r)

    ' Verifica identidade: q*d + r = a
    If BN_cmp(check, a) = 0 Then
        passed = passed + 1
        Debug.Print "APROVADO: Regressão identidade de divisão"
    Else
        Debug.Print "FALHOU: Regressão identidade de divisão - BUG RETORNOU!"
        Debug.Print "  a = ", BN_bn2hex(a)
        Debug.Print "  q*d+r = ", BN_bn2hex(check)
    End If
    total = total + 1

    ' Caso extremo adicional que era problemático
    a = BN_hex2bn("-1")
    d = BN_hex2bn("FFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFEFFFFFC2F")

    Call BN_div(q, r, a, d)
    Call BN_mul(check, q, d)
    Call BN_add(check, check, r)

    If BN_cmp(check, a) = 0 And Not r.neg Then
        passed = passed + 1
        Debug.Print "APROVADO: Regressão divisão -1 mod p"
    Else
        Debug.Print "FALHOU: Regressão divisão -1 mod p"
    End If
    total = total + 1
End Sub

' Testa regressão do bug de propagação de carry
Private Sub Test_Carry_Propagation_Bug(ByRef passed As Long, ByRef total As Long)
    Debug.Print "Testando regressão propagação de carry..."

    ' Caso de teste que poderia causar problemas de propagação de carry
    Dim a As BIGNUM_TYPE, b As BIGNUM_TYPE, result As BIGNUM_TYPE, expected As BIGNUM_TYPE
    a = BN_hex2bn("FFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFF")
    b = BN_hex2bn("FFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFF")
    result = BN_new()
    expected = BN_hex2bn("1FFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFE")

    Call BN_add(result, a, b)
    If BN_cmp(result, expected) = 0 Then
        passed = passed + 1
        Debug.Print "APROVADO: Regressão propagação carry máximo"
    Else
        Debug.Print "FALHOU: Regressão propagação carry máximo"
        Debug.Print "  Obtido: ", BN_bn2hex(result)
        Debug.Print "  Esperado: ", BN_bn2hex(expected)
    End If
    total = total + 1

    ' Teste de carry em cadeia
    a = BN_hex2bn("FFFFFFFF")
    b = BN_hex2bn("1")
    expected = BN_hex2bn("100000000")

    Call BN_add(result, a, b)
    If BN_cmp(result, expected) = 0 Then
        passed = passed + 1
        Debug.Print "APROVADO: Regressão carry em cadeia"
    Else
        Debug.Print "FALHOU: Regressão carry em cadeia"
    End If
    total = total + 1
End Sub

' Testa regressão do bug de manipulação de sinais
Private Sub Test_Sign_Handling_Bug(ByRef passed As Long, ByRef total As Long)
    Debug.Print "Testando regressão manipulação de sinais..."

    ' Testa o caso específico que falhava em BN_add
    Dim a As BIGNUM_TYPE, b As BIGNUM_TYPE, result As BIGNUM_TYPE, expected As BIGNUM_TYPE
    a = BN_hex2bn("-123456789ABCDEF0138138138138138")
    b = BN_hex2bn("14CE19AE67B349")
    expected = BN_hex2bn("-123456789ABCDEF0123456789ABCDEF")
    result = BN_new()

    Call BN_add(result, a, b)
    If BN_cmp(result, expected) = 0 Then
        passed = passed + 1
        Debug.Print "APROVADO: Regressão manipulação de sinais"
    Else
        Debug.Print "FALHOU: Regressão manipulação de sinais - BUG RETORNOU!"
        Debug.Print "  Obtido: ", BN_bn2hex(result)
        Debug.Print "  Esperado: ", BN_bn2hex(expected)
    End If
    total = total + 1

    ' Testa resultado zero de sinais opostos
    a = BN_hex2bn("123456789ABCDEF")
    b = BN_hex2bn("-123456789ABCDEF")

    Call BN_add(result, a, b)
    If BN_is_zero(result) And Not result.neg Then
        passed = passed + 1
        Debug.Print "APROVADO: Regressão sinal resultado zero"
    Else
        Debug.Print "FALHOU: Regressão sinal resultado zero"
    End If
    total = total + 1
End Sub

' Testa regressão do bug de palavras reservadas
Private Sub Test_Reserved_Word_Bugs(ByRef passed As Long, ByRef total As Long)
    Debug.Print "Testando regressão palavras reservadas..."

    ' Testa que não usamos 'mod' ou 'rem' como nomes de variáveis
    Dim modulus As BIGNUM_TYPE, remainder As BIGNUM_TYPE, quotient As BIGNUM_TYPE
    Dim dividend As BIGNUM_TYPE

    modulus = BN_hex2bn("FFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFEFFFFFC2F")
    dividend = BN_hex2bn("123456789ABCDEF0123456789ABCDEF0123456789ABCDEF0123456789ABCDEF0")
    quotient = BN_new() : remainder = BN_new()

    ' Isto deve compilar e executar sem erros de palavras reservadas VBA
    Call BN_div(quotient, remainder, dividend, modulus)

    If quotient.top > 0 Or remainder.top >= 0 Then
        passed = passed + 1
        Debug.Print "APROVADO: Evitar palavras reservadas"
    Else
        Debug.Print "FALHOU: Evitar palavras reservadas"
    End If
    total = total + 1

    ' Testa contexto Montgomery com parâmetro modulus
    Dim ctx As MONT_CTX
    ctx = BN_MONT_CTX_new()

    If BN_MONT_CTX_set(ctx, modulus) Then
        passed = passed + 1
        Debug.Print "APROVADO: Parâmetro modulus Montgomery"
    Else
        Debug.Print "FALHOU: Parâmetro modulus Montgomery"
    End If
    total = total + 1
End Sub

' Testa regressão de casos extremos de limite
Private Sub Test_Boundary_Edge_Cases(ByRef passed As Long, ByRef total As Long)
    Debug.Print "Testando regressão casos extremos de limite..."

    ' Casos de teste que estavam no limite entre funcionar/falhar
    Dim p As BIGNUM_TYPE, zero As BIGNUM_TYPE, one As BIGNUM_TYPE, p_minus_1 As BIGNUM_TYPE
    Dim result As BIGNUM_TYPE

    p = BN_hex2bn("FFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFEFFFFFC2F")
    zero = BN_new() : one = BN_new() : p_minus_1 = BN_new() : result = BN_new()
    Call BN_set_word(one, 1)
    Call BN_sub(p_minus_1, p, one)

    ' Testa p - 1 + 1 = 0 (mod p) - isto era complicado
    Call BN_mod_add(result, p_minus_1, one, p)
    If BN_is_zero(result) Then
        passed = passed + 1
        Debug.Print "APROVADO: Regressão wraparound de limite"
    Else
        Debug.Print "FALHOU: Regressão wraparound de limite"
    End If
    total = total + 1

    ' Testa inverso de 1
    Dim inv_one As BIGNUM_TYPE : inv_one = BN_new()
    If BN_mod_inverse(inv_one, one, p) Then
        If BN_cmp(inv_one, one) = 0 Then
            passed = passed + 1
            Debug.Print "APROVADO: Regressão inverso de 1"
        Else
            Debug.Print "FALHOU: Regressão inverso de 1"
        End If
    Else
        Debug.Print "FALHOU: Computação inverso de 1"
    End If
    total = total + 1

    ' Testa números muito pequenos
    Dim small As BIGNUM_TYPE : small = BN_new()
    Call BN_set_word(small, 2)
    Call BN_mod_mul(result, small, small, p)

    Dim four As BIGNUM_TYPE : four = BN_new()
    Call BN_set_word(four, 4)

    If BN_cmp(result, four) = 0 Then
        passed = passed + 1
        Debug.Print "APROVADO: Regressão número pequeno"
    Else
        Debug.Print "FALHOU: Regressão número pequeno"
    End If
    total = total + 1
End Sub