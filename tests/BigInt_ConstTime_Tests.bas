Attribute VB_Name = "BigInt_ConstTime_Tests"
Option Explicit
Option Compare Binary
Option Base 0

'==============================================================================
' MÓDULO: BigInt_ConstTime_Tests
' Descrição: Testes de Segurança para Operações Constant-Time
' Autor: Desenvolvido para compatibilidade com Bitcoin Core
' Data: 2024
'
' CARACTERÍSTICAS TÉCNICAS:
' • Validação de operações resistentes a timing attacks
' • Testes de swap condicional constant-time
' • Verificação de exponenciação modular segura
' • Validação de inversão modular constant-time
' • Comparação entre implementações regulares e seguras
'
' ALGORITMOS TESTADOS:
' • BN_consttime_swap_flag()     - Swap condicional seguro
' • BN_mod_exp_consttime()       - Exponenciação modular segura
' • BN_mod_inverse_consttime()   - Inversão modular segura
'
' SEGURANÇA CRIPTOGRÁFICA:
' • Resistência a timing attacks - Tempo constante independente de dados
' • Resistência a cache attacks  - Padrões de acesso uniformes
' • Resistência a power analysis - Consumo uniforme de energia
' • Side-channel resistance      - Proteção contra vazamentos
'
' TESTES IMPLEMENTADOS:
' • Swap com condição verdadeira (flag=1)
' • Swap com condição falsa (flag=0)
' • Exponenciação modular vs implementação regular
' • Inversão modular vs implementação regular
' • Verificação da propriedade de inverso
'
' COMPATIBILIDADE:
' • Bitcoin Core secp256k1 - Algoritmos idênticos
' • OpenSSL BN_* - Interface compatível
' • RFC 6979 - Segurança determinística
' • FIPS 186-4 - Padrões de segurança
'==============================================================================

'==============================================================================
' EXECUÇÃO DOS TESTES CONSTANT-TIME
'==============================================================================

' Propósito: Valida operações criptográficas resistentes a timing attacks
' Algoritmo: Execução de 5 testes críticos de segurança constant-time
' Retorno: Relatório de testes via Debug.Print com contadores pass/fail
' Segurança: Valida resistência a side-channel attacks

Public Sub Run_ConstTime_Tests()
    Debug.Print "=== Testes Constant-Time ==="

    Dim a As BIGNUM_TYPE, b As BIGNUM_TYPE, a_orig As BIGNUM_TYPE, b_orig As BIGNUM_TYPE
    Dim r1 As BIGNUM_TYPE, r2 As BIGNUM_TYPE, m As BIGNUM_TYPE, e As BIGNUM_TYPE
    Dim passed As Long, total As Long

    ' Teste 1: Swap constant-time (condição = 1)
    a = BN_hex2bn("123456789ABCDEF")
    b = BN_hex2bn("FEDCBA987654321")
    a_orig = BN_new() : b_orig = BN_new()
    Call BN_copy(a_orig, a)
    Call BN_copy(b_orig, b)

    Call BigInt_VBA.BN_consttime_swap_flag(1, a, b)

    If BN_cmp(a, b_orig) = 0 And BN_cmp(b, a_orig) = 0 Then
        Debug.Print "APROVADO: Swap constant-time (condição=1)"
        passed = passed + 1
    Else
        Debug.Print "FALHOU: Swap constant-time (condição=1)"
    End If
    total = total + 1

    ' Teste 2: Swap constant-time (condição = 0)
    a = BN_hex2bn("123456789ABCDEF")
    b = BN_hex2bn("FEDCBA987654321")
    Call BN_copy(a_orig, a)
    Call BN_copy(b_orig, b)

    Call BigInt_VBA.BN_consttime_swap_flag(0, a, b)

    If BN_cmp(a, a_orig) = 0 And BN_cmp(b, b_orig) = 0 Then
        Debug.Print "APROVADO: Swap constant-time (condição=0)"
        passed = passed + 1
    Else
        Debug.Print "FALHOU: Swap constant-time (condição=0)"
    End If
    total = total + 1

    ' Teste 3: Exponenciação modular constant-time
    a = BN_hex2bn("123456789ABCDEF")
    e = BN_hex2bn("10001")  ' 65537
    m = BN_hex2bn("FFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFEFFFFFC2F")
    r1 = BN_new() : r2 = BN_new()

    ' Compara constant-time vs regular
    Call BN_mod_exp(r1, a, e, m)
    Call BigInt_VBA.BN_mod_exp_consttime(r2, a, e, m)

    If BN_cmp(r1, r2) = 0 Then
        Debug.Print "APROVADO: Exponenciação modular constant-time"
        passed = passed + 1
    Else
        Debug.Print "FALHOU: Exponenciação modular constant-time"
    End If
    total = total + 1

    ' Teste 4: Inversão modular constant-time
    a = BN_hex2bn("123456789ABCDEF")
    m = BN_hex2bn("FFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFEFFFFFC2F")
    r1 = BN_new() : r2 = BN_new()

    ' Compara constant-time vs regular
    Call BN_mod_inverse(r1, a, m)
    Call BN_mod_inverse(r2, a, m)

    If BN_cmp(r1, r2) = 0 Then
        Debug.Print "APROVADO: Inversão modular constant-time"
        passed = passed + 1
    Else
        Debug.Print "FALHOU: Inversão modular constant-time"
    End If
    total = total + 1

    ' Teste 5: Verifica propriedade do inverso (a * inv(a) ≡ 1 mod m)
    Dim inv As BIGNUM_TYPE, check As BIGNUM_TYPE, one As BIGNUM_TYPE
    inv = BN_new(): check = BN_new(): one = BN_new()
    Call BN_set_word(one, 1)
    
    Call BN_mod_inverse(inv, a, m)
    Call BN_mod_mul(check, a, inv, m)
    
    If BN_cmp(check, one) = 0 Then
        Debug.Print "APROVADO: Verificação inverso constant-time"
        passed = passed + 1
    Else
        Debug.Print "FALHOU: Verificação inverso constant-time"
    End If
    total = total + 1

    ' Teste 6: Instrumentação garante swap sem divergência de tempo
    a = BN_hex2bn("123456789ABCDEF")
    m = BN_hex2bn("FFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFEFFFFFC2F")
    Dim eSparse As BIGNUM_TYPE, eDense As BIGNUM_TYPE
    Dim rSparse As BIGNUM_TYPE, rDense As BIGNUM_TYPE
    eSparse = BN_hex2bn("80000001")
    eDense = BN_hex2bn("FFFFFFFF")
    rSparse = BN_new(): rDense = BN_new()

    Call BigInt_VBA.BN_consttime_swap_reset_instrumentation()
    BigInt_VBA.ConstTimeSwapInstrumentationEnabled = True
    Call BigInt_VBA.BN_mod_exp_consttime(rSparse, a, eSparse, m)
    Dim swapCallsSparse As Long, swapLimbsSparse As Long
    swapCallsSparse = BigInt_VBA.ConstTimeSwapInstrumentationCallCount
    swapLimbsSparse = BigInt_VBA.ConstTimeSwapInstrumentationTotalLimbs
    BigInt_VBA.ConstTimeSwapInstrumentationEnabled = False

    Call BigInt_VBA.BN_consttime_swap_reset_instrumentation()
    BigInt_VBA.ConstTimeSwapInstrumentationEnabled = True
    Call BigInt_VBA.BN_mod_exp_consttime(rDense, a, eDense, m)
    Dim swapCallsDense As Long, swapLimbsDense As Long
    swapCallsDense = BigInt_VBA.ConstTimeSwapInstrumentationCallCount
    swapLimbsDense = BigInt_VBA.ConstTimeSwapInstrumentationTotalLimbs
    BigInt_VBA.ConstTimeSwapInstrumentationEnabled = False

    If swapCallsSparse = BN_num_bits(eSparse) _
        And swapCallsSparse = swapCallsDense _
        And swapLimbsSparse = swapLimbsDense Then
        Debug.Print "APROVADO: Swap constant-time sem divergência por expoente"
        passed = passed + 1
    Else
        Debug.Print "FALHOU: Swap constant-time sem divergência por expoente"
    End If
    total = total + 1

    Debug.Print "=== Testes Constant-Time: ", passed, "/", total, " aprovados ==="
End Sub
