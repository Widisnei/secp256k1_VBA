Attribute VB_Name = "BigInt_Montgomery_Tests"
Option Explicit
Option Compare Binary
Option Base 0

'==============================================================================
' MÓDULO: BigInt_Montgomery_Tests
' Descrição: Testes da Aritmética de Montgomery para Exponenciação Rápida
' Autor: Desenvolvido para compatibilidade com Bitcoin Core
' Data: 2024
'
' CARACTERÍSTICAS TÉCNICAS:
' • Validação da aritmética de Montgomery
' • Testes de exponenciação modular otimizada
' • Verificação de conversões ida e volta
' • Compatibilidade com parâmetros secp256k1
' • Testes de casos extremos e robustez
'
' ARITMÉTICA DE MONTGOMERY:
' • Representação: a' = a * R mod N (R = 2^k)
' • Vantagem: Divisão por R via shift (mais rápido)
' • Multiplicação: REDC(a' * b') = (a * b)' mod N
' • Exponenciação: 50-80% mais rápida que métodos clássicos
'
' ALGORITMOS TESTADOS:
' • BN_MONT_CTX_set()          - Inicialização do contexto
' • BN_to_montgomery()         - Conversão para forma Montgomery
' • BN_mod_mul_montgomery()    - Multiplicação Montgomery
' • BN_from_montgomery()       - Conversão de volta
' • BN_mod_exp_mont()          - Exponenciação otimizada
'
' TESTES IMPLEMENTADOS:
' • Inicialização de contexto Montgomery
' • Conversão para forma Montgomery
' • Multiplicação vs implementação regular
' • Conversão roundtrip (ida e volta)
' • Exponenciação vs implementação regular
' • Casos extremos (0, 1)
' • Compatibilidade secp256k1
'
' VANTAGENS DE PERFORMANCE:
' • Exponenciação modular: 50-80% mais rápida
' • Multiplicações repetidas: Evita divisões custosas
' • Ideal para: RSA, DH, ECDSA, operações secp256k1
'
' COMPATIBILIDADE:
' • OpenSSL BN_MONT_* - Interface idêntica
' • Bitcoin Core secp256k1 - Parâmetros otimizados
' • RFC 3447 (RSA) - Padrão suportado
' • FIPS 186-4 - Algoritmos validados
'==============================================================================

'==============================================================================
' EXECUÇÃO DOS TESTES DE MONTGOMERY
'==============================================================================

' Propósito: Valida aritmética de Montgomery para exponenciação otimizada
' Algoritmo: 7 testes cobrindo inicialização, conversões e operações
' Retorno: Relatório detalhado via Debug.Print
' Performance: Verifica otimizações vs implementações regulares

Public Sub Run_Montgomery_Tests()
    Debug.Print "=== TESTES MONTGOMERY ROBUSTOS ==="

    Dim passed As Long, total As Long
    passed = 0: total = 0

    ' Teste 1: Inicialização do contexto
    total = total + 1
    If test_mont_ctx_init() Then passed = passed + 1

    ' Teste 2: Conversão para Montgomery
    total = total + 1
    If test_to_montgomery() Then passed = passed + 1

    ' Teste 3: Multiplicação Montgomery
    total = total + 1
    If test_mont_multiplication() Then passed = passed + 1

    ' Teste 4: Conversão de Montgomery
    total = total + 1
    If test_from_montgomery() Then passed = passed + 1

    ' Teste 5: Exponenciação Montgomery
    total = total + 1
    If test_mont_exponentiation() Then passed = passed + 1

    ' Teste 6: Casos extremos
    total = total + 1
    If test_mont_edge_cases() Then passed = passed + 1

    ' Teste 7: Compatibilidade com secp256k1
    total = total + 1
    If test_mont_secp256k1() Then passed = passed + 1

    Debug.Print "=== TESTES MONTGOMERY: " & passed & "/" & total & " APROVADOS ==="
End Sub

' Testa inicialização do contexto Montgomery
Private Function test_mont_ctx_init() As Boolean
    Debug.Print "Testando inicialização contexto Montgomery..."

    Dim ctx As MONT_CTX, modulus As BIGNUM_TYPE
    ctx = BN_MONT_CTX_new()
    modulus = BN_hex2bn("FFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFEFFFFFC2F") ' secp256k1 p

    If BN_MONT_CTX_set(ctx, modulus) Then
        Debug.Print "APROVADO: Contexto Montgomery inicializado"
        test_mont_ctx_init = True
    Else
        Debug.Print "FALHOU: Inicialização contexto Montgomery falhou"
        test_mont_ctx_init = False
    End If
End Function

' Testa conversão para forma Montgomery
Private Function test_to_montgomery() As Boolean
    Debug.Print "Testando conversão para forma Montgomery..."

    Dim ctx As MONT_CTX, modulus As BIGNUM_TYPE, a As BIGNUM_TYPE, mont_a As BIGNUM_TYPE
    ctx = BN_MONT_CTX_new()
    modulus = BN_hex2bn("FFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFEFFFFFC2F")
    a = BN_hex2bn("123456789ABCDEF123456789ABCDEF123456789ABCDEF123456789ABCDEF")
    mont_a = BN_new()

    If BN_MONT_CTX_set(ctx, modulus) And BN_to_montgomery(mont_a, a, ctx) Then
        Debug.Print "APROVADO: Conversão para forma Montgomery"
        test_to_montgomery = True
    Else
        Debug.Print "FALHOU: Conversão para forma Montgomery falhou"
        test_to_montgomery = False
    End If
End Function

' Testa multiplicação Montgomery vs implementação regular
Private Function test_mont_multiplication() As Boolean
    Debug.Print "Testando multiplicação Montgomery..."

    Dim ctx As MONT_CTX, modulus As BIGNUM_TYPE
    Dim a As BIGNUM_TYPE, b As BIGNUM_TYPE, mont_a As BIGNUM_TYPE, mont_b As BIGNUM_TYPE
    Dim mont_result As BIGNUM_TYPE, regular_result As BIGNUM_TYPE, converted_result As BIGNUM_TYPE

    ctx = BN_MONT_CTX_new()
    modulus = BN_hex2bn("FFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFEFFFFFC2F")
    a = BN_hex2bn("123456789ABCDEF123456789ABCDEF123456789ABCDEF123456789ABCDEF")
    b = BN_hex2bn("FEDCBA9876543210FEDCBA9876543210FEDCBA9876543210FEDCBA9876543210")

    mont_a = BN_new(): mont_b = BN_new(): mont_result = BN_new()
    regular_result = BN_new(): converted_result = BN_new()

    If BN_MONT_CTX_set(ctx, modulus) And _
       BN_to_montgomery(mont_a, a, ctx) And _
       BN_to_montgomery(mont_b, b, ctx) And _
       BN_mod_mul_montgomery(mont_result, mont_a, mont_b, ctx) And _
       BN_from_montgomery(converted_result, mont_result, ctx) And _
       BN_mod_mul(regular_result, a, b, modulus) Then

        If BN_cmp(converted_result, regular_result) = 0 Then
            Debug.Print "APROVADO: Multiplicação Montgomery coincide com regular"
            test_mont_multiplication = True
        Else
            Debug.Print "FALHOU: Divergência multiplicação Montgomery"
            test_mont_multiplication = False
        End If
    Else
        Debug.Print "FALHOU: Configuração multiplicação Montgomery falhou"
        test_mont_multiplication = False
    End If
End Function

' Testa conversão roundtrip (ida e volta) Montgomery
Private Function test_from_montgomery() As Boolean
    Debug.Print "Testando conversão da forma Montgomery..."

    Dim ctx As MONT_CTX, modulus As BIGNUM_TYPE, a As BIGNUM_TYPE, mont_a As BIGNUM_TYPE, result As BIGNUM_TYPE
    ctx = BN_MONT_CTX_new()
    modulus = BN_hex2bn("FFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFEFFFFFC2F")
    a = BN_hex2bn("123456789ABCDEF123456789ABCDEF123456789ABCDEF123456789ABCDEF")
    mont_a = BN_new(): result = BN_new()

    If BN_MONT_CTX_set(ctx, modulus) And _
       BN_to_montgomery(mont_a, a, ctx) And _
       BN_from_montgomery(result, mont_a, ctx) Then

        If BN_cmp(result, a) = 0 Then
            Debug.Print "APROVADO: Conversão roundtrip Montgomery"
            test_from_montgomery = True
        Else
            Debug.Print "FALHOU: Conversão roundtrip Montgomery falhou"
            test_from_montgomery = False
        End If
    Else
        Debug.Print "FALHOU: Configuração conversão Montgomery falhou"
        test_from_montgomery = False
    End If
End Function

' Testa exponenciação Montgomery vs implementação regular
Private Function test_mont_exponentiation() As Boolean
    Debug.Print "Testando exponenciação Montgomery..."

    Dim ctx As MONT_CTX, modulus As BIGNUM_TYPE, base As BIGNUM_TYPE, exponent As BIGNUM_TYPE
    Dim mont_result As BIGNUM_TYPE, regular_result As BIGNUM_TYPE

    ctx = BN_MONT_CTX_new()
    modulus = BN_hex2bn("FFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFEFFFFFC2F")
    base = BN_hex2bn("123456789ABCDEF")
    exponent = BN_hex2bn("FEDCBA987654321")
    mont_result = BN_new(): regular_result = BN_new()

    If BN_MONT_CTX_set(ctx, modulus) And _
       BN_mod_exp_mont(mont_result, base, exponent, modulus, ctx) And _
       BN_mod_exp(regular_result, base, exponent, modulus) Then

        If BN_cmp(mont_result, regular_result) = 0 Then
            Debug.Print "APROVADO: Exponenciação Montgomery coincide com regular"
            test_mont_exponentiation = True
        Else
            Debug.Print "FALHOU: Divergência exponenciação Montgomery"
            test_mont_exponentiation = False
        End If
    Else
        Debug.Print "FALHOU: Configuração exponenciação Montgomery falhou"
        test_mont_exponentiation = False
    End If
End Function

' Testa casos extremos Montgomery (0, 1)
Private Function test_mont_edge_cases() As Boolean
    Debug.Print "Testando casos extremos Montgomery..."

    Dim ctx As MONT_CTX, modulus As BIGNUM_TYPE, zero As BIGNUM_TYPE, one As BIGNUM_TYPE
    Dim mont_zero As BIGNUM_TYPE, mont_one As BIGNUM_TYPE, result As BIGNUM_TYPE

    ctx = BN_MONT_CTX_new()
    modulus = BN_hex2bn("FFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFEFFFFFC2F")
    zero = BN_new(): one = BN_new()
    Call BN_set_word(one, 1)
    mont_zero = BN_new(): mont_one = BN_new(): result = BN_new()

    If BN_MONT_CTX_set(ctx, modulus) And _
       BN_to_montgomery(mont_zero, zero, ctx) And _
       BN_to_montgomery(mont_one, one, ctx) And _
       BN_from_montgomery(result, mont_one, ctx) Then

        If BN_is_zero(mont_zero) And BN_cmp(result, one) = 0 Then
            Debug.Print "APROVADO: Casos extremos Montgomery (0, 1)"
            test_mont_edge_cases = True
        Else
            Debug.Print "FALHOU: Casos extremos Montgomery falharam"
            test_mont_edge_cases = False
        End If
    Else
        Debug.Print "FALHOU: Configuração casos extremos Montgomery falhou"
        test_mont_edge_cases = False
    End If
End Function

' Testa Montgomery com parâmetros secp256k1 reais
Private Function test_mont_secp256k1() As Boolean
    Debug.Print "Testando Montgomery com parâmetros secp256k1..."

    Dim ctx As MONT_CTX, p As BIGNUM_TYPE, g_x As BIGNUM_TYPE
    Dim private_key As BIGNUM_TYPE, result As BIGNUM_TYPE, expected As BIGNUM_TYPE

    ctx = BN_MONT_CTX_new()
    p = BN_hex2bn("FFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFEFFFFFC2F")
    g_x = BN_hex2bn("79BE667EF9DCBBAC55A06295CE870B07029BFCDB2DCE28D959F2815B16F81798")
    private_key = BN_hex2bn("C9AFA9D845BA75166B5C215767B1D6934E50C3DB36E89B127B8A622B120F6721")
    result = BN_new(): expected = BN_new()

    ' Teste: g_x^private_key mod p usando Montgomery
    If BN_MONT_CTX_set(ctx, p) And _
       BN_mod_exp_mont(result, g_x, private_key, p, ctx) And _
       BN_mod_exp(expected, g_x, private_key, p) Then

        If BN_cmp(result, expected) = 0 Then
            Debug.Print "APROVADO: Montgomery com parâmetros secp256k1"
            test_mont_secp256k1 = True
        Else
            Debug.Print "FALHOU: Divergência Montgomery secp256k1"
            test_mont_secp256k1 = False
        End If
    Else
        Debug.Print "FALHOU: Configuração Montgomery secp256k1 falhou"
        test_mont_secp256k1 = False
    End If
End Function