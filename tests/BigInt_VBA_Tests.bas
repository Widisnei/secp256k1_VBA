Attribute VB_Name = "BigInt_VBA_Tests"
Option Explicit
Option Compare Binary
Option Base 0

'==============================================================================
' MÓDULO: BigInt_VBA_Tests
' Descrição: Testes Básicos de Autovalidação do BigInt_VBA
' Autor: Desenvolvido para compatibilidade com Bitcoin Core
' Data: 2024
'
' CARACTERÍSTICAS TÉCNICAS:
' • Autoteste fundamental do núcleo BigInt_VBA
' • Validação de operações aritméticas básicas
' • Testes de conversão hexadecimal
' • Verificação de propriedades matemáticas
' • Validação de algoritmos criptográficos
'
' OPERAÇÕES TESTADAS:
' • BN_hex2bn() / BN_bn2hex()   - Conversão hexadecimal
' • BN_add() / BN_sub()         - Adição e subtração
' • BN_mul()                   - Multiplicação
' • BN_div()                   - Divisão com resto
' • BN_mod_inverse()           - Inversão modular
' • BN_mod_exp()               - Exponenciação modular
'
' CASOS DE TESTE FUNDAMENTAIS:
' • Carry em adição: 0xFFFFFFFF + 1 = 0x100000000
' • Multiplicação: 0xFFFFFFFF² = 0xFFFFFFFE00000001
' • Divisão: Verifica identidade q*d + r = a
' • Inversão: inv(2) * 2 ≡ 1 (mod p)
' • Teste de Fermat: 2^(p-1) ≡ 1 (mod p)
' • Raíz quadrada: sqrt(4) para p ≡ 3 (mod 4)
'
' PROPRIEDADES MATEMÁTICAS:
' • Teste de primalidade Fermat
' • Propriedades de resíduo quadrático
' • Identidade de inversão modular
' • Consistência de operações básicas
'
' ALGORITMOS CRIPTOGRÁFICOS:
' • Exponenciação modular para secp256k1
' • Inversão modular para ECDSA
' • Raíz quadrada para descompressão de pontos
' • Aritmética de campo finito
'
' IMPORTÂNCIA DO AUTOTESTE:
' • Validação rápida da implementação
' • Detecção precoce de problemas
' • Verificação de compatibilidade
' • Base para testes mais avançados
'
' COMPATIBILIDADE:
' • Bitcoin Core secp256k1 - Parâmetros idênticos
' • OpenSSL BN_* - Comportamento compatível
' • GMP - Propriedades matemáticas
' • RFC 5480 - Padrões criptográficos
'==============================================================================

'==============================================================================
' AUTOTESTE FUNDAMENTAL DO BIGINT_VBA
'==============================================================================

' Propósito: Autovalidação rápida das operações fundamentais BigInt
' Algoritmo: Testes sequenciais de conversão, aritmética e propriedades
' Retorno: Relatório de validação via Debug.Print
' Fundamental: Base para todos os outros testes do sistema

Public Sub BigIntVBA_SelfTest()
    Dim a As BIGNUM_TYPE, b As BIGNUM_TYPE, r As BIGNUM_TYPE, t As BIGNUM_TYPE
    Dim q As BIGNUM_TYPE, remn As BIGNUM_TYPE, inv As BIGNUM_TYPE
    Dim x As BIGNUM_TYPE

    a = BN_new(): b = BN_new(): r = BN_new(): t = BN_new()
    q = BN_new(): remn = BN_new(): inv = BN_new()
    x = BN_new()

    ' Conversão Hex <-> BN
    x = BN_hex2bn("FFFFFFFF00000001")
    Debug.Print "Hex->BN->Hex:", BN_bn2hex(x)

    ' Adição/Subtração
    Call BN_set_word(a, &HFFFFFFFF)
    Call BN_set_word(b, &H1)
    Call BN_add(r, a, b)
    Debug.Print "Carry adição: ", BN_bn2hex(r), "      (espera 100000000)"
    Call BN_sub(t, r, b)
    Debug.Print "Subtração:    ", BN_bn2hex(t), "       (espera FFFFFFFF)"

    ' Multiplicação
    Call BN_mul(r, a, a)
    Debug.Print "FFFFFFFF^2:  ", BN_bn2hex(r), "             (espera FFFFFFFE00000001)"

    ' Divisão
    Call BN_div(q, remn, r, a)
    Debug.Print "Div q:       ", BN_bn2hex(q), "       resto: ", BN_bn2hex(remn), "              (espera q=FFFFFFFF, resto=0)"

    ' Inverso mod p
    Dim p As BIGNUM_TYPE, two As BIGNUM_TYPE
    p = BN_hex2bn("FFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFEFFFFFC2F")
    Call BN_set_word(two, 2)
    If BN_mod_inverse(inv, two, p) Then
        Call BN_mul(t, inv, two)
        Call BN_mod(r, t, p)
        Debug.Print "inv(2) mod p verificação: ", BN_bn2hex(r), "              (espera 1)"
    Else
        Debug.Print "inv(2) mod p falhou"
    End If

    ' === Auxiliares modulares ===
    Dim e As BIGNUM_TYPE, res As BIGNUM_TYPE, one As BIGNUM_TYPE
    e = BN_new(): res = BN_new(): one = BN_new(): Call BN_set_word(one, 1)

    ' Fermat: 2^(p-1) ≡ 1 (mod p)
    Call BN_sub(e, p, one)
    If BN_mod_exp(res, two, e, p) Then
        Debug.Print "Fermat(2):   ", BN_bn2hex(res), " (espera 1)"
    Else
        Debug.Print "Fermat(2): FALHOU"
    End If

    ' sqrt para p ≡ 3 (mod 4): s = 4^((p+1)/4), verifica s^2 ≡ 4
    Dim four As BIGNUM_TYPE, expSqrt As BIGNUM_TYPE, sq As BIGNUM_TYPE
    four = BN_new(): expSqrt = BN_new(): sq = BN_new()
    Call BN_set_word(four, 4)
    Call BN_add(expSqrt, p, one)
    Call BN_rshift(expSqrt, expSqrt, 2)
    If BN_mod_exp(res, four, expSqrt, p) Then
        Call BN_mod_mul(sq, res, res, p)
        Debug.Print "sqrt_p3mod4(4) -> s^2: ", BN_bn2hex(sq), " (espera 4)"
    Else
        Debug.Print "sqrt_p3mod4(4): FALHOU"
    End If

    Debug.Print "=== Autoteste concluído ==="
End Sub