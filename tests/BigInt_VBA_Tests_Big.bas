Attribute VB_Name = "BigInt_VBA_Tests_Big"
Option Explicit
Option Compare Binary
Option Base 0

'==============================================================================
' MÓDULO: BigInt_VBA_Tests_Big
' Descrição: Testes Robustos com Números Grandes (256-bit e 512-bit)
' Autor: Desenvolvido para compatibilidade com Bitcoin Core
' Data: 2024
'
' CARACTERÍSTICAS TÉCNICAS:
' • Testes robustos com números de 256-bit e 512-bit
' • Validação de operações aritméticas complexas
' • Testes de precisão com valores pré-calculados
' • Verificação de consistência em operações modulares
' • Validação de algoritmos com números grandes
'
' SUITE 256-BIT:
' • Módulo: secp256k1 field prime (2²⁵⁶ - 2³² - 977)
' • Operandos: Números aleatórios de 256-bit
' • Operações: Add, Sub, Mul, Div, ModAdd, ModSub, ModMul, ModExp, ModInv
' • Expoente: 65537 (0x10001)
' • Validação: Resultados pré-calculados com precisão
'
' SUITE 512-BIT:
' • Módulo: Primo de 512-bit (0xFFFF...FFDC7)
' • Operandos: Padrões repetitivos (0xA3A3... e 0x5C5C...)
' • Operações: Conjunto completo de aritmética modular
' • Expoente: 0xDEADBEEFCAFEBABE1234567890
' • Complexidade: Testa limites de precisão do sistema
'
' ALGORITMOS TESTADOS:
' • BN_add() / BN_usub()        - Aritmética básica com carries
' • BN_mul() / BN_div()         - Multiplicação e divisão precisas
' • BN_mod_add/sub/mul()        - Aritmética modular otimizada
' • BN_mod_exp()               - Exponenciação modular avançada
' • BN_mod_inverse()           - Inversão modular com Euclides estendido
'
' FUNÇÕES AUXILIARES:
' • h()                        - Concatenação segura de strings longas
' • ExpectEqHex()              - Comparação hexadecimal com relatório
' • ExpectEqBN()               - Comparação BIGNUM com relatório
' • ExpectTrue()               - Asserção booleana com relatório
'
' VALIDAÇÃO DE PRECISÃO:
' • Resultados pré-calculados com ferramentas externas
' • Verificação de identidades matemáticas
' • Testes de consistência entre operações
' • Validação de propriedades de campo finito
'
' IMPORTÂNCIA DOS TESTES:
' • Valida precisão em números grandes
' • Detecta erros de truncamento
' • Verifica estabilidade numérica
' • Garante compatibilidade com Bitcoin
'
' COMPATIBILIDADE:
' • Bitcoin Core secp256k1 - Parâmetros e resultados idênticos
' • OpenSSL BN_* - Precisão equivalente
' • GMP - Resultados matemáticos compatíveis
' • Sage Math - Validação independente
'==============================================================================

'==============================================================================
' FUNÇÕES AUXILIARES DE TESTE
'==============================================================================

' Auxiliar para concatenar strings longas sem continuações de linha
Private Function h(ParamArray s()) As String
    h = Join(s, "")
End Function

' Compara BIGNUM com valor hexadecimal esperado
Private Sub ExpectEqHex(ByVal label As String, ByRef got As BIGNUM_TYPE, ByVal expectHex As String)
    Dim gotHex As String : gotHex = BN_bn2hex(got)
    If UCase$(gotHex) = UCase$(expectHex) Then
        Debug.Print "APROVADO:", label, "=", gotHex
    Else
        Debug.Print "FALHOU:", label, "obtido=", gotHex, " esperado=", expectHex
    End If
End Sub

' Compara dois BIGNUMs
Private Sub ExpectEqBN(ByVal label As String, ByRef a As BIGNUM_TYPE, ByRef b As BIGNUM_TYPE)
    If BN_cmp(a, b) = 0 Then
        Debug.Print "APROVADO:", label
    Else
        Debug.Print "FALHOU:", label, " obtido=", BN_bn2hex(a), " esperado=", BN_bn2hex(b)
    End If
End Sub

' Asserção booleana
Private Sub ExpectTrue(ByVal label As String, ByVal cond As Boolean)
    If cond Then
        Debug.Print "APROVADO:", label
    Else
        Debug.Print "FALHOU:", label
    End If
End Sub

'==============================================================================
' TESTES ROBUSTOS 256-BIT E 512-BIT
'==============================================================================

' Propósito: Valida precisão e robustez com números grandes
' Algoritmo: Suítes 256-bit (secp256k1) e 512-bit com resultados pré-calculados
' Retorno: Relatório detalhado via Debug.Print
' Precisão: Validação com valores de referência externos

Public Sub BigIntVBA_RobustTests256_512()
    Dim a As BIGNUM_TYPE, b As BIGNUM_TYPE, r As BIGNUM_TYPE, t As BIGNUM_TYPE, q As BIGNUM_TYPE, remn As BIGNUM_TYPE
    Dim p As BIGNUM_TYPE, m As BIGNUM_TYPE, e As BIGNUM_TYPE, inv As BIGNUM_TYPE, one As BIGNUM_TYPE

    a = BN_new() : b = BN_new() : r = BN_new() : t = BN_new() : q = BN_new() : remn = BN_new()
    p = BN_new() : m = BN_new() : e = BN_new() : inv = BN_new() : one = BN_new() : Call BN_set_word(one, 1)

    'Suíte 256-bit
    p = BN_hex2bn(h("FFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFEFFFFFC2F"))
    a = BN_hex2bn(h("D1B2A3C4D5E6F7089A0B1C2D3E4F5061728394A5B6C7D8E9F0A1B2C3D4E5F60A"))
    b = BN_hex2bn(h("0F1E2D3C4B5A69788796A5B4C3D2E1F00112233445566778899AABBCCDDEEFF0"))

    Call BN_add(r, a, b)
    Call ExpectEqHex("256:add", r, h("E0D0D1012141608121A1C1E2022232517395B7D9FC1E40627A3C5E80A2C4E5FA"))

    Call BN_usub(r, a, b)
    Call ExpectEqHex("256:usub", r, h("C29476888A8C8D90127476787A7C6E717171717171717171670707070707061A"))

    Call BN_mul(t, a, b)
    Call BN_div(q, remn, t, a)
    ' Compara BN vs BN para evitar incompatibilidade de zeros à esquerda
    Call ExpectEqBN("256:div.q", q, b)
    Call ExpectTrue("256:div.rem==0", BN_is_zero(remn))

    Call BN_mod_add(r, a, b, p)
    Call ExpectEqHex("256:modadd", r, h("E0D0D1012141608121A1C1E2022232517395B7D9FC1E40627A3C5E80A2C4E5FA"))
    Call BN_mod_sub(r, a, b, p)
    Call ExpectEqHex("256:modsub", r, h("C29476888A8C8D90127476787A7C6E717171717171717171670707070707061A"))
    Call BN_mod_mul(r, a, b, p)
    Call ExpectEqHex("256:modmul", r, h("43578667F2AC0BD75F7AA1E473DF5173E99BB4BCA8451B15B2ABA920C9099FDE"))

    e = BN_hex2bn("10001")
    Call BN_mod_exp(r, a, e, p)
    Call ExpectEqHex("256:modexp", r, h("9EA85856814F9E2B29A162A8D222A50146758B9AC134DC1D59A544639E9B000E"))

    If BN_mod_inverse(inv, a, p) Then
        Call BN_mod_mul(r, inv, a, p)
        Call ExpectEqHex("256:inv* a == 1", r, "1")
        Call ExpectEqHex("256:inv exact", inv, h("AFF55B8056E94D9437B793E0531188BDD80799BBB4E8BC2E4974120AAABA4413"))
    Else
        Debug.Print "FALHOU: 256:mod_inverse não existiu"
    End If

    'Suíte 512-bit
    m = BN_hex2bn(h("FFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFDC7"))
    a = BN_hex2bn(h("A3A3A3A3A3A3A3A3A3A3A3A3A3A3A3A3A3A3A3A3A3A3A3A3A3A3A3A3A3A3A3A3A3A3A3A3A3A3A3A3A3A3A3A3A3A3A3A3A3A3A3A3A3A3A3A3A3A3A3A3A3A3A3A3"))
    b = BN_hex2bn(h("5C5C5C5C5C5C5C5C5C5C5C5C5C5C5C5C5C5C5C5C5C5C5C5C5C5C5C5C5C5C5C5C5C5C5C5C5C5C5C5C5C5C5C5C5C5C5C5C5C5C5C5C5C5C5C5C5C5C5C5C5C5C5C5C"))

    Call BN_add(r, a, b)
    Call ExpectEqHex("512:add", r, h("FFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFF"))
    Call BN_usub(r, a, b)
    Call ExpectEqHex("512:usub", r, h("47474747474747474747474747474747474747474747474747474747474747474747474747474747474747474747474747474747474747474747474747474747"))

    Call BN_mul(t, a, b)
    Call BN_div(q, remn, t, b)
    Call ExpectEqBN("512:div.q", q, a)
    Call ExpectTrue("512:div.rem==0", BN_is_zero(remn))

    Call BN_mod_add(r, a, b, m)
    Call ExpectEqHex("512:modadd", r, h("238"))
    Call BN_mod_sub(r, a, b, m)
    Call ExpectEqHex("512:modsub", r, h("47474747474747474747474747474747474747474747474747474747474747474747474747474747474747474747474747474747474747474747474747474747"))
    Call BN_mod_mul(r, a, b, m)
    ' Esperado computado previamente
    Call ExpectEqHex("512:modmul", r, h("B18C67421CF7D2AD88633E18F3CEA9845F3A14EFCAA5805B3610EBC6A17C5732", "0CE7C29D78532E08E3BE99744F2A04DFBA95704B2600DBB6916C4721FCD8D3A6"))

    e = BN_hex2bn(h("DEADBEEFCAFEBABE1234567890"))
    Call BN_mod_exp(r, a, e, m)
    Call ExpectEqHex("512:modexp", r, h("8AEDCBAE2DE53BC49ED34A8C9538F47FA65868D03EC068F3E622C54FE903D185", "2DBA97EE537F100E02F1B0F61BB6AF7AB7E7084F300912E206126A138C2DF9E2"))

    Dim a_mod As BIGNUM_TYPE
    a_mod = BN_new()
    Call BN_mod(a_mod, a, m)
    If BN_mod_inverse(inv, a_mod, m) Then
        Call BN_mod_mul(r, inv, a_mod, m)
        Call ExpectEqHex("512:inv * a_mod == 1", r, "1")
        Call ExpectEqHex("512:inv exact", inv, h("25452BB019D6313D79056011B244CB6B7064DE8821179A3EF46E5639BC007141", "B849160285903406FD74C61C5BC1718002D4D7CEA08CDCF067B35FEFB8271B92"))
    Else
        Debug.Print "FALHOU: 512:mod_inverse não existiu"
    End If

    Debug.Print "=== Testes robustos 256/512 concluídos ==="
End Sub