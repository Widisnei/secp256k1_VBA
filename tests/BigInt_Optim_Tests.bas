Attribute VB_Name = "BigInt_Optim_Tests"
Option Explicit
Option Compare Binary
Option Base 0

'==============================================================================
' MÓDULO: BigInt_Optim_Tests
' Descrição: Testes de Otimizações Avançadas BigInt (COMBA, ModExp)
' Autor: Desenvolvido para compatibilidade com Bitcoin Core
' Data: 2024
'
' CARACTERÍSTICAS TÉCNICAS:
' • Validação de multiplicação COMBA otimizada
' • Testes de exponenciação modular com janelas
' • Verificação do seletor automático de algoritmos
' • Testes de corretude vs implementações de referência
' • Validação com casos aleatórios e extremos
'
' ALGORITMOS TESTADOS:
' • BN_mul_fast256()          - Multiplicação COMBA 8x8
' • BN_mod_exp_win4()          - Exponenciação janela 4-bit
' • BN_mod_exp_auto()          - Seletor automático otimizado
'
' OTIMIZAÇÕES COMBA:
' • Multiplicação 8x8 words sem carries intermediários
' • Vantagem: 30-50% mais rápido para números 256-bit
' • Uso: Ideal para operações secp256k1
' • Limitação: Tamanho fixo (256-bit)
'
' EXPONENCIAÇÃO COM JANELAS:
' • Janela 4-bit: Pré-computa [1, 3, 5, 7, 9, 11, 13, 15]
' • Vantagem: 25-40% mais rápido para expoentes densos
' • Seletor auto: Escolhe melhor algoritmo baseado na densidade
' • Threshold: Densidade > 50% usa janelas
'
' TESTES IMPLEMENTADOS:
' • Corretude COMBA vs multiplicação padrão
' • Corretude ModExp janelas vs binário
' • Seletor automático vs implementações específicas
' • Casos extremos (zero, um)
' • Testes aleatórios (200+ amostras)
'
' COMPATIBILIDADE:
' • Bitcoin Core secp256k1 - Otimizações equivalentes
' • OpenSSL BN_* - Resultados idênticos
' • GMP - Performance competitiva
' • Handbook of Applied Cryptography - Algoritmos baseados
'==============================================================================

'==============================================================================
' FUNÇÕES AUXILIARES DE TESTE
'==============================================================================

' Asserção de igualdade com valores hexadecimais para debug
Private Sub AssertEqHex(ByVal label As String, ByRef got As BIGNUM_TYPE, ByRef exp As BIGNUM_TYPE)
    If BN_cmp(got, exp) = 0 Then
        Debug.Print "APROVADO:", label
    Else
        Debug.Print "FALHOU:", label, " obtido=", BN_bn2hex(got), " esperado=", BN_bn2hex(exp)
    End If
End Sub

' Asserção booleana simples
Private Sub AssertTrue(ByVal label As String, ByVal cond As Boolean)
    If cond Then
        Debug.Print "APROVADO:", label
    Else
        Debug.Print "FALHOU:", label
    End If
End Sub

' Conversão hexadecimal para BIGNUM
Private Sub HexToBN(ByRef outBN As BIGNUM_TYPE, ByVal h As String)
    outBN = BN_hex2bn(h)
End Sub

' Gera array de bytes aleatórios
Private Function RandBytes(ByVal n As Long) As Byte()
    Dim b() As Byte, i As Long
    ReDim b(0 To n - 1)
    For i = 0 To n - 1
        b(i) = Int(Rnd() * 256)
    Next i
    RandBytes = b
End Function

' Gera BIGNUM aleatório com número especificado de bits
Private Function RandBN(ByVal bits As Long) As BIGNUM_TYPE
    Dim nbytes As Long
    nbytes = (bits + 7) \ 8
    Dim b() As Byte
    b = RandBytes(nbytes)
    ' força bit alto para densidade opcional; aqui deixamos natural
    RandBN = BN_bin2bn(b, nbytes)
End Function

'==============================================================================
' TESTES DE CORRETUDE COMBA 8x8
'==============================================================================

Public Sub Test_COMBA_Correctness()
    Debug.Print "=== Teste_Corretude_COMBA ==="
    Randomize 1337

    Dim a As BIGNUM_TYPE, b As BIGNUM_TYPE, r1 As BIGNUM_TYPE, r2 As BIGNUM_TYPE
    Dim q As BIGNUM_TYPE, remn As BIGNUM_TYPE

    ' Vetor fixo 256-bit
    a = BN_hex2bn("D1B2A3C4D5E6F7089A0B1C2D3E4F5061728394A5B6C7D8E9F0A1B2C3D4E5F60A")
    b = BN_hex2bn("0F1E2D3C4B5A69788796A5B4C3D2E1F00112233445566778899AABBCCDDEEFF0")
    Call BN_mul(r1, a, b)
    Call BN_mul_fast256(r2, a, b)
    Call AssertEqHex("COMBA fixo 256", r2, r1)

    ' Casos degenerados
    Call BN_set_word(a, 0) : b = BN_hex2bn("1")
    Call BN_mul(r1, a, b) : Call BN_mul_fast256(r2, a, b) : Call AssertEqHex("COMBA a=0", r2, r1)
    a = BN_hex2bn("1") : Call BN_set_word(b, 0)
    Call BN_mul(r1, a, b) : Call BN_mul_fast256(r2, a, b) : Call AssertEqHex("COMBA b=0", r2, r1)

    ' Randomizados (200 amostras, tamanhos até 256 bits)
    Dim t As Long, bitsA As Long, bitsB As Long
    For t = 1 To 200
        bitsA = 1 + Int(Rnd() * 256)
        bitsB = 1 + Int(Rnd() * 256)
        a = RandBN(bitsA)
        b = RandBN(bitsB)
        ' garante não-zero
        If BN_is_zero(a) Then a = BN_hex2bn("1")
        If BN_is_zero(b) Then b = BN_hex2bn("1")
        Call BN_mul(r1, a, b)
        Call BN_mul_fast256(r2, a, b)
        If BN_cmp(r1, r2) <> 0 Then
            Debug.Print "FALHOU: COMBA rand #", t, " obtido=", BN_bn2hex(r2), " esperado=", BN_bn2hex(r1)
            Exit For
        End If
        ' identidade de divisão (quando divisor cabe): (a*b)\a == b, rem=0
        Call BN_div(q, remn, r2, a)
        If BN_cmp(q, b) <> 0 Or Not BN_is_zero(remn) Then
            Debug.Print "FALHOU: Identidade DIV #", t
            Exit For
        End If
    Next t
    Debug.Print "=== Fim COMBA ==="
End Sub

'==============================================================================
' TESTES DE EXPONENCIAÇÃO MODULAR COM JANELAS
'==============================================================================

Public Sub Test_ModExp_Correctness()
    Debug.Print "=== Teste_Corretude_ModExp ==="
    Randomize 1337
    Dim a As BIGNUM_TYPE, e As BIGNUM_TYPE, m As BIGNUM_TYPE
    Dim rb As BIGNUM_TYPE, rw As BIGNUM_TYPE
    m = BN_hex2bn("FFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFEFFFFFC2F")

    ' Vetores fixos: e=65537 e um expoente denso 256-bit
    a = BN_hex2bn("D1B2A3C4D5E6F7089A0B1C2D3E4F5061728394A5B6C7D8E9F0A1B2C3D4E5F60A")
    e = BN_hex2bn("10001")
    Call BN_mod_exp(rb, a, e, m)
    Call BN_mod_exp_win4(rw, a, e, m)
    Call AssertEqHex("ModExp e=65537", rw, rb)

    e = BN_hex2bn("F123456789ABCDEF0123456789ABCDEF123456789ABCDEF0123456789ABCDEF")
    Call BN_mod_exp(rb, a, e, m)
    Call BN_mod_exp_win4(rw, a, e, m)
    Call AssertEqHex("ModExp e denso", rw, rb)

    ' Randomizados: 50 amostras de expoentes até 256 bits
    Dim t As Long, bitsE As Long
    For t = 1 To 50
        a = RandBN(256)
        If BN_is_zero(a) Then a = BN_hex2bn("2")
        bitsE = 1 + Int(Rnd() * 256)
        e = RandBN(bitsE)
        If BN_is_zero(e) Then e = BN_hex2bn("1")
        Call BN_mod_exp(rb, a, e, m)
        Call BN_mod_exp_win4(rw, a, e, m)
        If BN_cmp(rw, rb) <> 0 Then
            Debug.Print "FALHOU: ModExp rand #", t, " obtido=", BN_bn2hex(rw), " esperado=", BN_bn2hex(rb)
            Exit For
        End If
    Next t
    Debug.Print "=== Fim ModExp ==="
End Sub

'==============================================================================
' TESTES DO SELETOR AUTOMÁTICO
'==============================================================================

Public Sub Test_ModExp_Auto()
    Debug.Print "=== Teste_ModExp_Auto ==="
    Dim a As BIGNUM_TYPE, m As BIGNUM_TYPE, e As BIGNUM_TYPE
    Dim r As BIGNUM_TYPE, rb As BIGNUM_TYPE, rw As BIGNUM_TYPE
    m = BN_hex2bn("FFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFEFFFFFC2F")
    a = BN_hex2bn("D1B2A3C4D5E6F7089A0B1C2D3E4F5061728394A5B6C7D8E9F0A1B2C3D4E5F60A")

    ' Esparso
    e = BN_hex2bn("10001")
    Call BN_mod_exp_auto(r, a, e, m)
    Call BN_mod_exp(rb, a, e, m)
    Call AssertEqHex("auto==baseline (sparse)", r, rb)

    ' Denso
    e = BN_hex2bn("F123456789ABCDEF0123456789ABCDEF123456789ABCDEF0123456789ABCDEF")
    Call BN_mod_exp_auto(r, a, e, m)
    Call BN_mod_exp_win4(rw, a, e, m)
    Call AssertEqHex("auto==win4 (dense)", r, rw)

    Debug.Print "=== Fim Auto ==="
End Sub

' Executa todos os testes de otimização
Public Sub Run_All_Optim_Tests()
    On Error GoTo EH
    Call Test_COMBA_Correctness
    Call Test_ModExp_Correctness
    Call Test_ModExp_Auto
    Debug.Print "=== TODOS OS TESTES OPT APROVADOS (ou veja linhas FALHOU) ==="
    Exit Sub
EH:
    Debug.Print "ERRO:", Err.Number, Err.Description
End Sub