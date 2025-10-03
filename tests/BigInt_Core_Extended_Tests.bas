Attribute VB_Name = "BigInt_Core_Extended_Tests"
Option Explicit
Option Compare Binary
Option Base 0

'==============================================================================
' MÓDULO: BigInt_Core_Extended_Tests
' Descrição: Testes Avançados de Robustez para BigInt_VBA
' Autor: Desenvolvido para compatibilidade com Bitcoin Core
' Data: 2024
'
' CARACTERÍSTICAS TÉCNICAS:
' • Testes de segurança alias (operações in-place)
' • Validação de casos extremos e limites
' • Testes de inversão modular com gcd > 1
' • Divisão e módulo com números negativos
' • Testes de roundtrip com dados aleatórios
'
' ALGORITMOS TESTADOS:
' • Alias_Add_Sub()           - Segurança de adição/subtração in-place
' • Alias_Mul_Mod()           - Segurança de multiplicação/módulo in-place
' • Inverse_NonExistence()    - Detecção de inversos inexistentes
' • Div_Mod_Signed()          - Divisão com números negativos
' • Roundtrip_Random()        - Conversão binária ida e volta
'
' CASOS DE TESTE CRÍTICOS:
' • Operações com alias (r = r op b)
' • Inversão quando gcd(a,m) > 1
' • Divisão com dividendo negativo
' • Conversão binária com 64 casos aleatórios
' • Verificação de propriedades matemáticas
'
' FUNÇÕES AUXILIARES:
' • Pass()                    - Registra teste bem-sucedido
' • FailHex()                 - Registra falha com valores hex
' • AssertEq()                - Asserção de igualdade
' • BN_FromHex()              - Conversão hex simplificada
' • SafeByteArrayLen()        - Tamanho seguro de array
'
' ROBUSTEZ E SEGURANÇA:
' • Detecção de casos extremos
' • Validação de propriedades matemáticas
' • Testes com dados aleatórios
' • Verificação de integridade de dados
'
' COMPATIBILIDADE:
' • OpenSSL BN_* - Comportamento idêntico
' • GMP - Propriedades matemáticas
' • Bitcoin Core - Casos de uso reais
'==============================================================================

'==============================================================================
' FUNÇÕES AUXILIARES DE TESTE
'==============================================================================

' Registra teste bem-sucedido
Private Sub Pass(ByVal label As String): Debug.Print "APROVADO:", label: End Sub

' Registra falha com valores hexadecimais para debug
Private Sub FailHex(ByVal label As String, ByRef got As BIGNUM_TYPE, ByRef exp As BIGNUM_TYPE)
    Debug.Print "FALHOU:", label, " obtido=", BN_bn2hex(got), " esperado=", BN_bn2hex(exp)
End Sub

' Asserção de igualdade com relatório automático
Private Sub AssertEq(ByVal label As String, ByRef a As BIGNUM_TYPE, ByRef b As BIGNUM_TYPE)
    If BN_cmp(a, b) = 0 Then Pass label Else FailHex label, a, b
End Sub

' Conversão hexadecimal simplificada
Private Function BN_FromHex(ByVal h As String) As BIGNUM_TYPE
    BN_FromHex = BN_hex2bn(h)
End Function

' Obtém tamanho seguro de array de bytes (trata arrays não inicializados)
Private Function SafeByteArrayLen(ByRef b() As Byte) As Long
    On Error GoTo Uninit
    SafeByteArrayLen = UBound(b) - LBound(b) + 1
    Exit Function
Uninit:
    SafeByteArrayLen = 0
End Function

'==============================================================================
' TESTES DE SEGURANÇA ALIAS
'==============================================================================

' Testa segurança de operações in-place (alias) para adição e subtração
Private Sub Alias_Add_Sub()
    Dim a As BIGNUM_TYPE, b As BIGNUM_TYPE, r As BIGNUM_TYPE, expect As BIGNUM_TYPE
    a = BN_FromHex("FFFFFFFF00000001")
    b = BN_FromHex("89ABCDEF01234567")
    r = a: Call BN_add(r, r, b): Call BN_add(expect, a, b): Call AssertEq("alias add r=a", r, expect)
    r = b: Call BN_add(r, a, r): Call AssertEq("alias add r=b", r, expect)
    r = a: Call BN_sub(r, r, b): Call BN_sub(expect, a, b): Call AssertEq("alias sub r=a", r, expect)
End Sub

' Testa segurança de operações in-place (alias) para multiplicação e módulo
Private Sub Alias_Mul_Mod()
    Dim a As BIGNUM_TYPE, b As BIGNUM_TYPE, r As BIGNUM_TYPE, m As BIGNUM_TYPE, expect As BIGNUM_TYPE
    a = BN_FromHex("D1B2A3C4D5E6F7089A0B1C2D3E4F5061728394A5B6C7D8E9F0A1B2C3D4E5F60A")
    b = BN_FromHex("0F1E2D3C4B5A69788796A5B4C3D2E1F00112233445566778899AABBCCDDEEFF0")
    m = BN_FromHex("FFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFEFFFFFC2F")
    r = a: Call BN_mul(r, r, b): Call BN_mul(expect, a, b): Call AssertEq("alias mul r=a", r, expect)
    r = a: Call BN_mod(r, r, m): Call BN_mod(expect, a, m): Call AssertEq("alias mod r=a", r, expect)
End Sub

' Testa detecção de casos onde o inverso modular não existe (gcd > 1)
Private Sub Inverse_NonExistence()
    Dim m As BIGNUM_TYPE, a As BIGNUM_TYPE, inv As BIGNUM_TYPE
    ' Composto ímpar garantido divisível por 3 (todos Fs -> soma dígitos hex ≡ 0 mod 3), então gcd(3,m)≠1
    m = BN_FromHex("FFFFFFFFFFFFFFFFFFFFFFFF")
    a = BN_FromHex("3")
    If BN_mod_inverse(inv, a, m) Then
        Debug.Print "FALHOU: inverso não deveria existir (gcd>1)"
    Else
        Pass "inexistência de inverso (gcd>1)"
    End If
End Sub

' Testa divisão e módulo com números negativos
Private Sub Div_Mod_Signed()
    Dim a As BIGNUM_TYPE, d As BIGNUM_TYPE, q As BIGNUM_TYPE, r As BIGNUM_TYPE, check As BIGNUM_TYPE
    a = BN_FromHex("-123456789ABCDEF0123456789ABCDEF")
    d = BN_FromHex("FEDCBA987654321")
    Call BN_div(q, r, a, d)
    Call BN_mul(check, q, d)
    Call BN_add(check, check, r)
    Call AssertEq("signed div identity", check, a)
    If r.neg Then Debug.Print "FALHOU: resto negativo" Else Pass "resto não-negativo"
End Sub

' Testa conversão binária ida e volta com dados aleatórios
Private Sub Roundtrip_Random()
    Dim i As Long, bits As Long, x As BIGNUM_TYPE, y As BIGNUM_TYPE
    Dim byts() As Byte, n As Long
    Randomize 20250917
    For i = 1 To 64
        bits = 1 + Int(Rnd() * 512)
        n = (bits + 7) \ 8
        ReDim byts(0 To n - 1)
        Dim j As Long
        For j = 0 To n - 1
            byts(j) = Int(Rnd() * 256)
        Next j
        x = BN_bin2bn(byts, n)
        byts = BN_bn2bin(x)
        n = SafeByteArrayLen(byts)
        y = BN_bin2bn(byts, n)
        Call AssertEq("roundtrip rand #" & i, x, y)
        If (i Mod 8) = 0 Then Debug.Print "progresso roundtrip:", i: DoEvents
    Next i
End Sub

'==============================================================================
' EXECUÇÃO DOS TESTES ESTENDIDOS
'==============================================================================

' Propósito: Executa testes avançados de robustez do BigInt_VBA
' Algoritmo: Execução sequencial de 5 categorias de testes críticos
' Retorno: Relatório detalhado via Debug.Print
' Segurança: Valida comportamento em casos extremos e limites

Public Sub Run_Core_Extended_Tests()
    On Error GoTo EH
    Debug.Print "=== Testes Estendidos do Núcleo ==="
    Call Alias_Add_Sub
    Call Alias_Mul_Mod
    Call Inverse_NonExistence
    Call Div_Mod_Signed
    Call Roundtrip_Random
    Debug.Print "=== Fim dos Testes do Núcleo ==="
    Exit Sub
EH:
    Debug.Print "ERRO(Núcleo):", Err.Number, Err.Description
End Sub