Attribute VB_Name = "Test_ECDSA_Direct_Fix"
Option Explicit

'==============================================================================
' TESTE ECDSA DIRETO SEM CONVERSÃO DER
'==============================================================================
'
' PROPÓSITO:
' • Teste de assinatura e verificação ECDSA sem conversão DER
' • Validação direta usando estruturas ECDSA_SIGNATURE nativas
' • Verificação de integridade com hash correto e incorreto
' • Debug de implementação Bitcoin Core sem overhead DER
'
' CARACTERÍSTICAS TÉCNICAS:
' • Chave privada: C9AFA9D845BA75166B5C215767B1D6934E50C3DB36E89B127B8A622B120F6721
' • Mensagem: "Hello, secp256k1!"
' • Algoritmo: ecdsa_sign_bitcoin_core() e ecdsa_verify_bitcoin_core()
' • Estrutura: ECDSA_SIGNATURE com r e s como BIGNUM_TYPE
' • Validação: Teste positivo e negativo (hash errado)
'
' ALGORITMOS IMPLEMENTADOS:
' • test_ecdsa_direct_fix() - Teste ECDSA direto completo
' • test_ecdsa_low_s_adjustment() - Vetor determinístico com ajuste low-s
'
' VANTAGENS:
' • Performance superior (sem overhead DER)
' • Debug direto da implementação core
' • Validação de integridade matemática
' • Teste de robustez com casos negativos
'
' COMPATIBILIDADE:
' • Bitcoin Core secp256k1 - Algoritmo idêntico
' • RFC 6979 - Geração determinística de k
' • FIPS 186-4 - Padrão ECDSA
' • OpenSSL - Comportamento compatível
'==============================================================================

'==============================================================================
' TESTE ECDSA DIRETO
'==============================================================================

' Propósito: Teste de assinatura e verificação ECDSA sem conversão DER
' Algoritmo: Assina com ecdsa_sign_bitcoin_core, verifica com ecdsa_verify_bitcoin_core
' Retorno: Relatório completo via Debug.Print com validação positiva e negativa

Public Sub test_ecdsa_direct_fix()
    Debug.Print "=== TESTE ECDSA DIRETO (SEM DER) ==="
    
    Dim ctx As SECP256K1_CTX
    ctx = secp256k1_context_create()
    
    Dim private_key As String, message As String, hash As String
    private_key = "C9AFA9D845BA75166B5C215767B1D6934E50C3DB36E89B127B8A622B120F6721"
    message = "Hello, secp256k1!"
    hash = SHA256_VBA.SHA256_String(message)
    
    ' Assinar usando Bitcoin Core
    Dim sig As ECDSA_SIGNATURE
    sig = ecdsa_sign_bitcoin_core(hash, private_key, ctx)
    
    ' Gerar chave pública
    Dim private_bn As BIGNUM_TYPE, public_key As EC_POINT
    private_bn = BN_hex2bn(private_key)
    Call ec_point_mul_generator(public_key, private_bn, ctx)
    
    Debug.Print "r: " & BN_bn2hex(sig.r)
    Debug.Print "s: " & BN_bn2hex(sig.s)
    
    ' Verificar diretamente
    Dim valid As Boolean
    valid = ecdsa_verify_bitcoin_core(hash, sig, public_key, ctx)
    Debug.Print "Verificação direta: " & valid
    
    ' Teste com hash errado
    Dim wrong_hash As String
    wrong_hash = SHA256_VBA.SHA256_String("Wrong message")
    Dim wrong_valid As Boolean
    wrong_valid = ecdsa_verify_bitcoin_core(wrong_hash, sig, public_key, ctx)
    Debug.Print "Hash errado: " & wrong_valid
    
    Debug.Print "=== TESTE DIRETO CONCLUÍDO ==="
End Sub

Public Sub test_ecdsa_low_s_adjustment()
    Debug.Print "=== TESTE ECDSA LOW-S (AJUSTE CANÔNICO) ==="

    Dim ctx As SECP256K1_CTX
    ctx = secp256k1_context_create()

    Dim private_key As String
    private_key = "C9AFA9D845BA75166B5C215767B1D6934E50C3DB36E89B127B8A622B120F6721"

    Dim message As String, hash As String
    message = "low-s-search-2"
    hash = SHA256_VBA.SHA256_String(message)

    Dim sig As ECDSA_SIGNATURE
    sig = ecdsa_sign_bitcoin_core(hash, private_key, ctx)

    Dim actual_r As String, actual_s As String
    actual_r = BN_bn2hex(sig.r)
    actual_s = BN_bn2hex(sig.s)

    Dim expected_r As String, expected_s As String
    expected_r = "EBDCFD8A91922EBE94D666667E3C37B045CAF00DF5693CA97E8E083BA28A7D31"
    expected_s = "50F6F77DA35FCF9E634A54066AEEA8085AD00789522FDE821E4CC74E5B152414"

    Debug.Print "r esperado corresponde: " & (actual_r = expected_r)
    Debug.Print "s esperado corresponde: " & (actual_s = expected_s)
    Debug.Print "r (obtido): " & actual_r
    Debug.Print "s (obtido): " & actual_s

    Dim half_n As BIGNUM_TYPE, two As BIGNUM_TYPE, remainder As BIGNUM_TYPE
    half_n = BN_new()
    two = BN_new()
    remainder = BN_new()
    Call BN_set_word(two, 2)
    Call BN_div(half_n, remainder, ctx.n, two)
    Call BN_free(remainder)

    Debug.Print "s em formato low-s: " & (BN_ucmp(sig.s, half_n) <= 0)

    Dim private_bn As BIGNUM_TYPE, public_key As EC_POINT
    private_bn = BN_hex2bn(private_key)
    Call ec_point_mul_generator(public_key, private_bn, ctx)

    Dim valid As Boolean
    valid = ecdsa_verify_bitcoin_core(hash, sig, public_key, ctx)
    Debug.Print "Verificação assinatura low-s: " & valid

    Call BN_free(half_n)
    Call BN_free(two)
    Call BN_free(private_bn)

    Debug.Print "=== TESTE LOW-S CONCLUÍDO ==="
End Sub

Public Sub test_ecdsa_sign_invalid_hash_inputs()
    Debug.Print "=== TESTE HASH INVÁLIDO NA ASSINATURA ECDSA ==="

    Dim ctx As SECP256K1_CTX
    ctx = secp256k1_context_create()

    Dim private_key As String
    private_key = "C9AFA9D845BA75166B5C215767B1D6934E50C3DB36E89B127B8A622B120F6721"

    Dim expected_error As Long
    expected_error = vbObjectError + &H1002&

    Dim err_len_short As Long
    err_len_short = attempt_sign_with_hash(String$(63, "A"), private_key, ctx)
    Debug.Print "Comprimento inferior rejeitado: " & (err_len_short = expected_error)

    Dim err_len_long As Long
    err_len_long = attempt_sign_with_hash(String$(65, "A"), private_key, ctx)
    Debug.Print "Comprimento superior rejeitado: " & (err_len_long = expected_error)

    Dim err_invalid_char As Long
    err_invalid_char = attempt_sign_with_hash(String$(6, "G") & String$(58, "0"), private_key, ctx)
    Debug.Print "Caractere inválido rejeitado: " & (err_invalid_char = expected_error)

    Debug.Print "=== TESTE HASH INVÁLIDO CONCLUÍDO ==="
End Sub

Private Function attempt_sign_with_hash(ByVal hash_value As String, ByVal private_key As String, ByRef ctx As SECP256K1_CTX) As Long
    On Error GoTo Handler

    Dim sig As ECDSA_SIGNATURE
    sig = ecdsa_sign_bitcoin_core(hash_value, private_key, ctx)

    attempt_sign_with_hash = 0
    On Error GoTo 0
    Exit Function

Handler:
    attempt_sign_with_hash = Err.Number
    Err.Clear
    On Error GoTo 0
End Function
