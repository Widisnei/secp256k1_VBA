Attribute VB_Name = "Test_ECDSA_Batch_Verify"
Option Explicit

'==============================================================================
' TESTES PARA VERIFICAÇÃO EM LOTE DE ASSINATURAS ECDSA
'==============================================================================
'
' Este módulo valida a rotina ecdsa_batch_verify com cenários positivos e
' negativos. Os testes montam lotes com assinaturas válidas para confirmar que
' são aceitas e lotes com assinaturas adulteradas para garantir que a rotina
' rejeita entradas inválidas.
'
' Cada teste executa duas verificações:
'   1) Um lote com assinaturas corretas deve retornar True.
'   2) Após adulterar uma assinatura ou o hash associado, o lote deve ser
'      rejeitado (retornar False).
'
Public Sub test_ecdsa_batch_verify_invalid_batches()
    Debug.Print "=== TESTE: ECDSA BATCH VERIFY - CASOS INVÁLIDOS ==="

    Dim ctx As SECP256K1_CTX
    ctx = secp256k1_context_create()

    Dim private_key As String
    private_key = "C9AFA9D845BA75166B5C215767B1D6934E50C3DB36E89B127B8A622B120F6721"

    Dim private_bn As BIGNUM_TYPE
    private_bn = BN_hex2bn(private_key)

    Dim public_key As EC_POINT
    Call ec_point_mul_generator(public_key, private_bn, ctx)

    Dim batch(0 To 1) As BATCH_SIGNATURE
    Dim message1 As String, message2 As String
    message1 = "Batch verify mensagem 1"
    message2 = "Batch verify mensagem 2"

    Dim hash1 As String, hash2 As String
    hash1 = SHA256_VBA.SHA256_String(message1)
    hash2 = SHA256_VBA.SHA256_String(message2)

    batch(0).message_hash = hash1
    batch(0).signature = ecdsa_sign_bitcoin_core(hash1, private_key, ctx)
    batch(0).public_key = public_key

    batch(1).message_hash = hash2
    batch(1).signature = ecdsa_sign_bitcoin_core(hash2, private_key, ctx)
    batch(1).public_key = public_key

    Dim valid_result As Boolean
    valid_result = ecdsa_batch_verify(batch, ctx)
    Debug.Print "Lote válido aceito: ", valid_result
    If Not valid_result Then
        Err.Raise vbObjectError + &H2100&, "test_ecdsa_batch_verify_invalid_batches", _
                  "Falha: lote de assinaturas válidas foi rejeitado pela verificação em lote."
    End If

    Dim one As BIGNUM_TYPE
    one = BN_new()
    Call BN_set_word(one, 1)

    ' Adulterar componente r da segunda assinatura para torná-la inválida
    Call BN_add(batch(1).signature.r, batch(1).signature.r, one)
    Call BN_mod(batch(1).signature.r, batch(1).signature.r, ctx.n)
    If BN_is_zero(batch(1).signature.r) Then
        Call BN_add(batch(1).signature.r, batch(1).signature.r, one)
        Call BN_mod(batch(1).signature.r, batch(1).signature.r, ctx.n)
    End If

    Dim tampered_result As Boolean
    tampered_result = ecdsa_batch_verify(batch, ctx)
    Debug.Print "Lote com assinatura adulterada aceito: ", tampered_result
    If tampered_result Then
        Err.Raise vbObjectError + &H2101&, "test_ecdsa_batch_verify_invalid_batches", _
                  "Falha: lote com assinatura adulterada foi aceito pela verificação em lote."
    End If

    ' Segundo cenário: hash da mensagem não corresponde à assinatura
    Dim mismatch_batch(0 To 0) As BATCH_SIGNATURE
    Dim signed_hash As String, wrong_hash As String
    signed_hash = SHA256_VBA.SHA256_String("Batch verify mensagem 3")
    wrong_hash = SHA256_VBA.SHA256_String("Batch verify mensagem incorreta")

    mismatch_batch(0).message_hash = wrong_hash
    mismatch_batch(0).signature = ecdsa_sign_bitcoin_core(signed_hash, private_key, ctx)
    mismatch_batch(0).public_key = public_key

    Dim mismatch_result As Boolean
    mismatch_result = ecdsa_batch_verify(mismatch_batch, ctx)
    Debug.Print "Lote com hash divergente aceito: ", mismatch_result
    If mismatch_result Then
        Err.Raise vbObjectError + &H2102&, "test_ecdsa_batch_verify_invalid_batches", _
                  "Falha: lote com hash divergente foi aceito pela verificação em lote."
    End If

    Debug.Print "--- RESUMO ---"
    Debug.Print "Validação lote válido: ", IIf(valid_result, "OK", "ERRO")
    Debug.Print "Rejeição assinatura adulterada: ", IIf(tampered_result, "ERRO", "OK")
    Debug.Print "Rejeição hash divergente: ", IIf(mismatch_result, "ERRO", "OK")
    Debug.Print "==============================="
End Sub

Public Sub test_ecdsa_batch_verify_rejects_invalid_hash()
    Debug.Print "=== TESTE: ECDSA BATCH VERIFY - REJEITA HASH INVÁLIDO ==="

    Dim ctx As SECP256K1_CTX
    ctx = secp256k1_context_create()

    Dim private_key As String
    private_key = "C9AFA9D845BA75166B5C215767B1D6934E50C3DB36E89B127B8A622B120F6721"

    Dim private_bn As BIGNUM_TYPE
    private_bn = BN_hex2bn(private_key)

    Dim public_key As EC_POINT
    Call ec_point_mul_generator(public_key, private_bn, ctx)

    Dim message As String, valid_hash As String
    message = "Batch verify rejeição hash inválido"
    valid_hash = SHA256_VBA.SHA256_String(message)

    Dim signature_valid As ECDSA_SIGNATURE
    signature_valid = ecdsa_sign_bitcoin_core(valid_hash, private_key, ctx)

    Dim invalid_length_batch(0 To 0) As BATCH_SIGNATURE
    invalid_length_batch(0).message_hash = Left$(valid_hash, Len(valid_hash) - 2)
    invalid_length_batch(0).signature = signature_valid
    invalid_length_batch(0).public_key = public_key

    Dim invalid_length_result As Boolean
    invalid_length_result = ecdsa_batch_verify(invalid_length_batch, ctx)
    Debug.Print "Lote com hash de tamanho inválido aceito: ", invalid_length_result
    If invalid_length_result Then
        Err.Raise vbObjectError + &H2105&, "test_ecdsa_batch_verify_rejects_invalid_hash", _
                  "Falha: verificação em lote aceitou hash com comprimento incorreto."
    End If

    Dim invalid_char_batch(0 To 0) As BATCH_SIGNATURE
    invalid_char_batch(0).message_hash = Left$(valid_hash, Len(valid_hash) - 1) & "Z"
    invalid_char_batch(0).signature = signature_valid
    invalid_char_batch(0).public_key = public_key

    Dim invalid_char_result As Boolean
    invalid_char_result = ecdsa_batch_verify(invalid_char_batch, ctx)
    Debug.Print "Lote com hash contendo caractere inválido aceito: ", invalid_char_result
    If invalid_char_result Then
        Err.Raise vbObjectError + &H2106&, "test_ecdsa_batch_verify_rejects_invalid_hash", _
                  "Falha: verificação em lote aceitou hash com caractere não-hexadecimal."
    End If

    Debug.Print "==============================="
End Sub

Public Sub test_ecdsa_batch_verify_rejects_zero_s()
    Debug.Print "=== TESTE: ECDSA BATCH VERIFY - REJEITA S = 0 ==="

    Dim ctx As SECP256K1_CTX
    ctx = secp256k1_context_create()

    Dim private_key As String
    private_key = "C9AFA9D845BA75166B5C215767B1D6934E50C3DB36E89B127B8A622B120F6721"

    Dim private_bn As BIGNUM_TYPE
    private_bn = BN_hex2bn(private_key)

    Dim public_key As EC_POINT
    Call ec_point_mul_generator(public_key, private_bn, ctx)

    Dim batch(0 To 0) As BATCH_SIGNATURE
    Dim message As String, hash As String
    message = "Batch verify rejeição s zero"
    hash = SHA256_VBA.SHA256_String(message)

    batch(0).message_hash = hash
    batch(0).signature = ecdsa_sign_bitcoin_core(hash, private_key, ctx)
    batch(0).public_key = public_key

    Call BN_zero(batch(0).signature.s)

    Dim zero_s_result As Boolean
    zero_s_result = ecdsa_batch_verify(batch, ctx)

    Debug.Print "Lote com s = 0 aceito: ", zero_s_result
    If zero_s_result Then
        Err.Raise vbObjectError + &H2103&, "test_ecdsa_batch_verify_rejects_zero_s", _
                  "Falha: verificação em lote aceitou assinatura com componente s = 0."
    End If

    Debug.Print "==============================="
End Sub

Public Sub test_ecdsa_batch_verify_rejects_invalid_public_key()
    Debug.Print "=== TESTE: ECDSA BATCH VERIFY - REJEITA CHAVE PÚBLICA INVÁLIDA ==="

    Dim ctx As SECP256K1_CTX
    ctx = secp256k1_context_create()

    Dim private_key As String
    private_key = "C9AFA9D845BA75166B5C215767B1D6934E50C3DB36E89B127B8A622B120F6721"

    Dim batch(0 To 0) As BATCH_SIGNATURE
    Dim message As String, hash As String
    message = "Batch verify rejeição chave inválida"
    hash = SHA256_VBA.SHA256_String(message)

    batch(0).message_hash = hash
    batch(0).signature = ecdsa_sign_bitcoin_core(hash, private_key, ctx)

    Dim invalid_point As EC_POINT
    invalid_point = ec_point_new()
    Call ec_point_set_infinity(invalid_point)
    batch(0).public_key = invalid_point

    Dim invalid_result As Boolean
    invalid_result = ecdsa_batch_verify(batch, ctx)

    Debug.Print "Lote com chave pública inválida aceito: ", invalid_result
    If invalid_result Then
        Err.Raise vbObjectError + &H2104&, "test_ecdsa_batch_verify_rejects_invalid_public_key", _
                  "Falha: verificação em lote aceitou chave pública inválida."
    End If

    Debug.Print "==============================="
End Sub
