Attribute VB_Name = "Test_Point_Validation_FailFast"
Option Explicit

'==============================================================================
' TESTES DE FALHA RÁPIDA PARA VALIDAÇÃO DE PONTOS
'==============================================================================
'
' Propósito: Garante que as rotinas de validação de pontos abortam imediatamente
'            quando a multiplicação escalar de subgrupo falha.
' Escopo:    secp256k1_point_decompress (API principal) e ecdsa_batch_verify
'            (validação com contexto explícito).
'
Public Sub test_point_decompress_rejects_on_mul_failure()
    Debug.Print "=== TESTE: DESCOMPRESSÃO REJEITA QUANDO MULTIPLICAÇÃO FALHA ==="

    On Error GoTo Handler

    If Not secp256k1_init() Then
        Err.Raise vbObjectError + &H5100&, "test_point_decompress_rejects_on_mul_failure", _
                  "Falha ao inicializar o contexto secp256k1."
    End If

    Dim generator As String
    generator = secp256k1_get_generator()

    Dim baseline As String
    baseline = secp256k1_point_decompress(generator)
    If baseline = "" Then
        Err.Raise vbObjectError + &H5101&, "test_point_decompress_rejects_on_mul_failure", _
                  "Pré-condição inválida: descompressão do gerador deveria funcionar."
    End If

    Dim originalUltimate As Boolean
    Dim originalPlain As Boolean
    originalUltimate = EC_Multiplication_Dispatch.ec_point_mul_ultimate_force_failure
    originalPlain = EC_secp256k1_Arithmetic.ec_point_mul_force_failure

    EC_Multiplication_Dispatch.ec_point_mul_ultimate_force_failure = True
    EC_secp256k1_Arithmetic.ec_point_mul_force_failure = True

    Dim failedCoords As String
    failedCoords = secp256k1_point_decompress(generator)
    Debug.Print "Descompressão abortada após falha forçada: ", (failedCoords = "")
    If failedCoords <> "" Then
        Err.Raise vbObjectError + &H5102&, "test_point_decompress_rejects_on_mul_failure", _
                  "secp256k1_point_decompress não rejeitou ponto após falha em ec_point_mul."
    End If

    Dim errCode As SECP256K1_ERROR
    errCode = secp256k1_get_last_error()
    Debug.Print "Erro propagado corretamente: ", (errCode = SECP256K1_ERROR_POINT_NOT_ON_CURVE)
    If errCode <> SECP256K1_ERROR_POINT_NOT_ON_CURVE Then
        Err.Raise vbObjectError + &H5103&, "test_point_decompress_rejects_on_mul_failure", _
                  "Código de erro inesperado após falha de multiplicação na validação."
    End If

    GoTo Cleanup

Handler:
    Debug.Print "FALHOU: " & Err.Description

Cleanup:
    EC_Multiplication_Dispatch.ec_point_mul_ultimate_force_failure = originalUltimate
    EC_secp256k1_Arithmetic.ec_point_mul_force_failure = originalPlain

    If Err.Number <> 0 Then
        Dim errNumber As Long, errSource As String, errDescription As String
        errNumber = Err.Number
        errSource = Err.Source
        errDescription = Err.Description
        Err.Clear
        Debug.Print "=== TESTE ABORTADO ==="
        Err.Raise errNumber, errSource, errDescription
    Else
        Debug.Print "=== TESTE CONCLUÍDO ==="
    End If
End Sub

Public Sub test_batch_validator_rejects_on_mul_failure()
    Debug.Print "=== TESTE: BATCH VERIFY REJEITA QUANDO MULTIPLICAÇÃO FALHA ==="

    On Error GoTo Handler

    Dim ctx As SECP256K1_CTX
    ctx = secp256k1_context_create()

    If Not secp256k1_init() Then
        Err.Raise vbObjectError + &H5200&, "test_batch_validator_rejects_on_mul_failure", _
                  "Falha ao inicializar o contexto global secp256k1."
    End If

    Dim private_key As String
    private_key = "C9AFA9D845BA75166B5C215767B1D6934E50C3DB36E89B127B8A622B120F6721"

    Dim private_bn As BIGNUM_TYPE
    private_bn = BN_hex2bn(private_key)

    Dim public_key As EC_POINT
    If Not ec_point_mul_generator(public_key, private_bn, ctx) Then
        Err.Raise vbObjectError + &H5201&, "test_batch_validator_rejects_on_mul_failure", _
                  "Falha ao derivar chave pública para o teste de lote."
    End If

    Dim message As String, hash As String
    message = "Batch verify rejeição por falha forçada"
    hash = SHA256_VBA.SHA256_String(message)

    Dim batch(0 To 0) As BATCH_SIGNATURE
    batch(0).message_hash = hash
    batch(0).signature = ecdsa_sign_bitcoin_core(hash, private_key, ctx)
    batch(0).public_key = public_key

    Dim baseline_result As Boolean
    baseline_result = ecdsa_batch_verify(batch, ctx)
    Debug.Print "Lote válido aceito antes da falha forçada: ", baseline_result
    If Not baseline_result Then
        Err.Raise vbObjectError + &H5202&, "test_batch_validator_rejects_on_mul_failure", _
                  "Pré-condição inválida: lote válido deveria ser aceito."
    End If

    Dim originalUltimate As Boolean
    Dim originalPlain As Boolean
    originalUltimate = EC_Multiplication_Dispatch.ec_point_mul_ultimate_force_failure
    originalPlain = EC_secp256k1_Arithmetic.ec_point_mul_force_failure

    EC_Multiplication_Dispatch.ec_point_mul_ultimate_force_failure = True
    EC_secp256k1_Arithmetic.ec_point_mul_force_failure = True

    Dim failure_result As Boolean
    failure_result = ecdsa_batch_verify(batch, ctx)
    Debug.Print "Lote rejeitado após falha forçada: ", (Not failure_result)
    If failure_result Then
        Err.Raise vbObjectError + &H5203&, "test_batch_validator_rejects_on_mul_failure", _
                  "ecdsa_batch_verify aceitou lote mesmo com falha em ec_point_mul durante validação."
    End If

    GoTo Cleanup

Handler:
    Debug.Print "FALHOU: " & Err.Description

Cleanup:
    EC_Multiplication_Dispatch.ec_point_mul_ultimate_force_failure = originalUltimate
    EC_secp256k1_Arithmetic.ec_point_mul_force_failure = originalPlain

    If Err.Number <> 0 Then
        Dim errNumber As Long, errSource As String, errDescription As String
        errNumber = Err.Number
        errSource = Err.Source
        errDescription = Err.Description
        Err.Clear
        Debug.Print "=== TESTE ABORTADO ==="
        Err.Raise errNumber, errSource, errDescription
    Else
        Debug.Print "=== TESTE CONCLUÍDO ==="
    End If
End Sub
