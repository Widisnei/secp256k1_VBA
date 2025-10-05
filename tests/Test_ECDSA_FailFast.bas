Attribute VB_Name = "Test_ECDSA_FailFast"
Option Explicit

'==============================================================================
' TESTES DE FALHA RÁPIDA PARA MULTIPLICAÇÃO ESCALAR
'==============================================================================
'
' Propósito: Garantir que as rotinas de assinatura e verificação abortam
'            imediatamente quando ec_point_mul_ultimate sinaliza erro.
' Escopo:   ecdsa_sign_bitcoin_core, ecdsa_verify_bitcoin_core, secp256k1_sign,
'            secp256k1_verify.
'
Public Sub test_ecdsa_fail_fast_on_mul_failure()
    Debug.Print "=== TESTE: FALHA RÁPIDA QUANDO MULTIPLICAÇÃO ESCALAR FALHA ==="

    On Error GoTo Handler

    Dim ctx As SECP256K1_CTX
    ctx = secp256k1_context_create()

    Dim apiInitialized As Boolean
    apiInitialized = secp256k1_init()
    Debug.Print "API inicializada: ", apiInitialized

    Dim private_key As String
    private_key = "C9AFA9D845BA75166B5C215767B1D6934E50C3DB36E89B127B8A622B120F6721"

    Dim message As String, hash As String
    message = "Fail fast test message"
    hash = SHA256_VBA.SHA256_String(message)

    Dim private_bn As BIGNUM_TYPE
    private_bn = BN_hex2bn(private_key)

    Dim public_key As EC_POINT
    If Not ec_point_mul_generator(public_key, private_bn, ctx) Then
        Err.Raise vbObjectError + &H3100&, "test_ecdsa_fail_fast_on_mul_failure", _
                  "Falha inesperada ao derivar chave pública de teste."
    End If

    Dim sig As ECDSA_SIGNATURE
    sig = ecdsa_sign_bitcoin_core(hash, private_key, ctx)

    Dim signature_der As String
    signature_der = ecdsa_signature_to_der(sig)

    Dim compressed_pub As String
    compressed_pub = ec_point_compress(public_key, ctx)

    Dim originalFlag As Boolean
    originalFlag = EC_Multiplication_Dispatch.ec_point_mul_ultimate_force_failure
    EC_Multiplication_Dispatch.ec_point_mul_ultimate_force_failure = True

    Dim verify_result As Boolean
    verify_result = ecdsa_verify_bitcoin_core(hash, sig, public_key, ctx)
    Debug.Print "Verificação direta abortou: ", (Not verify_result)
    If verify_result Then
        Err.Raise vbObjectError + &H3101&, "test_ecdsa_fail_fast_on_mul_failure", _
                  "ecdsa_verify_bitcoin_core não abortou quando a multiplicação falhou."
    End If

    Dim api_verify_result As Boolean
    api_verify_result = secp256k1_verify(hash, signature_der, compressed_pub)
    Debug.Print "Verificação via API abortou: ", (Not api_verify_result)
    If api_verify_result Then
        Err.Raise vbObjectError + &H3102&, "test_ecdsa_fail_fast_on_mul_failure", _
                  "secp256k1_verify retornou sucesso mesmo com falha de multiplicação."
    End If

    Dim api_verify_error As SECP256K1_ERROR
    api_verify_error = secp256k1_get_last_error()
    Debug.Print "Código de erro via API (falha multiplicação): ", _
                (api_verify_error = SECP256K1_ERROR_COMPUTATION_FAILED)
    If api_verify_error <> SECP256K1_ERROR_COMPUTATION_FAILED Then
        Err.Raise vbObjectError + &H3106&, "test_ecdsa_fail_fast_on_mul_failure", _
                  "secp256k1_verify não propagou SECP256K1_ERROR_COMPUTATION_FAILED."
    End If

    Dim sign_error As Long
    On Error Resume Next
    Dim fail_sig As ECDSA_SIGNATURE
    fail_sig = ecdsa_sign_bitcoin_core(hash, private_key, ctx)
    sign_error = Err.Number
    Err.Clear
    On Error GoTo 0
    On Error GoTo Handler
    Debug.Print "Assinatura direta gerou erro: ", (sign_error = vbObjectError + &H1103&)
    If sign_error <> vbObjectError + &H1103& Then
        Err.Raise vbObjectError + &H3103&, "test_ecdsa_fail_fast_on_mul_failure", _
                  "ecdsa_sign_bitcoin_core não propagou erro esperado na falha de multiplicação."
    End If

    Dim api_signature As String
    api_signature = secp256k1_sign(hash, private_key)
    Dim api_error As SECP256K1_ERROR
    api_error = secp256k1_get_last_error()
    Debug.Print "Assinatura via API retornou vazio: ", (api_signature = "")
    Debug.Print "Assinatura via API sinalizou falha de computação: ", _
                (api_error = SECP256K1_ERROR_COMPUTATION_FAILED)
    If api_signature <> "" Then
        Err.Raise vbObjectError + &H3104&, "test_ecdsa_fail_fast_on_mul_failure", _
                  "secp256k1_sign retornou assinatura mesmo após falha de multiplicação."
    End If
    If api_error <> SECP256K1_ERROR_COMPUTATION_FAILED Then
        Err.Raise vbObjectError + &H3105&, "test_ecdsa_fail_fast_on_mul_failure", _
                  "secp256k1_sign não definiu o código de erro correto após falha de multiplicação."
    End If

    GoTo Cleanup

Handler:
    Debug.Print "FALHOU: " & Err.Description

Cleanup:
    EC_Multiplication_Dispatch.ec_point_mul_ultimate_force_failure = originalFlag
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
