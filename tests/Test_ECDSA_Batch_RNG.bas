Attribute VB_Name = "Test_ECDSA_Batch_RNG"
Option Explicit

Public Sub test_ecdsa_batch_rng_custom_provider()
    Debug.Print "=== TESTE: RNG personalizado para coeficientes do batch ==="

    Dim ctx As SECP256K1_CTX
    Dim provider As MockBatchRNG
    Dim coeff As BIGNUM_TYPE
    Dim coeffHex As String

    On Error GoTo Fail

    Call secp256k1_init
    ctx = secp256k1_context_create()

    Set provider = New MockBatchRNG
    Call provider.SetFixedHex("A1B2C3D4E5F60718293A4B5C6D7E8F90")
    provider.RaiseErrorAfter = 0
    provider.ShouldRaiseError = False

    Call ecdsa_batch_set_rng_provider(provider)

    coeff = ecdsa_batch_debug_generate_coefficient(ctx)
    coeffHex = BN_bn2hex(coeff)

    Debug.Print "Coeficiente determinístico: ", coeffHex
    Debug.Print "Chamadas ao provedor: ", provider.CallCount

    If provider.CallCount <> 1 Then
        Err.Raise vbObjectError + &H4310&, "test_ecdsa_batch_rng_custom_provider", _
                  "O provedor deveria ter sido chamado exatamente uma vez."
    End If

    If coeffHex <> "A1B2C3D4E5F60718293A4B5C6D7E8F90" Then
        Err.Raise vbObjectError + &H4311&, "test_ecdsa_batch_rng_custom_provider", _
                  "O coeficiente retornado não corresponde ao padrão injetado."
    End If

    If BN_is_zero(coeff) Then
        Err.Raise vbObjectError + &H4312&, "test_ecdsa_batch_rng_custom_provider", _
                  "O coeficiente gerado não pode ser zero."
    End If

    Debug.Print "[OK] Provedor customizado alimentou coeficiente determinístico"

    Call ecdsa_batch_set_rng_provider(Nothing)
    Set provider = Nothing
    Exit Sub

Fail:
    Dim errNumber As Long, errSource As String, errDescription As String
    errNumber = Err.Number
    errSource = Err.Source
    errDescription = Err.Description

    On Error Resume Next
    Call ecdsa_batch_set_rng_provider(Nothing)
    Set provider = Nothing
    On Error GoTo 0

    Err.Raise errNumber, errSource, errDescription
End Sub

Public Sub test_ecdsa_batch_rng_provider_fallback()
    Debug.Print "=== TESTE: Fallback para entropia padrão no batch ==="

    Dim ctx As SECP256K1_CTX
    Dim provider As MockBatchRNG
    Dim coeff As BIGNUM_TYPE

    On Error GoTo Fail

    Call secp256k1_init
    ctx = secp256k1_context_create()

    Set provider = New MockBatchRNG
    provider.ShouldRaiseError = True
    provider.RaiseErrorAfter = 1

    Call ecdsa_batch_set_rng_provider(provider)

    coeff = ecdsa_batch_debug_generate_coefficient(ctx)

    Debug.Print "Chamadas ao provedor antes do fallback: ", provider.CallCount
    Debug.Print "Erros simulados pelo provedor: ", provider.ErrorCount
    Debug.Print "Coeficiente pós-fallback (hex): ", BN_bn2hex(coeff)

    If provider.CallCount <> 1 Then
        Err.Raise vbObjectError + &H4313&, "test_ecdsa_batch_rng_provider_fallback", _
                  "O provedor deveria ter sido invocado exatamente uma vez antes do fallback."
    End If

    If provider.ErrorCount <> 1 Then
        Err.Raise vbObjectError + &H4314&, "test_ecdsa_batch_rng_provider_fallback", _
                  "Era esperado exatamente um erro simulado antes do fallback."
    End If

    If BN_is_zero(coeff) Then
        Err.Raise vbObjectError + &H4315&, "test_ecdsa_batch_rng_provider_fallback", _
                  "O coeficiente obtido após o fallback não pode ser zero."
    End If

    Debug.Print "[OK] Falha do provedor personalizado acionou fallback seguro"

    Call ecdsa_batch_set_rng_provider(Nothing)
    Set provider = Nothing
    Exit Sub

Fail:
    Dim errNumber As Long, errSource As String, errDescription As String
    errNumber = Err.Number
    errSource = Err.Source
    errDescription = Err.Description

    On Error Resume Next
    Call ecdsa_batch_set_rng_provider(Nothing)
    Set provider = Nothing
    On Error GoTo 0

    Err.Raise errNumber, errSource, errDescription
End Sub

Public Sub test_ecdsa_batch_rng_provider_returns_false()
    Debug.Print "=== TESTE: Fallback quando provedor retorna False ==="

    Dim ctx As SECP256K1_CTX
    Dim provider As MockBatchRNG
    Dim coeff As BIGNUM_TYPE

    On Error GoTo Fail

    Call secp256k1_init
    ctx = secp256k1_context_create()

    Set provider = New MockBatchRNG
    provider.ShouldReturnFalse = True
    provider.ReturnFalseAfter = 1

    Call ecdsa_batch_set_rng_provider(provider)

    coeff = ecdsa_batch_debug_generate_coefficient(ctx)

    Debug.Print "Chamadas ao provedor antes do fallback (False): ", provider.CallCount
    Debug.Print "Retornos False simulados: ", provider.FalseCount
    Debug.Print "Coeficiente pós-fallback (hex): ", BN_bn2hex(coeff)

    If provider.CallCount <> 1 Then
        Err.Raise vbObjectError + &H4316&, "test_ecdsa_batch_rng_provider_returns_false", _
                  "O provedor deveria ter sido chamado exatamente uma vez antes do fallback."
    End If

    If provider.FalseCount <> 1 Then
        Err.Raise vbObjectError + &H4317&, "test_ecdsa_batch_rng_provider_returns_false", _
                  "Era esperado exatamente um retorno False simulado antes do fallback."
    End If

    If BN_is_zero(coeff) Then
        Err.Raise vbObjectError + &H4318&, "test_ecdsa_batch_rng_provider_returns_false", _
                  "O coeficiente obtido após o fallback não pode ser zero."
    End If

    Debug.Print "[OK] Retorno False acionou fallback seguro"

    Call ecdsa_batch_set_rng_provider(Nothing)
    Set provider = Nothing
    Exit Sub

Fail:
    Dim errNumber As Long, errSource As String, errDescription As String
    errNumber = Err.Number
    errSource = Err.Source
    errDescription = Err.Description

    On Error Resume Next
    Call ecdsa_batch_set_rng_provider(Nothing)
    Set provider = Nothing
    On Error GoTo 0

    Err.Raise errNumber, errSource, errDescription
End Sub
