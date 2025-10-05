Attribute VB_Name = "Test_RNG_Override"
Option Explicit

Public Sub Run_RNG_Override_Tests()
    Debug.Print "=== TESTES RNG OVERRIDE ==="
    Call Test_RNG_Override_Deterministico
    Call Test_RNG_Override_BufferVazio
    Call test_ecdsa_batch_rng_custom_provider
    Call test_ecdsa_batch_rng_provider_fallback
    Debug.Print "=== TESTES RNG OVERRIDE CONCLUÍDOS ==="
End Sub

Public Sub Test_RNG_Override_Deterministico()
    Debug.Print "=== TESTE RNG OVERRIDE DETERMINÍSTICO ==="

    Dim seed(0 To 63) As Byte
    Dim i As Long
    For i = 0 To 63
        seed(i) = i Mod &H100
    Next i

    Call ecdsa_rng_override_seed(seed)

    Dim bloco32(0 To 31) As Byte
    If Not ecdsa_collect_secure_entropy(bloco32) Then
        Err.Raise vbObjectError + &H1182&, "Test_RNG_Override_Deterministico", _
                  "Coleta de entropia retornou False durante override."
    End If

    For i = 0 To 31
        If bloco32(i) <> i Then
            Err.Raise vbObjectError + &H1182&, "Test_RNG_Override_Deterministico", _
                      "Byte fora da sequência esperada na primeira leitura."
        End If
    Next i

    Dim bloco16(0 To 15) As Byte
    If Not ecdsa_collect_secure_entropy(bloco16) Then
        Err.Raise vbObjectError + &H1182&, "Test_RNG_Override_Deterministico", _
                  "Segunda coleta de entropia retornou False durante override."
    End If

    For i = 0 To 15
        If bloco16(i) <> (32 + i) Then
            Err.Raise vbObjectError + &H1182&, "Test_RNG_Override_Deterministico", _
                      "Byte fora da sequência esperada na segunda leitura."
        End If
    Next i

    On Error Resume Next
    Dim blocoOverflow(0 To 31) As Byte
    Call ecdsa_collect_secure_entropy(blocoOverflow)
    Dim erroOverflow As Long
    erroOverflow = Err.Number
    On Error GoTo 0

    If erroOverflow <> ecdsa_rng_override_error_exhausted() Then
        Err.Raise vbObjectError + &H1183&, "Test_RNG_Override_Deterministico", _
                  "Esperava erro de entropia esgotada após consumir o buffer de teste."
    End If

    Call ecdsa_rng_override_disable()

    If ecdsa_rng_override_is_enabled() Then
        Err.Raise vbObjectError + &H1184&, "Test_RNG_Override_Deterministico", _
                  "O modo override deveria ser desativado."
    End If

    Debug.Print "[OK] RNG override entrega bytes determinísticos e sinaliza esgotamento"
End Sub

Public Sub Test_RNG_Override_BufferVazio()
    Debug.Print "=== TESTE RNG OVERRIDE BUFFER VAZIO ==="

    Dim vazio() As Byte

    On Error Resume Next
    Call ecdsa_rng_override_seed(vazio)
    Dim erro As Long
    erro = Err.Number
    On Error GoTo 0

    If erro <> ecdsa_rng_override_error_empty() Then
        Err.Raise vbObjectError + &H1185&, "Test_RNG_Override_BufferVazio", _
                  "Esperava erro específico para buffer vazio ao injetar RNG de teste."
    End If

    Debug.Print "[OK] RNG override rejeita buffers vazios"
End Sub
