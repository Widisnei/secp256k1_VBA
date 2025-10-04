Attribute VB_Name = "Constant_Time_Dispatch_Tests"
Option Explicit

Public Sub Run_Constant_Time_Dispatch_Tests()
    Debug.Print "=== Testes Constant-Time Dispatch ==="

    Call secp256k1_init

    Dim ctx As SECP256K1_CTX
    ctx = secp256k1_context_create()

    Dim scalar As BIGNUM_TYPE
    scalar = BN_new()
    Call BN_set_word(scalar, 12345&)

    Dim result As EC_POINT
    result = ec_point_new()

    Dim arbitrary_point As EC_POINT
    arbitrary_point = ec_point_new()
    Call ec_point_double(arbitrary_point, ctx.g, ctx)

    Call enable_security_mode
    Call reset_ladder_call_counter

    Dim before As Long
    Dim after As Long

    before = get_ladder_call_counter()
    If Not ec_point_mul_ultimate(result, scalar, ctx.g, ctx) Then
        Debug.Print "[ERRO] Multiplicação k*G falhou em modo seguro"
    Else
        after = get_ladder_call_counter()
        If after - before = 1 Then
            Debug.Print "[OK] k*G direcionado para Montgomery ladder"
        Else
            Debug.Print "[ERRO] k*G não passou pela Montgomery ladder (delta=" & (after - before) & ")"
        End If
    End If

    before = get_ladder_call_counter()
    If Not ec_point_mul_ultimate(result, scalar, arbitrary_point, ctx) Then
        Debug.Print "[ERRO] Multiplicação k*P falhou em modo seguro"
    Else
        after = get_ladder_call_counter()
        If after - before = 1 Then
            Debug.Print "[OK] k*P direcionado para Montgomery ladder"
        Else
            Debug.Print "[ERRO] k*P não passou pela Montgomery ladder (delta=" & (after - before) & ")"
        End If
    End If

    Call disable_security_mode
End Sub
