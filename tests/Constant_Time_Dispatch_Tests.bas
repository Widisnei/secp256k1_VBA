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

    Dim scalar_reg0 As BIGNUM_TYPE
    Dim scalar_reg1 As BIGNUM_TYPE
    scalar_reg0 = BN_hex2bn("A5")
    scalar_reg1 = BN_hex2bn("9A")

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

    Debug.Print "--- Regression: Montgomery ladder constant-time path instrumentation ---"

    Dim iterRef As Long
    Dim iterAlt As Long
    Dim cswapRef As Long
    Dim cswapAlt As Long
    Dim zeroCount As Long
    Dim oneCount As Long

    Call reset_ladder_call_counter()
    If Not ec_point_mul_ladder(result, scalar_reg0, ctx.g, ctx) Then
        Debug.Print "[ERRO] Montgomery ladder falhou para escalar base"
    Else
        iterRef = get_ladder_iteration_counter()
        cswapRef = get_ladder_cswap_counter()
        zeroCount = get_ladder_bit_count(0)
        oneCount = get_ladder_bit_count(1)

        If zeroCount > 0 And oneCount > 0 Then
            Debug.Print "[OK] Ladder executou caminhos de bit 0 e 1 (0s=" & zeroCount & ", 1s=" & oneCount & ")"
        Else
            Debug.Print "[ERRO] Ladder não percorreu ambos os caminhos de bits (0s=" & zeroCount & ", 1s=" & oneCount & ")"
        End If
    End If

    Call reset_ladder_call_counter()
    If Not ec_point_mul_ladder(result, scalar_reg1, ctx.g, ctx) Then
        Debug.Print "[ERRO] Montgomery ladder falhou para escalar alternativa"
    Else
        iterAlt = get_ladder_iteration_counter()
        cswapAlt = get_ladder_cswap_counter()

        If iterRef = iterAlt And cswapRef = cswapAlt Then
            Debug.Print "[OK] Contagem de iterações/cswap idêntica para escalares distintos " & _
                        "(iter=" & iterAlt & ", cswap=" & cswapAlt & ")"
        Else
            Debug.Print "[ERRO] Contagens divergem: iterRef=" & iterRef & _
                        ", iterAlt=" & iterAlt & ", csRef=" & cswapRef & ", csAlt=" & cswapAlt & ")"
        End If
    End If

    Call disable_security_mode
End Sub
