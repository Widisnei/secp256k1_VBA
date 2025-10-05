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

    Dim scalar_hex As String
    Dim api_point_hex As String
    Dim api_result As String

    Call enable_security_mode

    Dim diagnosticsAvailable As Boolean
    diagnosticsAvailable = ladder_diagnostics_available()
    If diagnosticsAvailable Then
        Call ladder_set_diagnostics_enabled(True)
    End If

    Call reset_ladder_call_counter

    Dim before As Long
    Dim after As Long

    before = get_ladder_call_counter()
    If Not ec_point_mul_ultimate(result, scalar, ctx.g, ctx) Then
        Debug.Print "[ERRO] Multiplicação k*G falhou em modo seguro"
    Else
        after = get_ladder_call_counter()
        If ladder_diagnostics_active() Then
            If after - before = 1 Then
                Debug.Print "[OK] k*G direcionado para Montgomery ladder"
            Else
                Debug.Print "[ERRO] k*G não passou pela Montgomery ladder (delta=" & (after - before) & ")"
            End If
        ElseIf after = 0 And before = 0 Then
            Debug.Print "[INFO] Instrumentação indisponível: counters permanecem zerados para k*G"
        Else
            Debug.Print "[ERRO] Counters alterados sem instrumentação ativa para k*G"
        End If
    End If

    before = get_ladder_call_counter()
    If Not ec_point_mul_ultimate(result, scalar, arbitrary_point, ctx) Then
        Debug.Print "[ERRO] Multiplicação k*P falhou em modo seguro"
    Else
        after = get_ladder_call_counter()
        If ladder_diagnostics_active() Then
            If after - before = 1 Then
                Debug.Print "[OK] k*P direcionado para Montgomery ladder"
            Else
                Debug.Print "[ERRO] k*P não passou pela Montgomery ladder (delta=" & (after - before) & ")"
            End If
        ElseIf after = 0 And before = 0 Then
            Debug.Print "[INFO] Instrumentação indisponível: counters permanecem zerados para k*P"
        Else
            Debug.Print "[ERRO] Counters alterados sem instrumentação ativa para k*P"
        End If
    End If

    Debug.Print "--- Constant-time dispatch: BN_mod_exp_ultimate ---"

    Dim modexpBase As BIGNUM_TYPE
    Dim modexpExp As BIGNUM_TYPE
    Dim modexpMod As BIGNUM_TYPE
    Dim modexpResult As BIGNUM_TYPE
    modexpBase = BN_hex2bn("02")
    modexpExp = BN_hex2bn("FFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFF")
    modexpMod = BN_hex2bn("FFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFEFFFFFC2F")
    modexpResult = BN_new()

    Call reset_mod_exp_dispatch_diagnostics()
    If Not BN_mod_exp_ultimate(modexpResult, modexpBase, modexpExp, modexpMod) Then
        Debug.Print "[ERRO] BN_mod_exp_ultimate falhou em modo seguro"
    ElseIf mod_exp_dispatch_last_algorithm = "CONSTTIME" Then
        Debug.Print "[OK] BN_mod_exp_ultimate roteado para backend constant-time"
    Else
        Debug.Print "[ERRO] BN_mod_exp_ultimate não usou backend constant-time (" & mod_exp_dispatch_last_algorithm & ")"
    End If

    Debug.Print "--- Constant-time dispatch: BN_mod_exp_auto ---"

    Dim modexpAutoResult As BIGNUM_TYPE
    modexpAutoResult = BN_new()

    Call BN_mod_exp_auto_reset_diagnostics()
    If Not BN_mod_exp_auto(modexpAutoResult, modexpBase, modexpExp, modexpMod) Then
        Debug.Print "[ERRO] BN_mod_exp_auto falhou em modo seguro"
    ElseIf BN_mod_exp_auto_last_algorithm = "CONSTTIME" Then
        Debug.Print "[OK] BN_mod_exp_auto roteado para backend constant-time"
    Else
        Debug.Print "[ERRO] BN_mod_exp_auto não usou backend constant-time (" & BN_mod_exp_auto_last_algorithm & ")"
    End If

    Debug.Print "--- Regression: API secp256k1_point_multiply constant-time path ---"

    scalar_hex = BN_bn2hex(scalar)
    api_point_hex = ec_point_compress(arbitrary_point, ctx)

    Call reset_ladder_call_counter()
    before = get_ladder_call_counter()
    api_result = secp256k1_point_multiply(scalar_hex, api_point_hex)

    If api_result = "" Then
        Debug.Print "[ERRO] API secp256k1_point_multiply falhou em modo seguro"
    Else
        after = get_ladder_call_counter()
        If ladder_diagnostics_active() Then
            If after - before = 1 Then
                Debug.Print "[OK] API secp256k1_point_multiply roteada para Montgomery ladder"
            Else
                Debug.Print "[ERRO] API secp256k1_point_multiply não passou pela ladder (delta=" & (after - before) & ")"
            End If
        ElseIf after = 0 And before = 0 Then
            Debug.Print "[INFO] Instrumentação indisponível: counters zerados para secp256k1_point_multiply"
        Else
            Debug.Print "[ERRO] Counters alterados sem instrumentação ativa para secp256k1_point_multiply"
        End If
    End If

    Debug.Print "--- Regression: Montgomery ladder constant-time path instrumentation ---"

    Dim iterRef As Long
    Dim iterAlt As Long
    Dim cswapRef As Long
    Dim cswapAlt As Long
    Dim zeroCount As Long
    Dim oneCount As Long

    If ladder_diagnostics_active() Then
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
    Else
        zeroCount = get_ladder_bit_count(0)
        oneCount = get_ladder_bit_count(1)
        If zeroCount = 0 And oneCount = 0 Then
            Debug.Print "[OK] Build padrão: get_ladder_bit_count retorna zero sem instrumentação"
        Else
            Debug.Print "[ERRO] Contadores de bits alterados em build padrão (0s=" & zeroCount & ", 1s=" & oneCount & ")"
        End If
    End If

    Debug.Print "--- Validando correção da ladder em modo constante ---"

    Dim idx As Long
    Dim randScalar As BIGNUM_TYPE
    Dim ladderPoint As EC_POINT
    Dim referencePoint As EC_POINT
    ladderPoint = ec_point_new()
    referencePoint = ec_point_new()

    Randomize 42
    For idx = 1 To 8
        randScalar = random_scalar_mod_n(ctx)

        If BN_is_zero(randScalar) Then
            Call BN_set_word(randScalar, idx)
        End If

        If Not ec_point_mul_ultimate(ladderPoint, randScalar, arbitrary_point, ctx) Then
            Debug.Print "[ERRO] Multiplicação ladder falhou para escalar aleatório #" & idx
        ElseIf Not ec_point_mul(referencePoint, randScalar, arbitrary_point, ctx) Then
            Debug.Print "[ERRO] Multiplicação de referência falhou para escalar aleatório #" & idx
        ElseIf ec_point_cmp(ladderPoint, referencePoint, ctx) = 0 Then
            Debug.Print "[OK] Ladder corresponde à referência para escalar aleatório #" & idx
        Else
            Debug.Print "[ERRO] Divergência ladder vs referência para escalar aleatório #" & idx
        End If
    Next idx

    Debug.Print "--- Casos especiais: R + R e R + (-R) ---"

    Dim scalarTwo As BIGNUM_TYPE
    Dim ladderDouble As EC_POINT
    Dim expectedDouble As EC_POINT
    scalarTwo = BN_new()
    Call BN_set_word(scalarTwo, 2)

    ladderDouble = ec_point_new()
    expectedDouble = ec_point_new()
    Call ec_point_double(expectedDouble, ctx.g, ctx)

    If Not ec_point_mul_ladder(ladderDouble, scalarTwo, ctx.g, ctx) Then
        Debug.Print "[ERRO] Ladder falhou ao duplicar ponto base (R + R)"
    ElseIf ec_point_cmp(ladderDouble, expectedDouble, ctx) = 0 Then
        Debug.Print "[OK] Ladder retornou duplicação correta para R + R"
    Else
        Debug.Print "[ERRO] Ladder divergente na duplicação R + R"
    End If

    Dim orderScalar As BIGNUM_TYPE
    Dim ladderInfinity As EC_POINT
    orderScalar = BN_new()
    Call BN_copy(orderScalar, ctx.n)
    ladderInfinity = ec_point_new()

    If Not ec_point_mul_ladder(ladderInfinity, orderScalar, ctx.g, ctx) Then
        Debug.Print "[ERRO] Ladder falhou ao multiplicar por ordem (R + (-R))"
    ElseIf ladderInfinity.infinity Then
        Debug.Print "[OK] Ladder retornou ponto no infinito para R + (-R)"
    Else
        Debug.Print "[ERRO] Ladder não retornou infinito para R + (-R)"
    End If

    If diagnosticsAvailable Then
        Call ladder_set_diagnostics_enabled(False)
    End If

    Call disable_security_mode
End Sub

Private Function random_scalar_mod_n(ByRef ctx As SECP256K1_CTX) As BIGNUM_TYPE
    Dim hex_str As String
    Dim j As Long

    hex_str = ""
    For j = 1 To 32
        Dim byte_val As Long
        byte_val = Int(Rnd() * 256)
        hex_str = hex_str & Right$("0" & Hex$(byte_val), 2)
    Next j

    Dim scalar As BIGNUM_TYPE
    scalar = BN_hex2bn(hex_str)

    If Not BN_mod(scalar, scalar, ctx.n) Then
        Call BN_zero(scalar)
    End If

    random_scalar_mod_n = scalar
End Function
