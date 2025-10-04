Attribute VB_Name = "EC_Optimizations_Advanced"
Option Explicit

' =============================================================================
' EC OPTIMIZATIONS ADVANCED - TÉCNICAS AVANÇADAS BITCOIN CORE
' =============================================================================

Public Function ec_point_mul_generator_optimized(ByRef result As EC_POINT, ByRef scalar As BIGNUM_TYPE, ByRef ctx As SECP256K1_CTX) As Boolean
    ' Multiplicação do gerador com seleção automática da melhor técnica
    If require_constant_time() Then
        ec_point_mul_generator_optimized = ec_point_mul_ladder(result, scalar, ctx.g, ctx)
        Exit Function
    End If

    If use_precomputed_gen_tables() Then
        ec_point_mul_generator_optimized = ec_generator_mul_precomputed_naf(result, scalar, ctx)
    Else
        ec_point_mul_generator_optimized = ec_point_mul_jacobian_optimized(result, scalar, ctx.g, ctx)
    End If
End Function

Public Function ec_generator_mul_precomputed_naf(ByRef result As EC_POINT, ByRef scalar As BIGNUM_TYPE, ByRef ctx As SECP256K1_CTX) As Boolean
    ' Multiplicação com Windowed Non-Adjacent Form baseada na tabela de 1760 entradas
    Const window_size As Long = 4
    Dim naf() As Long
    Dim i As Long

    If require_constant_time() Then
        ec_generator_mul_precomputed_naf = ec_point_mul_ladder(result, scalar, ctx.g, ctx)
        Exit Function
    End If

    Call scalar_to_naf(naf, scalar, window_size)

    Call ec_point_set_infinity(result)

    For i = UBound(naf) To 0 Step -1
        If Not ec_point_double(result, result, ctx) Then
            ec_generator_mul_precomputed_naf = False
            Exit Function
        End If

        Dim digit As Long
        digit = naf(i)

        If digit <> 0 Then
            Dim tableIndex As Long
            tableIndex = (Abs(digit) - 1) \ 2

            Dim precomp_point As EC_POINT
            precomp_point = get_precomputed_point(tableIndex, ctx)

            If precomp_point.infinity Then
                ec_generator_mul_precomputed_naf = False
                Exit Function
            End If

            If digit < 0 Then
                If Not ec_point_negate(precomp_point, precomp_point, ctx) Then
                    ec_generator_mul_precomputed_naf = False
                    Exit Function
                End If
            End If

            If Not ec_point_add(result, result, precomp_point, ctx) Then
                ec_generator_mul_precomputed_naf = False
                Exit Function
            End If
        End If
    Next i

    ec_generator_mul_precomputed_naf = True
End Function

Private Sub scalar_to_naf(ByRef naf() As Long, ByRef scalar As BIGNUM_TYPE, ByVal w As Long)
    ' Converte escalar para wNAF com dígitos ímpares em [-2^{w-1}+1, 2^{w-1}-1]
    Dim pow_w As Long: pow_w = CLng(2 ^ w)
    Dim half_pow As Long: half_pow = pow_w \ 2

    Dim k As BIGNUM_TYPE: k = BN_new()
    Call BN_copy(k, scalar)

    Dim was_negative As Boolean
    was_negative = scalar.neg
    k.neg = False

    Dim remainder As BIGNUM_TYPE: remainder = BN_new()
    Dim magnitude As BIGNUM_TYPE: magnitude = BN_new()
    Dim twoPow As BIGNUM_TYPE: twoPow = BN_new()

    If Not BN_set_word(twoPow, pow_w) Then GoTo wnaf_error

    If BN_is_zero(k) Then
        ReDim naf(0 To 0)
        naf(0) = 0
        GoTo wnaf_finish
    End If

    Dim maxDigits As Long
    maxDigits = BN_num_bits(k) + w + 1
    If maxDigits < 1 Then maxDigits = 1
    ReDim naf(0 To maxDigits - 1)

    Dim used As Long: used = 0

    Do While Not BN_is_zero(k)
        If used > UBound(naf) Then ReDim Preserve naf(0 To used)

        Dim digit As Long

        If BN_is_odd(k) Then
            If Not BN_mod(remainder, k, twoPow) Then GoTo wnaf_error

            If remainder.top > 0 Then
                digit = remainder.d(0)
            Else
                digit = 0
            End If

            If digit >= half_pow Then digit = digit - pow_w

            If (digit And 1) = 0 Then
                If digit >= 0 Then
                    digit = digit + 1 - pow_w
                Else
                    digit = digit - 1 + pow_w
                End If
            End If

            naf(used) = digit

            If digit > 0 Then
                If Not BN_set_word(magnitude, digit) Then GoTo wnaf_error
                If Not BN_sub(k, k, magnitude) Then GoTo wnaf_error
            Else
                If Not BN_set_word(magnitude, -digit) Then GoTo wnaf_error
                If Not BN_add(k, k, magnitude) Then GoTo wnaf_error
            End If
        Else
            naf(used) = 0
        End If

        If Not BN_rshift(k, k, 1) Then GoTo wnaf_error
        used = used + 1
    Loop

    If used = 0 Then
        ReDim naf(0 To 0)
        naf(0) = 0
    Else
        ReDim Preserve naf(0 To used - 1)
    End If

wnaf_finish:
    If was_negative Then
        Dim idx As Long
        For idx = LBound(naf) To UBound(naf)
            naf(idx) = -naf(idx)
        Next idx
    End If
    Exit Sub

wnaf_error:
    ReDim naf(0 To 0)
    naf(0) = 0
    GoTo wnaf_finish
End Sub

Private Function get_precomputed_point(ByVal index As Long, ByRef ctx As SECP256K1_CTX) As EC_POINT
    Dim point As EC_POINT
    point = ec_point_new()

    If index < 0 Then
        Call ec_point_set_infinity(point)
        get_precomputed_point = point
        Exit Function
    End If

    Dim entry As String
    entry = get_precomputed_gen_point(index)

    If Len(entry) = 0 Then
        Call ec_point_set_infinity(point)
        get_precomputed_point = point
        Exit Function
    End If

    If Not decode_precomputed_entry(entry, point, ctx) Then
        Call ec_point_set_infinity(point)
    End If

    get_precomputed_point = point
End Function

Private Function decode_precomputed_entry(ByVal entry As String, ByRef point As EC_POINT, ByRef ctx As SECP256K1_CTX) As Boolean
    Dim coords() As String
    coords = Split(entry, ",")

    If UBound(coords) < 15 Then Exit Function

    Dim x_hex As String
    Dim y_hex As String
    Dim i As Long

    For i = 7 To 0 Step -1
        x_hex = x_hex & normalize_word(coords(i))
    Next i

    For i = 7 To 0 Step -1
        y_hex = y_hex & normalize_word(coords(i + 8))
    Next i

    On Error GoTo decode_error

    point.x = BN_hex2bn(x_hex)
    point.y = BN_hex2bn(y_hex)
    Call BN_set_word(point.z, 1)
    point.infinity = False

    If Not ec_point_is_on_curve(point, ctx) Then GoTo decode_error

    decode_precomputed_entry = True
    On Error GoTo 0
    Exit Function

decode_error:
    On Error GoTo 0
    decode_precomputed_entry = False
End Function

Private Function normalize_word(ByVal wordHex As String) As String
    Dim trimmed As String
    trimmed = UCase$(Replace$(Trim$(wordHex), " ", ""))

    If Len(trimmed) = 0 Then trimmed = "0"

    If Len(trimmed) > 8 Then
        trimmed = Right$(trimmed, 8)
    End If

    normalize_word = Right$("00000000" & trimmed, 8)
End Function
