Attribute VB_Name = "Test_Generator_wNAF"
Option Explicit
Option Compare Binary
Option Base 0

' =============================================================================
' TESTES DE REGRESSÃO - MULTIPLICAÇÃO DO GERADOR COM WNAF PRÉ-COMPUTADO
' =============================================================================
'
' OBJETIVOS:
'   • Validar que ec_generator_mul_precomputed_naf produz o mesmo resultado
'     que ec_point_mul para um conjunto amplo de escalares.
'   • Garantir que a representação wNAF utilizada emprega dígitos fora de ±1
'     e inclui dígitos negativos, confirmando o uso da tabela pré-computada.
'
Public Sub test_generator_wnaf_regression()
    Debug.Print "=== TESTE WNAF PRECOMPUTED × MULTIPLICAÇÃO GENÉRICA ==="

    Call secp256k1_init

    Dim ctx As SECP256K1_CTX
    ctx = secp256k1_context_create()

    Dim scalars() As String
    scalars = Array( _
        "01", _
        "03", _
        "05", _
        "09", _
        "1D", _
        "11223344556677889900AABBCCDDEEFF", _
        "C9AFA9D845BA75166B5C215767B1D6934E50C3DB36E89B127B8A622B120F6721", _
        "FFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFEBAAEDCE6AF48A03BBFD25E8CD0364140")

    Dim totalChecks As Long: totalChecks = 0
    Dim passedChecks As Long: passedChecks = 0
    Dim idx As Long
    Dim maxAbsDigit As Long: maxAbsDigit = 0
    Dim sawNegativeDigit As Boolean: sawNegativeDigit = False

    For idx = LBound(scalars) To UBound(scalars)
        Dim scalar_bn As BIGNUM_TYPE
        scalar_bn = BN_hex2bn(scalars(idx))

        Dim nafDigits() As Long
        nafDigits = compute_wnaf_digits_for_test(scalar_bn, 4)

        Dim j As Long
        For j = LBound(nafDigits) To UBound(nafDigits)
            If Abs(nafDigits(j)) > maxAbsDigit Then maxAbsDigit = Abs(nafDigits(j))
            If nafDigits(j) < 0 Then sawNegativeDigit = True
        Next j

        Dim result_opt As EC_POINT
        Dim result_ref As EC_POINT
        result_opt = ec_point_new()
        result_ref = ec_point_new()

        totalChecks = totalChecks + 1

        If Not ec_generator_mul_precomputed_naf(result_opt, scalar_bn, ctx) Then
            Debug.Print "✗ Falha na multiplicação otimizada para escalar " & scalars(idx)
            GoTo next_scalar
        End If

        If Not ec_point_mul(result_ref, scalar_bn, ctx.g, ctx) Then
            Debug.Print "✗ Falha na multiplicação de referência para escalar " & scalars(idx)
            GoTo next_scalar
        End If

        If ec_point_cmp(result_opt, result_ref, ctx) = 0 Then
            passedChecks = passedChecks + 1
            Debug.Print "✓ Escalar " & scalars(idx) & " validado"
        Else
            Debug.Print "✗ Divergência entre caminhos para escalar " & scalars(idx)
        End If

next_scalar:
        ' continuar testes mesmo após falhas
    Next idx

    totalChecks = totalChecks + 1
    If maxAbsDigit > 1 Then
        passedChecks = passedChecks + 1
        Debug.Print "✓ Dígito máximo wNAF |d| = " & maxAbsDigit
    Else
        Debug.Print "✗ Nenhum dígito wNAF excedeu ±1 (|d|max = " & maxAbsDigit & ")"
    End If

    totalChecks = totalChecks + 1
    If sawNegativeDigit Then
        passedChecks = passedChecks + 1
        Debug.Print "✓ Representação wNAF apresentou dígitos negativos"
    Else
        Debug.Print "✗ Representação wNAF não gerou dígitos negativos"
    End If

    Debug.Print "--- RESUMO: " & passedChecks & " / " & totalChecks & " verificações aprovadas ---"
End Sub

Private Function compute_wnaf_digits_for_test(ByRef scalar As BIGNUM_TYPE, ByVal window_size As Long) As Long()
    Dim result() As Long

    Dim pow_w As Long: pow_w = CLng(2 ^ window_size)
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
        ReDim result(0 To 0)
        result(0) = 0
        GoTo wnaf_finish
    End If

    Dim maxDigits As Long
    maxDigits = BN_num_bits(k) + window_size + 1
    If maxDigits < 1 Then maxDigits = 1
    ReDim result(0 To maxDigits - 1)

    Dim used As Long: used = 0

    Do While Not BN_is_zero(k)
        If used > UBound(result) Then ReDim Preserve result(0 To used)

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

            result(used) = digit

            If digit > 0 Then
                If Not BN_set_word(magnitude, digit) Then GoTo wnaf_error
                If Not BN_sub(k, k, magnitude) Then GoTo wnaf_error
            Else
                If Not BN_set_word(magnitude, -digit) Then GoTo wnaf_error
                If Not BN_add(k, k, magnitude) Then GoTo wnaf_error
            End If
        Else
            result(used) = 0
        End If

        If Not BN_rshift(k, k, 1) Then GoTo wnaf_error
        used = used + 1
    Loop

    If used = 0 Then
        ReDim result(0 To 0)
        result(0) = 0
    Else
        ReDim Preserve result(0 To used - 1)
    End If

wnaf_finish:
    If was_negative Then
        Dim idxDigit As Long
        For idxDigit = LBound(result) To UBound(result)
            result(idxDigit) = -result(idxDigit)
        Next idxDigit
    End If

    compute_wnaf_digits_for_test = result
    Exit Function

wnaf_error:
    ReDim result(0 To 0)
    result(0) = 0
    GoTo wnaf_finish
End Function
