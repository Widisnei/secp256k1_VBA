Attribute VB_Name = "EC_Precomputed_Cache"
Option Explicit

' =============================================================================
' PRECOMPUTED CACHE - CACHE DINÂMICO DE MÚLTIPLOS
' =============================================================================

Private Type CACHE_ENTRY
    point_hash As String
    multiples(1 To 15) As EC_POINT
    initialized As Boolean
End Type

Private cache(0 To 7) As CACHE_ENTRY ' Cache para 8 pontos mais usados
Private cache_hits As Long
Private cache_misses As Long

Public Function ec_point_mul_cached(ByRef result As EC_POINT, ByRef scalar As BIGNUM_TYPE, ByRef point As EC_POINT, ByRef ctx As SECP256K1_CTX) As Boolean
    ' Multiplicação com cache dinâmico de múltiplos ímpares

    If require_constant_time() Then
        ec_point_mul_cached = ec_point_mul_ladder(result, scalar, point, ctx)
        Exit Function
    End If

    If BN_is_zero(scalar) Or point.infinity Then
        Call ec_point_set_infinity(result)
        ec_point_mul_cached = True
        Exit Function
    End If
    
    ' Verificar cache
    Dim cache_idx As Long
    cache_idx = find_or_create_cache_entry(point, ctx)
    
    If cache_idx >= 0 Then
        ' Usar multiplicação com cache
        ec_point_mul_cached = multiply_with_cache(result, scalar, cache_idx, ctx)
        cache_hits = cache_hits + 1
    Else
        ' Fallback para multiplicação regular
        ec_point_mul_cached = ec_point_mul(result, scalar, point, ctx)
        cache_misses = cache_misses + 1
    End If
End Function

Private Function find_or_create_cache_entry(ByRef point As EC_POINT, ByRef ctx As SECP256K1_CTX) As Long
    ' Encontra entrada no cache ou cria nova
    Dim point_hash As String, i As Long
    point_hash = compute_point_hash(point, ctx)

    If Len(point_hash) = 0 Then
        find_or_create_cache_entry = -1
        Exit Function
    End If

    ' Invalida entradas legadas com formato antigo de hash
    Dim expectedLen As Long
    expectedLen = Len(point_hash)

    For i = 0 To 7
        If cache(i).initialized Then
            If Len(cache(i).point_hash) <> expectedLen Then
                Call invalidate_cache_entry(i)
            End If
        End If
    Next i

    ' Procurar entrada existente
    For i = 0 To 7
        If cache(i).initialized And cache(i).point_hash = point_hash Then
            find_or_create_cache_entry = i
            Exit Function
        End If
    Next i
    
    ' Criar nova entrada (substituir LRU)
    Dim lru_idx As Long
    lru_idx = find_lru_entry()
    
    If create_cache_entry(lru_idx, point, point_hash, ctx) Then
        find_or_create_cache_entry = lru_idx
    Else
        find_or_create_cache_entry = -1
    End If
End Function

Private Sub invalidate_cache_entry(ByVal idx As Long)
    Dim j As Long

    For j = LBound(cache(idx).multiples) To UBound(cache(idx).multiples)
        cache(idx).multiples(j) = ec_point_new()
    Next j

    cache(idx).point_hash = ""
    cache(idx).initialized = False
End Sub

Private Function create_cache_entry(ByVal idx As Long, ByRef point As EC_POINT, ByVal point_hash As String, ByRef ctx As SECP256K1_CTX) As Boolean
    ' Cria entrada no cache com múltiplos ímpares pré-computados
    cache(idx).point_hash = point_hash
    
    ' Inicializar múltiplos ímpares: P, 3P, 5P, ..., 15P
    Dim i As Long
    For i = 1 To 15
        cache(idx).multiples(i) = ec_point_new()
    Next i
    
    Call ec_point_copy(cache(idx).multiples(1), point) ' P
    
    Dim double_p As EC_POINT
    double_p = ec_point_new()
    Call ec_point_double(double_p, point, ctx) ' 2P
    
    For i = 3 To 15 Step 2
        Call ec_point_add(cache(idx).multiples(i), cache(idx).multiples(i - 2), double_p, ctx)
    Next i
    
    cache(idx).initialized = True
    create_cache_entry = True
End Function

Private Function compute_windowed_wnaf_digits(ByRef scalar As BIGNUM_TYPE, ByRef digits() As Long) As Long
    Const window_size As Long = 4

    Dim pow_w As Long: pow_w = CLng(2 ^ window_size)
    Dim half_pow As Long: half_pow = pow_w \ 2

    Dim k As BIGNUM_TYPE: k = BN_new()
    Call BN_copy(k, scalar)
    k.neg = False

    Dim maxDigits As Long
    maxDigits = BN_num_bits(scalar) + window_size + 1
    If maxDigits < 1 Then maxDigits = 1
    ReDim digits(0 To maxDigits - 1)

    Dim remainder As BIGNUM_TYPE: remainder = BN_new()
    Dim magnitude As BIGNUM_TYPE: magnitude = BN_new()
    Dim twoPow As BIGNUM_TYPE: twoPow = BN_new()
    If Not BN_set_word(twoPow, pow_w) Then
        compute_windowed_wnaf_digits = -1
        Exit Function
    End If

    Dim used As Long: used = 0
    Dim success As Boolean: success = True

    Do While Not BN_is_zero(k)
        If used > UBound(digits) Then
            ReDim Preserve digits(0 To used)
        End If

        Dim digit As Long
        If BN_is_odd(k) Then
            If Not BN_mod(remainder, k, twoPow) Then success = False: Exit Do

            If remainder.top > 0 Then
                digit = remainder.d(0)
            Else
                digit = 0
            End If

            If digit >= half_pow Then
                digit = digit - pow_w
            End If

            If (digit And 1) = 0 Then
                If digit >= 0 Then
                    digit = digit + 1 - pow_w
                Else
                    digit = digit - 1 + pow_w
                End If
            End If

            digits(used) = digit

            If digit > 0 Then
                If Not BN_set_word(magnitude, digit) Then success = False: Exit Do
                If Not BN_sub(k, k, magnitude) Then success = False: Exit Do
            Else
                If Not BN_set_word(magnitude, -digit) Then success = False: Exit Do
                If Not BN_add(k, k, magnitude) Then success = False: Exit Do
            End If
        Else
            digits(used) = 0
        End If

        If Not BN_rshift(k, k, 1) Then success = False: Exit Do
        used = used + 1
    Loop

    If Not success Then
        compute_windowed_wnaf_digits = -1
        Exit Function
    End If

    If used = 0 Then
        ReDim digits(0 To 0)
        digits(0) = 0
        compute_windowed_wnaf_digits = -1
        Exit Function
    End If

    ReDim Preserve digits(0 To used - 1)

    Dim highest As Long
    For highest = used - 1 To 0 Step -1
        If digits(highest) <> 0 Then
            compute_windowed_wnaf_digits = highest
            Exit Function
        End If
    Next highest

    compute_windowed_wnaf_digits = -1
End Function

Private Function multiply_with_cache(ByRef result As EC_POINT, ByRef scalar As BIGNUM_TYPE, ByVal cache_idx As Long, ByRef ctx As SECP256K1_CTX) As Boolean
    Call ec_point_set_infinity(result)

    Dim negateResult As Boolean
    negateResult = scalar.neg

    Dim digits() As Long
    Dim highest As Long
    highest = compute_windowed_wnaf_digits(scalar, digits)

    If highest < 0 Then
        multiply_with_cache = BN_is_zero(scalar)
        Exit Function
    End If

    Dim addPoint As EC_POINT
    addPoint = ec_point_new()

    Dim started As Boolean
    Dim i As Long

    For i = highest To 0 Step -1
        If started Then
            If Not ec_point_double(result, result, ctx) Then multiply_with_cache = False: Exit Function
        End If

        Dim digit As Long
        digit = digits(i)

        If digit <> 0 Then
            Dim multipleIndex As Long
            multipleIndex = CLng(Abs(digit))

            If multipleIndex < 1 Or multipleIndex > 15 Then
                multiply_with_cache = False
                Exit Function
            End If

            Call ec_point_copy(addPoint, cache(cache_idx).multiples(multipleIndex))

            If digit < 0 Then
                If Not ec_point_negate(addPoint, addPoint, ctx) Then multiply_with_cache = False: Exit Function
            End If

            If Not started Then
                Call ec_point_copy(result, addPoint)
                started = True
            Else
                If Not ec_point_add(result, result, addPoint, ctx) Then multiply_with_cache = False: Exit Function
            End If
        End If
    Next i

    If Not started Then
        Call ec_point_set_infinity(result)
    ElseIf negateResult Then
        If Not ec_point_is_infinity(result) Then
            If Not ec_point_negate(result, result, ctx) Then multiply_with_cache = False: Exit Function
        End If
    End If

    multiply_with_cache = True
End Function

Private Function compute_point_hash(ByRef point As EC_POINT, ByRef ctx As SECP256K1_CTX) As String
    ' Utiliza codificação SEC comprimida (inclui paridade de y) como identificador do ponto
    If point.infinity Then
        compute_point_hash = "00"
        Exit Function
    End If

    Dim compressed As String
    compressed = ec_point_compress(point, ctx)

    compute_point_hash = compressed
End Function

Private Function find_lru_entry() As Long
    ' Encontra entrada menos recentemente usada (implementação simples)
    Static next_replace As Long
    find_lru_entry = next_replace
    next_replace = (next_replace + 1) Mod 8
End Function

Public Sub get_cache_stats()
    Debug.Print "=== ESTATÍSTICAS DO CACHE ==="
    Debug.Print "Cache hits: " & cache_hits
    Debug.Print "Cache misses: " & cache_misses
    If cache_hits + cache_misses > 0 Then
        Debug.Print "Hit rate: " & Format(cache_hits / (cache_hits + cache_misses) * 100, "0.0") & "%"
    End If
End Sub