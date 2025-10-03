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
    point_hash = compute_point_hash(point)
    
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

Private Function multiply_with_cache(ByRef result As EC_POINT, ByRef scalar As BIGNUM_TYPE, ByVal cache_idx As Long, ByRef ctx As SECP256K1_CTX) As Boolean
    ' Multiplicação usando múltiplos em cache (método windowed)
    Const window_size As Long = 4
    
    Call ec_point_set_infinity(result)
    
    Dim i As Long, nbits As Long, window_val As Long
    nbits = BN_num_bits(scalar)
    
    i = nbits - 1
    Do While i >= 0
        ' Extrair janela de 4 bits
        window_val = 0
        Dim j As Long
        For j = 0 To window_size - 1
            If i - j >= 0 And BN_is_bit_set(scalar, i - j) Then
                window_val = window_val Or (1 * (2 ^ j))
            End If
        Next j
        
        ' Deslocar resultado
        For j = 1 To window_size
            Call ec_point_double(result, result, ctx)
        Next j
        
        ' Adicionar múltiplo do cache se necessário
        If window_val > 0 And window_val <= 15 And (window_val Mod 2 = 1) Then
            Call ec_point_add(result, result, cache(cache_idx).multiples(window_val), ctx)
        End If
        
        i = i - window_size
    Loop
    
    multiply_with_cache = True
End Function

Private Function compute_point_hash(ByRef point As EC_POINT) As String
    ' Computa hash simples do ponto para identificação
    compute_point_hash = Left$(BN_bn2hex(point.x), 16)
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