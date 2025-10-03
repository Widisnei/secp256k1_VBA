Attribute VB_Name = "Final_Optimizations_Test"
Option Explicit

Public Sub test_final_optimizations()
    Debug.Print "=== TESTE OTIMIZAÇÕES FINAIS ==="
    
    Call secp256k1_init
    
    ' Teste Sliding Window NAF
    Debug.Print "Testando Sliding Window NAF..."
    Dim scalar As BIGNUM_TYPE, point As EC_POINT, result As EC_POINT, ctx As SECP256K1_CTX
    scalar = BN_hex2bn("C9AFA9D845BA75166B5C215767B1D6934E50C3DB36E89B127B8A622B120F6721")
    ctx = secp256k1_context_create()
    point = ctx.g
    result = ec_point_new()
    
    If ec_point_mul_sliding_naf(result, scalar, point, ctx) Then
        Debug.Print "[OK] Sliding Window NAF funcionando"
    Else
        Debug.Print "[ERRO] Sliding Window NAF falhou"
    End If
    
    ' Teste Cache Dinâmico
    Debug.Print "Testando cache dinâmico..."
    If ec_point_mul_cached(result, scalar, point, ctx) Then
        Debug.Print "[OK] Cache dinâmico funcionando"
    Else
        Debug.Print "[ERRO] Cache dinâmico falhou"
    End If
    
    Call get_cache_stats
    
    Debug.Print "=== OTIMIZAÇÕES FINAIS TESTADAS ==="
End Sub

Public Function is_frequent_point(ByRef point As EC_POINT) As Boolean
    ' Determina se ponto é usado frequentemente (heurística simples)
    is_frequent_point = (BN_num_bits(point.x) > 200) ' Pontos "grandes" são candidatos a cache
End Function