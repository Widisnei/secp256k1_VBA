Attribute VB_Name = "EC_Sliding_Window_NAF"
Option Explicit

Public Function ec_point_mul_sliding_naf(ByRef result As EC_POINT, ByRef scalar As BIGNUM_TYPE, ByRef point As EC_POINT, ByRef ctx As SECP256K1_CTX) As Boolean
    ' Sliding Window NAF - 15-20% melhoria sobre NAF fixo
    Const max_window As Long = 6
    
    If BN_is_zero(scalar) Or point.infinity Then
        Call ec_point_set_infinity(result)
        ec_point_mul_sliding_naf = True
        Exit Function
    End If
    
    ' Pré-computar ímpares: P, 3P, 5P, ..., (2^w-1)P
    Dim precomp(1 To 31) As EC_POINT ' Até 2^6-1 = 63, só ímpares
    Dim i As Long
    
    For i = 1 To 31 Step 2
        precomp(i) = ec_point_new()
    Next i
    
    Call ec_point_copy(precomp(1), point) ' P
    
    Dim double_p As EC_POINT
    double_p = ec_point_new()
    Call ec_point_double(double_p, point, ctx) ' 2P
    
    For i = 3 To 31 Step 2
        Call ec_point_add(precomp(i), precomp(i - 2), double_p, ctx)
    Next i
    
    ' Converter para Sliding Window NAF
    Dim naf() As Long, naf_len As Long
    Call scalar_to_sliding_naf(naf, naf_len, scalar, max_window)
    
    ' Multiplicação usando NAF
    Call ec_point_set_infinity(result)
    
    For i = naf_len - 1 To 0 Step -1
        Call ec_point_double(result, result, ctx)
        
        If naf(i) <> 0 Then
            Dim abs_val As Long, temp As EC_POINT
            abs_val = Abs(naf(i))
            temp = ec_point_new()
            
            Call ec_point_copy(temp, precomp(abs_val))
            If naf(i) < 0 Then Call ec_point_negate(temp, temp, ctx)
            Call ec_point_add(result, result, temp, ctx)
        End If
    Next i
    
    ec_point_mul_sliding_naf = True
End Function

Private Sub scalar_to_sliding_naf(ByRef naf() As Long, ByRef naf_len As Long, ByRef scalar As BIGNUM_TYPE, ByVal w As Long)
    ' Converte escalar para Sliding Window NAF
    Dim bits As Long, i As Long
    bits = BN_num_bits(scalar)
    ReDim naf(0 To bits + 1)
    naf_len = 0
    
    Dim k As BIGNUM_TYPE
    k = BN_new()
    Call BN_copy(k, scalar)
    
    i = 0
    Do While Not BN_is_zero(k)
        If BN_is_odd(k) Then
            Dim window_size As Long, val As Long
            window_size = 1
            
            ' Encontrar maior janela possível
            Do While window_size < w And BN_is_bit_set(k, window_size)
                window_size = window_size + 1
            Loop
            
            ' Extrair valor da janela
            val = extract_window_value(k, window_size)
            naf(i) = val
            
            ' Subtrair valor extraído
            Dim sub_val As BIGNUM_TYPE
            sub_val = BN_new()
            Call BN_set_word(sub_val, Abs(val))
            If val < 0 Then Call BN_add(k, k, sub_val) Else Call BN_sub(k, k, sub_val)
        Else
            naf(i) = 0
        End If
        
        Call BN_rshift(k, k, 1)
        i = i + 1
    Loop
    
    naf_len = i
End Sub

Private Function extract_window_value(ByRef k As BIGNUM_TYPE, ByVal window_size As Long) As Long
    ' Extrai valor da janela para NAF
    Dim val As Long, i As Long
    
    For i = 0 To window_size - 1
        If BN_is_bit_set(k, i) Then val = val + (2 ^ i)
    Next i
    
    ' Converter para forma NAF (reduzir se necessário)
    If val > (2 ^ (window_size - 1)) Then
        val = val - (2 ^ window_size)
    End If
    
    extract_window_value = val
End Function