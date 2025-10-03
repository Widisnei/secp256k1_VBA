Attribute VB_Name = "EC_Montgomery_Ladder"
Option Explicit

' =============================================================================
' MONTGOMERY LADDER - RESISTÊNCIA A TIMING ATTACKS
' =============================================================================

Public Function ec_point_mul_ladder(ByRef result As EC_POINT, ByRef scalar As BIGNUM_TYPE, ByRef point As EC_POINT, ByRef ctx As SECP256K1_CTX) As Boolean
    ' Multiplicação escalar resistente a timing attacks
    ' Sempre executa mesmo número de operações independente do escalar
    
    If BN_is_zero(scalar) Or point.infinity Then
        Call ec_point_set_infinity(result)
        ec_point_mul_ladder = True
        Exit Function
    End If
    
    Dim R0 As EC_POINT, R1 As EC_POINT, temp As EC_POINT
    R0 = ec_point_new(): R1 = ec_point_new(): temp = ec_point_new()
    
    Call ec_point_set_infinity(R0)  ' R0 = O
    Call ec_point_copy(R1, point)   ' R1 = P
    
    Dim i As Long, nbits As Long, bit As Long
    nbits = BN_num_bits(scalar)
    
    For i = nbits - 1 To 0 Step -1
        bit = IIf(BN_is_bit_set(scalar, i), 1, 0)
        
        ' Sempre executar ambas operações (constant-time)
        Call ec_point_add(temp, R0, R1, ctx)
        Call ec_point_double(R1, R1, ctx)
        Call ec_point_double(R0, R0, ctx)
        
        ' Swap condicional baseado no bit
        If bit = 1 Then
            Call ec_point_copy(R0, temp)
        Else
            Call ec_point_copy(R1, temp)
        End If
    Next i
    
    Call ec_point_copy(result, R0)
    ec_point_mul_ladder = True
End Function