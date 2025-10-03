Attribute VB_Name = "BigInt_Optimizer"
Option Explicit

' =============================================================================
' BIGINT OPTIMIZER - DISPATCHER INTELIGENTE GLOBAL
' =============================================================================
' Integra todas as técnicas BigInt disponíveis e seleciona automaticamente
' a melhor implementação baseada no tamanho dos operandos e tipo de operação
' =============================================================================

Public Function BN_mul_auto(ByRef r As BIGNUM_TYPE, ByRef a As BIGNUM_TYPE, ByRef b As BIGNUM_TYPE) As Boolean
    ' Seleção automática do melhor algoritmo de multiplicação
    Dim total_bits As Long
    total_bits = BN_num_bits(a) + BN_num_bits(b)
    
    If total_bits <= 512 Then
        BN_mul_auto = BN_mul_fast256(r, a, b)      ' COMBA 8x8
    ElseIf total_bits <= 2048 Then
        BN_mul_auto = BN_mul_karatsuba(r, a, b)    ' Karatsuba
    Else
        BN_mul_auto = BN_mul(r, a, b)              ' Clássico
    End If
End Function

Public Function BN_mod_auto(ByRef r As BIGNUM_TYPE, ByRef a As BIGNUM_TYPE, ByRef m As BIGNUM_TYPE) As Boolean
    ' Seleção automática de redução modular
    If is_secp256k1_prime(m) Then
        BN_mod_auto = BN_mod_secp256k1_fast(r, a)  ' Redução especializada
    Else
        BN_mod_auto = BN_mod(r, a, m)              ' Redução genérica
    End If
End Function

Private Function is_secp256k1_prime(ByRef m As BIGNUM_TYPE) As Boolean
    Dim secp_p As BIGNUM_TYPE
    secp_p = BN_hex2bn("FFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFEFFFFFC2F")
    is_secp256k1_prime = (BN_cmp(m, secp_p) = 0)
End Function