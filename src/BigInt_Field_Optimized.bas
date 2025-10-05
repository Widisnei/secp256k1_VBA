Attribute VB_Name = "BigInt_Field_Optimized"
Option Explicit

' =============================================================================
' FIELD-SPECIFIC OPTIMIZATIONS - SECP256K1 ESPECIALIZADO
' =============================================================================

Public Function BN_mod_sqr_secp256k1(ByRef r As BIGNUM_TYPE, ByRef a As BIGNUM_TYPE) As Boolean
    ' Quadrado modular especializado para secp256k1: r = a² mod p
    ' 10-15% mais rápido que quadrado genérico
    
    Dim temp As BIGNUM_TYPE
    temp = BN_new()
    
    Dim ok As Boolean

    ' Usar multiplicação COMBA para quadrado
    If a.top <= 8 Then
        ok = BN_sqr_fast256(temp, a)
    Else
        ok = BN_mul(temp, a, a)
    End If

    If Not ok Then
        BN_mod_sqr_secp256k1 = False
        Exit Function
    End If

    ' Redução modular rápida secp256k1
    BN_mod_sqr_secp256k1 = BN_mod_secp256k1_fast(r, temp)
End Function

Public Function BN_mod_mul_secp256k1(ByRef r As BIGNUM_TYPE, ByRef a As BIGNUM_TYPE, ByRef b As BIGNUM_TYPE) As Boolean
    ' Multiplicação modular especializada para secp256k1
    Dim temp As BIGNUM_TYPE
    temp = BN_new()
    
    If Not BN_mul_auto(temp, a, b) Then
        BN_mod_mul_secp256k1 = False
        Exit Function
    End If

    BN_mod_mul_secp256k1 = BN_mod_secp256k1_fast(r, temp)
End Function

Public Function BN_mod_add_secp256k1(ByRef r As BIGNUM_TYPE, ByRef a As BIGNUM_TYPE, ByRef b As BIGNUM_TYPE) As Boolean
    ' Adição modular especializada para secp256k1
    Dim temp As BIGNUM_TYPE
    temp = BN_new()
    
    If Not BN_add(temp, a, b) Then
        BN_mod_add_secp256k1 = False
        Exit Function
    End If

    BN_mod_add_secp256k1 = BN_mod_secp256k1_fast(r, temp)
End Function