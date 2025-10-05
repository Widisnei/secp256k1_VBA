Attribute VB_Name = "Advanced_Features"
Option Explicit

' =============================================================================
' ADVANCED FEATURES - RECURSOS AVANÇADOS IMPLEMENTADOS
' =============================================================================

Private security_mode As Boolean
Private security_mode_initialized As Boolean

Public Sub initialize_security_mode()
    ' Garantir que o modo de segurança seja ativado sempre que a inicialização ocorrer
    If Not security_mode_initialized Then
        security_mode_initialized = True
    End If

    security_mode = True
End Sub

Private Sub ensure_security_mode_initialized()
    If Not security_mode_initialized Then
        Call initialize_security_mode
    End If
End Sub

Public Sub enable_security_mode()
    ' Ativa modo de segurança máxima (constant-time operations)
    Call ensure_security_mode_initialized()
    security_mode = True
    Debug.Print "[SECURITY] Modo constant-time ativado"
End Sub

Public Sub disable_security_mode()
    ' Desativa modo de segurança (máxima performance)
    Call ensure_security_mode_initialized()
    security_mode = False
    Debug.Print "[PERFORMANCE] Modo máxima performance ativado"
End Sub

Public Function require_constant_time() As Boolean
    ' Verifica se operações constant-time são necessárias
    Call ensure_security_mode_initialized()
    require_constant_time = security_mode
End Function

Public Sub test_advanced_features()
    Debug.Print "=== TESTE RECURSOS AVANÇADOS ==="
    
    Call secp256k1_init
    
    ' Teste Montgomery Ladder
    Debug.Print "Testando Montgomery Ladder (constant-time)..."
    Call enable_security_mode
    
    Dim scalar As BIGNUM_TYPE, point As EC_POINT, result As EC_POINT, ctx As SECP256K1_CTX
    scalar = BN_hex2bn("C9AFA9D845BA75166B5C215767B1D6934E50C3DB36E89B127B8A622B120F6721")
    ctx = secp256k1_context_create()
    point = ctx.g
    result = ec_point_new()
    
    If ec_point_mul_ladder(result, scalar, point, ctx) Then
        Debug.Print "[OK] Montgomery Ladder funcionando"
    Else
        Debug.Print "[ERRO] Montgomery Ladder falhou"
    End If
    
    ' Teste Field-Specific Operations
    Debug.Print "Testando operações de campo especializadas..."
    Dim a As BIGNUM_TYPE, b As BIGNUM_TYPE, r As BIGNUM_TYPE
    a = BN_hex2bn("FFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFF")
    b = BN_hex2bn("EEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEE")
    r = BN_new()
    
    If BN_mod_mul_secp256k1(r, a, b) Then
        Debug.Print "[OK] Multiplicação modular secp256k1 especializada"
    Else
        Debug.Print "[ERRO] Multiplicação modular falhou"
    End If
    
    Debug.Print "=== RECURSOS AVANÇADOS TESTADOS ==="
End Sub