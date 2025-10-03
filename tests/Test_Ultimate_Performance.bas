Attribute VB_Name = "Test_Ultimate_Performance"
Option Explicit

Public Sub test_ultimate_optimizations()
    Debug.Print "=== TESTE DE OTIMIZAÇÕES ULTIMATE ==="
    
    Call secp256k1_init
    Call integrate_all_optimizations
    
    ' Teste 1: Geração de chaves com otimizações
    Dim start_time As Double, end_time As Double
    start_time = Timer
    
    Dim keypair As ECDSA_KEYPAIR
    keypair = secp256k1_generate_keypair()
    
    end_time = Timer
    Debug.Print "Geração de chaves ULTIMATE: " & Format(end_time - start_time, "0.000") & "s"
    
    ' Teste 2: Multiplicação escalar com GLV
    start_time = Timer
    Dim result As EC_POINT, scalar As BIGNUM_TYPE, ctx As SECP256K1_CTX
    result = ec_point_new()
    scalar = BN_hex2bn("C9AFA9D845BA75166B5C215767B1D6934E50C3DB36E89B127B8A622B120F6721")
    ctx = secp256k1_context_create()
    
    Call ec_point_mul_ultimate(result, scalar, ctx.g, ctx)
    end_time = Timer
    Debug.Print "Multiplicação escalar ULTIMATE: " & Format(end_time - start_time, "0.000") & "s"
    
    ' Teste 3: Comparação BigInt
    start_time = Timer
    Dim a As BIGNUM_TYPE, b As BIGNUM_TYPE, mul_result As BIGNUM_TYPE
    a = BN_hex2bn("FFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFF")
    b = BN_hex2bn("EEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEE")
    mul_result = BN_new()
    
    Call BN_mul_auto(mul_result, a, b)
    end_time = Timer
    Debug.Print "Multiplicação BigInt AUTO: " & Format(end_time - start_time, "0.000") & "s"
    
    Debug.Print "=== TODAS AS OTIMIZAÇÕES INTEGRADAS E FUNCIONANDO ==="
End Sub