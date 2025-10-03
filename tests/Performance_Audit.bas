Attribute VB_Name = "Performance_Audit"
Option Explicit
Option Compare Binary
Option Base 0

'==============================================================================
' AUDITORIA DE PERFORMANCE SECP256K1_EXCEL
'==============================================================================

Public Sub Run_Performance_Audit()
    Debug.Print "=== AUDITORIA DE PERFORMANCE SECP256K1_EXCEL ==="
    
    Call secp256k1_init
    
    ' 1. Benchmark operações básicas
    Call Benchmark_Basic_Operations
    
    ' 2. Benchmark otimizações
    Call Benchmark_Optimizations
    
    ' 3. Benchmark casos reais
    Call Benchmark_Real_World
    
    Debug.Print "=== AUDITORIA DE PERFORMANCE CONCLUÍDA ==="
End Sub

Private Sub Benchmark_Basic_Operations()
    Debug.Print "--- Benchmark: Operações Básicas ---"
    
    Dim start_time As Double, end_time As Double
    Dim i As Long, iterations As Long: iterations = 10
    
    ' Geração de chaves
    start_time = Timer
    For i = 1 To iterations
        Dim kp As ECDSA_KEYPAIR: kp = secp256k1_generate_keypair()
    Next i
    end_time = Timer
    Debug.Print "Geração de chaves (x" & iterations & "): " & Format((end_time - start_time) * 1000, "0.0") & " ms"
    
    ' Derivação de chave pública
    Dim priv As String: priv = "C9AFA9D845BA75166B5C215767B1D6934E50C3DB36E89B127B8A622B120F6721"
    start_time = Timer
    For i = 1 To iterations
        Dim pub As String: pub = secp256k1_public_key_from_private(priv, True)
    Next i
    end_time = Timer
    Debug.Print "Derivação chave pública (x" & iterations & "): " & Format((end_time - start_time) * 1000, "0.0") & " ms"
    
    ' Assinatura
    Dim hash As String: hash = "A665A45920422F9D417E4867EFDC4FB8A04A1F3FFF1FA07E998E86F7F7A27AE3"
    start_time = Timer
    For i = 1 To iterations
        Dim sig As String: sig = secp256k1_sign(hash, priv)
    Next i
    end_time = Timer
    Debug.Print "Assinatura ECDSA (x" & iterations & "): " & Format((end_time - start_time) * 1000, "0.0") & " ms"
    
    ' Verificação
    Dim signature As String: signature = secp256k1_sign(hash, priv)
    Dim public_key As String: public_key = secp256k1_public_key_from_private(priv, True)
    start_time = Timer
    For i = 1 To iterations
        Dim valid As Boolean: valid = secp256k1_verify(hash, signature, public_key)
    Next i
    end_time = Timer
    Debug.Print "Verificação ECDSA (x" & iterations & "): " & Format((end_time - start_time) * 1000, "0.0") & " ms"
End Sub

Private Sub Benchmark_Optimizations()
    Debug.Print "--- Benchmark: Otimizações ---"
    
    Dim ctx As SECP256K1_CTX: ctx = secp256k1_context_create()
    Dim scalar As BIGNUM_TYPE: scalar = BN_hex2bn("C9AFA9D845BA75166B5C215767B1D6934E50C3DB36E89B127B8A622B120F6721")
    Dim point As EC_POINT: point = ec_point_new()
    Dim start_time As Double, end_time As Double
    
    ' Multiplicação Jacobiana vs Regular
    start_time = Timer
    Call ec_point_mul_jacobian_optimized(point, scalar, ctx.g, ctx)
    end_time = Timer
    Debug.Print "Multiplicação Jacobiana: " & Format((end_time - start_time) * 1000, "0.0") & " ms"
    
    start_time = Timer
    Call ec_point_mul(point, scalar, ctx.g, ctx)
    end_time = Timer
    Debug.Print "Multiplicação Regular: " & Format((end_time - start_time) * 1000, "0.0") & " ms"
    
    ' Tabelas pré-computadas vs Regular
    start_time = Timer
    Call EC_Precomputed_Manager.ec_generator_mul_fast(point, scalar, ctx)
    end_time = Timer
    Debug.Print "Tabelas pré-computadas: " & Format((end_time - start_time) * 1000, "0.0") & " ms"
    
    ' Redução modular rápida vs Regular
    Dim a As BIGNUM_TYPE, result As BIGNUM_TYPE
    a = BN_hex2bn("123456789ABCDEF0123456789ABCDEF0123456789ABCDEF0123456789ABCDEF012345")
    result = BN_new()
    
    start_time = Timer
    Call BN_mod_secp256k1_fast(result, a)
    end_time = Timer
    Debug.Print "Redução modular rápida: " & Format((end_time - start_time) * 1000, "0.0") & " ms"
    
    start_time = Timer
    Call BN_mod(result, a, ctx.p)
    end_time = Timer
    Debug.Print "Redução modular regular: " & Format((end_time - start_time) * 1000, "0.0") & " ms"
End Sub

Private Sub Benchmark_Real_World()
    Debug.Print "--- Benchmark: Casos Reais ---"
    
    Dim start_time As Double, end_time As Double
    Dim i As Long
    
    ' Simulação carteira Bitcoin (100 chaves)
    start_time = Timer
    For i = 1 To 100
        Dim addr As BitcoinAddress: addr = Bitcoin_Address_Generation.generate_bitcoin_address(LEGACY_P2PKH, "mainnet")
    Next i
    end_time = Timer
    Debug.Print "Geração 100 endereços Bitcoin: " & Format((end_time - start_time) * 1000, "0.0") & " ms"
    
    ' Simulação transação Bitcoin (10 assinaturas)
    Dim priv As String: priv = "C9AFA9D845BA75166B5C215767B1D6934E50C3DB36E89B127B8A622B120F6721"
    Dim hash As String: hash = "A665A45920422F9D417E4867EFDC4FB8A04A1F3FFF1FA07E998E86F7F7A27AE3"
    
    start_time = Timer
    For i = 1 To 10
        Dim sig As String: sig = secp256k1_sign(hash, priv)
    Next i
    end_time = Timer
    Debug.Print "Assinatura 10 inputs Bitcoin: " & Format((end_time - start_time) * 1000, "0.0") & " ms"
    
    ' Simulação validação bloco (50 verificações)
    Dim signature As String: signature = secp256k1_sign(hash, priv)
    Dim public_key As String: public_key = secp256k1_public_key_from_private(priv, True)
    
    start_time = Timer
    For i = 1 To 50
        Dim valid As Boolean: valid = secp256k1_verify(hash, signature, public_key)
    Next i
    end_time = Timer
    Debug.Print "Verificação 50 assinaturas: " & Format((end_time - start_time) * 1000, "0.0") & " ms"
End Sub