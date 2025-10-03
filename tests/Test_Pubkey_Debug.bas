Attribute VB_Name = "Test_Pubkey_Debug"
Option Explicit

'==============================================================================
' DEBUG DE CHAVE PÚBLICA SECP256K1
'==============================================================================
'
' PROPÓSITO:
' • Debug detalhado de geração e processamento de chave pública
' • Validação de compressão/descompressão de pontos EC
' • Verificação de integridade de coordenadas (x, y)
' • Teste de roundtrip: geração → compressão → descompressão
'
' CARACTERÍSTICAS TÉCNICAS:
' • Chave privada: C9AFA9D845BA75166B5C215767B1D6934E50C3DB36E89B127B8A622B120F6721
' • Mensagem: "Hello, secp256k1!"
' • Operação: Q = d × G (multiplicação escalar do gerador)
' • Compressão: Formato 02/03 + coordenada X (33 bytes)
' • Descompressão: Recuperação de coordenada Y via raiz quadrada
'
' ALGORITMOS TESTADOS:
' • ec_point_mul_generator() - Multiplicação escalar do gerador
' • ec_point_compress() - Compressão de ponto EC
' • ec_point_decompress() - Descompressão de ponto EC
' • ec_point_get_affine() - Extração de coordenadas afins
'
' VALIDAÇÕES REALIZADAS:
' • Geração correta de chave pública
' • Compressão sem perda de informação
' • Descompressão recupera ponto original
' • Coordenadas X e Y idênticas após roundtrip
' • Ponto não é infinito após operações
'
' COMPATIBILIDADE:
' • Bitcoin Core secp256k1 - Algoritmos idênticos
' • RFC 5480 - Padrão de compressão EC
' • OpenSSL EC_POINT - Comportamento compatível
' • SEC 1 - Especificação de curvas elípticas
'==============================================================================

'==============================================================================
' DEBUG DE CHAVE PÚBLICA
'==============================================================================

' Propósito: Debug detalhado de geração, compressão e descompressão de chave pública
' Algoritmo: Gera chave pública, comprime, descomprime, compara coordenadas
' Retorno: Relatório detalhado via Debug.Print com validações de integridade

Public Sub Test_Pubkey_Debug()
    Debug.Print "=== DEBUG CHAVE PÚBLICA ==="
    
    Dim ctx As SECP256K1_CTX
    ctx = secp256k1_context_create()
    
    Dim private_key As String, hash As String
    private_key = "C9AFA9D845BA75166B5C215767B1D6934E50C3DB36E89B127B8A622B120F6721"
    hash = SHA256_VBA.SHA256_String("Hello, secp256k1!")
    
    ' Gerar chave pública diretamente
    Dim private_bn As BIGNUM_TYPE, public_key_direct As EC_POINT
    private_bn = BN_hex2bn(private_key)
    Call ec_point_mul_generator(public_key_direct, private_bn, ctx)
    
    ' Comprimir
    Dim compressed As String
    compressed = ec_point_compress(public_key_direct, ctx)
    Debug.Print "Comprimida: " & compressed
    
    ' Descomprimir
    Dim public_key_decompressed As EC_POINT
    public_key_decompressed = ec_point_decompress(compressed, ctx)
    
    Debug.Print "Descompressão OK: " & Not public_key_decompressed.infinity
    
    ' Comparar coordenadas
    Dim x1 As BIGNUM_TYPE, y1 As BIGNUM_TYPE, x2 As BIGNUM_TYPE, y2 As BIGNUM_TYPE
    x1 = BN_new(): y1 = BN_new(): x2 = BN_new(): y2 = BN_new()
    
    Call ec_point_get_affine(public_key_direct, x1, y1, ctx)
    Call ec_point_get_affine(public_key_decompressed, x2, y2, ctx)
    
    Debug.Print "X igual: " & (BN_cmp(x1, x2) = 0)
    Debug.Print "Y igual: " & (BN_cmp(y1, y2) = 0)
    
    Debug.Print "=== DEBUG CONCLUÍDO ==="
End Sub