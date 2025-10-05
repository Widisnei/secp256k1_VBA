Attribute VB_Name = "secp256k1_API"
Option Explicit
Option Compare Binary
Option Base 0

' =============================================================================
' SECP256K1 VBA - API PRINCIPAL
' =============================================================================
' Implementação completa da curva elíptica secp256k1 em VBA
' Compatível com Bitcoin Core e OpenSSL
' Suporte a assinatura ECDSA, geração de chaves e endereços Bitcoin
' =============================================================================

' Códigos de erro da API
Public Enum SECP256K1_ERROR
    SECP256K1_OK = 0
    SECP256K1_ERROR_INIT_FAILED = 1
    SECP256K1_ERROR_INVALID_PRIVATE_KEY = 2
    SECP256K1_ERROR_INVALID_PUBLIC_KEY = 3
    SECP256K1_ERROR_INVALID_SIGNATURE = 4
    SECP256K1_ERROR_INVALID_HASH = 5
    SECP256K1_ERROR_POINT_NOT_ON_CURVE = 6
    SECP256K1_ERROR_COMPUTATION_FAILED = 7
End Enum

Private ctx As SECP256K1_CTX
Private last_error As SECP256K1_ERROR
Private is_initialized As Boolean

Public Function secp256k1_get_last_error() As SECP256K1_ERROR
    ' Retorna o último código de erro ocorrido
    secp256k1_get_last_error = last_error
End Function

Public Function secp256k1_error_string(ByVal error_code As SECP256K1_ERROR) As String
    ' Converte código de erro para mensagem descritiva
    Select Case error_code
        Case SECP256K1_OK: secp256k1_error_string = "Sucesso"
        Case SECP256K1_ERROR_INIT_FAILED: secp256k1_error_string = "Falha na inicialização"
        Case SECP256K1_ERROR_INVALID_PRIVATE_KEY: secp256k1_error_string = "Chave privada inválida"
        Case SECP256K1_ERROR_INVALID_PUBLIC_KEY: secp256k1_error_string = "Chave pública inválida"
        Case SECP256K1_ERROR_INVALID_SIGNATURE: secp256k1_error_string = "Assinatura inválida"
        Case SECP256K1_ERROR_INVALID_HASH: secp256k1_error_string = "Hash inválido"
        Case SECP256K1_ERROR_POINT_NOT_ON_CURVE: secp256k1_error_string = "Ponto fora da curva"
        Case SECP256K1_ERROR_COMPUTATION_FAILED: secp256k1_error_string = "Falha na computação"
        Case Else: secp256k1_error_string = "Erro desconhecido"
    End Select
End Function

Private Function secp256k1_is_hex_string(ByVal value As String) As Boolean
    ' Verifica se a string contém apenas caracteres hexadecimais válidos (0-9, A-F, a-f)
    Dim i As Long, code As Long
    If Len(value) = 0 Then Exit Function

    For i = 1 To Len(value)
        code = Asc(Mid$(value, i, 1))
        Select Case code
            Case 48 To 57, 65 To 70, 97 To 102
                ' Válido
            Case Else
                Exit Function
        End Select
    Next i

    secp256k1_is_hex_string = True
End Function

' =============================================================================
' INICIALIZAÇÃO DO CONTEXTO SECP256K1
' =============================================================================

Public Function secp256k1_init() As Boolean
    ' Inicializa o contexto secp256k1 com tabelas pré-computadas para performance otimizada
    If is_initialized Then
        secp256k1_init = True
        Exit Function
    End If
    
    last_error = SECP256K1_OK
    ctx = secp256k1_context_create()
    
    ' Carregar tabelas pré-computadas para aceleração das operações
    If init_precomputed_tables() Then
        secp256k1_init = True
        is_initialized = True
        Debug.Print "secp256k1 inicializado com tabelas pré-computadas"
    Else
        secp256k1_init = False
        last_error = SECP256K1_ERROR_INIT_FAILED
        Debug.Print "Erro ao inicializar tabelas pré-computadas"
    End If
End Function

' =============================================================================
' GERAÇÃO E MANIPULAÇÃO DE CHAVES CRIPTOGRÁFICAS
' =============================================================================

Public Function secp256k1_generate_keypair() As ECDSA_KEYPAIR
    ' Gera um novo par de chaves (privada/pública) criptograficamente seguro
    ' Usa geração otimizada com tabelas pré-computadas (alias para versão otimizada)
    secp256k1_generate_keypair = secp256k1_generate_keypair_internal(True)
End Function

Public Function secp256k1_generate_keypair_optimized() As ECDSA_KEYPAIR
    ' Exposição explícita da geração otimizada com tratamento de erros consistente
    secp256k1_generate_keypair_optimized = secp256k1_generate_keypair_internal(True)
End Function

Private Function secp256k1_generate_keypair_internal(ByVal useOptimized As Boolean) As ECDSA_KEYPAIR
    last_error = SECP256K1_OK

    On Error GoTo ComputationFailed

    If useOptimized Then
        secp256k1_generate_keypair_internal = ecdsa_generate_keypair_optimized(ctx)
    Else
        secp256k1_generate_keypair_internal = ecdsa_generate_keypair(ctx)
    End If
    Exit Function

ComputationFailed:
    last_error = SECP256K1_ERROR_COMPUTATION_FAILED
    secp256k1_generate_keypair_internal = secp256k1_empty_keypair()
End Function

Private Function secp256k1_empty_keypair() As ECDSA_KEYPAIR
    Dim empty As ECDSA_KEYPAIR
    empty.private_key = BN_new()
    Call BN_zero(empty.private_key)
    empty.public_key = ec_point_new()
    Call ec_point_set_infinity(empty.public_key)
    secp256k1_empty_keypair = empty
End Function

Public Function secp256k1_private_key_from_hex(ByVal private_key_hex As String) As ECDSA_KEYPAIR
    ' Cria par de chaves a partir de uma chave privada em formato hexadecimal
    last_error = SECP256K1_OK

    On Error GoTo HandleError
    secp256k1_private_key_from_hex = ecdsa_set_private_key(private_key_hex, ctx)
    Exit Function

HandleError:
    Dim empty_keypair As ECDSA_KEYPAIR
    empty_keypair = secp256k1_empty_keypair()
    secp256k1_private_key_from_hex = empty_keypair

    Select Case Err.Number
        Case vbObjectError + &H1002&
            last_error = SECP256K1_ERROR_INVALID_PRIVATE_KEY
        Case vbObjectError + &H1102&
            last_error = SECP256K1_ERROR_COMPUTATION_FAILED
        Case Else
            last_error = SECP256K1_ERROR_COMPUTATION_FAILED
    End Select

    Err.Clear
End Function

Public Function secp256k1_public_key_from_private(ByVal private_key_hex As String, Optional ByVal compressed As Boolean = True) As String
    ' Deriva chave pública a partir de chave privada
    ' Parâmetros:
    '   private_key_hex: Chave privada em formato hexadecimal (64 caracteres)
    '   compressed: True para formato comprimido (33 bytes), False para descomprimido (65 bytes)
    ' Retorna: Chave pública em formato hexadecimal

    ' Validar chave privada primeiro
    last_error = SECP256K1_OK
    If Not secp256k1_validate_private_key(private_key_hex) Then
        last_error = SECP256K1_ERROR_INVALID_PRIVATE_KEY
        secp256k1_public_key_from_private = ""
        Exit Function
    End If
 
    ' Usar derivação otimizada sem revalidação
    Dim derived_public As String
    derived_public = secp256k1_derive_public_key_fast(private_key_hex, compressed)

    If derived_public = "" Then
        If last_error = SECP256K1_OK Then
            last_error = SECP256K1_ERROR_COMPUTATION_FAILED
        End If
        secp256k1_public_key_from_private = ""
        Exit Function
    End If

    secp256k1_public_key_from_private = derived_public
End Function

Public Function secp256k1_derive_public_key_fast(ByVal private_key_hex As String, ByVal compressed As Boolean) As String
    ' Derivação rápida sem revalidação (uso interno)
    Dim private_bn As BIGNUM_TYPE, public_point As EC_POINT
    private_bn = BN_hex2bn(private_key_hex)
    public_point = ec_point_new()
    
    ' Multiplicação otimizada do gerador usando técnicas avançadas
    If Not ec_point_mul_ultimate(public_point, private_bn, ctx.g, ctx) Then
        last_error = SECP256K1_ERROR_COMPUTATION_FAILED
        secp256k1_derive_public_key_fast = ""
        Exit Function
    End If
    
    ' Retornar no formato solicitado
    If compressed Then
        secp256k1_derive_public_key_fast = ec_point_compress(public_point, ctx)
    Else
        secp256k1_derive_public_key_fast = ec_point_to_uncompressed(public_point)
    End If
End Function

Public Function secp256k1_compress_public_key(ByVal uncompressed_hex As String) As String
    ' Converte chave pública descomprimida (04+x+y) para formato comprimido (02/03+x)
    If Len(uncompressed_hex) <> 130 Or Left$(uncompressed_hex, 2) <> "04" Then
        secp256k1_compress_public_key = "" : Exit Function
    End If

    Dim x_hex As String, y_hex As String
    x_hex = Mid$(uncompressed_hex, 3, 64)
    y_hex = Mid$(uncompressed_hex, 67, 64)

    secp256k1_compress_public_key = secp256k1_point_compress(x_hex, y_hex)
End Function

Public Function secp256k1_uncompress_public_key(ByVal compressed_hex As String) As String
    ' Converte chave pública comprimida (02/03+x) para formato descomprimido (04+x+y)
    Dim coords As String
    coords = secp256k1_point_decompress(compressed_hex)

    If coords = "" Then
        secp256k1_uncompress_public_key = "" : Exit Function
    End If

    Dim comma_pos As Long
    comma_pos = InStr(coords, ",")
    If comma_pos = 0 Then
        secp256k1_uncompress_public_key = "" : Exit Function
    End If

    Dim x_hex As String, y_hex As String
    x_hex = Left$(coords, comma_pos - 1)
    y_hex = Mid$(coords, comma_pos + 1)

    ' Garantir 64 caracteres para x e y
    Do While Len(x_hex) < 64 : x_hex = "0" & x_hex : Loop
    Do While Len(y_hex) < 64 : y_hex = "0" & y_hex : Loop

    secp256k1_uncompress_public_key = "04" & x_hex & y_hex
End Function

' =============================================================================
' ASSINATURA DIGITAL ECDSA (RFC 6979 + BIP 62)
' =============================================================================

Public Function secp256k1_sign(ByVal message_hash As String, ByVal private_key_hex As String) As String
    ' Assina um hash de mensagem usando ECDSA determinístico (RFC 6979)
    ' Parâmetros:
    '   message_hash: Hash SHA-256 da mensagem (64 caracteres hex)
    '   private_key_hex: Chave privada (64 caracteres hex)
    ' Retorna: Assinatura em formato DER hexadecimal ou string vazia se erro

    ' Validar entradas
    last_error = SECP256K1_OK
    If Len(message_hash) <> 64 Then
        last_error = SECP256K1_ERROR_INVALID_HASH
        secp256k1_sign = "": Exit Function
    End If
    If Not secp256k1_is_hex_string(message_hash) Then
        last_error = SECP256K1_ERROR_INVALID_HASH
        secp256k1_sign = "": Exit Function
    End If
    If Len(private_key_hex) <> 64 Then
        last_error = SECP256K1_ERROR_INVALID_PRIVATE_KEY
        secp256k1_sign = "": Exit Function
    End If
    If Not secp256k1_validate_private_key(private_key_hex) Then
        last_error = SECP256K1_ERROR_INVALID_PRIVATE_KEY
        secp256k1_sign = "": Exit Function
    End If

    On Error GoTo SigningFailed

    Dim sig As ECDSA_SIGNATURE
    sig = ecdsa_sign_bitcoin_core(message_hash, private_key_hex, ctx)

    secp256k1_sign = ecdsa_signature_to_der(sig)
    On Error GoTo 0
    Exit Function

SigningFailed:
    last_error = SECP256K1_ERROR_COMPUTATION_FAILED
    secp256k1_sign = ""
    Err.Clear
    On Error GoTo 0
End Function

Public Function secp256k1_verify(ByVal message_hash As String, ByVal signature_der As String, ByVal public_key_compressed As String) As Boolean
    ' Verifica uma assinatura ECDSA
    ' Parâmetros:
    '   message_hash: Hash SHA-256 da mensagem (64 caracteres hex)
    '   signature_der: Assinatura em formato DER hexadecimal
    '   public_key_compressed: Chave pública comprimida (66 caracteres hex)
    ' Retorna: True se a assinatura for válida
    
    ' Validar entradas
    last_error = SECP256K1_OK
    If Len(message_hash) <> 64 Then
        last_error = SECP256K1_ERROR_INVALID_HASH
        secp256k1_verify = False: Exit Function
    End If
    If Not secp256k1_is_hex_string(message_hash) Then
        last_error = SECP256K1_ERROR_INVALID_HASH
        secp256k1_verify = False: Exit Function
    End If
    If Len(signature_der) < 8 Then
        last_error = SECP256K1_ERROR_INVALID_SIGNATURE
        secp256k1_verify = False: Exit Function
    End If
    If Not secp256k1_validate_public_key(public_key_compressed) Then
        last_error = SECP256K1_ERROR_INVALID_PUBLIC_KEY
        secp256k1_verify = False: Exit Function
    End If
    
    Dim sig As ECDSA_SIGNATURE
    If Not ecdsa_signature_from_der(sig, signature_der) Then
        secp256k1_verify = False
        Exit Function
    End If

    Dim public_key As EC_POINT
    public_key = ec_point_decompress(public_key_compressed, ctx)
    If Not secp256k1_validate_affine_point(public_key) Then
        last_error = SECP256K1_ERROR_INVALID_PUBLIC_KEY
        secp256k1_verify = False
        Exit Function
    End If

    secp256k1_verify = ecdsa_verify_bitcoin_core(message_hash, sig, public_key, ctx)
End Function

' =============================================================================
' UTILITÁRIOS DE MANIPULAÇÃO DE PONTOS DA CURVA ELÍPTICA
' =============================================================================

Public Function secp256k1_point_compress(ByVal x_hex As String, ByVal y_hex As String) As String
    ' Comprime um ponto da curva elíptica a partir das coordenadas x,y
    Dim point As EC_POINT
    point = ec_point_new()
    
    Dim x As BIGNUM_TYPE, y As BIGNUM_TYPE
    x = BN_hex2bn(x_hex)
    y = BN_hex2bn(y_hex)
    
    Call ec_point_set_affine(point, x, y)
    
    If Not ec_point_is_on_curve(point, ctx) Then
        secp256k1_point_compress = ""
        Exit Function
    End If
    
    secp256k1_point_compress = ec_point_compress(point, ctx)
End Function

Private Function secp256k1_validate_affine_point(ByRef point As EC_POINT) As Boolean
    ' Valida ponto descomprimido garantindo que pertence ao subgrupo gerado por G
    If point.infinity Then Exit Function

    If Not ec_point_is_on_curve(point, ctx) Then Exit Function

    Dim subgroup_order As BIGNUM_TYPE
    If ctx.n.top = 0 Then
        subgroup_order = BN_hex2bn("FFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFEBAAEDCE6AF48A03BBFD25E8CD0364141")
    Else
        subgroup_order = ctx.n
    End If

    Dim n_point As EC_POINT
    n_point = ec_point_new()

    Dim mul_succeeded As Boolean
    mul_succeeded = ec_point_mul(n_point, subgroup_order, point, ctx)
    If Not mul_succeeded Then
        secp256k1_validate_affine_point = False
        Exit Function
    End If
    If Not n_point.infinity Then Exit Function

    secp256k1_validate_affine_point = True
End Function

Public Function secp256k1_point_decompress(ByVal compressed_hex As String) As String
    ' Descomprime um ponto da curva elíptica retornando coordenadas x,y
    Dim point As EC_POINT
    last_error = SECP256K1_OK
    point = ec_point_decompress(compressed_hex, ctx)

    If Not secp256k1_validate_affine_point(point) Then
        last_error = SECP256K1_ERROR_POINT_NOT_ON_CURVE
        secp256k1_point_decompress = ""
        Exit Function
    End If

    Dim x As BIGNUM_TYPE, y As BIGNUM_TYPE
    x = BN_new(): y = BN_new()

    Call ec_point_get_affine(point, x, y, ctx)
    
    secp256k1_point_decompress = BN_bn2hex(x) & "," & BN_bn2hex(y)
End Function

' =============================================================================
' OPERAÇÕES ARITMÉTICAS COM PONTOS DA CURVA ELÍPTICA
' =============================================================================

Public Function secp256k1_point_add(ByVal point1_compressed As String, ByVal point2_compressed As String) As String
    ' Realiza adição de dois pontos da curva elíptica: P1 + P2
    Dim p1 As EC_POINT, p2 As EC_POINT, result As EC_POINT
    last_error = SECP256K1_OK
    p1 = ec_point_decompress(point1_compressed, ctx)
    p2 = ec_point_decompress(point2_compressed, ctx)
    result = ec_point_new()

    If Not secp256k1_validate_affine_point(p1) Then
        last_error = SECP256K1_ERROR_POINT_NOT_ON_CURVE
        secp256k1_point_add = ""
        Exit Function
    End If

    If Not secp256k1_validate_affine_point(p2) Then
        last_error = SECP256K1_ERROR_POINT_NOT_ON_CURVE
        secp256k1_point_add = ""
        Exit Function
    End If

    If Not ec_point_add(result, p1, p2, ctx) Then
        secp256k1_point_add = ""
        Exit Function
    End If

    If result.infinity Then
        secp256k1_point_add = "00"
        Exit Function
    End If

    If Not secp256k1_validate_affine_point(result) Then
        last_error = SECP256K1_ERROR_POINT_NOT_ON_CURVE
        secp256k1_point_add = ""
        Exit Function
    End If

    secp256k1_point_add = ec_point_compress(result, ctx)
End Function

Public Function secp256k1_point_multiply(ByVal scalar_hex As String, ByVal point_compressed As String) As String
    ' Realiza multiplicação escalar de um ponto: k * P
    Dim scalar As BIGNUM_TYPE, point As EC_POINT, result As EC_POINT
    Dim zero As BIGNUM_TYPE, curve_order As BIGNUM_TYPE

    last_error = SECP256K1_OK

    If Len(scalar_hex) <> 64 Or Not secp256k1_is_hex_string(scalar_hex) Then
        last_error = SECP256K1_ERROR_INVALID_PRIVATE_KEY
        secp256k1_point_multiply = ""
        Exit Function
    End If

    scalar = BN_hex2bn(scalar_hex)
    zero = BN_new()

    If ctx.n.top = 0 Then
        curve_order = BN_hex2bn("FFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFEBAAEDCE6AF48A03BBFD25E8CD0364141")
    Else
        curve_order = ctx.n
    End If

    If BN_ucmp(scalar, zero) <= 0 Or BN_ucmp(scalar, curve_order) >= 0 Then
        last_error = SECP256K1_ERROR_INVALID_PRIVATE_KEY
        secp256k1_point_multiply = ""
        Exit Function
    End If

    point = ec_point_decompress(point_compressed, ctx)
    result = ec_point_new()

    If Not secp256k1_validate_affine_point(point) Then
        last_error = SECP256K1_ERROR_POINT_NOT_ON_CURVE
        secp256k1_point_multiply = ""
        Exit Function
    End If

    If Not ec_point_mul_ultimate(result, scalar, point, ctx) Then
        last_error = SECP256K1_ERROR_COMPUTATION_FAILED
        secp256k1_point_multiply = ""
        Exit Function
    End If

    If result.infinity Then
        secp256k1_point_multiply = "00"
        Exit Function
    End If

    If Not secp256k1_validate_affine_point(result) Then
        last_error = SECP256K1_ERROR_POINT_NOT_ON_CURVE
        secp256k1_point_multiply = ""
        Exit Function
    End If

    secp256k1_point_multiply = ec_point_compress(result, ctx)
End Function

Public Function secp256k1_generator_multiply(ByVal scalar_hex As String) As String
    ' Multiplica o ponto gerador da curva por um escalar: k * G
    Dim scalar As BIGNUM_TYPE, result As EC_POINT
    Dim zero As BIGNUM_TYPE, curve_order As BIGNUM_TYPE

    last_error = SECP256K1_OK

    If Len(scalar_hex) <> 64 Or Not secp256k1_is_hex_string(scalar_hex) Then
        last_error = SECP256K1_ERROR_INVALID_PRIVATE_KEY
        secp256k1_generator_multiply = ""
        Exit Function
    End If

    scalar = BN_hex2bn(scalar_hex)
    zero = BN_new()

    If ctx.n.top = 0 Then
        curve_order = BN_hex2bn("FFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFEBAAEDCE6AF48A03BBFD25E8CD0364141")
    Else
        curve_order = ctx.n
    End If

    If BN_ucmp(scalar, zero) <= 0 Or BN_ucmp(scalar, curve_order) >= 0 Then
        last_error = SECP256K1_ERROR_INVALID_PRIVATE_KEY
        secp256k1_generator_multiply = ""
        Exit Function
    End If

    result = ec_point_new()

    If Not ec_point_mul_generator(result, scalar, ctx) Then
        secp256k1_generator_multiply = ""
        Exit Function
    End If
    
    If result.infinity Then
        secp256k1_generator_multiply = "00"
        Exit Function
    End If
    
    secp256k1_generator_multiply = ec_point_compress(result, ctx)
End Function

' =============================================================================
' VALIDAÇÃO DE CHAVES CRIPTOGRÁFICAS
' =============================================================================

Public Function secp256k1_validate_private_key(ByVal private_key_hex As String) As Boolean
    ' Valida se uma chave privada está no intervalo válido [1, n-1]
    
    ' Validar formato da entrada
    If Len(private_key_hex) <> 64 Then secp256k1_validate_private_key = False: Exit Function
    
    Dim priv_key As BIGNUM_TYPE, zero As BIGNUM_TYPE, curve_order As BIGNUM_TYPE
    priv_key = BN_hex2bn(private_key_hex)
    zero = BN_new()
    
    ' Usar ordem da curva secp256k1 diretamente se contexto não inicializado
    If ctx.n.top = 0 Then
        curve_order = BN_hex2bn("FFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFEBAAEDCE6AF48A03BBFD25E8CD0364141")
    Else
        curve_order = ctx.n
    End If
    
    secp256k1_validate_private_key = (BN_ucmp(priv_key, zero) > 0 And BN_ucmp(priv_key, curve_order) < 0)
End Function

Public Function secp256k1_validate_public_key(ByVal public_key_compressed As String) As Boolean
    ' Validação rigorosa de chave pública compatível com Bitcoin Core
    
    ' Verificar formato
    If Len(public_key_compressed) <> 66 Then secp256k1_validate_public_key = False: Exit Function
    
    Dim prefix As String: prefix = left$(public_key_compressed, 2)
    If prefix <> "02" And prefix <> "03" Then secp256k1_validate_public_key = False: Exit Function
    
    ' Extrair coordenada x
    Dim x_hex As String: x_hex = mid$(public_key_compressed, 3)
    Dim x As BIGNUM_TYPE: x = BN_hex2bn(x_hex)
    
    ' Validar se x está no campo [0, p-1]
    If BN_ucmp(x, ctx.p) >= 0 Then secp256k1_validate_public_key = False: Exit Function
    
    ' Descomprimir e validar
    Dim point As EC_POINT: point = ec_point_decompress(public_key_compressed, ctx)

    If Not secp256k1_validate_affine_point(point) Then
        secp256k1_validate_public_key = False
        Exit Function
    End If

    secp256k1_validate_public_key = True
End Function

' =============================================================================
' INFORMAÇÕES DO CONTEXTO SECP256K1
' =============================================================================

Public Function secp256k1_get_field_prime() As String
    ' Retorna o número primo do campo finito da curva secp256k1
    secp256k1_get_field_prime = BN_bn2hex(ctx.p)
End Function

Public Function secp256k1_get_curve_order() As String
    ' Retorna a ordem da curva secp256k1 (número de pontos)
    secp256k1_get_curve_order = BN_bn2hex(ctx.n)
End Function

Public Function secp256k1_get_generator() As String
    ' Retorna o ponto gerador da curva secp256k1 em formato comprimido
    secp256k1_get_generator = ec_point_compress(ctx.g, ctx)
End Function

' =============================================================================
' UTILITÁRIOS DE HASH (FUNÇÃO DEMONSTRATIVA)
' =============================================================================

Public Function secp256k1_hash_sha256(ByVal message As String) As String
    ' Hash SHA-256 usando implementação completa e segura
    ' Integração com módulo SHA256_Hash.bas
    secp256k1_hash_sha256 = SHA256_VBA.SHA256_String(message)
End Function

' =============================================================================
' FUNÇÃO DE DEMONSTRAÇÃO DO SISTEMA SECP256K1
' =============================================================================

Public Sub secp256k1_demo()
    Debug.Print "=== DEMONSTRAÇÃO SECP256K1 ==="
    
    ' Inicializar contexto
    If Not secp256k1_init() Then
        Debug.Print "ERRO: Falha na inicialização"
        Exit Sub
    End If
    Debug.Print "Contexto inicializado com sucesso"
    
    ' Gerar par de chaves
    Dim keypair As ECDSA_KEYPAIR
    keypair = secp256k1_generate_keypair()
    Debug.Print "Chave privada: ", BN_bn2hex(keypair.private_key)
    Debug.Print "Chave pública: ", ec_point_compress(keypair.public_key, ctx)
    
    ' Assinar mensagem
    Dim message As String, message_hash As String, signature As String
    message = "Olá, secp256k1!"
    message_hash = secp256k1_hash_sha256(message)
    signature = secp256k1_sign(message_hash, BN_bn2hex(keypair.private_key))
    
    If signature = "" Then
        Debug.Print "ERRO: Falha na assinatura"
        Exit Sub
    End If
    
    Debug.Print "Mensagem: ", message
    Debug.Print "Hash: ", message_hash
    Debug.Print "Assinatura: ", signature
    
    ' Verificar assinatura
    Dim public_key_compressed As String, is_valid As Boolean
    public_key_compressed = ec_point_compress(keypair.public_key, ctx)
    is_valid = secp256k1_verify(message_hash, signature, public_key_compressed)
    
    Debug.Print "Assinatura válida: ", is_valid
    
    Debug.Print "=== DEMONSTRAÇÃO CONCLUÍDA ==="
End Sub

Public Sub secp256k1_demo_bitcoin_address()
    Debug.Print "=== DEMONSTRAÇÃO ENDEREÇO BITCOIN ==="
    
    Call secp256k1_init
    
    ' Gerar chaves
    Dim keypair As ECDSA_KEYPAIR
    keypair = secp256k1_generate_keypair()
    
    Dim private_hex As String, public_compressed As String
    private_hex = BN_bn2hex(keypair.private_key)
    public_compressed = ec_point_compress(keypair.public_key, ctx)
    
    Debug.Print "Chave privada: ", private_hex
    Debug.Print "Chave pública comprimida: ", public_compressed
    
    ' Gerar endereço Bitcoin
    Dim address As String
    address = generate_bitcoin_address(public_compressed, True) ' Legacy P2PKH
    Debug.Print "Endereço Bitcoin: ", address
    
    Debug.Print "=== DEMONSTRAÇÃO ENDEREÇO CONCLUÍDA ==="
End Sub

Public Sub secp256k1_demo_key_import()
    Debug.Print "=== DEMONSTRAÇÃO IMPORTAÇÃO DE CHAVE ==="
    
    Call secp256k1_init
    
    ' Chave privada conhecida (exemplo)
    Dim known_private As String
    known_private = "C9AFA9D845BA75166B5C215767B1D6934E50C3DB36E89B127B8A622B120F6721"
    
    ' Validar chave privada
    If Not secp256k1_validate_private_key(known_private) Then
        Debug.Print "ERRO: Chave privada inválida"
        Exit Sub
    End If
    
    ' Derivar chave pública
    Dim public_key As String
    public_key = secp256k1_public_key_from_private(known_private, True)
    
    If public_key = "" Then
        Debug.Print "ERRO: Falha na derivação da chave pública"
        Exit Sub
    End If
    
    ' Validar chave pública
    If Not secp256k1_validate_public_key(public_key) Then
        Debug.Print "ERRO: Chave pública inválida"
        Exit Sub
    End If
    
    Debug.Print "Chave privada importada: ", known_private
    Debug.Print "Chave pública derivada: ", public_key
    Debug.Print "Validação: APROVADA"
    
    Debug.Print "=== DEMONSTRAÇÃO IMPORTAÇÃO CONCLUÍDA ==="
End Sub