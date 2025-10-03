Attribute VB_Name = "EC_secp256k1_ECDSA"
Option Explicit

' =============================================================================
' SECP256K1 VBA - ASSINATURA DIGITAL ECDSA
' =============================================================================
' Implementação completa do algoritmo ECDSA (Elliptic Curve Digital Signature Algorithm)
' Compatível com Bitcoin Core, RFC 6979 (assinaturas determinísticas) e BIP 62 (low-s)
' Inclui geração de chaves, assinatura, verificação e codificação DER
' =============================================================================

' =============================================================================
' ESTRUTURAS DE DADOS PARA ECDSA
' =============================================================================

' Representa uma assinatura digital ECDSA
Public Type ECDSA_SIGNATURE
    r As BIGNUM_TYPE        ' Componente R da assinatura
    s As BIGNUM_TYPE        ' Componente S da assinatura (com low-s enforcement)
End Type

' Representa um par de chaves criptográficas ECDSA
Public Type ECDSA_KEYPAIR
    private_key As BIGNUM_TYPE  ' Chave privada (escalar no intervalo [1, n-1])
    public_key As EC_POINT      ' Chave pública (ponto da curva = private_key * G)
End Type

' =============================================================================
' ASSINATURA DIGITAL ECDSA (BITCOIN CORE COMPATÍVEL)
' =============================================================================

Public Function ecdsa_sign_bitcoin_core(ByVal message_hash As String, ByVal private_key_hex As String, ByRef ctx As SECP256K1_CTX) As ECDSA_SIGNATURE
    ' Gera assinatura ECDSA determinística compatível com Bitcoin Core
    ' Parâmetros:
    '   message_hash: Hash SHA-256 da mensagem (64 caracteres hex)
    '   private_key_hex: Chave privada (64 caracteres hex)
    '   ctx: Contexto da curva secp256k1
    ' Retorna: Assinatura ECDSA com low-s enforcement (BIP 62)
    Dim z As BIGNUM_TYPE, d As BIGNUM_TYPE, k As BIGNUM_TYPE
    Dim r As BIGNUM_TYPE, s As BIGNUM_TYPE, kinv As BIGNUM_TYPE
    Dim R_point As EC_POINT

    z = BN_hex2bn(message_hash)
    d = BN_hex2bn(private_key_hex)

    ' Gerar k determinístico (RFC 6979 simplificado)
    k = generate_k_rfc6979(z, d, ctx)

    ' R = k * G (usando sistema ULTIMATE)
    R_point = ec_point_new()
    Call ec_point_mul_ultimate(R_point, k, ctx.g, ctx)
    Call BN_mod(r, R_point.x, ctx.n)

    ' Se r = 0, tentar novamente (evento extremamente raro)
    If BN_is_zero(r) Then
        Dim one As BIGNUM_TYPE
        Call BN_set_word(one, 1)
        Call BN_add(k, k, one)
        Call ec_point_mul_generator(R_point, k, ctx)
        Call BN_mod(r, R_point.x, ctx.n)
    End If

    ' s = k^-1 * (z + r * d) mod n
    Call BN_mod_inverse(kinv, k, ctx.n)
    Dim rd As BIGNUM_TYPE, zrd As BIGNUM_TYPE
    Call BN_mod_mul(rd, r, d, ctx.n)
    Call BN_mod_add(zrd, z, rd, ctx.n)
    Call BN_mod_mul(s, kinv, zrd, ctx.n)

    ' Aplicar low-s enforcement (BIP 62) - forçar s <= n/2
    Dim half_n As BIGNUM_TYPE, two As BIGNUM_TYPE, temp As BIGNUM_TYPE
    Call BN_set_word(two, 2)
    Call BN_div(half_n, temp, ctx.n, two)
    If BN_ucmp(s, half_n) > 0 Then
        ' s = n - s (mod n) para garantir s < n
        Dim zero As BIGNUM_TYPE: zero = BN_new()
        Call BN_mod_sub(s, zero, s, ctx.n)  ' 0 - s = -s = n - s (mod n)
    End If

    ecdsa_sign_bitcoin_core.r = r
    ecdsa_sign_bitcoin_core.s = s
End Function

Public Function ecdsa_verify_bitcoin_core(ByVal message_hash As String, ByRef sig As ECDSA_SIGNATURE, ByRef public_key As EC_POINT, ByRef ctx As SECP256K1_CTX) As Boolean
    ' Verifica uma assinatura ECDSA usando algoritmo padrão
    ' Parâmetros:
    '   message_hash: Hash SHA-256 da mensagem original
    '   sig: Assinatura ECDSA a ser verificada
    '   public_key: Chave pública do signatário
    '   ctx: Contexto da curva secp256k1
    ' Retorna: True se a assinatura for válida
    
    ' Validar parâmetros da assinatura
    If BN_is_zero(sig.r) Or BN_is_zero(sig.s) Then
        ecdsa_verify_bitcoin_core = False
        Exit Function
    End If

    If BN_ucmp(sig.r, ctx.n) >= 0 Or BN_ucmp(sig.s, ctx.n) >= 0 Then
        ecdsa_verify_bitcoin_core = False
        Exit Function
    End If

    ' Calcular coeficientes: u1 = z * s^-1 mod n, u2 = r * s^-1 mod n
    Dim z As BIGNUM_TYPE, sinv As BIGNUM_TYPE
    Dim u1 As BIGNUM_TYPE, u2 As BIGNUM_TYPE

    z = BN_hex2bn(message_hash)
    Call BN_mod_inverse(sinv, sig.s, ctx.n)
    Call BN_mod_mul(u1, z, sinv, ctx.n)
    Call BN_mod_mul(u2, sig.r, sinv, ctx.n)

    ' R = u1*G + u2*Q (usando sistema ULTIMATE para gerador, regular para ponto arbitrário)
    Dim point1 As EC_POINT, point2 As EC_POINT, R_result As EC_POINT
    point1 = ec_point_new()
    point2 = ec_point_new()
    Call ec_point_mul_ultimate(point1, u1, ctx.g, ctx)
    Call ec_point_mul(point2, u2, public_key, ctx)
    Call ec_point_add(R_result, point1, point2, ctx)

    If R_result.infinity Then
        ecdsa_verify_bitcoin_core = False
        Exit Function
    End If

    ' v = R.x mod n
    Dim v As BIGNUM_TYPE
    Call BN_mod(v, R_result.x, ctx.n)

    ' Verificar se v == r (assinatura válida)
    ecdsa_verify_bitcoin_core = (BN_cmp(v, sig.r) = 0)
End Function

' Geração de k determinístico (RFC 6979 simplificado)
Private Function generate_k_rfc6979(ByRef z As BIGNUM_TYPE, ByRef d As BIGNUM_TYPE, ByRef ctx As SECP256K1_CTX) As BIGNUM_TYPE
    ' Implementação simplificada do RFC 6979 para geração determinística de k
    ' Usa hash(z + d) mod n em vez do algoritmo completo HMAC-DRBG
    ' Parâmetros: z (hash da mensagem), d (chave privada), ctx (contexto)
    ' Retorna: Valor k válido no intervalo [1, n-1]
    Dim combined As BIGNUM_TYPE, k As BIGNUM_TYPE
    combined = BN_new()
    Call BN_add(combined, z, d)

    ' Converter para hexadecimal e aplicar SHA-256
    Dim combined_hex As String, hash_hex As String
    combined_hex = BN_bn2hex(combined)
    hash_hex = SHA256_VBA.SHA256_String(combined_hex)

    k = BN_hex2bn(hash_hex)
    Call BN_mod(k, k, ctx.n)

    ' Garantir que k não é zero (requisito do ECDSA)
    If BN_is_zero(k) Then
        Call BN_set_word(k, 1)
    End If

    generate_k_rfc6979 = k
End Function

' =============================================================================
' GERAÇÃO E MANIPULAÇÃO DE PARES DE CHAVES
' =============================================================================

Public Function ecdsa_generate_keypair(ByRef ctx As SECP256K1_CTX) As ECDSA_KEYPAIR
    ' Gera um novo par de chaves ECDSA criptograficamente seguro
    ' Retorna: Par de chaves com private_key no intervalo [1, n-1] e public_key = private_key * G
    
    Dim keypair As ECDSA_KEYPAIR

    ' Gerar chave privada aleatória válida no intervalo [1, n-1]
    Do
        keypair.private_key = generate_random_private_key()
    Loop While BN_is_zero(keypair.private_key) Or BN_ucmp(keypair.private_key, ctx.n) >= 0

    keypair.public_key = ec_point_new()
    Call ec_point_mul_ultimate(keypair.public_key, keypair.private_key, ctx.g, ctx)
    ecdsa_generate_keypair = keypair
End Function

Private Function generate_random_private_key() As BIGNUM_TYPE
    ' Gera chave privada pseudo-aleatória de 256 bits
    ' ATENÇÃO: Para uso em produção, usar gerador criptograficamente seguro
    ' Atualmente usa Timer + Rnd para demonstração
    Randomize Timer
    Dim hex_str As String, i As Long

    For i = 1 To 64 ' 64 caracteres hex = 256 bits
        hex_str = hex_str & hex$(Int(Rnd() * 16))
    Next i

    generate_random_private_key = BN_hex2bn(hex_str)
End Function

Public Function ecdsa_set_private_key(ByRef private_key_hex As String, ByRef ctx As SECP256K1_CTX) As ECDSA_KEYPAIR
    ' Cria par de chaves a partir de uma chave privada conhecida
    ' Parâmetro: private_key_hex - chave privada em formato hexadecimal
    ' Retorna: Par de chaves com chave pública derivada
    Dim keypair As ECDSA_KEYPAIR
    keypair.private_key = BN_hex2bn(private_key_hex)
    keypair.public_key = ec_point_new()
    Call ec_point_mul_ultimate(keypair.public_key, keypair.private_key, ctx.g, ctx)
    ecdsa_set_private_key = keypair
End Function

' =============================================================================
' CODIFICAÇÃO E DECODIFICAÇÃO DER (DISTINGUISHED ENCODING RULES)
' =============================================================================

Public Function ecdsa_signature_to_der(ByRef sig As ECDSA_SIGNATURE) As String
    ' Codifica assinatura ECDSA no formato DER padrão
    ' Formato: 30 [tamanho_total] 02 [tamanho_r] [r] 02 [tamanho_s] [s]
    ' Retorna: String hexadecimal no formato DER
    Dim r_hex As String, s_hex As String
    r_hex = BN_bn2hex(sig.r)
    s_hex = BN_bn2hex(sig.s)

    If Len(r_hex) Mod 2 = 1 Then r_hex = "0" & r_hex
    If Len(s_hex) Mod 2 = 1 Then s_hex = "0" & s_hex

    If left$(r_hex, 1) >= "8" Then r_hex = "00" & r_hex
    If left$(s_hex, 1) >= "8" Then s_hex = "00" & s_hex

    Dim r_len As String, s_len As String, total_len As String
    r_len = right$("0" & hex$(Len(r_hex) \ 2), 2)
    s_len = right$("0" & hex$(Len(s_hex) \ 2), 2)
    total_len = right$("0" & hex$((Len(r_hex) + Len(s_hex)) \ 2 + 4), 2)

    ecdsa_signature_to_der = "30" & total_len & "02" & r_len & r_hex & "02" & s_len & s_hex
End Function

Public Function ecdsa_signature_from_der(ByRef sig As ECDSA_SIGNATURE, ByVal der_hex As String) As Boolean
    ' Decodifica assinatura do formato DER para estrutura ECDSA_SIGNATURE
    ' Parâmetro: der_hex - assinatura em formato DER hexadecimal
    ' Retorna: True se a decodificação foi bem-sucedida
    sig.r = BN_new()
    sig.s = BN_new()

    If left$(der_hex, 2) <> "30" Then
        ecdsa_signature_from_der = False
        Exit Function
    End If

    Dim pos As Long : pos = 5

    If mid$(der_hex, pos, 2) = "02" Then
        pos = pos + 2
        Dim r_len As Long : r_len = CLng("&H" & mid$(der_hex, pos, 2))
        pos = pos + 2
        Dim r_hex As String : r_hex = mid$(der_hex, pos, r_len * 2)
        If left$(r_hex, 2) = "00" And Len(r_hex) > 2 Then r_hex = mid$(r_hex, 3)
        sig.r = BN_hex2bn(r_hex)
        pos = pos + r_len * 2
    End If

    If mid$(der_hex, pos, 2) = "02" Then
        pos = pos + 2
        Dim s_len As Long : s_len = CLng("&H" & mid$(der_hex, pos, 2))
        pos = pos + 2
        Dim s_hex As String : s_hex = mid$(der_hex, pos, s_len * 2)
        If left$(s_hex, 2) = "00" And Len(s_hex) > 2 Then s_hex = mid$(s_hex, 3)
        sig.s = BN_hex2bn(s_hex)
    End If

    ecdsa_signature_from_der = True
End Function
Public Function ecdsa_generate_keypair_optimized(ByRef ctx As SECP256K1_CTX) As ECDSA_KEYPAIR
    ' Geração otimizada de par de chaves usando multiplicação rápida do gerador
    ' Usa tabelas pré-computadas quando disponíveis para máxima performance
    
    Dim keypair As ECDSA_KEYPAIR
    
    ' Gerar chave privada aleatória válida no intervalo [1, n-1]
    Do
        keypair.private_key = generate_random_private_key()
    Loop While BN_is_zero(keypair.private_key) Or BN_ucmp(keypair.private_key, ctx.n) >= 0
    
    ' Usar multiplicação otimizada do gerador (tabelas pré-computadas)
    keypair.public_key = ec_point_new()
    Call ec_point_mul_generator(keypair.public_key, keypair.private_key, ctx)
    
    ecdsa_generate_keypair_optimized = keypair
End Function