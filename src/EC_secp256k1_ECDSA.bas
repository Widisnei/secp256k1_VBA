Attribute VB_Name = "EC_secp256k1_ECDSA"
Option Explicit

#If VBA7 Then
    Private Declare PtrSafe Function BCryptGenRandom Lib "bcrypt.dll" ( _
        ByVal hAlgorithm As LongPtr, _
        ByRef pbBuffer As Byte, _
        ByVal cbBuffer As Long, _
        ByVal dwFlags As Long) As Long
#Else
    Private Declare Function BCryptGenRandom Lib "bcrypt.dll" ( _
        ByVal hAlgorithm As Long, _
        ByRef pbBuffer As Byte, _
        ByVal cbBuffer As Long, _
        ByVal dwFlags As Long) As Long
#End If

Private Const BCRYPT_USE_SYSTEM_PREFERRED_RNG As Long = &H2&
Private Const STATUS_SUCCESS As Long = 0&

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

' Variáveis de teste expostas para instrumentação controlada em cenários de regressão.
' Não afetam o comportamento normal quando mantidas em zero.
Public RFC6979_Test_RejectNextCandidates As Long
Public RFC6979_Test_Rejections As Long
Public RFC6979_Test_ForceRetryCount As Long

Private Const RFC6979_HOLEN As Long = 32
Private Const RFC6979_ROLEN As Long = 32
Private Const ERR_KEYPAIR_POINT_MUL_FAILED As Long = vbObjectError + &H1102&

Private Type RFC6979_STATE
    K() As Byte
    V() As Byte
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
    Dim rd As BIGNUM_TYPE, zrd As BIGNUM_TYPE
    Dim R_point As EC_POINT
    Dim rng_state As RFC6979_STATE
    Dim success As Boolean

    z = BN_hex2bn(message_hash)
    d = BN_hex2bn(private_key_hex)

    Call rfc6979_initialize_state(rng_state, z, d, ctx)

    R_point = ec_point_new()

    Do
        success = True
        k = rfc6979_generate_candidate(rng_state, ctx)

        Call ec_point_mul_ultimate(R_point, k, ctx.g, ctx)
        Call BN_mod(r, R_point.x, ctx.n)

        If BN_is_zero(r) Then
            Call rfc6979_state_reject(rng_state)
            success = False
        ElseIf RFC6979_Test_ForceRetryCount > 0 Then
            RFC6979_Test_ForceRetryCount = RFC6979_Test_ForceRetryCount - 1
            Call rfc6979_state_reject(rng_state)
            success = False
        ElseIf Not BN_mod_inverse(kinv, k, ctx.n) Then
            Call rfc6979_state_reject(rng_state)
            success = False
        Else
            Call BN_mod_mul(rd, r, d, ctx.n)
            Call BN_mod_add(zrd, z, rd, ctx.n)
            Call BN_mod_mul(s, kinv, zrd, ctx.n)

            If BN_is_zero(s) Then
                Call rfc6979_state_reject(rng_state)
                success = False
            End If
        End If

        If success Then Exit Do
    Loop

    ' Aplicar low-s enforcement (BIP 62) - forçar s <= n/2
    Dim half_n As BIGNUM_TYPE, two As BIGNUM_TYPE, temp As BIGNUM_TYPE
    half_n = BN_new()
    two = BN_new()
    temp = BN_new()
    Call BN_set_word(two, 2)
    Call BN_div(half_n, temp, ctx.n, two)
    Call BN_free(temp)
    If BN_ucmp(s, half_n) > 0 Then
        ' s = n - s (mod n) para garantir s < n
        Dim zero As BIGNUM_TYPE: zero = BN_new()
        Call BN_mod_sub(s, zero, s, ctx.n)  ' 0 - s = -s = n - s (mod n)
        Call BN_free(zero)
    End If
    Call BN_free(half_n)
    Call BN_free(two)

    ecdsa_sign_bitcoin_core.r = r
    ecdsa_sign_bitcoin_core.s = s
End Function

Public Function ecdsa_signature_is_valid(ByRef sig As ECDSA_SIGNATURE, ByRef ctx As SECP256K1_CTX) As Boolean
    ' Helper compartilhado para validar componentes r e s de uma assinatura
    If BN_is_zero(sig.r) Or BN_is_zero(sig.s) Then
        ecdsa_signature_is_valid = False
        Exit Function
    End If

    If BN_ucmp(sig.r, ctx.n) >= 0 Or BN_ucmp(sig.s, ctx.n) >= 0 Then
        ecdsa_signature_is_valid = False
        Exit Function
    End If

    ecdsa_signature_is_valid = True
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
    If Not ecdsa_signature_is_valid(sig, ctx) Then
        ecdsa_verify_bitcoin_core = False
        Exit Function
    End If

    ' Calcular coeficientes: u1 = z * s^-1 mod n, u2 = r * s^-1 mod n
    Dim z As BIGNUM_TYPE, sinv As BIGNUM_TYPE
    Dim u1 As BIGNUM_TYPE, u2 As BIGNUM_TYPE

    z = BN_hex2bn(message_hash)
    If Not BN_mod_inverse(sinv, sig.s, ctx.n) Then
        ecdsa_verify_bitcoin_core = False
        Exit Function
    End If
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

Private Sub rfc6979_initialize_state(ByRef state As RFC6979_STATE, ByRef z As BIGNUM_TYPE, ByRef d As BIGNUM_TYPE, ByRef ctx As SECP256K1_CTX)
    Dim i As Long

    ReDim state.K(0 To RFC6979_HOLEN - 1)
    ReDim state.V(0 To RFC6979_HOLEN - 1)

    For i = 0 To RFC6979_HOLEN - 1
        state.K(i) = 0
        state.V(i) = &H1
    Next i

    Dim x_octets() As Byte, h1_octets() As Byte
    x_octets = bn_to_octets(d, RFC6979_ROLEN)
    h1_octets = bits2octets(z, ctx.n, RFC6979_ROLEN)

    Dim zeroByte(0 To 0) As Byte, oneByte(0 To 0) As Byte
    zeroByte(0) = 0
    oneByte(0) = 1

    Dim temp() As Byte
    temp = ByteArrayConcat(state.V, zeroByte)
    temp = ByteArrayConcat(temp, x_octets)
    temp = ByteArrayConcat(temp, h1_octets)
    state.K = SHA256_VBA.SHA256_HMAC(state.K, temp)

    state.V = SHA256_VBA.SHA256_HMAC(state.K, state.V)

    temp = ByteArrayConcat(state.V, oneByte)
    temp = ByteArrayConcat(temp, x_octets)
    temp = ByteArrayConcat(temp, h1_octets)
    state.K = SHA256_VBA.SHA256_HMAC(state.K, temp)

    state.V = SHA256_VBA.SHA256_HMAC(state.K, state.V)

    RFC6979_Test_Rejections = 0
End Sub

Private Sub rfc6979_state_reject(ByRef state As RFC6979_STATE)
    Dim zeroByte(0 To 0) As Byte
    zeroByte(0) = 0

    Dim temp() As Byte
    temp = ByteArrayConcat(state.V, zeroByte)
    state.K = SHA256_VBA.SHA256_HMAC(state.K, temp)
    state.V = SHA256_VBA.SHA256_HMAC(state.K, state.V)

    RFC6979_Test_Rejections = RFC6979_Test_Rejections + 1
End Sub

Private Function rfc6979_generate_candidate(ByRef state As RFC6979_STATE, ByRef ctx As SECP256K1_CTX) As BIGNUM_TYPE
    Dim candidate As BIGNUM_TYPE

    Do
        Dim T() As Byte, generated() As Byte
        Erase T
        Do While ByteArrayLength(T) < RFC6979_ROLEN
            state.V = SHA256_VBA.SHA256_HMAC(state.K, state.V)
            T = ByteArrayConcat(T, state.V)
        Loop

        generated = ByteArrayLeft(T, RFC6979_ROLEN)
        candidate = BN_bin2bn(generated, ByteArrayLength(generated))

        Dim forceReject As Boolean
        If RFC6979_Test_RejectNextCandidates > 0 Then
            RFC6979_Test_RejectNextCandidates = RFC6979_Test_RejectNextCandidates - 1
            forceReject = True
        End If

        If Not forceReject Then
            If (Not BN_is_zero(candidate)) And BN_ucmp(candidate, ctx.n) < 0 Then
                rfc6979_generate_candidate = candidate
                Exit Function
            End If
        End If

        Call rfc6979_state_reject(state)
    Loop
End Function

Private Function generate_k_rfc6979(ByRef z As BIGNUM_TYPE, ByRef d As BIGNUM_TYPE, ByRef ctx As SECP256K1_CTX) As BIGNUM_TYPE
    Dim state As RFC6979_STATE
    Call rfc6979_initialize_state(state, z, d, ctx)
    generate_k_rfc6979 = rfc6979_generate_candidate(state, ctx)
End Function

Private Function ByteArrayLength(ByRef arr() As Byte) As Long
    On Error GoTo EmptyArray
    ByteArrayLength = UBound(arr) - LBound(arr) + 1
    Exit Function
EmptyArray:
    ByteArrayLength = 0
End Function

Private Function ByteArrayConcat(ByRef a() As Byte, ByRef b() As Byte) As Byte()
    Dim lenA As Long, lenB As Long
    lenA = ByteArrayLength(a)
    lenB = ByteArrayLength(b)

    Dim total As Long
    total = lenA + lenB

    Dim result() As Byte
    If total <= 0 Then
        ByteArrayConcat = result
        Exit Function
    End If

    ReDim result(0 To total - 1)

    Dim baseA As Long, baseB As Long
    Dim i As Long

    If lenA > 0 Then
        baseA = LBoundSafe(a)
        For i = 0 To lenA - 1
            result(i) = a(baseA + i)
        Next i
    End If

    If lenB > 0 Then
        baseB = LBoundSafe(b)
        For i = 0 To lenB - 1
            result(lenA + i) = b(baseB + i)
        Next i
    End If

    ByteArrayConcat = result
End Function

Private Function ByteArrayLeft(ByRef arr() As Byte, ByVal count As Long) As Byte()
    Dim lengthBytes As Long
    lengthBytes = ByteArrayLength(arr)

    Dim result() As Byte
    If count <= 0 Or lengthBytes = 0 Then
        ByteArrayLeft = result
        Exit Function
    End If

    If count > lengthBytes Then count = lengthBytes

    ReDim result(0 To count - 1)

    Dim baseArr As Long
    baseArr = LBoundSafe(arr)

    Dim i As Long
    For i = 0 To count - 1
        result(i) = arr(baseArr + i)
    Next i

    ByteArrayLeft = result
End Function

Private Function LBoundSafe(ByRef arr() As Byte) As Long
    On Error GoTo EmptyArray
    LBoundSafe = LBound(arr)
    Exit Function
EmptyArray:
    LBoundSafe = 0
End Function

Private Function bn_to_octets(ByRef value As BIGNUM_TYPE, ByVal rolen As Long) As Byte()
    Dim raw() As Byte
    raw = BN_bn2bin(value)

    Dim lengthBytes As Long
    lengthBytes = ByteArrayLength(raw)

    Dim result() As Byte
    If rolen <= 0 Then
        bn_to_octets = result
        Exit Function
    End If

    ReDim result(0 To rolen - 1)

    Dim baseRaw As Long
    baseRaw = LBoundSafe(raw)

    Dim offset As Long
    offset = rolen - lengthBytes
    If offset < 0 Then offset = 0

    Dim i As Long
    If lengthBytes > 0 Then
        Dim startRaw As Long
        startRaw = lengthBytes - (rolen - offset)
        If startRaw < 0 Then startRaw = 0
        For i = 0 To rolen - offset - 1
            result(offset + i) = raw(baseRaw + startRaw + i)
        Next i
    End If

    bn_to_octets = result
End Function

Private Function bits2octets(ByRef z As BIGNUM_TYPE, ByRef n As BIGNUM_TYPE, ByVal rolen As Long) As Byte()
    Dim reduced As BIGNUM_TYPE
    reduced = BN_new()
    Call BN_mod(reduced, z, n)
    bits2octets = bn_to_octets(reduced, rolen)
End Function

' =============================================================================
' GERAÇÃO E MANIPULAÇÃO DE PARES DE CHAVES
' =============================================================================

Private Function fill_random_bytes(ByRef buffer() As Byte) As Boolean
    Dim length As Long
    length = UBound(buffer) - LBound(buffer) + 1

    If length <= 0 Then
        fill_random_bytes = True
        Exit Function
    End If

    Dim status As Long
    status = BCryptGenRandom(0, buffer(LBound(buffer)), length, BCRYPT_USE_SYSTEM_PREFERRED_RNG)
    fill_random_bytes = (status = STATUS_SUCCESS)
End Function

Public Function ecdsa_collect_secure_entropy(ByRef buffer() As Byte) As Boolean
    ' Expõe a rotina de coleta de entropia para outros módulos
    ecdsa_collect_secure_entropy = fill_random_bytes(buffer)
End Function

Public Function ecdsa_generate_keypair(ByRef ctx As SECP256K1_CTX) As ECDSA_KEYPAIR
    ' Gera um novo par de chaves ECDSA criptograficamente seguro
    ' Retorna: Par de chaves com private_key no intervalo [1, n-1] e public_key = private_key * G

    Dim keypair As ECDSA_KEYPAIR

    ' Gerar chave privada aleatória válida no intervalo [1, n-1]
    Do
        keypair.private_key = generate_random_private_key(ctx)
    Loop While BN_is_zero(keypair.private_key) Or BN_ucmp(keypair.private_key, ctx.n) >= 0

    keypair.public_key = ec_point_new()
    If Not ec_point_mul_ultimate(keypair.public_key, keypair.private_key, ctx.g, ctx) Then
        Err.Raise ERR_KEYPAIR_POINT_MUL_FAILED, "ecdsa_generate_keypair", _
                  "Falha ao calcular a chave pública durante a geração do par de chaves."
    End If
    ecdsa_generate_keypair = keypair
End Function

Private Function generate_random_private_key(ByRef ctx As SECP256K1_CTX) As BIGNUM_TYPE
    ' Gera chave privada aleatória criptograficamente segura no intervalo [1, n-1]
    Const MAX_ATTEMPTS As Long = 128
    Dim random_bytes(0 To 31) As Byte
    Dim attempt As Long

    For attempt = 1 To MAX_ATTEMPTS
        If Not fill_random_bytes(random_bytes) Then
            Err.Raise vbObjectError + &H1000&, "generate_random_private_key", _
                      "Falha ao coletar entropia criptográfica do sistema."
        End If

        Dim candidate As BIGNUM_TYPE
        candidate = BN_bin2bn(random_bytes, UBound(random_bytes) - LBound(random_bytes) + 1)

        If Not BN_is_zero(candidate) Then
            If BN_ucmp(candidate, ctx.n) < 0 Then
                generate_random_private_key = candidate
                Exit Function
            End If
        End If
    Next attempt

    Err.Raise vbObjectError + &H1001&, "generate_random_private_key", _
              "Não foi possível gerar uma chave privada válida após múltiplas tentativas."
End Function

Public Function ecdsa_set_private_key(ByRef private_key_hex As String, ByRef ctx As SECP256K1_CTX) As ECDSA_KEYPAIR
    ' Cria par de chaves a partir de uma chave privada conhecida
    ' Parâmetro: private_key_hex - chave privada em formato hexadecimal
    ' Retorna: Par de chaves com chave pública derivada
    Dim keypair As ECDSA_KEYPAIR
    keypair.private_key = BN_hex2bn(private_key_hex)

    Dim zero As BIGNUM_TYPE, curve_order As BIGNUM_TYPE
    zero = BN_new()

    If ctx.n.top = 0 Then
        curve_order = BN_hex2bn("FFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFEBAAEDCE6AF48A03BBFD25E8CD0364141")
    Else
        curve_order = ctx.n
    End If

    If BN_ucmp(keypair.private_key, zero) <= 0 Then
        Err.Raise vbObjectError + &H1002&, "ecdsa_set_private_key", _
                  "Chave privada inválida: deve ser maior que zero."
    End If

    If BN_ucmp(keypair.private_key, curve_order) >= 0 Then
        Err.Raise vbObjectError + &H1002&, "ecdsa_set_private_key", _
                  "Chave privada inválida: deve ser menor que a ordem da curva."
    End If

    keypair.public_key = ec_point_new()
    If Not ec_point_mul_ultimate(keypair.public_key, keypair.private_key, ctx.g, ctx) Then
        Err.Raise ERR_KEYPAIR_POINT_MUL_FAILED, "ecdsa_set_private_key", _
                  "Falha ao calcular a chave pública a partir da chave privada fornecida."
    End If
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
    ecdsa_signature_from_der = False

    sig.r = BN_new()
    sig.s = BN_new()

    If Len(der_hex) < 4 Or Len(der_hex) Mod 2 <> 0 Then Exit Function
    If left$(der_hex, 2) <> "30" Then Exit Function

    Dim pos As Long: pos = 3
    Dim length_byte As Long
    Dim seq_len As Long

    length_byte = CLng("&H" & Mid$(der_hex, pos, 2))
    pos = pos + 2

    If (length_byte And &H80) <> 0 Then
        Dim len_octets As Long
        len_octets = length_byte And &H7F
        If len_octets = 0 Or len_octets > 4 Then Exit Function
        If Len(der_hex) < pos + len_octets * 2 - 1 Then Exit Function
        seq_len = CLng("&H" & Mid$(der_hex, pos, len_octets * 2))
        pos = pos + len_octets * 2
    Else
        seq_len = length_byte
    End If

    If seq_len <= 0 Then Exit Function
    If Len(der_hex) - (pos - 1) <> seq_len * 2 Then Exit Function

    If pos > Len(der_hex) - 1 Then Exit Function
    If Mid$(der_hex, pos, 2) <> "02" Then Exit Function
    pos = pos + 2

    If pos > Len(der_hex) - 1 Then Exit Function
    Dim r_len As Long: r_len = CLng("&H" & Mid$(der_hex, pos, 2))
    pos = pos + 2
    If r_len <= 0 Then Exit Function
    If Len(der_hex) < pos + r_len * 2 - 1 Then Exit Function

    Dim r_hex_raw As String: r_hex_raw = Mid$(der_hex, pos, r_len * 2)
    Dim r_hex_normalized As String
    pos = pos + r_len * 2
    If Not normalize_der_integer(r_hex_raw, r_hex_normalized) Then Exit Function
    sig.r = BN_hex2bn(r_hex_normalized)

    If pos > Len(der_hex) - 1 Then Exit Function
    If Mid$(der_hex, pos, 2) <> "02" Then Exit Function
    pos = pos + 2

    If pos > Len(der_hex) - 1 Then Exit Function
    Dim s_len As Long: s_len = CLng("&H" & Mid$(der_hex, pos, 2))
    pos = pos + 2
    If s_len <= 0 Then Exit Function
    If Len(der_hex) < pos + s_len * 2 - 1 Then Exit Function

    Dim s_hex_raw As String: s_hex_raw = Mid$(der_hex, pos, s_len * 2)
    Dim s_hex_normalized As String
    pos = pos + s_len * 2
    If Not normalize_der_integer(s_hex_raw, s_hex_normalized) Then Exit Function
    sig.s = BN_hex2bn(s_hex_normalized)

    If pos <> Len(der_hex) + 1 Then Exit Function

    ecdsa_signature_from_der = True
End Function

Private Function normalize_der_integer(ByVal int_hex As String, ByRef normalized_hex As String) As Boolean
    normalize_der_integer = False
    normalized_hex = ""

    If Len(int_hex) <= 0 Or Len(int_hex) Mod 2 <> 0 Then Exit Function

    int_hex = UCase$(int_hex)

    Dim first_byte As Long
    first_byte = CLng("&H" & Left$(int_hex, 2))

    If first_byte = 0 Then
        If Len(int_hex) <= 2 Then Exit Function

        Dim second_byte As Long
        second_byte = CLng("&H" & Mid$(int_hex, 3, 2))
        If second_byte < &H80 Then Exit Function

        normalized_hex = Mid$(int_hex, 3)
        If Len(normalized_hex) = 0 Then Exit Function
        If Left$(normalized_hex, 2) = "00" Then Exit Function
    Else
        If first_byte >= &H80 Then Exit Function
        normalized_hex = int_hex
    End If

    normalize_der_integer = True
End Function
Public Function ecdsa_generate_keypair_optimized(ByRef ctx As SECP256K1_CTX) As ECDSA_KEYPAIR
    ' Geração otimizada de par de chaves usando multiplicação rápida do gerador
    ' Usa tabelas pré-computadas quando disponíveis para máxima performance
    
    Dim keypair As ECDSA_KEYPAIR
    
    ' Gerar chave privada aleatória válida no intervalo [1, n-1]
    Do
        keypair.private_key = generate_random_private_key(ctx)
    Loop While BN_is_zero(keypair.private_key) Or BN_ucmp(keypair.private_key, ctx.n) >= 0
    
    ' Usar multiplicação otimizada do gerador (tabelas pré-computadas)
    keypair.public_key = ec_point_new()
    If Not ec_point_mul_generator(keypair.public_key, keypair.private_key, ctx) Then
        Err.Raise ERR_KEYPAIR_POINT_MUL_FAILED, "ecdsa_generate_keypair_optimized", _
                  "Falha ao calcular a chave pública usando a multiplicação otimizada do gerador."
    End If

    ecdsa_generate_keypair_optimized = keypair
End Function
