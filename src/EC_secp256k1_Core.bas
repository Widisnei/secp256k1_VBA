Attribute VB_Name = "EC_secp256k1_Core"
Option Explicit
Option Compare Binary
Option Base 0

' =============================================================================
' SECP256K1 VBA - NÚCLEO DA CURVA ELÍPTICA
' =============================================================================
' Implementação das operações fundamentais da curva elíptica secp256k1
' Inclui estruturas de dados, operações de ponto e validações
' Compatível com especificações SEC 2 e Bitcoin Core
' =============================================================================

' =============================================================================
' ESTRUTURAS DE DADOS FUNDAMENTAIS
' =============================================================================

' Ponto da curva elíptica
Public Type EC_POINT
    x As BIGNUM_TYPE
    y As BIGNUM_TYPE
    z As BIGNUM_TYPE        ' Coordenada Z para projetivas (1 = afim)
    infinity As Boolean     ' Ponto no infinito
End Type

' Contexto secp256k1
Public Type SECP256K1_CTX
    p As BIGNUM_TYPE        ' Número primo do campo finito (2^256 - 2^32 - 977)
    n As BIGNUM_TYPE        ' Ordem da curva (número de pontos válidos)
    g As EC_POINT           ' Ponto gerador da curva
    a As BIGNUM_TYPE        ' Parâmetro 'a' da equação y² = x³ + ax + b (0 para secp256k1)
    b As BIGNUM_TYPE        ' Parâmetro 'b' da equação y² = x³ + ax + b (7 para secp256k1)
End Type

' =============================================================================
' CONSTANTES MATEMÁTICAS DA CURVA SECP256K1
' =============================================================================

' Número primo do campo finito: p = 2^256 - 2^32 - 977
Private Const SECP256K1_P_HEX As String = "FFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFEFFFFFC2F"

' Ordem da curva (número total de pontos válidos)
Private Const SECP256K1_N_HEX As String = "FFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFEBAAEDCE6AF48A03BBFD25E8CD0364141"

' Coordenadas do ponto gerador G (ponto base para todas as operações)
Private Const SECP256K1_GX_HEX As String = "79BE667EF9DCBBAC55A06295CE870B07029BFCDB2DCE28D959F2815B16F81798"
Private Const SECP256K1_GY_HEX As String = "483ADA7726A3C4655DA4FBFC0E1108A8FD17B448A68554199C47D08FFB10D4B8"

' ===================== INICIALIZAÇÃO =====================

Public Function secp256k1_context_create() As SECP256K1_CTX
    ' Cria e inicializa o contexto completo da curva secp256k1
    ' Retorna: Estrutura SECP256K1_CTX com todos os parâmetros configurados

    Dim ctx As SECP256K1_CTX

    ' Inicializar número primo do campo finito
    ctx.p = BN_hex2bn(SECP256K1_P_HEX)

    ' Inicializar ordem da curva
    ctx.n = BN_hex2bn(SECP256K1_N_HEX)

    ' Inicializar parâmetros da curva: y² = x³ + 7
    ctx.a = BN_new()  ' a = 0 (secp256k1 não tem termo x²)
    ctx.b = BN_new() : Call BN_set_word(ctx.b, 7)  ' b = 7

    ' Inicializar ponto gerador G
    ctx.g = ec_point_new()
    ctx.g.x = BN_hex2bn(SECP256K1_GX_HEX)
    ctx.g.y = BN_hex2bn(SECP256K1_GY_HEX)
    Call BN_set_word(ctx.g.z, 1)  ' Usar coordenadas afins
    ctx.g.infinity = False

    secp256k1_context_create = ctx
End Function

' ===================== OPERAÇÕES DE PONTO =====================

Public Function ec_point_new() As EC_POINT
    ' Cria um novo ponto da curva elíptica inicializado com valores padrão
    ' Retorna: Ponto inicializado em coordenadas afins (z=1)

    Dim pt As EC_POINT
    pt.x = BN_new()
    pt.y = BN_new()
    pt.z = BN_new()
    Call BN_set_word(pt.z, 1)  ' Padrão para coordenadas afins
    pt.infinity = False
    ec_point_new = pt
End Function

Public Function ec_point_set_infinity(ByRef pt As EC_POINT) As Boolean
    ' Define um ponto como ponto no infinito (elemento neutro da adição)
    BN_zero pt.x
    BN_zero pt.y
    Call BN_set_word(pt.z, 1)
    pt.infinity = True
    ec_point_set_infinity = True
End Function

Public Function ec_point_is_infinity(ByRef pt As EC_POINT) As Boolean
    ' Verifica se um ponto é o ponto no infinito
    ec_point_is_infinity = pt.infinity
End Function

Public Function ec_point_copy(ByRef dest As EC_POINT, ByRef src As EC_POINT) As Boolean
    ' Copia todos os valores de um ponto para outro
    Call BN_copy(dest.x, src.x)
    Call BN_copy(dest.y, src.y)
    Call BN_copy(dest.z, src.z)
    dest.infinity = src.infinity
    ec_point_copy = True
End Function

Public Function ec_point_cmp(ByRef a As EC_POINT, ByRef b As EC_POINT, ByRef ctx As SECP256K1_CTX) As Long
    ' Compara dois pontos da curva elíptica
    ' Retorna: -1 se a < b, 0 se a = b, 1 se a > b

    If a.infinity And b.infinity Then ec_point_cmp = 0 : Exit Function
    If a.infinity Then ec_point_cmp = -1 : Exit Function
    If b.infinity Then ec_point_cmp = 1 : Exit Function

    ' Comparar coordenadas afins
    If BN_cmp(a.x, b.x) <> 0 Then ec_point_cmp = BN_cmp(a.x, b.x) : Exit Function
    ec_point_cmp = BN_cmp(a.y, b.y)
End Function

' ===================== VALIDAÇÃO DE PONTO =====================

Public Function ec_point_is_on_curve(ByRef pt As EC_POINT, ByRef ctx As SECP256K1_CTX) As Boolean
    ' Verifica se um ponto satisfaz a equação da curva secp256k1: y² = x³ + 7 (mod p)
    ' Retorna: True se o ponto está na curva, False caso contrário

    If pt.infinity Then ec_point_is_on_curve = True : Exit Function

    ' Verificar equação y² = x³ + 7 (mod p) com redução rápida
    Dim y2 As BIGNUM_TYPE, x3 As BIGNUM_TYPE, rhs As BIGNUM_TYPE, temp As BIGNUM_TYPE
    y2 = BN_new() : x3 = BN_new() : rhs = BN_new() : temp = BN_new()

    ' y^2 mod p
    Call BN_mod_sqr(y2, pt.y, ctx.p)

    ' x^3 + 7 mod p
    Call BN_mod_sqr(x3, pt.x, ctx.p)
    Call BN_mod_mul(x3, x3, pt.x, ctx.p)
    Call BN_mod_add(rhs, x3, ctx.b, ctx.p)

    ec_point_is_on_curve = (BN_cmp(y2, rhs) = 0)
End Function

' ===================== CONVERSÕES =====================

Public Function ec_point_set_affine(ByRef pt As EC_POINT, ByRef x As BIGNUM_TYPE, ByRef y As BIGNUM_TYPE) As Boolean
    ' Define as coordenadas afins (x, y) de um ponto da curva elíptica
    Call BN_copy(pt.x, x)
    Call BN_copy(pt.y, y)
    Call BN_set_word(pt.z, 1)
    pt.infinity = False
    ec_point_set_affine = True
End Function

Public Function ec_point_get_affine(ByRef pt As EC_POINT, ByRef x As BIGNUM_TYPE, ByRef y As BIGNUM_TYPE, ByRef ctx As SECP256K1_CTX) As Boolean
    ' Extrai as coordenadas afins (x, y) de um ponto, convertendo de projetivas se necessário
    ' Retorna: True se a conversão foi bem-sucedida

    If pt.infinity Then ec_point_get_affine = False : Exit Function

    ' Para coordenadas afins (z = 1), apenas copiar
    If BN_is_one(pt.z) Then
        Call BN_copy(x, pt.x)
        Call BN_copy(y, pt.y)
        ec_point_get_affine = True
        Exit Function
    End If

    ' Converter de coordenadas projetivas usando mesmo fluxo que ec_jacobian_to_affine
    Dim z_inv As BIGNUM_TYPE, z_inv2 As BIGNUM_TYPE, z_inv3 As BIGNUM_TYPE
    z_inv = BN_new() : z_inv2 = BN_new() : z_inv3 = BN_new()

    ' Calcular z_inv = Z⁻¹ mod p
    If Not BN_mod_inverse(z_inv, pt.z, ctx.p) Then ec_point_get_affine = False : Exit Function

    ' Calcular z_inv2 = z_inv² e z_inv3 = z_inv² * z_inv
    Call BN_mod_sqr(z_inv2, z_inv, ctx.p)
    Call BN_mod_mul(z_inv3, z_inv2, z_inv, ctx.p)

    ' Calcular coordenadas afins: x = X * z_inv², y = Y * z_inv³
    Call BN_mod_mul(x, pt.x, z_inv2, ctx.p)
    Call BN_mod_mul(y, pt.y, z_inv3, ctx.p)

    ec_point_get_affine = True
End Function

' ===================== COMPRESSÃO DE PONTO =====================

Public Function ec_point_compress(ByRef pt As EC_POINT, ByRef ctx As SECP256K1_CTX) As String
    ' Comprime um ponto da curva para formato SEC (33 bytes: prefixo + coordenada x)
    ' Retorna: String hexadecimal com prefixo 02 (y par) ou 03 (y ímpar) + coordenada x

    If pt.infinity Then ec_point_compress = "00" : Exit Function

    Dim x_affine As BIGNUM_TYPE, y_affine As BIGNUM_TYPE
    x_affine = BN_new() : y_affine = BN_new()

    If Not ec_point_get_affine(pt, x_affine, y_affine, ctx) Then ec_point_compress = "" : Exit Function

    ' Prefixo: 02 se y é par, 03 se y é ímpar
    Dim prefix As String
    If BN_is_odd(y_affine) Then prefix = "03" Else prefix = "02"

    Dim x_hex As String
    x_hex = BN_bn2hex(x_affine)

    ' Garantir que a coordenada X comprimida tenha sempre 32 bytes (64 caracteres hex)
    Do While Len(x_hex) < 64
        x_hex = "0" & x_hex
    Loop

    ec_point_compress = prefix & x_hex
End Function

Public Function ec_point_decompress(ByVal compressed As String, ByRef ctx As SECP256K1_CTX) As EC_POINT
    ' Descomprime um ponto do formato SEC para coordenadas completas (x, y)
    ' Parâmetro: String hexadecimal de 66 caracteres (prefixo + coordenada x)
    ' Retorna: Ponto da curva com coordenadas x e y calculadas

    Dim pt As EC_POINT : pt = ec_point_new()

    If Len(compressed) <> 66 Then  ' Deve ter 02/03 + 64 caracteres hex
        Call ec_point_set_infinity(pt)
        ec_point_decompress = pt
        Exit Function
    End If

    Dim prefix As String, x_hex As String
    prefix = left$(compressed, 2)
    x_hex = mid$(compressed, 3)

    If prefix <> "02" And prefix <> "03" Then
        Call ec_point_set_infinity(pt)
        ec_point_decompress = pt
        Exit Function
    End If

    If Not ec_is_hex_string(x_hex) Then
        Call ec_point_set_infinity(pt)
        ec_point_decompress = pt
        Exit Function
    End If

    Dim x As BIGNUM_TYPE, y As BIGNUM_TYPE
    x = BN_hex2bn(x_hex)
    y = BN_new()

    ' Rejeitar coordenadas fora do campo
    If BN_ucmp(x, ctx.p) >= 0 Then
        Call ec_point_set_infinity(pt)
        ec_point_decompress = pt
        Exit Function
    End If

    ' Calcular y² = x³ + 7
    Dim y_squared As BIGNUM_TYPE, x_cubed As BIGNUM_TYPE
    y_squared = BN_new() : x_cubed = BN_new()

    Call BN_mod_sqr(x_cubed, x, ctx.p)
    Call BN_mod_mul(x_cubed, x_cubed, x, ctx.p)
    Call BN_mod_add(y_squared, x_cubed, ctx.b, ctx.p)

    ' Calcular y = sqrt(y²) usando (p+1)/4 para secp256k1
    Dim sqrt_exp As BIGNUM_TYPE
    sqrt_exp = BN_hex2bn("3FFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFBFFFFF0C")
    Call BN_mod_exp(y, y_squared, sqrt_exp, ctx.p)

    ' Escolher raiz correta baseada na paridade
    Dim y_odd As Boolean: y_odd = BN_is_odd(y)
    Dim want_odd As Boolean: want_odd = (prefix = "03")
    
    If y_odd <> want_odd Then
        ' Calcular p - y diretamente
        Dim p_minus_y As BIGNUM_TYPE
        p_minus_y = BN_new()

        ' Usar BN_mod_sub: p - y mod p
        Dim zero As BIGNUM_TYPE
        zero = BN_new()
        Call BN_mod_sub(p_minus_y, zero, y, ctx.p)  ' 0 - y = -y, então mod p dá p - y
        Call BN_copy(y, p_minus_y)
    End If

    ' Se ainda houver inconsistência de paridade, a entrada é inválida
    If BN_is_odd(y) <> want_odd Then
        Call ec_point_set_infinity(pt)
        ec_point_decompress = pt
        Exit Function
    End If
    
    ' Verificar se y^2 mod p == x^3 + 7 mod p
    Dim verify As BIGNUM_TYPE
    verify = BN_new()
    Call BN_mod_sqr(verify, y, ctx.p)

    If BN_ucmp(verify, y_squared) <> 0 Then
        Call ec_point_set_infinity(pt)
        ec_point_decompress = pt
        Exit Function
    End If

    Call ec_point_set_affine(pt, x, y)
    ec_point_decompress = pt
End Function

Private Function ec_is_hex_string(ByVal value As String) As Boolean
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

    ec_is_hex_string = True
End Function

Public Function ec_point_to_uncompressed(ByRef point As EC_POINT) As String
    ' Converte ponto para formato descomprimido SEC (65 bytes: 04 + x + y)
    ' Retorna: String hexadecimal com prefixo 04 + coordenadas x e y completas
    
    If point.infinity Then
        ec_point_to_uncompressed = "": Exit Function
    End If
    
    Dim x_hex As String, y_hex As String
    x_hex = BN_bn2hex(point.x)
    y_hex = BN_bn2hex(point.y)
    
    ' Garantir 64 caracteres para x e y (padding com zeros à esquerda)
    Do While Len(x_hex) < 64: x_hex = "0" & x_hex: Loop
    Do While Len(y_hex) < 64: y_hex = "0" & y_hex: Loop
    
    ec_point_to_uncompressed = "04" & x_hex & y_hex
End Function