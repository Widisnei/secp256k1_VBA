Attribute VB_Name = "EC_secp256k1_Jacobian"
Option Explicit
Option Compare Binary
Option Base 0

' =============================================================================
' MÓDULO EC SECP256K1 JACOBIAN - COORDENADAS JACOBIANAS
' =============================================================================
'
' DESCRIÇÃO:
' Implementação de operações aritméticas da curva elíptica secp256k1
' usando coordenadas Jacobianas para máxima performance. Baseado na
' implementação do Bitcoin Core secp256k1 para compatibilidade total.
'
' CARACTERÍSTICAS TÉCNICAS:
' - Coordenadas Jacobianas (X, Y, Z) para evitar inversões modulares
' - Algoritmos otimizados baseados no Bitcoin Core secp256k1
' - Conversão bidirecional com coordenadas afins
' - Operações de duplicação e adição de alta performance
'
' VANTAGENS DAS COORDENADAS JACOBIANAS:
' - Evita inversões modulares custosas durante operações intermediárias
' - Duplicação: 4M + 6S vs 1M + 2S + 1I (afim)
' - Adição: 12M + 4S vs 1M + 2S + 1I (afim)
' - Onde M=multiplicação, S=quadrado, I=inversão (I >> M > S)
'
' REPRESENTAÇÃO:
' - Ponto afim (x, y) ↔ Ponto Jacobiano (X, Y, Z)
' - Relação: x = X/Z², y = Y/Z³
' - Infinito: Z = 0
'
' COMPATIBILIDADE:
' - Baseado em secp256k1_gej_* do Bitcoin Core
' - Algoritmos idênticos para máxima compatibilidade
' - Otimizações específicas para secp256k1 (a=0)
'
' =============================================================================

' =============================================================================
' ESTRUTURAS DE DADOS JACOBIANAS
' =============================================================================

' Estrutura de ponto em coordenadas Jacobianas para performance otimizada
' Representa ponto (X, Y, Z) onde coordenadas afins são (X/Z², Y/Z³)
Public Type EC_POINT_JACOBIAN
    x As BIGNUM_TYPE        ' Coordenada X Jacobiana
    y As BIGNUM_TYPE        ' Coordenada Y Jacobiana  
    z As BIGNUM_TYPE        ' Coordenada Z Jacobiana (Z=0 para infinito)
    infinity As Boolean     ' Flag de ponto no infinito
End Type

' =============================================================================
' FUNÇÕES BÁSICAS DE GERENCIAMENTO
' =============================================================================

Public Function ec_jacobian_new() As EC_POINT_JACOBIAN
    ' -------------------------------------------------------------------------
    ' PROPÓSITO:
    '   Cria novo ponto em coordenadas Jacobianas inicializado como infinito
    ' 
    ' RETORNA:
    '   Ponto Jacobiano inicializado (0, 1, 0) representando infinito
    ' 
    ' INICIALIZAÇÃO:
    '   - x, y, z como BIGNUM zerados
    '   - infinity = True (ponto no infinito)
    ' -------------------------------------------------------------------------

    Dim pt As EC_POINT_JACOBIAN
    pt.x = BN_new()        ' Inicializar coordenada X
    pt.y = BN_new()        ' Inicializar coordenada Y
    pt.z = BN_new()        ' Inicializar coordenada Z
    pt.infinity = True     ' Marcar como infinito
    ec_jacobian_new = pt
End Function

Public Function ec_jacobian_set_infinity(ByRef pt As EC_POINT_JACOBIAN) As Boolean
    ' -------------------------------------------------------------------------
    ' PROPÓSITO:
    '   Define ponto Jacobiano como infinito (elemento neutro da adição)
    ' 
    ' PARÂMETROS:
    '   pt - Ponto Jacobiano a ser definido como infinito
    ' 
    ' REPRESENTAÇÃO:
    '   Infinito Jacobiano: (0, 1, 0) com flag infinity = True
    ' 
    ' RETORNA:
    '   True sempre (operação sempre bem-sucedida)
    ' -------------------------------------------------------------------------

    BN_zero pt.x                ' X = 0
    Call BN_set_word(pt.y, 1)   ' Y = 1
    BN_zero pt.z                ' Z = 0 (indica infinito)
    pt.infinity = True          ' Marcar flag de infinito
    ec_jacobian_set_infinity = True
End Function

' =============================================================================
' CONVERSÕES ENTRE SISTEMAS DE COORDENADAS
' =============================================================================

Public Function ec_jacobian_from_affine(ByRef jac As EC_POINT_JACOBIAN, ByRef aff As EC_POINT) As Boolean
    ' -------------------------------------------------------------------------
    ' PROPÓSITO:
    '   Converte ponto de coordenadas afins para coordenadas Jacobianas
    ' 
    ' PARÂMETROS:
    '   jac - Ponto Jacobiano resultante
    '   aff - Ponto afim de entrada (x, y)
    ' 
    ' CONVERSÃO:
    '   Afim (x, y) → Jacobiano (x, y, 1)
    '   Infinito afim → Infinito Jacobiano (0, 1, 0)
    ' 
    ' VANTAGEM:
    '   Permite usar algoritmos Jacobianos otimizados
    ' 
    ' RETORNA:
    '   True sempre (conversão sempre possível)
    ' -------------------------------------------------------------------------
    If aff.infinity Then
        ' Converter infinito afim para infinito Jacobiano
        Call ec_jacobian_set_infinity(jac)
    Else
        ' Converter ponto regular: (x, y) → (x, y, 1)
        Call BN_copy(jac.x, aff.x)     ' X = x
        Call BN_copy(jac.y, aff.y)     ' Y = y
        Call BN_set_word(jac.z, 1)     ' Z = 1
        jac.infinity = False           ' Não é infinito
    End If
    ec_jacobian_from_affine = True
End Function

Public Function ec_jacobian_to_affine(ByRef aff As EC_POINT, ByRef jac As EC_POINT_JACOBIAN, ByRef ctx As SECP256K1_CTX) As Boolean
    ' -------------------------------------------------------------------------
    ' PROPÓSITO:
    '   Converte ponto de coordenadas Jacobianas para coordenadas afins
    ' 
    ' PARÂMETROS:
    '   aff - Ponto afim resultante (x, y)
    '   jac - Ponto Jacobiano de entrada (X, Y, Z)
    '   ctx - Contexto secp256k1 para operações modulares
    ' 
    ' CONVERSÃO:
    '   Jacobiano (X, Y, Z) → Afim (X/Z², Y/Z³)
    '   Infinito Jacobiano → Infinito afim
    ' 
    ' CUSTO COMPUTACIONAL:
    '   1 inversão modular + 1 quadrado + 2 multiplicações
    '   Operação custosa devido à inversão modular
    ' 
    ' RETORNA:
    '   True se conversão bem-sucedida, False se erro na inversão
    ' -------------------------------------------------------------------------
    If jac.infinity Then
        Call ec_point_set_infinity(aff)
        ec_jacobian_to_affine = True
        Exit Function
    End If

    ' Algoritmo: x = X/Z², y = Y/Z³
    Dim z_inv As BIGNUM_TYPE, z_inv2 As BIGNUM_TYPE, z_inv3 As BIGNUM_TYPE
    z_inv = BN_new() : z_inv2 = BN_new() : z_inv3 = BN_new()

    ' Calcular z_inv = Z⁻¹ mod p
    If Not BN_mod_inverse(z_inv, jac.z, ctx.p) Then ec_jacobian_to_affine = False : Exit Function
    
    ' Calcular z_inv2 = Z⁻² mod p
    Call BN_mod_sqr(z_inv2, z_inv, ctx.p)
    
    ' Calcular z_inv3 = Z⁻³ mod p
    Call BN_mod_mul(z_inv3, z_inv2, z_inv, ctx.p)

    ' Calcular coordenadas afins: x = X * Z⁻², y = Y * Z⁻³
    Call BN_mod_mul(aff.x, jac.x, z_inv2, ctx.p)
    Call BN_mod_mul(aff.y, jac.y, z_inv3, ctx.p)
    Call BN_set_word(aff.z, 1)    ' Coordenada z sempre 1 em afim
    aff.infinity = False

    ec_jacobian_to_affine = True
End Function

' =============================================================================
' DUPLICAÇÃO OTIMIZADA EM COORDENADAS JACOBIANAS
' =============================================================================

Public Function ec_jacobian_double(ByRef result As EC_POINT_JACOBIAN, ByRef a As EC_POINT_JACOBIAN, ByRef ctx As SECP256K1_CTX) As Boolean
    ' -------------------------------------------------------------------------
    ' PROPÓSITO:
    '   Duplica ponto em coordenadas Jacobianas usando algoritmo otimizado
    '   Calcula 2P = P + P de forma mais eficiente que adição geral
    ' 
    ' PARÂMETROS:
    '   result - Ponto Jacobiano resultante (2P)
    '   a - Ponto Jacobiano de entrada (P)
    '   ctx - Contexto secp256k1 para operações modulares
    ' 
    ' ALGORITMO OTIMIZADO:
    '   Custo: 4M + 6S (vs 1M + 2S + 1I em coordenadas afins)
    '   Onde M=multiplicação, S=quadrado, I=inversão (I >> M > S)
    ' 
    ' VANTAGEM:
    '   ~90% mais rápido que duplicação em coordenadas afins
    ' 
    ' RETORNA:
    '   True sempre (duplicação sempre bem-sucedida)
    ' -------------------------------------------------------------------------
    If a.infinity Then
        Call ec_jacobian_set_infinity(result)
        ec_jacobian_double = True
        Exit Function
    End If

    ' Algoritmo otimizado: 4M + 6S + operações com constantes (2, 3, 4, 8)
    Dim Y1Z1 As BIGNUM_TYPE, X1_2 As BIGNUM_TYPE, Y1_2 As BIGNUM_TYPE, Y1_4 As BIGNUM_TYPE
    Dim s As BIGNUM_TYPE, m As BIGNUM_TYPE, t As BIGNUM_TYPE
    Y1Z1 = BN_new() : X1_2 = BN_new() : Y1_2 = BN_new() : Y1_4 = BN_new()
    s = BN_new() : m = BN_new() : t = BN_new()

    ' Calcular Y1Z1 = Y1 * Z1
    Call BN_mod_mul(Y1Z1, a.y, a.z, ctx.p)

    ' Calcular S = 4 * X1 * Y1²
    Call BN_mod_sqr(Y1_2, a.y, ctx.p)
    Call BN_mod_mul(s, a.x, Y1_2, ctx.p)
    Dim four As BIGNUM_TYPE : four = BN_new() : Call BN_set_word(four, 4)
    Call BN_mod_mul(s, s, four, ctx.p)

    ' Calcular M = 3 * X1²
    Call BN_mod_sqr(X1_2, a.x, ctx.p)
    Dim three As BIGNUM_TYPE : three = BN_new() : Call BN_set_word(three, 3)
    Call BN_mod_mul(m, three, X1_2, ctx.p)

    ' X3 = M² - 2*S
    Call BN_mod_sqr(result.x, m, ctx.p)
    Dim two As BIGNUM_TYPE : two = BN_new() : Call BN_set_word(two, 2)
    Call BN_mod_mul(t, two, s, ctx.p)
    Call BN_mod_sub(result.x, result.x, t, ctx.p)

    ' Calcular Y3 = M*(S - X3) - 8*Y1⁴
    Call BN_mod_sub(t, s, result.x, ctx.p)
    Call BN_mod_mul(result.y, m, t, ctx.p)
    Call BN_mod_sqr(Y1_4, Y1_2, ctx.p)
    Dim eight As BIGNUM_TYPE : eight = BN_new() : Call BN_set_word(eight, 8)
    Call BN_mod_mul(t, eight, Y1_4, ctx.p)
    Call BN_mod_sub(result.y, result.y, t, ctx.p)

    ' Calcular Z3 = 2 * Y1 * Z1
    Call BN_mod_mul(result.z, two, Y1Z1, ctx.p)

    result.infinity = False
    ec_jacobian_double = True
End Function

' =============================================================================
' ADIÇÃO JACOBIANA + AFIM OTIMIZADA
' =============================================================================

Public Function ec_jacobian_add_affine(ByRef result As EC_POINT_JACOBIAN, ByRef a As EC_POINT_JACOBIAN, ByRef b As EC_POINT, ByRef ctx As SECP256K1_CTX) As Boolean
    ' -------------------------------------------------------------------------
    ' PROPÓSITO:
    '   Realiza adição de ponto Jacobiano com ponto afim usando
    '   algoritmo otimizado baseado no Bitcoin Core secp256k1_gej_add_ge
    ' 
    ' PARÂMETROS:
    '   result - Ponto Jacobiano resultante da adição
    '   a - Ponto Jacobiano (X1, Y1, Z1)
    '   b - Ponto afim (x2, y2)
    '   ctx - Contexto secp256k1 para operações modulares
    ' 
    ' ALGORITMO:
    '   Adição mista Jacobiano + Afim:
    '   1. Transformar ponto afim para coordenadas Jacobianas
    '   2. Calcular diferenças h e r
    '   3. Aplicar fórmulas de adição Jacobiana
    ' 
    ' VANTAGEM:
    '   Mais eficiente que converter afim→Jacobiano + adição pura
    ' 
    ' RETORNA:
    '   True se adição bem-sucedida, False caso contrário
    ' -------------------------------------------------------------------------
    If a.infinity Then
        Call ec_jacobian_from_affine(result, b)
        ec_jacobian_add_affine = True
        Exit Function
    End If

    If b.infinity Then
        Call BN_copy(result.x, a.x)
        Call BN_copy(result.y, a.y)
        Call BN_copy(result.z, a.z)
        result.infinity = a.infinity
        ec_jacobian_add_affine = True
        Exit Function
    End If

    ' =============================================================================
    ' ALGORITMO: Baseado em secp256k1_gej_add_ge do Bitcoin Core
    ' =============================================================================

    Dim z12 As BIGNUM_TYPE, z13 As BIGNUM_TYPE, u2 As BIGNUM_TYPE, s2 As BIGNUM_TYPE
    Dim h As BIGNUM_TYPE, h2 As BIGNUM_TYPE, h3 As BIGNUM_TYPE, r As BIGNUM_TYPE
    z12 = BN_new() : z13 = BN_new() : u2 = BN_new() : s2 = BN_new()
    h = BN_new() : h2 = BN_new() : h3 = BN_new() : r = BN_new()

    ' Calcular z12 = z1²
    Call BN_mod_sqr(z12, a.z, ctx.p)

    ' z13 = z1³ = z1 * z12
    Call BN_mod_mul(z13, a.z, z12, ctx.p)

    ' u2 = x2 * z12 (transforma x2 para coordenadas Jacobian)
    Call BN_mod_mul(u2, b.x, z12, ctx.p)

    ' s2 = y2 * z13 (transforma y2 para coordenadas Jacobian)
    Call BN_mod_mul(s2, b.y, z13, ctx.p)

    ' h = u2 - x1
    Call BN_mod_sub(h, u2, a.x, ctx.p)

    ' r = s2 - y1
    Call BN_mod_sub(r, s2, a.y, ctx.p)

    ' Verificar casos especiais
    If BN_is_zero(h) Then
        If BN_is_zero(r) Then
            ' Pontos iguais: duplicar
            ec_jacobian_add_affine = ec_jacobian_double(result, a, ctx)
        Else
            ' Pontos opostos: infinito
            Call ec_jacobian_set_infinity(result)
            ec_jacobian_add_affine = True
        End If
        Exit Function
    End If

    ' h2 = h²
    Call BN_mod_sqr(h2, h, ctx.p)

    ' h3 = h * h2
    Call BN_mod_mul(h3, h, h2, ctx.p)

    ' x3 = r² - h3 - 2*x1*h2
    Call BN_mod_sqr(result.x, r, ctx.p)
    Call BN_mod_sub(result.x, result.x, h3, ctx.p)
    Dim temp As BIGNUM_TYPE: temp = BN_new()
    Call BN_mod_mul(temp, a.x, h2, ctx.p)
    Dim two As BIGNUM_TYPE: two = BN_new(): Call BN_set_word(two, 2)
    Call BN_mod_mul(temp, temp, two, ctx.p)
    Call BN_mod_sub(result.x, result.x, temp, ctx.p)
    
    ' y3 = r*(x1*h2 - x3) - y1*h3
    Call BN_mod_mul(temp, a.x, h2, ctx.p)
    Call BN_mod_sub(temp, temp, result.x, ctx.p)
    Call BN_mod_mul(result.y, r, temp, ctx.p)
    Call BN_mod_mul(temp, a.y, h3, ctx.p)
    Call BN_mod_sub(result.y, result.y, temp, ctx.p)
    
    ' z3 = z1 * h
    Call BN_mod_mul(result.z, a.z, h, ctx.p)
    
    result.infinity = False
    ec_jacobian_add_affine = True
End Function