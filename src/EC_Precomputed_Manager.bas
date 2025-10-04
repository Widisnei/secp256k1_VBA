Attribute VB_Name = "EC_Precomputed_Manager"
Option Explicit
Option Compare Binary
Option Base 0

' =============================================================================
' MÓDULO EC PRECOMPUTED MANAGER - TABELA PRÉ-COMPUTADA SECP256K1
' =============================================================================
'
' DESCRIÇÃO:
' Gerenciador central das tabelas pré-computadas secp256k1 para otimização
' de operações de curva elíptica. Coordena inicialização, carregamento
' e acesso às tabelas do gerador e multiplicação geral.
'
' CARACTERÍSTICAS TÉCNICAS:
' - Gerenciamento centralizado de todas as tabelas pré-computadas
' - Inicialização sob demanda (lazy loading) para economia de memória
' - Interface unificada para multiplicação escalar otimizada
' - Compatibilidade total com Bitcoin Core secp256k1
'
' TABELAS GERENCIADAS:
' - Tabela do gerador: 1760 pontos pré-computados para k*G
' - Tabelas ecmult: 2×8192 pontos para multiplicação geral k*P
' - Cache de status para evitar reinicializações desnecessárias
'
' OTIMIZAÇÕES:
' - Reduz multiplicação escalar de O(k) para O(log k)
' - Essencial para performance em ECDSA e ECDH
' - Memória gerenciada eficientemente
'
' =============================================================================

' =============================================================================
' VARIÁVEIS GLOBAIS E ESTADO
' =============================================================================

' Flag de controle de inicialização das tabelas pré-computadas
' Evita carregamento duplicado e melhora performance
Private precomputed_initialized As Boolean

' =============================================================================
' INICIALIZAÇÃO E GERENCIAMENTO DE TABELAS
' =============================================================================

Public Function init_precomputed_tables() As Boolean
    ' -------------------------------------------------------------------------
    ' PROPÓSITO:
    '   Inicializa todas as tabelas pré-computadas do secp256k1 de forma
    '   coordenada, garantindo carregamento único e eficiente
    ' 
    ' FUNCIONAMENTO:
    '   - Verifica se já foram inicializadas (lazy loading)
    '   - Carrega tabela do gerador (1760 pontos)
    '   - Carrega tabelas de multiplicação geral (2×8192 pontos)
    '   - Define flag de inicialização para evitar recarregamento
    ' 
    ' RETORNA:
    '   True se inicialização foi bem-sucedida, False caso contrário
    ' 
    ' OTIMIZAÇÃO:
    '   - Carregamento sob demanda reduz uso inicial de memória
    '   - Cache de status evita operações redundantes
    ' -------------------------------------------------------------------------
    ' Verificar se já foram inicializadas (evita recarregamento)
    If precomputed_initialized Then
        init_precomputed_tables = True
        Exit Function
    End If

    ' Carregar tabelas do gerador (k*G otimizado)
    Call load_precomputed_gen_tables()

    ' Carregar tabelas de multiplicação geral (k*P otimizado)
    Call load_precomputed_ecmult_tables()

    ' Marcar como inicializadas
    precomputed_initialized = True
    init_precomputed_tables = True
End Function

' =============================================================================
' MULTIPLICAÇÃO ESCALAR OTIMIZADA
' =============================================================================

Public Function ec_generator_mul_fast(ByRef result As EC_POINT, ByRef scalar As BIGNUM_TYPE, ByRef ctx As SECP256K1_CTX) As Boolean
    ' -------------------------------------------------------------------------
    ' PROPÓSITO:
    '   Realiza multiplicação escalar k*G otimizada usando tabelas
    '   pré-computadas do gerador secp256k1
    ' 
    ' PARÂMETROS:
    '   result - Ponto resultante da operação k*G
    '   scalar - Escalar k (256 bits)
    '   ctx - Contexto secp256k1 inicializado
    ' 
    ' OTIMIZAÇÃO:
    '   - Usa 1760 pontos pré-computados do gerador
    '   - Reduz complexidade de O(k) para O(log k)
    '   - Essencial para geração rápida de chaves públicas
    ' 
    ' RETORNA:
    '   True se multiplicação foi bem-sucedida, False caso contrário
    ' -------------------------------------------------------------------------
    If require_constant_time() Then
        ec_generator_mul_fast = ec_point_mul_ladder(result, scalar, ctx.g, ctx)
        Exit Function
    End If

    ' Garantir que tabelas estejam inicializadas
    If Not precomputed_initialized Then
        If Not init_precomputed_tables() Then
            ec_generator_mul_fast = False
            Exit Function
        End If
    End If

    ' Usar tabela do gerador (1760 entradas) para máxima performance
    ec_generator_mul_fast = ec_generator_mul_precomputed_table(result, scalar, ctx)
End Function

Public Function ec_point_mul_fast(ByRef result As EC_POINT, ByRef scalar As BIGNUM_TYPE, ByRef point As EC_POINT, ByRef ctx As SECP256K1_CTX) As Boolean
    ' -------------------------------------------------------------------------
    ' PROPÓSITO:
    '   Realiza multiplicação escalar k*P otimizada usando tabelas
    '   pré-computadas para pontos arbitrários da curva
    ' 
    ' PARÂMETROS:
    '   result - Ponto resultante da operação k*P
    '   scalar - Escalar k (256 bits)
    '   point - Ponto base P da curva secp256k1
    '   ctx - Contexto secp256k1 inicializado
    ' 
    ' OTIMIZAÇÃO:
    '   - Usa 2×8192 pontos pré-computados para multiplicação geral
    '   - Algoritmo sliding window para máxima eficiência
    '   - Essencial para verificação rápida de assinaturas ECDSA
    ' 
    ' RETORNA:
    '   True se multiplicação foi bem-sucedida, False caso contrário
    ' -------------------------------------------------------------------------
    If require_constant_time() Then
        ec_point_mul_fast = ec_point_mul_ladder(result, scalar, point, ctx)
        Exit Function
    End If

    ' Garantir que tabelas estejam inicializadas
    If Not precomputed_initialized Then
        If Not init_precomputed_tables() Then
            ec_point_mul_fast = False
            Exit Function
        End If
    End If

    ' Usar tabelas de multiplicação geral (2×8192 entradas)
    ec_point_mul_fast = ec_point_mul_precomputed_table(result, scalar, point, ctx)
End Function

' =============================================================================
' IMPLEMENTAÇÕES INTERNAS DAS TABELAS
' =============================================================================

Private Function ec_generator_mul_precomputed_table(ByRef result As EC_POINT, ByRef scalar As BIGNUM_TYPE, ByRef ctx As SECP256K1_CTX) As Boolean
    ' -------------------------------------------------------------------------
    ' PROPÓSITO:
    '   Implementação interna da multiplicação do gerador usando
    '   tabelas pré-computadas com integração Bitcoin Core
    ' 
    ' ALGORITMO:
    '   - Usa função corrigida do módulo EC_Precomputed_Integration
    '   - Método COMB com janelas de 4 bits
    '   - Conversão de formato Bitcoin Core para VBA
    ' -------------------------------------------------------------------------

    ' Usar implementação corrigida do Bitcoin Core
    ec_generator_mul_precomputed_table = EC_Precomputed_Integration.ec_generator_mul_precomputed_correct(result, scalar, ctx)
End Function

Private Function ec_point_mul_precomputed_table(ByRef result As EC_POINT, ByRef scalar As BIGNUM_TYPE, ByRef point As EC_POINT, ByRef ctx As SECP256K1_CTX) As Boolean
    ' -------------------------------------------------------------------------
    ' PROPÓSITO:
    '   Multiplicação k*P utilizando tabelas pré-computadas. As tabelas
    '   disponíveis cobrem múltiplos do gerador (G) e da componente deslocada
    '   2^128*G. Para pontos fora desse conjunto a rotina recorre à versão
    '   genérica de multiplicação.
    ' -------------------------------------------------------------------------

    If point.infinity Then
        Call ec_point_set_infinity(result)
        ec_point_mul_precomputed_table = True
        Exit Function
    End If

    ' Somente o gerador possui tabela global pré-computada.
    If ec_point_cmp(point, ctx.g, ctx) <> 0 Then
        ec_point_mul_precomputed_table = ec_point_mul(result, scalar, point, ctx)
        Exit Function
    End If

    Dim baseTable As PRECOMP_TABLE_KIND
    Dim highTable As PRECOMP_TABLE_KIND
    Dim windowSize As Long

    If Not select_precomputed_tables(baseTable, highTable, windowSize) Then
        ec_point_mul_precomputed_table = ec_point_mul(result, scalar, point, ctx)
        Exit Function
    End If

    Dim kNorm As BIGNUM_TYPE
    kNorm = BN_new()

    If Not BN_mod(kNorm, scalar, ctx.n) Then
        ec_point_mul_precomputed_table = False
        Exit Function
    End If

    Dim negateResult As Boolean
    negateResult = kNorm.neg
    kNorm.neg = False

    If BN_is_zero(kNorm) Then
        Call ec_point_set_infinity(result)
        ec_point_mul_precomputed_table = True
        Exit Function
    End If

    Dim kLow As BIGNUM_TYPE
    Dim kHigh As BIGNUM_TYPE
    Dim twoPow128 As BIGNUM_TYPE

    kLow = BN_new()
    kHigh = BN_new()
    twoPow128 = BN_new()

    Call BN_set_word(twoPow128, 1)
    If Not BN_lshift(twoPow128, twoPow128, 128) Then
        ec_point_mul_precomputed_table = False
        Exit Function
    End If

    If Not BN_mod(kLow, kNorm, twoPow128) Then
        ec_point_mul_precomputed_table = False
        Exit Function
    End If

    Call BN_copy(kHigh, kNorm)
    If Not BN_rshift(kHigh, kHigh, 128) Then
        ec_point_mul_precomputed_table = False
        Exit Function
    End If

    Dim digitsLow() As Long
    Dim digitsHigh() As Long
    Dim highestLow As Long
    Dim highestHigh As Long

    highestLow = compute_wnaf_digits(kLow, windowSize, digitsLow)
    highestHigh = compute_wnaf_digits(kHigh, windowSize, digitsHigh)

    If highestLow < 0 And highestHigh < 0 Then
        Call ec_point_set_infinity(result)
        ec_point_mul_precomputed_table = True
        Exit Function
    End If

    Call ec_point_set_infinity(result)

    Dim maxIndex As Long
    maxIndex = highestLow
    If highestHigh > maxIndex Then maxIndex = highestHigh

    Dim started As Boolean
    Dim addLow As EC_POINT
    Dim addHigh As EC_POINT
    addLow = ec_point_new()
    addHigh = ec_point_new()

    Dim i As Long
    For i = maxIndex To 0 Step -1
        If started Then
            If Not ec_point_double(result, result, ctx) Then
                ec_point_mul_precomputed_table = False
                Exit Function
            End If
        End If

        If i <= highestLow Then
            If Not apply_precomputed_digit(result, started, digitsLow(i), baseTable, addLow, ctx) Then
                ec_point_mul_precomputed_table = False
                Exit Function
            End If
        End If

        If i <= highestHigh Then
            If Not apply_precomputed_digit(result, started, digitsHigh(i), highTable, addHigh, ctx) Then
                ec_point_mul_precomputed_table = False
                Exit Function
            End If
        End If
    Next i

    If negateResult Then
        If Not ec_point_negate(result, result, ctx) Then
            ec_point_mul_precomputed_table = False
            Exit Function
        End If
    End If

    ec_point_mul_precomputed_table = True
End Function

' =============================================================================
' FUNÇÕES AUXILIARES E DIAGNÓSTICO
' =============================================================================

Public Function use_precomputed_gen_tables() As Boolean
    ' -------------------------------------------------------------------------
    ' PROPÓSITO:
    '   Verifica se as tabelas pré-computadas do gerador estão disponíveis
    '   para uso em operações otimizadas
    ' 
    ' RETORNA:
    '   True se tabelas estão carregadas e prontas para uso
    ' -------------------------------------------------------------------------

    use_precomputed_gen_tables = precomputed_initialized
End Function

Public Sub load_precomputed_gen_tables()
    ' -------------------------------------------------------------------------
    ' PROPÓSITO:
    '   Carrega tabela pré-computada do gerador secp256k1 com 1760 pontos
    '   para otimização de operações k*G
    ' 
    ' FUNCIONAMENTO:
    '   - Chama inicializador do módulo EC_Precomputed_Gen
    '   - Carrega 1760 pontos pré-computados em memória
    '   - Exibe confirmação no Debug para diagnóstico
    ' -------------------------------------------------------------------------

    Call init_precomputed_gen_data
    Debug.Print "Tabela gerador carregada: 1760 entradas"
End Sub

Public Sub load_precomputed_ecmult_tables()
    ' -------------------------------------------------------------------------
    ' PROPÓSITO:
    '   Carrega tabelas pré-computadas para multiplicação geral com
    '   2×8192 pontos para otimização de operações k*P
    ' 
    ' FUNCIONAMENTO:
    '   - Chama inicializador do módulo EC_Precomputed_Ecmult
    '   - Carrega tabelas G13/G14 e auxiliares 128-bit
    '   - Exibe confirmação no Debug para diagnóstico
    ' -------------------------------------------------------------------------

    Call init_precomputed_ecmult_data
    Debug.Print "Tabelas ecmult carregadas: 2x8192 entradas"
End Sub

Public Function get_precomputed_status() As String
    ' -------------------------------------------------------------------------
    ' PROPÓSITO:
    '   Retorna status atual das tabelas pré-computadas para diagnóstico
    '   e monitoramento do sistema
    ' 
    ' RETORNA:
    '   String descritiva com estado das tabelas carregadas
    ' 
    ' FORMATO:
    '   - "Tabelas não inicializadas" se ainda não carregadas
    '   - "Tabelas carregadas: Gen(1760) + Ecmult(2x8192)" se ativas
    ' -------------------------------------------------------------------------

    If Not precomputed_initialized Then
        get_precomputed_status = "Tabelas não inicializadas"
    Else
        get_precomputed_status = "Tabelas carregadas: Gen(1760) + Ecmult(2x8192)"
    End If
End Function

Private Enum PRECOMP_TABLE_KIND
    PRECOMP_NONE = 0
    PRECOMP_G13_BASE = 1
    PRECOMP_G14_BASE = 2
    PRECOMP_G13_128 = 3
    PRECOMP_G14_128 = 4
End Enum

Private Function select_precomputed_tables(ByRef baseTable As PRECOMP_TABLE_KIND, ByRef highTable As PRECOMP_TABLE_KIND, ByRef windowSize As Long) As Boolean
    Dim hasG14 As Boolean
    Dim hasG14High As Boolean
    Dim hasG13 As Boolean
    Dim hasG13High As Boolean

    hasG14 = Len(EC_Precomputed_Ecmult.get_precomputed_point_g14(0)) > 0
    hasG14High = Len(EC_Precomputed_Ecmult.get_precomputed_point_g14_128(0)) > 0
    hasG13 = Len(EC_Precomputed_Ecmult.get_precomputed_point_g13(0)) > 0
    hasG13High = Len(EC_Precomputed_Ecmult.get_precomputed_point_g13_128(0)) > 0

    If hasG14 And hasG14High Then
        baseTable = PRECOMP_G14_BASE
        highTable = PRECOMP_G14_128
        windowSize = 14
        select_precomputed_tables = True
    ElseIf hasG13 And hasG13High Then
        baseTable = PRECOMP_G13_BASE
        highTable = PRECOMP_G13_128
        windowSize = 13
        select_precomputed_tables = True
    Else
        baseTable = PRECOMP_NONE
        highTable = PRECOMP_NONE
        windowSize = 0
        select_precomputed_tables = False
    End If
End Function

Private Function compute_wnaf_digits(ByRef scalar As BIGNUM_TYPE, ByVal windowSize As Long, ByRef digits() As Long) As Long
    Dim k As BIGNUM_TYPE
    k = BN_new()
    Call BN_copy(k, scalar)
    k.neg = False

    Dim powW As Long
    powW = 1
    Dim i As Long
    For i = 1 To windowSize
        powW = powW * 2
    Next i

    Dim halfPow As Long
    halfPow = CLng(powW / 2)

    Dim twoPow As BIGNUM_TYPE
    Dim remainder As BIGNUM_TYPE
    Dim magnitude As BIGNUM_TYPE
    twoPow = BN_new()
    remainder = BN_new()
    magnitude = BN_new()

    If Not BN_set_word(twoPow, powW) Then
        compute_wnaf_digits = -1
        Exit Function
    End If

    Dim used As Long
    Dim success As Boolean
    used = 0
    success = True

    ReDim digits(0 To 0)

    Do While Not BN_is_zero(k)
        If used > UBound(digits) Then
            ReDim Preserve digits(0 To used)
        End If

        Dim digit As Long
        digit = 0

        If BN_is_odd(k) Then
            If Not BN_mod(remainder, k, twoPow) Then
                success = False
                Exit Do
            End If

            If remainder.top > 0 Then
                digit = remainder.d(0)
            End If

            If digit >= halfPow Then
                digit = digit - powW
            End If

            If (digit And 1) = 0 Then
                If digit >= 0 Then
                    digit = digit + 1 - powW
                Else
                    digit = digit - 1 + powW
                End If
            End If

            digits(used) = digit

            If digit > 0 Then
                If Not BN_set_word(magnitude, digit) Then
                    success = False
                    Exit Do
                End If
                If Not BN_sub(k, k, magnitude) Then
                    success = False
                    Exit Do
                End If
            Else
                If Not BN_set_word(magnitude, -digit) Then
                    success = False
                    Exit Do
                End If
                If Not BN_add(k, k, magnitude) Then
                    success = False
                    Exit Do
                End If
            End If
        Else
            digits(used) = 0
        End If

        If Not BN_rshift(k, k, 1) Then
            success = False
            Exit Do
        End If

        used = used + 1
    Loop

    If Not success Then
        compute_wnaf_digits = -1
        Exit Function
    End If

    If used = 0 Then
        ReDim digits(0 To 0)
        digits(0) = 0
        compute_wnaf_digits = -1
        Exit Function
    End If

    ReDim Preserve digits(0 To used - 1)

    Dim highest As Long
    For highest = used - 1 To 0 Step -1
        If digits(highest) <> 0 Then
            compute_wnaf_digits = highest
            Exit Function
        End If
    Next highest

    compute_wnaf_digits = -1
End Function

Private Function apply_precomputed_digit(ByRef result As EC_POINT, ByRef started As Boolean, ByVal digit As Long, ByVal tableKind As PRECOMP_TABLE_KIND, ByRef scratch As EC_POINT, ByRef ctx As SECP256K1_CTX) As Boolean
    If digit = 0 Or tableKind = PRECOMP_NONE Then
        apply_precomputed_digit = True
        Exit Function
    End If

    Dim index As Long
    index = CLng((Abs(digit) - 1) / 2)

    If Not load_precomputed_point(tableKind, index, scratch, ctx) Then
        apply_precomputed_digit = False
        Exit Function
    End If

    If digit < 0 Then
        If Not ec_point_negate(scratch, scratch, ctx) Then
            apply_precomputed_digit = False
            Exit Function
        End If
    End If

    If Not started Then
        If Not ec_point_copy(result, scratch) Then
            apply_precomputed_digit = False
            Exit Function
        End If
        started = True
    Else
        If Not ec_point_add(result, result, scratch, ctx) Then
            apply_precomputed_digit = False
            Exit Function
        End If
    End If

    apply_precomputed_digit = True
End Function

Private Function load_precomputed_point(ByVal tableKind As PRECOMP_TABLE_KIND, ByVal index As Long, ByRef point As EC_POINT, ByRef ctx As SECP256K1_CTX) As Boolean
    Dim entry As String

    Select Case tableKind
        Case PRECOMP_G14_BASE
            entry = EC_Precomputed_Ecmult.get_precomputed_point_g14(index)
        Case PRECOMP_G14_128
            entry = EC_Precomputed_Ecmult.get_precomputed_point_g14_128(index)
        Case PRECOMP_G13_BASE
            entry = EC_Precomputed_Ecmult.get_precomputed_point_g13(index)
        Case PRECOMP_G13_128
            entry = EC_Precomputed_Ecmult.get_precomputed_point_g13_128(index)
        Case Else
            entry = ""
    End Select

    If Len(entry) = 0 Then
        load_precomputed_point = False
        Exit Function
    End If

    load_precomputed_point = EC_Precomputed_Integration.convert_bitcoin_core_point(entry, point, ctx)
End Function
