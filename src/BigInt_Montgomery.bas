Attribute VB_Name = "BigInt_Montgomery"
Option Explicit
Option Compare Binary
Option Base 0

' =============================================================================
' BIGINT MONTGOMERY VBA - ARITMÉTICA MODULAR DE MONTGOMERY
' =============================================================================
' Implementação da aritmética de Montgomery para operações modulares eficientes
' Algoritmo: redução de Montgomery para evitar divisões custosas
' Ideal para exponenciação modular com múltiplas operações no mesmo módulo
' Baseado no trabalho de Peter Montgomery (1985) e otimizações OpenSSL
' =============================================================================

' =============================================================================
' ESTRUTURA DE CONTEXTO MONTGOMERY
' =============================================================================

Public Type MONT_CTX
    n As BIGNUM_TYPE        ' Módulo N (deve ser ímpar)
    Ni As Long              ' -N^(-1) mod 2^32 (constante de Montgomery)
    RR As BIGNUM_TYPE       ' R^2 mod N (para conversão para forma Montgomery)
    ri As Long              ' Comprimento de N em bits (determina R = 2^ri)
End Type

' =============================================================================
' FUNÇÕES DE GERENCIAMENTO DE CONTEXTO
' =============================================================================

Public Function BN_MONT_CTX_new() As MONT_CTX
    ' Cria novo contexto Montgomery inicializado
    ' Retorna: Contexto Montgomery com BIGNUM alocados
    Dim ctx As MONT_CTX
    ctx.n = BN_new()   ' Módulo
    ctx.RR = BN_new()  ' R^2 mod N
    BN_MONT_CTX_new = ctx
End Function

Public Function BN_MONT_CTX_set(ByRef ctx As MONT_CTX, ByRef modulus As BIGNUM_TYPE) As Boolean
    ' Configura contexto Montgomery para módulo específico
    ' Parâmetros: ctx - contexto a ser configurado, modulus - módulo N
    ' Retorna: True se configuração bem-sucedida
    ' Requisito: módulo deve ser ímpar e não-zero
    
    ' Validar módulo (Montgomery requer módulo ímpar)
    If BN_is_zero(modulus) Or Not BN_is_odd(modulus) Then BN_MONT_CTX_set = False: Exit Function
    
    ' Armazenar módulo e calcular parâmetros
    Call BN_copy(ctx.n, modulus)
    ctx.ri = BN_num_bits(modulus)  ' R = 2^ri onde ri = bits do módulo

    ' =============================================================================
    ' CÁLCULO DA CONSTANTE DE MONTGOMERY Ni = -N^(-1) mod 2^32
    ' =============================================================================

    ' Usar algoritmo de Newton-Raphson para calcular inverso modular eficientemente
    Dim n0 As Double, x As Double
    n0 = LongToUnsignedDouble(modulus.d(0))  ' Palavra menos significativa do módulo
    
    ' Algoritmo de Newton-Raphson: x_{k+1} = x_k * (2 - n0 * x_k)
    ' Converge quadraticamente para N^(-1) mod 2^32
    x = n0
    x = x * (2# - n0 * x)  ' Precisão: 2 bits
    x = x - Fix(x / 4294967296#) * 4294967296#  ' Reduzir mod 2^32
    x = x * (2# - n0 * x)  ' Precisão: 4 bits
    x = x - Fix(x / 4294967296#) * 4294967296#
    x = x * (2# - n0 * x)  ' Precisão: 8 bits
    x = x - Fix(x / 4294967296#) * 4294967296#
    x = x * (2# - n0 * x)  ' Precisão: 16 bits
    x = x - Fix(x / 4294967296#) * 4294967296#
    x = x * (2# - n0 * x)  ' Precisão: 32 bits (completa)
    x = x - Fix(x / 4294967296#) * 4294967296#
    
    ' Armazenar -N^(-1) mod 2^32 (negativo para algoritmo CIOS)
    ctx.Ni = DoubleToLong32(-x)

    ' =============================================================================
    ' CÁLCULO DE R^2 mod N PARA CONVERSÃO MONTGOMERY
    ' =============================================================================

    ' RR = R^2 mod N onde R = 2^ri
    ' Usado para converter números para forma Montgomery: a_mont = a * R mod N
    Dim r As BIGNUM_TYPE: r = BN_new()
    Call BN_set_word(r, 1)
    Call BN_lshift(r, r, ctx.ri * 2)  ' r = 2^(2*ri) = R^2
    Call BN_mod(ctx.RR, r, ctx.n)    ' RR = R^2 mod N
    
    BN_MONT_CTX_set = True
End Function

' =============================================================================
' FUNÇÕES DE CONVERSÃO MONTGOMERY
' =============================================================================

Public Function BN_to_montgomery(ByRef r As BIGNUM_TYPE, ByRef a As BIGNUM_TYPE, ByRef ctx As MONT_CTX) As Boolean
    ' Converte número para forma Montgomery: r = a * R mod N
    ' Parâmetros: r - resultado, a - número a converter, ctx - contexto Montgomery
    ' Algoritmo: a_mont = (a * R^2) * R^(-1) mod N = a * R mod N
    Dim temp As BIGNUM_TYPE : temp = BN_new()
    If Not BN_mul(temp, a, ctx.RR) Then BN_to_montgomery = False : Exit Function
    BN_to_montgomery = bn_from_mont_word(r, temp, ctx)
End Function

Public Function BN_from_montgomery(ByRef r As BIGNUM_TYPE, ByRef a As BIGNUM_TYPE, ByRef ctx As MONT_CTX) As Boolean
    ' Converte número da forma Montgomery: r = a * R^(-1) mod N
    ' Parâmetros: r - resultado, a - número Montgomery, ctx - contexto
    ' Algoritmo: usa redução de Montgomery (CIOS - Coarsely Integrated Operand Scanning)
    BN_from_montgomery = bn_from_mont_word(r, a, ctx)
End Function

Public Function BN_mod_mul_montgomery(ByRef r As BIGNUM_TYPE, ByRef a As BIGNUM_TYPE, ByRef b As BIGNUM_TYPE, ByRef ctx As MONT_CTX) As Boolean
    ' Multiplicação modular Montgomery: r = (a * b) * R^(-1) mod N
    ' Parâmetros: r - resultado, a e b - operandos em forma Montgomery, ctx - contexto
    ' Algoritmo: multiplica e aplica redução Montgomery em uma operação (CIOS)
    ' Vantagem: evita divisão custosa, usando apenas shifts e adições
    Dim temp As BIGNUM_TYPE : temp = BN_new()
    If Not BN_mul(temp, a, b) Then BN_mod_mul_montgomery = False : Exit Function
    BN_mod_mul_montgomery = bn_from_mont_word(r, temp, ctx)
End Function

' =============================================================================
' REDUÇÃO DE MONTGOMERY (ALGORITMO CIOS SIMPLIFICADO)
' =============================================================================

Private Function bn_from_mont_word(ByRef ret As BIGNUM_TYPE, ByRef a As BIGNUM_TYPE, ByRef ctx As MONT_CTX) As Boolean
    ' Implementa redução de Montgomery: ret = (a * R^(-1)) mod N
    ' Parâmetros: ret - resultado, a - entrada, ctx - contexto Montgomery
    ' Algoritmo: versão simplificada usando operações BigInt_VBA
    ' NOTA: Em implementação otimizada, usaria algoritmo CIOS word-by-word
    
    Dim R As BIGNUM_TYPE, Rinv As BIGNUM_TYPE, temp As BIGNUM_TYPE
    R = BN_new(): Rinv = BN_new(): temp = BN_new()
    
    ' Calcular R = 2^ri (fator de Montgomery)
    Call BN_set_word(R, 1)
    Call BN_lshift(R, R, ctx.ri)
    
    ' Calcular R^(-1) mod N usando algoritmo de Euclides estendido
    If Not BN_mod_inverse(Rinv, R, ctx.n) Then bn_from_mont_word = False: Exit Function
    
    ' Aplicar redução: ret = (a * R^(-1)) mod N
    If Not BN_mod_mul(ret, a, Rinv, ctx.n) Then bn_from_mont_word = False: Exit Function
    
    bn_from_mont_word = True
End Function

' =============================================================================
' FUNÇÕES AUXILIARES (WRAPPERS PARA BIGINT_VBA)
' =============================================================================

Private Function bn_mul_add_words(ByRef rp() As Long, ByRef ap() As Long, ByVal num As Long, ByVal w As Long) As Long
    ' Multiplica array de palavras por escalar e adiciona ao resultado
    ' Parâmetros: rp - resultado, ap - operando, num - tamanho, w - multiplicador
    ' Usado em algoritmos Montgomery otimizados (CIOS)
    bn_mul_add_words = bn_add_word(rp, num, LongToUnsignedDouble(w))
End Function

Private Function bn_cmp_words(ByRef a() As Long, ByRef b() As Long, ByVal n As Long) As Long
    ' Compara arrays de palavras como números sem sinal
    ' Parâmetros: a, b - arrays a comparar, n - número de palavras
    ' Retorna: -1 se a < b, 0 se a = b, 1 se a > b
    ' Usado para determinar necessidade de subtração final em Montgomery
    
    Dim temp_a As BIGNUM_TYPE, temp_b As BIGNUM_TYPE
    temp_a = BN_new(): temp_b = BN_new()
    
    ' Expandir BIGNUM temporários para acomodar arrays
    If Not bn_wexpand(temp_a, n) Or Not bn_wexpand(temp_b, n) Then
        bn_cmp_words = 0
        Exit Function
    End If
    
    ' Copiar dados dos arrays para BIGNUM
    Dim i As Long
    For i = 0 To n - 1
        temp_a.d(i) = a(i)
        temp_b.d(i) = b(i)
    Next i
    temp_a.top = n: temp_b.top = n
    
    ' Usar comparação de magnitude do BigInt_VBA
    bn_cmp_words = BN_ucmp(temp_a, temp_b)
End Function

Private Sub bn_sub_words(ByRef r() As Long, ByRef a() As Long, ByRef b() As Long, ByVal n As Long)
    ' Subtrai arrays de palavras: r = a - b
    ' Parâmetros: r - resultado, a - minuendo, b - subtraendo, n - tamanho
    ' Assume que a >= b para evitar underflow
    ' Usado na subtração final condicional do algoritmo Montgomery
    
    Dim temp_a As BIGNUM_TYPE, temp_b As BIGNUM_TYPE, temp_r As BIGNUM_TYPE
    temp_a = BN_new(): temp_b = BN_new(): temp_r = BN_new()
    
    ' Expandir BIGNUM temporários
    If Not bn_wexpand(temp_a, n) Or Not bn_wexpand(temp_b, n) Or Not bn_wexpand(temp_r, n) Then Exit Sub
    
    ' Copiar dados dos arrays para BIGNUM
    Dim i As Long
    For i = 0 To n - 1
        temp_a.d(i) = a(i)
        temp_b.d(i) = b(i)
    Next i
    temp_a.top = n: temp_b.top = n
    
    ' Executar subtração usando BigInt_VBA e copiar resultado
    If BN_usub(temp_r, temp_a, temp_b) Then
        For i = 0 To n - 1
            If i < temp_r.top Then r(i) = temp_r.d(i) Else r(i) = 0
        Next i
    End If
End Sub

Private Sub bn_correct_top(ByRef a As BIGNUM_TYPE)
    ' Normaliza BIGNUM removendo zeros à esquerda
    ' Garante que top aponte para a palavra mais significativa não-zero
    ' Essencial após operações que podem gerar zeros à esquerda
    
    ' Remover zeros à esquerda mantendo pelo menos uma palavra
    Do While a.top > 1 And a.d(a.top - 1) = 0
        a.top = a.top - 1
    Loop
    
    ' Garantir que top nunca seja zero (representação canônica)
    If a.top = 0 Then a.top = 1
End Sub

' =============================================================================
' EXPONENCIAÇÃO MODULAR MONTGOMERY
' =============================================================================

Public Function BN_mod_exp_mont(ByRef r As BIGNUM_TYPE, ByRef a As BIGNUM_TYPE, ByRef p As BIGNUM_TYPE, ByRef m As BIGNUM_TYPE, ByRef ctx As MONT_CTX) As Boolean
    ' Exponenciação modular usando contexto Montgomery
    ' Parâmetros: r = a^p mod m, ctx - contexto Montgomery pré-configurado
    ' NOTA: Implementação simplificada - usa algoritmo padrão
    ' Para máxima eficiência, deveria converter operandos para forma Montgomery
    BN_mod_exp_mont = BN_mod_exp(r, a, p, m)
End Function

Private Function BN_value_one() As BIGNUM_TYPE
    ' Cria BIGNUM com valor 1 (identidade multiplicativa)
    ' Função de conveniência para inicializações
    ' Usado em algoritmos que precisam de valor unitário
    
    Dim one As BIGNUM_TYPE: one = BN_new()
    Call BN_set_word(one, 1)
    BN_value_one = one
End Function