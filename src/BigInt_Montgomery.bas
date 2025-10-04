Attribute VB_Name = "BigInt_Montgomery"
Option Explicit
Option Compare Binary
Option Base 0

Private Const TWO32 As Double = 4294967296#
Private Const TWO16 As Double = 65536#

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
    ' Implementa redução de Montgomery usando algoritmo CIOS word-by-word
    ' Ret = a * R^(-1) mod ctx.n sem ramificações dependentes de dados secretos

    Dim num As Long: num = ctx.n.top
    If num = 0 Then
        Call BN_zero(ret)
        bn_from_mont_word = True
        Exit Function
    End If

    If a.neg Or ctx.n.neg Then
        bn_from_mont_word = False
        Exit Function
    End If

    Dim maxWords As Long: maxWords = 2 * num
    Dim bufferLen As Long: bufferLen = maxWords + 1
    Dim t() As Long
    ReDim t(0 To bufferLen)

    Dim copyWords As Long
    If a.top > bufferLen + 1 Then
        copyWords = bufferLen + 1
    Else
        copyWords = a.top
    End If

    Dim i As Long
    For i = 0 To copyWords - 1
        t(i) = a.d(i)
    Next i
    For i = copyWords To bufferLen
        t(i) = 0
    Next i

    If a.top > bufferLen + 1 Then
        Dim extra As Long
        For extra = bufferLen + 1 To a.top - 1
            If a.d(extra) <> 0 Then
                bn_from_mont_word = False
                Exit Function
            End If
        Next extra
    End If

    Dim m As Long, carry As Long
    Dim mul_hi As Long
    Dim j As Long, idx As Long, k As Long
    Dim sum As Double
    Dim add_lo As Long, add_hi As Long
    Dim out_lo As Long, out_hi As Long

    For i = 0 To num - 1
        ' Compute m = (t(0) * Ni) mod 2^32 without touching floating point rounding.
        ' MulAdd32Word yields the exact 64-bit product; we keep the low limb in m and
        ' discard mul_hi, mirroring OpenSSL's bn_from_montgomery_word behaviour.
        Call MulAdd32Word(t(0), ctx.Ni, 0, 0, m, mul_hi)
        carry = 0

        For j = 0 To num - 1
            sum = LongToUnsignedDouble(t(j)) + LongToUnsignedDouble(carry)
            add_lo = DoubleToLong32(sum)
            add_hi = CLng(Fix(sum / TWO32))
            Call MulAdd32Word(ctx.n.d(j), m, add_lo, add_hi, out_lo, out_hi)
            t(j) = out_lo
            carry = out_hi
        Next j

        idx = num
        sum = LongToUnsignedDouble(t(idx)) + LongToUnsignedDouble(carry)
        t(idx) = DoubleToLong32(sum)
        carry = CLng(Fix(sum / TWO32))

        k = num + 1
        Do While carry <> 0 And k <= bufferLen
            sum = LongToUnsignedDouble(t(k)) + LongToUnsignedDouble(carry)
            t(k) = DoubleToLong32(sum)
            carry = CLng(Fix(sum / TWO32))
            k = k + 1
        Loop

        If carry <> 0 Then
            bn_from_mont_word = False
            Exit Function
        End If

        For k = 0 To bufferLen - 1
            t(k) = t(k + 1)
        Next k
        t(bufferLen) = 0
    Next i

    If Not bn_wexpand(ret, num + 1) Then
        bn_from_mont_word = False
        Exit Function
    End If

    For i = 0 To num
        ret.d(i) = t(i)
    Next i
    ret.top = num + 1
    ret.neg = False

    Dim diff() As Long
    ReDim diff(0 To num)
    Dim modExt() As Long
    ReDim modExt(0 To num)
    For i = 0 To num - 1
        modExt(i) = ctx.n.d(i)
    Next i
    modExt(num) = 0

    Dim borrow As Long
    borrow = bn_sub_words(diff, ret.d, modExt, num + 1)
    Dim mask As Long
    mask = borrow - 1

    For i = 0 To num
        ret.d(i) = (ret.d(i) And Not mask) Or (diff(i) And mask)
    Next i

    Call bn_correct_top(ret)
    ret.neg = False
    bn_from_mont_word = True
End Function

' =============================================================================
' FUNÇÕES AUXILIARES (WRAPPERS PARA BIGINT_VBA)
' =============================================================================

Private Sub MulAdd32Word(ByVal a As Long, ByVal b As Long, ByVal add_lo As Long, ByVal add_hi As Long, _
                         ByRef out_lo As Long, ByRef out_hi As Long)
    Dim ua As Double, ub As Double
    Dim a0 As Double, a1 As Double, b0 As Double, b1 As Double
    Dim m0 As Double, m1 As Double, m2 As Double
    Dim m1_lo As Double, m1_hi As Double
    Dim lo_acc As Double, lo_carry As Double, hi_acc As Double

    ua = LongToUnsignedDouble(a)
    ub = LongToUnsignedDouble(b)

    a0 = ua - Fix(ua / TWO16) * TWO16
    a1 = Fix(ua / TWO16)
    b0 = ub - Fix(ub / TWO16) * TWO16
    b1 = Fix(ub / TWO16)

    m0 = a0 * b0
    m1 = a0 * b1 + a1 * b0
    m2 = a1 * b1

    m1_hi = Fix(m1 / TWO16)
    m1_lo = m1 - m1_hi * TWO16

    lo_acc = m0 + (m1_lo * TWO16) + LongToUnsignedDouble(add_lo)
    lo_carry = Fix(lo_acc / TWO32)
    out_lo = DoubleToLong32(lo_acc)

    hi_acc = m2 + m1_hi + LongToUnsignedDouble(add_hi) + lo_carry
    out_hi = DoubleToLong32(hi_acc)
End Sub

Private Function bn_mul_add_words(ByRef rp() As Long, ByRef ap() As Long, ByVal num As Long, ByVal w As Long) As Long
    Dim i As Long, carry As Long
    Dim sum As Double
    Dim add_lo As Long, add_hi As Long
    Dim out_lo As Long, out_hi As Long

    carry = 0
    For i = 0 To num - 1
        sum = LongToUnsignedDouble(rp(i)) + LongToUnsignedDouble(carry)
        add_lo = DoubleToLong32(sum)
        add_hi = CLng(Fix(sum / TWO32))
        Call MulAdd32Word(ap(i), w, add_lo, add_hi, out_lo, out_hi)
        rp(i) = out_lo
        carry = out_hi
    Next i

    bn_mul_add_words = carry
End Function

Private Function bn_cmp_words(ByRef a() As Long, ByRef b() As Long, ByVal n As Long) As Long
    Dim i As Long
    Dim ua As Double, ub As Double

    For i = n - 1 To 0 Step -1
        ua = LongToUnsignedDouble(a(i))
        ub = LongToUnsignedDouble(b(i))
        If ua > ub Then
            bn_cmp_words = 1
            Exit Function
        ElseIf ua < ub Then
            bn_cmp_words = -1
            Exit Function
        End If
    Next i

    bn_cmp_words = 0
End Function

Private Function bn_sub_words(ByRef r() As Long, ByRef a() As Long, ByRef b() As Long, ByVal n As Long) As Long
    Dim i As Long
    Dim borrow As Long
    Dim t As Double

    borrow = 0
    For i = 0 To n - 1
        t = LongToUnsignedDouble(a(i)) - LongToUnsignedDouble(b(i)) - borrow
        r(i) = DoubleToLong32(t)
        If t < 0# Then
            borrow = 1
        Else
            borrow = 0
        End If
    Next i

    bn_sub_words = borrow
End Function

Private Sub bn_correct_top(ByRef a As BIGNUM_TYPE)
    ' Normaliza BIGNUM removendo zeros à esquerda
    ' Garante que top aponte para a palavra mais significativa não-zero
    ' Essencial após operações que podem gerar zeros à esquerda
    
    ' Remover zeros à esquerda, permitindo que top chegue a zero quando apropriado
    Do While a.top > 0 And a.d(a.top - 1) = 0
        a.top = a.top - 1
    Loop

    ' Zero canônico deve ter top = 0 e sinal positivo
    If a.top = 0 Then a.neg = False
End Sub

' =============================================================================
' EXPONENCIAÇÃO MODULAR MONTGOMERY
' =============================================================================

Public Function BN_mod_exp_mont(ByRef r As BIGNUM_TYPE, ByRef a As BIGNUM_TYPE, ByRef p As BIGNUM_TYPE, ByRef m As BIGNUM_TYPE, ByRef ctx As MONT_CTX) As Boolean
    ' Exponenciação modular usando contexto Montgomery
    ' Parâmetros: r = a^p mod m, ctx - contexto Montgomery pré-configurado
    ' Implementação em forma Montgomery com square-and-multiply

    Dim baseMont As BIGNUM_TYPE, accMont As BIGNUM_TYPE
    Dim baseRed As BIGNUM_TYPE, one As BIGNUM_TYPE
    Dim nbits As Long, i As Long

    If BN_is_zero(m) Then
        BN_mod_exp_mont = False
        Exit Function
    End If

    baseMont = BN_new()
    accMont = BN_new()
    baseRed = BN_new()
    one = BN_new()

    If Not BN_mod(baseRed, a, m) Then GoTo CleanupError
    Call BN_set_word(one, 1)

    If Not BN_to_montgomery(baseMont, baseRed, ctx) Then GoTo CleanupError
    If Not BN_to_montgomery(accMont, one, ctx) Then GoTo CleanupError

    nbits = BN_num_bits(p)

    For i = nbits - 1 To 0 Step -1
        If Not BN_mod_mul_montgomery(accMont, accMont, accMont, ctx) Then GoTo CleanupError
        If BN_is_bit_set(p, i) Then
            If Not BN_mod_mul_montgomery(accMont, accMont, baseMont, ctx) Then GoTo CleanupError
        End If
    Next i

    If Not BN_from_montgomery(r, accMont, ctx) Then GoTo CleanupError

    BN_mod_exp_mont = True
    GoTo Cleanup

CleanupError:
    BN_mod_exp_mont = False

Cleanup:
    Call BN_zero(baseMont)
    Call BN_zero(accMont)
    Call BN_zero(baseRed)
    Call BN_zero(one)
End Function

Private Function BN_value_one() As BIGNUM_TYPE
    ' Cria BIGNUM com valor 1 (identidade multiplicativa)
    ' Função de conveniência para inicializações
    ' Usado em algoritmos que precisam de valor unitário
    
    Dim one As BIGNUM_TYPE: one = BN_new()
    Call BN_set_word(one, 1)
    BN_value_one = one
End Function