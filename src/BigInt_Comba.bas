Attribute VB_Name = "BigInt_Comba"
Option Explicit
Option Compare Binary
Option Base 0

' =============================================================================
' BIGINT COMBA VBA - MULTIPLICAÇÃO RÁPIDA COMBA 8x8
' =============================================================================
' Implementação do algoritmo COMBA para multiplicação rápida de números 256-bit
' Acumulação diagonal com partes baixa (32-bit) e alta (Double) explícitas
' Usa decomposição 16x16 para produtos exatos 32x32→64 sem erros de arredondamento
' Otimizado para operações secp256k1 e criptografia de curva elíptica
' =============================================================================

' =============================================================================
' CONSTANTES MATEMÁTICAS
' =============================================================================

Private Const TWO16 As Double = 65536.0#        ' 2^16 para decomposição de palavras
Private Const TWO32 As Double = 4294967296.0#   ' 2^32 para aritmética de precisão

' =============================================================================
' FUNÇÕES AUXILIARES DE CONVERSÃO
' =============================================================================

Private Function U32(ByVal x As Long) As Double
    ' Converte Long com sinal para Double sem sinal (0 a 2^32-1)
    ' Necessário pois VBA não possui tipos unsigned nativos
    If x < 0 Then U32 = TWO32 + x Else U32 = x
End Function

Private Function D2L32(ByVal val As Double) As Long
    ' Mapeia Double para Long com sinal (complemento de dois 32-bit)
    ' Trata overflow e garante resultado válido no intervalo Long
    val = val - Fix(val / TWO32) * TWO32
    If val < 0# Then val = val + TWO32
    If val >= 2147483648.0# Then
        D2L32 = CLng(val - TWO32)
    Else
        D2L32 = CLng(val)
    End If
End Function

Private Function LongToUnsignedDouble(ByVal val As Long) As Double
    ' Converte Long com sinal para Double sem sinal
    ' Necessário pois VBA não tem tipos unsigned nativos
    If val < 0 Then
        LongToUnsignedDouble = TWO32 + CDbl(val)
    Else
        LongToUnsignedDouble = CDbl(val)
    End If
End Function

Private Sub Mul32x32_to64(ByVal a As Long, ByVal b As Long, ByRef lo As Long, ByRef hi As Long)
    ' Multiplicação 32x32 → 64 bits usando peças 16x16 (exato com Double)
    ' Implementa algoritmo clássico de multiplicação longa sem perda de precisão
    ' Evita erros de arredondamento do VBA decompondo em operações menores

    Dim ua As Double, ub As Double
    Dim a0 As Double, a1 As Double, b0 As Double, b1 As Double
    Dim m0 As Double, m1 As Double, m2 As Double
    Dim m1_lo As Double, m1_hi As Double
    Dim lo_acc As Double, lo_carry As Double, hi_acc As Double

    ' Converter para unsigned e decompor em partes de 16 bits
    ua = U32(a) : ub = U32(b)
    a0 = ua - Fix(ua / TWO16) * TWO16 : a1 = Fix(ua / TWO16)  ' a = a1*2^16 + a0
    b0 = ub - Fix(ub / TWO16) * TWO16 : b1 = Fix(ub / TWO16)  ' b = b1*2^16 + b0

    ' Calcular produtos parciais: (a1*2^16 + a0) * (b1*2^16 + b0)
    m0 = a0 * b0        ' Termo de ordem 0
    m1 = a0 * b1 + a1 * b0  ' Termos de ordem 1 (cruzados)
    m2 = a1 * b1        ' Termo de ordem 2

    ' Decompor termo m1 para evitar overflow
    m1_hi = Fix(m1 / TWO16)
    m1_lo = m1 - m1_hi * TWO16

    ' Montar resultado de 64 bits: m2*2^32 + m1*2^16 + m0
    lo_acc = m0 + m1_lo * TWO16
    lo_carry = Fix(lo_acc / TWO32)
    lo = D2L32(lo_acc)

    hi_acc = m2 + m1_hi + lo_carry
    hi = D2L32(hi_acc)
End Sub

Private Function BN_limb(ByRef bn As BIGNUM_TYPE, ByVal i As Long) As Long
    ' Extrai limb i (0..7) do BIGNUM; retorna 0 se fora do intervalo
    ' Função auxiliar para acesso seguro aos limbs durante multiplicação COMBA
    ' Trata automaticamente casos onde o índice excede o tamanho do número

    If i < 0 Then
        BN_limb = 0         ' Índice negativo = zero
    ElseIf i >= bn.top Then
        BN_limb = 0         ' Além do tamanho = zero (padding implícito)
    Else
        BN_limb = bn.d(i)   ' Limb válido
    End If
End Function

' =============================================================================
' MULTIPLICAÇÃO RÁPIDA COMBA 8x8 (256-BIT)
' =============================================================================

Public Function BN_mul_fast256(ByRef r As BIGNUM_TYPE, ByRef a As BIGNUM_TYPE, ByRef b As BIGNUM_TYPE) As Boolean
    ' Multiplicação rápida usando algoritmo COMBA para operandos até 8 limbs
    ' Parâmetros: r = a * b (suporta até 8 limbs cada; resultado até 16 limbs)
    ' Algoritmo: acumulação diagonal com propagação de carry otimizada
    ' Ideal para operações secp256k1 (256 bits = 8 limbs de 32 bits)

    Dim na As Long, nb As Long
    Dim k As Long, i As Long, j As Long
    Dim i_min As Long, i_max As Long
    Dim pl As Long, ph As Long
    Dim carry_lo As Double, carry_hi As Double
    Dim sum_lo As Double, sum_hi As Double
    Dim ai As Long, bj As Long
    Dim last_index As Long
    Dim tmp(0 To 15) As Long  ' Buffer seguro contra aliasing

    ' Segurança: fallback para multiplicação genérica se operando exceder 8 limbs
    If a.top > 8 Or b.top > 8 Then
        BN_mul_fast256 = BN_mul(r, a, b)
        Exit Function
    End If

    ' Determinar tamanhos efetivos (limitados a 8 limbs)
    na = IIf(a.top > 8, 8, a.top)
    nb = IIf(b.top > 8, 8, b.top)

    ' Tratar casos triviais (multiplicação por zero)
    If na = 0 Or nb = 0 Then
        Call BN_zero(r)
        BN_mul_fast256 = True
        Exit Function
    End If

    ' Inicializar buffer temporário e carries
    For i = 0 To 15 : tmp(i) = 0 : Next i
    carry_lo = 0# : carry_hi = 0#

    ' =============================================================================
    ' LOOP PRINCIPAL: ACUMULAÇÃO DIAGONAL COMBA
    ' =============================================================================

    ' Processar diagonais 0..(na+nb-2) da matriz de produtos
    For k = 0 To (na + nb) - 2
        ' Inicializar acumuladores com carry da diagonal anterior
        sum_lo = carry_lo
        sum_hi = carry_hi

        ' Calcular limites da diagonal k: i + j = k
        i_min = k - (nb - 1) : If i_min < 0 Then i_min = 0
        i_max = k : If i_max > (na - 1) Then i_max = na - 1

        ' Somar todos os produtos ai * bj onde i + j = k
        For i = i_min To i_max
            j = k - i
            ai = BN_limb(a, i)  ' Limb i do operando a
            bj = BN_limb(b, j)  ' Limb j do operando b

            ' Multiplicar e obter produto de 64 bits
            Call Mul32x32_to64(ai, bj, pl, ph)

            ' Adicionar parte baixa e propagar para parte alta
            sum_lo = sum_lo + LongToUnsignedDouble(pl)
            sum_hi = sum_hi + LongToUnsignedDouble(ph) + Fix(sum_lo / TWO32)
            sum_lo = sum_lo - Fix(sum_lo / TWO32) * TWO32
        Next i

        ' Armazenar resultado da diagonal k
        tmp(k) = D2L32(sum_lo)

        ' Preparar carry para próxima diagonal
        carry_lo = sum_hi - Fix(sum_hi / TWO32) * TWO32
        carry_hi = Fix(sum_hi / TWO32)
    Next k

    ' Armazenar carry final na última posição
    last_index = (na + nb) - 1
    tmp(last_index) = D2L32(carry_lo)

    ' =============================================================================
    ' FINALIZAÇÃO E NORMALIZAÇÃO
    ' =============================================================================

    ' Transferir resultado do buffer temporário para BIGNUM destino
    If Not bn_wexpand(r, last_index + 1) Then BN_mul_fast256 = False : Exit Function
    For i = 0 To last_index
        r.d(i) = tmp(i)
    Next i
    r.top = last_index + 1

    ' Remover zeros à esquerda (normalização defensiva)
    Do While r.top > 0
        If r.d(r.top - 1) = 0 Then r.top = r.top - 1 Else Exit Do
    Loop

    ' Definir sinal do resultado e tratar caso especial de zero
    r.neg = (a.neg Xor b.neg)
    If r.top = 0 Then r.neg = False
    BN_mul_fast256 = True
End Function

' =============================================================================
' QUADRADO RÁPIDO COMBA 256-BIT
' =============================================================================

Public Function BN_sqr_fast256(ByRef r As BIGNUM_TYPE, ByRef a As BIGNUM_TYPE) As Boolean
    ' Calcula quadrado usando multiplicação COMBA otimizada
    ' Parâmetro: r = a² (resultado), a = operando
    ' Implementação simples que reutiliza BN_mul_fast256
    ' Para máxima performance, poderia ser otimizado com algoritmo específico

    BN_sqr_fast256 = BN_mul_fast256(r, a, a)
End Function

Public Function BN_mul_comba8(ByRef r As BIGNUM_TYPE, ByRef a As BIGNUM_TYPE, ByRef b As BIGNUM_TYPE) As Boolean
    ' Alias para BN_mul_fast256 - multiplicação COMBA 8x8 limbs
    ' Mantém compatibilidade com chamadas existentes
    BN_mul_comba8 = BN_mul_fast256(r, a, b)
End Function