Attribute VB_Name = "BigInt_ConstTime"
Option Explicit
Option Compare Binary
Option Base 0

' =============================================================================
' BIGINT CONSTANT-TIME VBA - OPERAÇÕES RESISTENTES A TIMING ATTACKS
' =============================================================================
' Implementação de operações criptográficas em tempo constante
' Proteção contra ataques de canal lateral baseados em tempo de execução
' Essencial para segurança em operações com chaves privadas e dados sensíveis
' =============================================================================

' =============================================================================
' TROCA CONDICIONAL EM TEMPO CONSTANTE
' =============================================================================

Public Function BN_consttime_swap_flag(ByVal condition As Long, ByRef a As BIGNUM_TYPE, ByRef b As BIGNUM_TYPE) As Boolean
    ' Encaminha para implementação principal em BigInt_VBA
    Call BigInt_VBA.BN_consttime_swap_flag(condition, a, b)
    BN_consttime_swap_flag = True
End Function

' =============================================================================
' EXPONENCIAÇÃO MODULAR EM TEMPO CONSTANTE
' =============================================================================

Public Function BN_mod_exp_consttime(ByRef r As BIGNUM_TYPE, ByRef a As BIGNUM_TYPE, ByRef e As BIGNUM_TYPE, ByRef m As BIGNUM_TYPE) As Boolean
    ' Encaminha para implementação principal em BigInt_VBA para manter tempo constante
    BN_mod_exp_consttime = BigInt_VBA.BN_mod_exp_consttime(r, a, e, m)
End Function

' =============================================================================
' INVERSO MODULAR EM TEMPO CONSTANTE
' =============================================================================

Public Function BN_mod_inverse_consttime(ByRef r As BIGNUM_TYPE, ByRef a As BIGNUM_TYPE, ByRef n As BIGNUM_TYPE) As Boolean
    ' Calcula inverso modular de forma resistente a timing attacks
    ' Parâmetros: r = a^(-1) mod n
    ' Algoritmo: Euclides estendido binário com número fixo de iterações
    ' Tempo de execução independente dos valores de entrada

    ' Implementação constant-time do algoritmo binário de Euclides para inverso modular
    ' Sempre executa 2 * BN_num_bits(n) iterações e utiliza swaps mascarados

    If Not BN_is_odd(n) Then
        BN_mod_inverse_consttime = False
        Exit Function
    End If

    Dim u As BIGNUM_TYPE, v As BIGNUM_TYPE, x1 As BIGNUM_TYPE, x2 As BIGNUM_TYPE
    Dim u_base As BIGNUM_TYPE, v_base As BIGNUM_TYPE, x1_base As BIGNUM_TYPE, x2_base As BIGNUM_TYPE
    Dim u_case_u_even As BIGNUM_TYPE, v_case_u_even As BIGNUM_TYPE, x1_case_u_even As BIGNUM_TYPE, x2_case_u_even As BIGNUM_TYPE
    Dim u_case_v_even As BIGNUM_TYPE, v_case_v_even As BIGNUM_TYPE, x1_case_v_even As BIGNUM_TYPE, x2_case_v_even As BIGNUM_TYPE
    Dim u_case_sub_uv As BIGNUM_TYPE, v_case_sub_uv As BIGNUM_TYPE, x1_case_sub_uv As BIGNUM_TYPE, x2_case_sub_uv As BIGNUM_TYPE
    Dim u_case_sub_vu As BIGNUM_TYPE, v_case_sub_vu As BIGNUM_TYPE, x1_case_sub_vu As BIGNUM_TYPE, x2_case_sub_vu As BIGNUM_TYPE
    Dim tmpA As BIGNUM_TYPE, tmpB As BIGNUM_TYPE, tmpC As BIGNUM_TYPE, tmpD As BIGNUM_TYPE

    u = BN_new(): v = BN_new(): x1 = BN_new(): x2 = BN_new()
    u_base = BN_new(): v_base = BN_new(): x1_base = BN_new(): x2_base = BN_new()
    u_case_u_even = BN_new(): v_case_u_even = BN_new(): x1_case_u_even = BN_new(): x2_case_u_even = BN_new()
    u_case_v_even = BN_new(): v_case_v_even = BN_new(): x1_case_v_even = BN_new(): x2_case_v_even = BN_new()
    u_case_sub_uv = BN_new(): v_case_sub_uv = BN_new(): x1_case_sub_uv = BN_new(): x2_case_sub_uv = BN_new()
    u_case_sub_vu = BN_new(): v_case_sub_vu = BN_new(): x1_case_sub_vu = BN_new(): x2_case_sub_vu = BN_new()
    tmpA = BN_new(): tmpB = BN_new(): tmpC = BN_new(): tmpD = BN_new()

    Call BN_mod(u, a, n)
    Call BN_copy(v, n)
    Call BN_set_word(x1, 1)
    BN_zero x2

    Dim max_iterations As Long
    Dim iter As Long
    Dim swapCount As Long
    max_iterations = 2 * BN_num_bits(n)

    For iter = 0 To max_iterations - 1
        Call BN_copy(u_base, u)
        Call BN_copy(v_base, v)
        Call BN_copy(x1_base, x1)
        Call BN_copy(x2_base, x2)

        Dim u_is_odd As Long, v_is_odd As Long, x1_is_odd As Long, x2_is_odd As Long
        Dim u_ge_v As Long

        u_is_odd = 0 - (BN_is_odd(u_base) <> 0)
        v_is_odd = 0 - (BN_is_odd(v_base) <> 0)
        x1_is_odd = 0 - (BN_is_odd(x1_base) <> 0)
        x2_is_odd = 0 - (BN_is_odd(x2_base) <> 0)
        u_ge_v = 0 - (BN_ucmp(u_base, v_base) >= 0)

        Dim mask_case1 As Long, mask_case2 As Long, mask_case3 As Long, mask_case4 As Long
        mask_case1 = 1 - u_is_odd
        mask_case2 = u_is_odd * (1 - v_is_odd)
        mask_case3 = u_is_odd * v_is_odd * u_ge_v
        mask_case4 = u_is_odd * v_is_odd * (1 - u_ge_v)

        Call BN_copy(u_case_u_even, u_base)
        Call BN_rshift(u_case_u_even, u_case_u_even, 1)
        Call BN_copy(v_case_u_even, v_base)
        Call BN_copy(x2_case_u_even, x2_base)

        Call BN_copy(tmpA, x1_base)
        Call BN_rshift(tmpA, tmpA, 1)
        Call BN_copy(tmpB, x1_base)
        Call BN_add(tmpB, tmpB, n)
        Call BN_rshift(tmpB, tmpB, 1)
        Call BN_copy(x1_case_u_even, tmpA)
        Call BN_copy(tmpC, tmpB)
        swapCount = swapCount + 1
        Call BN_consttime_swap_flag(x1_is_odd, x1_case_u_even, tmpC)

        Call BN_copy(u_case_v_even, u_base)
        Call BN_copy(v_case_v_even, v_base)
        Call BN_rshift(v_case_v_even, v_case_v_even, 1)
        Call BN_copy(x1_case_v_even, x1_base)

        Call BN_copy(tmpA, x2_base)
        Call BN_rshift(tmpA, tmpA, 1)
        Call BN_copy(tmpB, x2_base)
        Call BN_add(tmpB, tmpB, n)
        Call BN_rshift(tmpB, tmpB, 1)
        Call BN_copy(x2_case_v_even, tmpA)
        Call BN_copy(tmpC, tmpB)
        swapCount = swapCount + 1
        Call BN_consttime_swap_flag(x2_is_odd, x2_case_v_even, tmpC)

        Call BN_copy(u_case_sub_uv, u_base)
        Call BN_sub(u_case_sub_uv, u_case_sub_uv, v_base)
        Call BN_copy(v_case_sub_uv, v_base)
        Call BN_copy(x1_case_sub_uv, x1_base)
        Call BN_sub(x1_case_sub_uv, x1_case_sub_uv, x2_base)
        Call BN_copy(x2_case_sub_uv, x2_base)

        Call BN_copy(u_case_sub_vu, u_base)
        Call BN_copy(v_case_sub_vu, v_base)
        Call BN_sub(v_case_sub_vu, v_case_sub_vu, u_base)
        Call BN_copy(x1_case_sub_vu, x1_base)
        Call BN_copy(x2_case_sub_vu, x2_base)
        Call BN_sub(x2_case_sub_vu, x2_case_sub_vu, x1_base)

        swapCount = swapCount + 1
        Call BN_consttime_swap_flag(mask_case1, u, u_case_u_even)
        swapCount = swapCount + 1
        Call BN_consttime_swap_flag(mask_case1, v, v_case_u_even)
        swapCount = swapCount + 1
        Call BN_consttime_swap_flag(mask_case1, x1, x1_case_u_even)
        swapCount = swapCount + 1
        Call BN_consttime_swap_flag(mask_case1, x2, x2_case_u_even)

        swapCount = swapCount + 1
        Call BN_consttime_swap_flag(mask_case2, u, u_case_v_even)
        swapCount = swapCount + 1
        Call BN_consttime_swap_flag(mask_case2, v, v_case_v_even)
        swapCount = swapCount + 1
        Call BN_consttime_swap_flag(mask_case2, x1, x1_case_v_even)
        swapCount = swapCount + 1
        Call BN_consttime_swap_flag(mask_case2, x2, x2_case_v_even)

        swapCount = swapCount + 1
        Call BN_consttime_swap_flag(mask_case3, u, u_case_sub_uv)
        swapCount = swapCount + 1
        Call BN_consttime_swap_flag(mask_case3, v, v_case_sub_uv)
        swapCount = swapCount + 1
        Call BN_consttime_swap_flag(mask_case3, x1, x1_case_sub_uv)
        swapCount = swapCount + 1
        Call BN_consttime_swap_flag(mask_case3, x2, x2_case_sub_uv)

        swapCount = swapCount + 1
        Call BN_consttime_swap_flag(mask_case4, u, u_case_sub_vu)
        swapCount = swapCount + 1
        Call BN_consttime_swap_flag(mask_case4, v, v_case_sub_vu)
        swapCount = swapCount + 1
        Call BN_consttime_swap_flag(mask_case4, x1, x1_case_sub_vu)
        swapCount = swapCount + 1
        Call BN_consttime_swap_flag(mask_case4, x2, x2_case_sub_vu)
    Next iter

    Dim u_is_one As Long, v_is_one As Long
    u_is_one = 0 - (BN_is_one(u) <> 0)
    v_is_one = 0 - (BN_is_one(v) <> 0)

    Call BN_copy(tmpA, x1)
    Call BN_copy(tmpB, x2)
    swapCount = swapCount + 1
    Call BN_consttime_swap_flag(v_is_one, tmpA, tmpB)

    Dim is_neg As Long
    is_neg = 0 - (tmpA.neg <> 0)

    Call BN_copy(tmpC, tmpA)
    Call BN_copy(tmpD, tmpA)
    Call BN_add(tmpD, tmpD, n)
    swapCount = swapCount + 1
    Call BN_consttime_swap_flag(is_neg, tmpC, tmpD)

    Dim ge_n As Long
    ge_n = 0 - (BN_ucmp(tmpC, n) >= 0)

    Call BN_copy(tmpD, tmpC)
    Call BN_sub(tmpD, tmpD, n)
    swapCount = swapCount + 1
    Call BN_consttime_swap_flag(ge_n, tmpC, tmpD)

    Call BN_copy(r, tmpC)

    If BigInt_VBA.ConstTimeInverseInstrumentationEnabled Then
        BigInt_VBA.ConstTimeInverseIterationCount = max_iterations
        BigInt_VBA.ConstTimeInverseSwapCalls = swapCount
    End If

    BN_mod_inverse_consttime = ((u_is_one Or v_is_one) <> 0)
End Function
