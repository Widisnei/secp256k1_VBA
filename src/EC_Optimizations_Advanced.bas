Attribute VB_Name = "EC_Optimizations_Advanced"
Option Explicit

' =============================================================================
' EC OPTIMIZATIONS ADVANCED - TÉCNICAS AVANÇADAS BITCOIN CORE
' =============================================================================

Public Function ec_point_mul_generator_optimized(ByRef result As EC_POINT, ByRef scalar As BIGNUM_TYPE, ByRef ctx As SECP256K1_CTX) As Boolean
    ' Multiplicação do gerador com seleção automática da melhor técnica
    If require_constant_time() Then
        ec_point_mul_generator_optimized = ec_point_mul_ladder(result, scalar, ctx.g, ctx)
        Exit Function
    End If

    If use_precomputed_gen_tables() Then
        ec_point_mul_generator_optimized = ec_generator_mul_precomputed_naf(result, scalar, ctx)
    Else
        ec_point_mul_generator_optimized = ec_point_mul_jacobian_optimized(result, scalar, ctx.g, ctx)
    End If
End Function

Public Function ec_generator_mul_precomputed_naf(ByRef result As EC_POINT, ByRef scalar As BIGNUM_TYPE, ByRef ctx As SECP256K1_CTX) As Boolean
    ' Multiplicação com Windowed Non-Adjacent Form delegando para o mapeamento COMB corrigido
    '
    ' A implementação anterior acessava a tabela pré-computada utilizando um índice linear
    ' derivado diretamente do dígito wNAF (Abs(digit) - 1) \ 2. Esse cálculo assumia uma
    ' estrutura compacta [1P, 3P, 5P, ...] que não corresponde à organização real do
    ' secp256k1_ecmult_gen_prec_table (55 blocos × 32 entradas). O resultado era a seleção
    ' incorreta de múltiplos ímpares e, consequentemente, pontos inválidos.
    '
    ' Para garantir que cada dígito wNAF acesse o múltiplo correto, agora delegamos a
    ' ec_generator_mul_precomputed_correct, que divide o escalar em janelas COMB e utiliza
    ' get_precomputed_point_fixed para mapear (bloco, dígito) → secp256k1_ecmult_gen_prec_table.
    ' Dessa forma, qualquer alteração no layout da tabela permanece centralizada em um único
    ' caminho e reaproveitamos o código de validação existente.

    If require_constant_time() Then
        ec_generator_mul_precomputed_naf = ec_point_mul_ladder(result, scalar, ctx.g, ctx)
        Exit Function
    End If

    ec_generator_mul_precomputed_naf = ec_generator_mul_precomputed_correct(result, scalar, ctx)
End Function
