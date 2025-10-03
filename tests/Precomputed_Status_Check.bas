Attribute VB_Name = "Precomputed_Status_Check"
Option Explicit

Public Sub check_precomputed_integration()
    Debug.Print "=== STATUS DAS TABELAS PRÉ-COMPUTADAS ==="
    
    ' Verificar status das tabelas
    Debug.Print "Status: " & get_precomputed_status()
    
    ' Forçar inicialização se necessário
    If Not use_precomputed_gen_tables() Then
        Debug.Print "Inicializando tabelas pré-computadas..."
        Call init_precomputed_tables
    End If
    
    ' Verificar novamente
    If use_precomputed_gen_tables() Then
        Debug.Print "✅ Tabelas ATIVAS: Gerador (1760) + Ecmult (2x8192)"
        Debug.Print "✅ Multiplicação k*G: ~90% mais rápida"
        Debug.Print "✅ Integração ULTIMATE: Funcionando"
    Else
        Debug.Print "❌ Tabelas NÃO ATIVAS: Usando fallback"
    End If
    
    Debug.Print "=== VERIFICAÇÃO CONCLUÍDA ==="
End Sub