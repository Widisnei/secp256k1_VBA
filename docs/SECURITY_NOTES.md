# Segurança da API secp256k1_VBA

Este documento resume medidas implementadas para evitar exposição acidental de materiais sensíveis e garantir a coleta confiável de entropia durante o uso da biblioteca.

## Demonstrações da API

As rotinas de demonstração (`secp256k1_demo`, `secp256k1_demo_bitcoin_address` e `secp256k1_demo_key_import`) agora aceitam o parâmetro opcional `reveal_secrets`. As chaves privadas, chaves públicas e assinaturas geradas só são impressas quando o chamador define explicitamente `reveal_secrets := True`, permitindo que ambientes de produção executem as demos sem vazar segredos.

## Modo constant-time

O modo constant-time passou a ser inicializado automaticamente durante `secp256k1_init`, através da rotina `initialize_security_mode`. Operações de multiplicação escalar e exponenciação modular adotam o caminho sem ramificações condicionadas a segredos por padrão, reduzindo vetores de ataque por canal lateral. Bancos de testes e benchmarks ainda podem desativar o modo chamando `disable_security_mode` quando necessário. Chamadas subsequentes a `secp256k1_reset_context_for_tests` restauram o modo seguro, evitando que execuções posteriores permaneçam com otimizações inseguras ativadas por engano.

## Coleta de entropia

A função `fill_random_bytes` agora se apoia em `GetSecureRandomBytes`, que seleciona o melhor provedor disponível em tempo de execução:

- `BCryptGenRandom` (Windows)
- `SystemFunction036` (Windows legacy)
- `SecRandomCopyBytes` (macOS)

Se todos os provedores falharem, a rotina gera um erro (`vbObjectError + &H1000&`) para forçar tratamento explícito em vez de prosseguir com entropia insuficiente.

### Override controlado para testes

Para cenários de teste e auditoria, a API expõe `ecdsa_rng_override_seed`, permitindo injetar um buffer finito de entropia determinística. Enquanto o override estiver ativo, `fill_random_bytes` consome bytes do buffer em ordem, sinalizando esgotamento com o erro `ecdsa_rng_override_error_exhausted`. Use `ecdsa_rng_override_disable` para retornar imediatamente ao provedor criptográfico do sistema. Buffers vazios são rejeitados com `ecdsa_rng_override_error_empty` para evitar execuções inadvertidas sem entropia.