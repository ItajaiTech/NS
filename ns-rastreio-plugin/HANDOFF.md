# NS Rastreio Plugin — Handoff para próxima IA

## Visão geral do projeto

Plugin WordPress (`ns-rastreio-plugin.php`, versão `1.5.0`) para:
1. Importar planilhas `.xlsx`/`.csv` com números de série (NS) de produtos.
2. Bipar NS por SKU a partir de PDF/XML de pedido de venda.
3. Consultar NS no frontend via shortcode `[ns_rastreio_consulta]`.
4. Enviar NS bipados para o campo `obs_interna` do pedido no **Tiny ERP** (API 2.0).

---

## Problema resolvido nesta sessão

### Sintoma
O campo `obs_interna` da API Tiny tem **limite de 100 caracteres**.  
Com NS de 13 dígitos (ex: `2025082001841`) + `\n` = **14 chars por NS**, apenas ~7 NS cabem em 100 chars.  
Pedidos com 20+ NS eram truncados na primeira chamada.

### Causa raiz — PHP
Função `nsr_tiny_merge_existing_observations()` (linha ~4183) define `$max_len = 100` fixo e corta tudo.  
Função `nsr_build_tiny_obs_text()` (linha ~4054) monta o texto completo mas já era limitado depois.

### Solução implementada — **envio em lotes sequenciais**

#### Novo endpoint AJAX: `nsr_ajax_send_tiny_serials_chunked`
Arquivo: `ns-rastreio-plugin.php`

Lógica do novo fluxo PHP:
1. Coleta todos os NS bipados em lista única (`nsr_collect_serials_by_sku`).
2. Divide em **lotes de NS que caibam em ~95 chars** (margem de segurança).
3. Para cada lote:
   - Busca o estado atual do pedido no Tiny (`pedido.obter`).
   - Verifica se `obs_interna` está vazia (ou contém apenas NS já enviados).
   - Se `obs_interna` vazia → escreve o lote ali.
   - Se `obs_interna` cheia → move conteúdo atual para `obs` (público), escreve lote em `obs_interna`.
   - Se ambos cheios → concatena ao campo menos cheio com `\n`.
4. Retorna JSON com `{ ok: true, chunks_sent: N, total_ns: N }`.

#### Novo handler JS: `nsrSendTinyChunked(tinySystem)`
Substitui a chamada única por um loop `async/await` que:
1. Agrupa todos NS em lotes de ~95 chars.
2. Faz POST sequencial para cada lote via `nsr_send_tiny_serials_chunked`.
3. Exibe progresso: "Enviando lote 1/4…", "Enviando lote 2/4…", etc.
4. Ao final: "Tiny KDT atualizado — 20 NS enviados em 3 lotes."

---

## Arquitetura do código (resumo)

### Tabelas WordPress
| Tabela | Uso |
|--------|-----|
| `wp_ns_rastreio` | Registros NS importados (NS, NF, pedido, SKU, etc.) |
| `wp_nsr_produtos` | Cadastro SKU × descrição |
| `wp_nsr_scan_sessions` | Sessões de bipagem em andamento (JSON) |

### Funções PHP chave
| Função | O que faz |
|--------|-----------|
| `nsr_collect_serials_by_sku($session)` | Retorna `[SKU => [NS, NS, ...]]` da sessão |
| `nsr_build_tiny_obs_text($session, $serials_by_sku)` | Monta texto completo de obs com todos NS |
| `nsr_tiny_merge_existing_observations($pedido, $new_obs)` | Mescla obs nova com existente, respeita 100 chars |
| `nsr_send_serials_to_tiny_order($session, $system_key)` | Envio único (legado, 100 chars) |
| `nsr_tiny_get_order_details($token, $order_id)` | GET `pedido.obter` para ler estado atual |
| `nsr_resolve_tiny_internal_order_id($token, $order_ref)` | Resolve número do pedido → ID interno Tiny |
| `nsr_build_tiny_dados_pedido_strict_layout(...)` | Monta payload com parcelas + pagamentos_integrados |

### Fluxo AJAX de bipagem
```
[Usuário clica SKU na tabela]
  → nsrSelectSku(row) — destaca linha, foca input NS

[Usuário digita/bipa NS + Enter]
  → nsrScanNs() → POST wp_ajax_nsr_scan_ns
    → nsr_ajax_scan_ns() PHP → salva na sessão BD

[Usuário clica "Enviar NS ao Tiny KDT"]
  → nsrSendTiny('kdt') → POST wp_ajax_nsr_send_tiny_serials
    → nsr_ajax_send_tiny_serials() → nsr_send_serials_to_tiny_order()
    → pedido.alterar.php (Tiny API 2.0)

[Usuário clica "Finalizar e salvar NS"]
  → nsrFinalize() → POST wp_ajax_nsr_finalize_session
    → nsr_ajax_finalize_session() → nsr_upsert_ns_record() p/ cada NS
```

### Autenticação Tiny
- Tokens por sistema: `kdt`, `teke`, `tech`
- Salvos em `wp_options` chave `nsr_tiny_tokens` (array)
- Fallback para constantes `NSR_TINY_TOKEN_KDT`, `NSR_TINY_TOKEN_TEKE`, `NSR_TINY_TOKEN_TECH`
- Constante opcional `NSR_TINY_ORDER_ID_SOURCE` (`pedido` ou `nota_fiscal`)

---

## Próximas tarefas pendentes

### PRIORIDADE ALTA
- [ ] **Testar lote sequencial em produção** — validar que os lotes 2, 3, etc. não sobrescrevem o lote 1.
- [ ] **Tiny campo `obs` também tem 100 chars** — quando `obs_interna` lotar, o overflow vai para `obs` que também lota. Avaliar usar campo de observação do item (`itens[].obs`) como alternativa, ou campo customizado.

### PRIORIDADE MÉDIA
- [ ] Adicionar opção admin para escolher o separador de NS no envio ao Tiny (espaço vs `\n` vs `,`) — alguns clientes preferem tudo em uma linha.
- [ ] Permitir reenvio parcial: botão "Reenviar lote N" quando falhar meio a meio.
- [ ] Adicionar feedback visual de progresso de lote no painel de bipagem (barra de progresso).

### MELHORIAS FUTURAS
- [ ] Suporte a múltiplas planilhas em paralelo (import queue).
- [ ] API REST própria (`/wp-json/ns-rastreio/v1/`) para integração com apps externos.
- [ ] Exportar relatório de bipagem por sessão (PDF ou XLSX).

---

## Arquivos do projeto
```
ns-rastreio-plugin/
├── ns-rastreio-plugin.php   ← arquivo principal (único arquivo PHP do plugin)
├── modeloimportacao.csv     ← modelo de planilha de importação
└── README.md                ← documentação de uso e API Tiny
```

## Configurações opcionais em wp-config.php
```php
define('NSR_TINY_ORDER_ID_SOURCE', 'nota_fiscal'); // usar NF como id (padrão: pedido)
define('NSR_TINY_WRITE_OBS_PUBLIC', true);          // replicar obs_interna em obs
define('NSR_TINY_PEDIDO_ALTERAR_URL', '...');       // URL customizada
define('NSR_TINY_OBS_MAX_LEN', 1800);               // limite obs completa (não campo Tiny)
define('NSR_TINY_TOKEN_KDT',  'xxx');               // tokens por constante (fallback)
define('NSR_TINY_TOKEN_TEKE', 'xxx');
define('NSR_TINY_TOKEN_TECH', 'xxx');
```

---

## Como aplicar o patch desta sessão

O arquivo `ns-rastreio-plugin.patch` contém as alterações necessárias.  
Alternativamente, aplicar manualmente:

1. **PHP** — Adicionar função `nsr_send_serials_to_tiny_order_chunked()` após `nsr_send_serials_to_tiny_order()`.
2. **PHP** — Adicionar action `wp_ajax_nsr_send_tiny_serials_chunked` apontando para o novo handler.
3. **JS** — Substituir chamada de `nsrSendTiny()` pela versão com loop de lotes.

Ver arquivo `ns-rastreio-plugin-tiny-chunked.patch` para o diff completo.
