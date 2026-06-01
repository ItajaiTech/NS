# NS Rastreio (Plugin WordPress)

Plugin para importar planilhas Excel/CSV e consultar numero de serie (NS) no navegador para obter informacoes completas do pedido.

Versao atual: `1.5.0`

## Requisitos

- WordPress 5.8+
- PHP 7.4+
- Extensao PHP `ZipArchive` habilitada (para importar `.xlsx`)

## Instalacao

1. Compacte a pasta `ns-rastreio-plugin` em um arquivo `.zip`.
2. No WordPress, acesse `Plugins > Adicionar novo > Enviar plugin`.
3. Envie o `.zip`, instale e ative.
4. Acesse o menu `NS Rastreio` no painel admin.

## Formato da planilha

### Colunas esperadas

O plugin detecta automaticamente os cabecalhos. As colunas principais sao:

**Obrigatorias:**
- **Observacoes internas** → contém o(s) NS (numero(s) de serie). Pode ter **múltiplos NSs** separados por virgula ou quebra de linha
- **Numero (Nota Fiscal)** → numero da nota fiscal
- **Numero** → numero do pedido

**Opcionais (para enriquecer resultados):**
- **Codigo (SKU)** → codigo SKU do produto
- **Descricao do produto** → descricao do item
- **Quantidade de produtos** → quantidade vendida
- **Valor total da venda** → valor total
- **Data da venda** → data da transacao

### Exemplo de layout

| Numero | Numero (Nota Fiscal) | Quantidade de produtos | Valor total da venda | Observacoes internas | Codigo (SKU) | Descricao do produto | Data da venda |
|--------|----------------------|------------------------|----------------------|---------------------|--------------|----------------------|---------------|
| 456789 | 987654 | 2 | R$ 150,00 | ABC12345 | SKU001 | Produto exemplo | 01/03/2026 |
| 456790 | 987655 | 1 | R$ 89.90 | DEF98765,GHI54321 | SKU002 | Produto exemplo | 02/03/2026 |

**Importante:** 
- O NS (numero de serie) vem da coluna "Observacoes internas", nao do SKU como nas versoes anteriores
- A coluna "Observacoes internas" pode conter **múltiplos NSs** separados por virgula ou quebra de linha (Enter)
- Cada NS gera um registro separado no banco, mesmo dentro de uma única linha da planilha

## Como importar

1. Acesse `NS Rastreio` no painel admin.
2. Selecione um ou mais arquivos `.xlsx` ou `.csv`.
3. Clique em `Importar Arquivos`.

## Como consultar no site

1. Crie uma pagina no WordPress.
2. Adicione o shortcode:

```text
[ns_rastreio_consulta]
```

3. Publique a pagina.
4. Abra a pagina no navegador e pesquise pelo NS.

Opcional:

- Marque `Busca parcial` para localizar NS por trecho (pesquisa parcial).

## Observacoes

- O plugin mantem historico: um mesmo NS pode ter varios registros (NF/Pedido diferentes).
- Reimportar exatamente a mesma combinacao `NS + NF + Pedido` nao duplica o dado.
- O plugin salva os dados em tabela propria no banco: `wp_ns_rastreio` (prefixo pode variar).
- A consulta exibe todos os campos importados (NF, Pedido, SKU, Descricao, Quantidade, Valor, Data).

## Integracao Tiny API 2.0 (pedido.alterar)

Documentacao oficial:

- https://api.tiny.com.br/api2/pedido.alterar.php
- REST: https://tiny.com.br/api-docs/api2-pedidos-alterar#rest-service

### Objetivo

Servico destinado a alterar alguns dados de pedidos de venda no Tiny.

### URL REST

`POST https://api.tiny.com.br/api2/pedido.alterar.php`

### Parametros do servico

- `token` (obrigatorio): chave gerada para identificar sua empresa.
- `id` (obrigatorio): id do pedido de venda a ser alterado.

### Conteudo do body

- `dados_pedido` (objeto obrigatorio): dados do pedido conforme layout.
- Layout: https://tiny.com.br/api-docs/api2-pedidos-alterar#layout-parametro-pedido

Exemplo JSON:

```json
{
	"dados_pedido": {
		"parcelas": [
			{
				"data": "20/01/2022",
				"valor": 5177.72,
				"obs": "",
				"destino": "Caixa",
				"forma_pagamento": "dinheiro"
			},
			{
				"data": "20/02/2022",
				"valor": 5200,
				"obs": "",
				"destino": "Caixa",
				"forma_pagamento": "boleto",
				"meio_pagamento": "Banco Inter"
			}
		],
		"data_prevista": "15/05/2022",
		"data_envio": "05/02/2022 08:00:00",
		"obs": "teste api",
		"obs_interna": "observacao interna teste api",
		"pagamentos_integrados": [
			{
				"tipo_pagamento": "17",
				"valor": "29.99",
				"cnpj_intermediador": "21018182000106",
				"codigo_autorizacao": "E0000020820250904130544849357542",
				"codigo_bandeira": "2"
			}
		]
	}
}
```

### Exemplo de chamada REST (PHP)

```php
$url = 'https://api.tiny.com.br/api2/pedido.alterar.php';
$token = 'coloque aqui a sua chave da api';
$id = '12345';
$data = "token=$token&id=$id";

enviarREST($url, $data);

function enviarREST($url, $data, $optional_headers = null) {
		$params = array('http' => array(
				'method' => 'POST',
				'content' => $data
		));

		if ($optional_headers !== null) {
				$params['http']['header'] = $optional_headers;
		}

		$ctx = stream_context_create($params);
		$fp = @fopen($url, 'rb', false, $ctx);
		if (!$fp) {
				throw new Exception("Problema com $url, $php_errormsg");
		}
		$response = @stream_get_contents($fp);
		if ($response === false) {
				throw new Exception("Problema obtendo retorno de $url, $php_errormsg");
		}

		return $response;
}
```

### Retorno do servico

- `retorno.status_processamento` (obrigatorio): conforme tabela de status.
- `retorno.status` (obrigatorio): `OK` ou `Erro`.
- `retorno.codigo_erro` (condicional): presente quando `status` for `Erro`.
- `retorno.erros[]` (condicional): lista de erros de validacao.

Tabelas de apoio:

- Status: https://tiny.com.br/api-docs/api2-tabelas-processamento
- Codigos de erro: https://tiny.com.br/api-docs/api2-tabelas-processamento

Exemplo de retorno com erro:

```json
{
	"retorno": {
		"status_processamento": "2",
		"status": "Erro",
		"codigo_erro": 10,
		"erros": [
			{
				"erro": "O valor total das parcelas deve ser igual ao valor total da venda.",
				"campo": "parcelas"
			}
		]
	}
}
```

Exemplo de retorno OK:

```json
{
	"retorno": {
		"status_processamento": "3",
		"status": "OK"
	}
}
```

### Envio automatico apos finalizar bipagem (NS -> observacoes do pedido)

O plugin pode enviar automaticamente os numeros de serie bipados para o Tiny ao concluir a acao de finalizar e salvar NS.

Configure no `wp-config.php`:

```php
define('NSR_TINY_TOKEN', 'SEU_TOKEN_TINY');
// Opcional: usar NF como id do pedido no Tiny. Padrao: pedido.
// define('NSR_TINY_ORDER_ID_SOURCE', 'nota_fiscal');
// Opcional: tambem escrever no campo publico obs (alem de obs_interna).
// define('NSR_TINY_WRITE_OBS_PUBLIC', true);
// Opcional: URL customizada do endpoint.
// define('NSR_TINY_PEDIDO_ALTERAR_URL', 'https://api.tiny.com.br/api2/pedido.alterar.php');
// Opcional: limite de caracteres da observacao enviada.
// define('NSR_TINY_OBS_MAX_LEN', 1800);
```

Com isso habilitado, ao finalizar a sessao de bipagem:

- os NS sao salvos no banco do plugin;
- o plugin chama `pedido.alterar.php` no Tiny;
- os NS sao enviados para `dados_pedido.obs_interna` (e opcionalmente `obs`).
