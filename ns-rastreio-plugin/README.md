# NS Rastreio (Plugin WordPress)

Plugin para importar planilhas Excel/CSV e consultar numero de serie (NS) no navegador para obter informacoes completas do pedido.

Na sessao de bipagem ativa, o botao **Exportar bipagem em Excel** gera um arquivo `.xlsx` somente daquela sessao. A planilha traz Pedido, Nota Fiscal, SKU, Produto, quantidade do item e uma linha para cada numero de serie bipado. A lista **Bipagens salvas** permite reabrir uma bipagem finalizada para consultar, corrigir, continuar ou exportar novamente.

A lista **Bipagens salvas** permite buscar pelo pedido ou nota fiscal (numero completo ou parte) e navegar por todas as bipagens salvas, em paginas de 50 registros. Use **Abrir bipagem** no resultado desejado para consultar ou alterar, inclusive registros antigos.

O painel administrativo esta organizado nas abas **Bipagem**, **Consultar NS**, **Produtos**, **Importar / Exportar** e **Integracao Tiny**. A troca de abas preserva a sessao de bipagem aberta na pagina.

Versao atual: `1.6.0`

## Layout para celular (1.6.0)

Em telas de ate 782px, as tabelas aparecem como cartoes com os nomes dos campos.
Toque no cartao do produto para selecionar o SKU. Quantidade, status e NS ficam
no proprio cartao; expanda Editar este produto para alterar a quantidade ou rebipar.
Campos de bipagem e botoes de sessao ficam em uma coluna. Finalizar e salvar NS
aparece primeiro no grupo de acoes. Cadastro manual e adicao de SKU iniciam
recolhidos no celular e podem ser abertos pelo titulo. No computador, as tabelas
continuam com colunas. A mudanca de tamanho adapta a apresentacao automaticamente.

## Uso no celular e camera (1.5.9)

Entre no mesmo site WordPress em HTTPS com um usuario autorizado a administrar
o NS Rastreio. Na aba Bipagem, selecione um PDF ou XML de NF-e 4.0 salvo em
Arquivos/Downloads e clique em Ler PDF / XML. Selecione o SKU na tabela ou
no cartao do produto em telas pequenas.

Na sessao aberta, toque em Abrir camera e permita o acesso. O leitor prefere
a camera traseira e pausa assim que encontra um QR. Confira os numeros no campo
editavel e o SKU selecionado, depois toque em Registrar NS no SKU selecionado.
Confira o resultado da gravacao antes de abrir a camera para a proxima etiqueta.
Aceita um ou varios NS separados por linhas, espacos, virgulas, ponto e virgula
ou pipe, preservando zeros iniciais. QR com links ou dados estruturados precisa
ser corrigido para conter somente NS. O mesmo QR consecutivo para o mesmo SKU
nao e reenviado; em caso de falha, confira a lista e use o campo manual para
reenviar somente os NS ausentes.

As imagens sao decodificadas no navegador com jsQR 1.4.0, incluido no plugin
(licenca Apache 2.0 em assets/jsQR.LICENSE). A camera fecha apos a leitura, ao
sair da pagina, trocar de aba ou colocar o navegador em segundo plano.
O fluxo exige conexao com o servidor para gravar. A bipagem manual permanece
disponivel. Esta versao foi validada com testes automatizados; a leitura e o
layout ainda precisam de validacao em aparelhos Android/iPhone reais.

## Importacao/exportacao 1.5.8

Em Importar / Exportar, selecione **Padrao Olist** (layout original preservado) ou
**Modelo NS**. Use Exportar CSV completo para os dados ou Baixar modelo vazio
para obter apenas os cabecalhos. A importacao reconhece os dois formatos.

O CSV do Modelo NS usa: NS; Nota fiscal; Pedido; SKU; Descricao; Data de venda;
GTIN/EAN; Expiracao; Quantidade; Valor. Datas validas sao exportadas em DD/MM/AAAA.
O nome e GTIN/EAN sao consultados no RMA pelo SKU. A expiracao e calculada em 2 anos.
A importacao aceita esse layout e os cabecalhos antigos. NS e pelo menos NF ou Pedido
sao obrigatorios. GTIN/EAN e Expiracao sao informativos e nao alteram o RMA na importacao.
Ao abrir CSV no Excel, importe NS, SKU, NF, Pedido e GTIN/EAN como texto para preservar zeros.

## Catalogo compartilhado com RMA

Com o RMA instalado no mesmo WordPress, o NS consulta automaticamente seu catalogo
(`rma_produtos`) pelo SKU. A aba Produtos indica se o catalogo esta conectado.
Cadastro/importacao de produtos passam a ser feitos no RMA. O NS usa esse catalogo
para validar os SKUs da bipagem e atualizar as descricoes nas consultas individuais,
por nota/pedido e em lote. A consulta em lote tambem mostra o GTIN/EAN do RMA.
O vinculo preserva espacos, pontuacao e zeros do SKU e ignora diferencas de caixa.
As notas, datas e numeros de serie continuam no NS. Registros sem SKU nao sao
vinculados automaticamente. Sem catalogo RMA disponivel, a base local continua funcionando.

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

## Obter Pedido API 2.0

Serviço destinado a obter os dados de um Pedido.

- [REST](https://tiny.com.br/api-docs/api2-pedidos-obter#rest-service)

REST URL

### Parâmetros do serviço

`POST https://api.tiny.com.br/api2/pedido.obter.php`

### Parametros do servico

- `token` (string, obrigatório): chave gerada para identificar sua empresa.
- `id` (int, obrigatório): número de identificação do pedido na Olist.
- `formato` (string, obrigatório): formato do retorno (`json`).

### Retorno do serviço

| Elemento | Tipo | Tamanho | Ocorrência | Descrição |
|---|---|---|---|---|
| retorno | - | - | obrigatório | Elemento raiz do retorno |
| retorno.status_processamento | int | - | obrigatório | Conforme tabela [Status de Processamento](https://tiny.com.br/api-docs/api2-tabelas-processamento) |
| retorno.status | string | - | obrigatório | Contém o status do retorno "OK" ou "Erro" |
| retorno.codigo_erro (1) | int | - | condicional | Conforme tabela [Códigos de erro](https://tiny.com.br/api-docs/api2-tabelas-processamento) |
| retorno.erros[] (1) (3) | list | - | condicional [0..n] | Lista de erros encontrados |
| retorno.erros[].erro | string | - | condicional | Mensagem de erro |
| retorno.pedido (2) | object | - | condicional | Elemento que representa o pedido |
| retorno.pedido.id | int | - | condicional | Identificação do pedido na Olist |
| retorno.pedido.numero | int | - | condicional | Número do pedido na Olist |
| retorno.pedido.numero_ecommerce | string | 50 | condicional | Número do pedido no ecommerce/sistema |
| retorno.pedido.data_pedido (4) | date | 10 | opcional | Data do pedido |
| retorno.pedido.data_prevista (4) | date | 10 | opcional | Data prevista |
| retorno.pedido.data_faturamento (4) | date | 10 | opcional | Data de faturamento |
| retorno.pedido.data_envio (4) | date | 10 | opcional | Data de envio |
| retorno.pedido.data_entrega (4) | date | 10 | opcional | Data de entrega |
| retorno.pedido.cliente | object | - | obrigatório | Dados do cliente |
| retorno.pedido.cliente.codigo | string | 30 | opcional | Código do cliente |
| retorno.pedido.cliente.nome | string | 30 | obrigatório | Nome do cliente |
| retorno.pedido.cliente.nome_fantasia | string | 60 | opcional | Nome fantasia |
| retorno.pedido.cliente.tipo_pessoa | string | 1 | opcional | F, J ou E |
| retorno.pedido.cliente.cpf_cnpj | string | 18 | opcional | CPF/CNPJ |
| retorno.pedido.cliente.ie | string | 18 | opcional | Inscrição estadual |
| retorno.pedido.cliente.rg | string | 10 | opcional | RG |
| retorno.pedido.cliente.endereco | string | 50 | opcional | Endereço |
| retorno.pedido.cliente.numero | string | 10 | opcional | Número |
| retorno.pedido.cliente.complemento | string | 50 | opcional | Complemento |
| retorno.pedido.cliente.bairro | string | 30 | opcional | Bairro |
| retorno.pedido.cliente.cep | string | 10 | opcional | CEP |
| retorno.pedido.cliente.cidade | string | 30 | opcional | Cidade (tabela de municípios) |
| retorno.pedido.cliente.uf | string | 30 | opcional | UF |
| retorno.pedido.cliente.pais | string | 50 | opcional | País (tabela de países) |
| retorno.pedido.cliente.fone | string | 40 | opcional | Telefone |
| retorno.pedido.cliente.email | string | 50 | opcional | E-mail |
| retorno.pedido.endereco_entrega | object | - | opcional | Endereço de entrega (se diferente) |
| retorno.pedido.itens[] | list | - | obrigatório | Lista de itens |
| retorno.pedido.itens[].item | object | - | obrigatório | Item do pedido |
| retorno.pedido.itens[].item.id_produto | int | - | opcional | ID do produto na Olist |
| retorno.pedido.itens[].item.codigo | string | 20 | opcional | Código do produto |
| retorno.pedido.itens[].item.descricao | string | 120 | obrigatório | Descrição do produto |
| retorno.pedido.itens[].item.unidade | string | 3 | obrigatório | Unidade |
| retorno.pedido.itens[].item.quantidade (5) | decimal | - | obrigatório | Quantidade |
| retorno.pedido.itens[].item.valor_unitario (5) | decimal | - | obrigatório | Valor unitário |
| retorno.pedido.parcelas[] | list | - | opcional | Lista de parcelas |
| retorno.pedido.marcadores[] | list | - | opcional | Lista de marcadores |
| retorno.pedido.forma_pagamento | string | 30 | obrigatório | Conforme tabela de formas de pagamento |
| retorno.pedido.meio_pagamento | string | 100 | opcional | Meio de pagamento |
| retorno.pedido.valor_frete (5) | decimal | - | opcional | Valor do frete |
| retorno.pedido.valor_desconto (5) | decimal | - | opcional | Valor do desconto |
| retorno.pedido.total_produtos (5) | decimal | - | opcional | Total dos produtos |
| retorno.pedido.total_pedido (5) | decimal | - | opcional | Total do pedido |
| retorno.pedido.situacao | string | 15 | opcional | Situação do pedido |
| retorno.pedido.obs | string | 100 | opcional | Observação |
| retorno.pedido.obs_interna | string | 100 | opcional | Observação interna |
| retorno.pedido.codigo_rastreamento | string | 20 | opcional | Código de rastreamento |
| retorno.pedido.url_rastreamento | string | 120 | opcional | URL de rastreamento |
| retorno.pedido.ecommerce | object | - | opcional | Dados do e-commerce |
| retorno.pedido.intermediador | object | - | opcional | Dados do intermediador |
| retorno.pedido.pagamentos_integrados[] | list [0..n] | - | obrigatório | Lista de pagamentos integrados |

Observações:

- (1) Somente estará presente no retorno caso o elemento "status" seja "Erro".
- (2) Somente estará presente no retorno caso o elemento "status" seja "OK".
- (3) Estes campos somente serão informados caso o retorno contenha erros.
- (4) Estes campos devem ser informados no formato dd/mm/yyyy, exemplo "01/01/2012".
- (5) Estes campos utilizam "." (ponto) como separador de decimais, exemplo "5.25".

Referência completa do layout:

- https://tiny.com.br/api-docs/api2-pedidos-obter

### Exemplos de chamada da API

#### [Exemplos da chamada em REST](https://tiny.com.br/api-docs/api2-pedidos-obter#exemplos-chamada-rest)

```php
$url = 'https://api.tiny.com.br/api2/pedido.obter.php';
$token = 'coloque aqui a sua chave da api';
$id = 'xxxxx';
$formato = 'JSON';
$data = "token=$token&id=$id&formato='$formato'";

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

### Exemplos de retorno da API

#### [Exemplos do retorno do serviço em JSON](https://tiny.com.br/api-docs/api2-pedidos-obter#exemplos-retorno-json)

```json
{
	"retorno": {
		"status_processamento": 1,
		"status": "Erro",
		"codigo_erro": 2,
		"erros": [
			{
				"erro": "token invalido"
			}
		]
	}
}
```

```json
{
	"retorno": {
		"status_processamento": 2,
		"status": "Erro",
		"codigo_erro": 32,
		"erros": [
			{
				"erro": "Pedido não localizado"
			}
		]
	}
}
```

```json
{
	"retorno": {
		"status_processamento": "3",
		"status": "OK",
		"pedido": {
			"id": "123456",
			"numero": "123",
			"data_pedido": "01/01/2012",
			"data_prevista": "10/01/2012",
			"data_faturamento": "09/01/2012",
			"cliente": {
				"codigo": "1235",
				"nome": "Contato Teste 2",
				"nome_fantasia": "Fantasia Contato Teste 2",
				"tipo_pessoa": "F",
				"cpf_cnpj": "22755777850",
				"ie": "",
				"rg": "1234567890",
				"endereco": "Rua Teste",
				"numero": "123",
				"complemento": "sala 2",
				"bairro": "Teste",
				"cep": "95700000",
				"cidade": "Bento Gonçalves",
				"uf": "RS",
				"fone": "5412345678"
			},
			"itens": [
				{
					"item": {
						"codigo": "1234",
						"descricao": "Produto Teste 1",
						"unidade": "UN",
						"quantidade": "2",
						"valor_unitario": "50.25"
					}
				},
				{
					"item": {
						"codigo": "1235",
						"descricao": "Produto Teste 2",
						"unidade": "UN",
						"quantidade": "4",
						"valor_unitario": "15.25"
					}
				}
			],
			"parcelas": [
				{
					"parcela": {
						"dias": "30",
						"data": "29/11/2012",
						"valor": "53.84",
						"obs": "Obs Parcela 1"
					}
				},
				{
					"parcela": {
						"dias": "60",
						"data": "29/12/2012",
						"valor": "53.83",
						"obs": "Obs Parcela 2"
					}
				},
				{
					"parcela": {
						"dias": "90",
						"data": "27/01/2013",
						"valor": "53.83",
						"obs": "Obs Parcela 3"
					}
				}
			],
			"marcadores": [
				{
					"marcador": {
						"id": "149238",
						"descricao": "Teste",
						"cor": "#808080"
					}
				}
			],
			"condicao_pagamento": "30 60 90",
			"forma_pagamento": "crediario",
			"meio_pagamento": "Dinheiro",
			"nome_transportador": "transportador teste",
			"frete_por_conta": "E",
			"valor_frete": "35.00",
			"valor_desconto": "35.00",
			"total_produtos": "161.50",
			"total_pedido": "161.50",
			"numero_ordem_compra": "123",
			"deposito": "Teste",
			"forma_envio": "C",
			"forma_frete": "SEDEX - CONTRATO (40436)",
			"situacao": "Em aberto",
			"obs": "Observação Teste",
			"id_vendedor": "0",
			"nome_vendedor": "",
			"codigo_rastreamento": "TINY90831920321BR",
			"url_rastreamento": "http://urlrastreamento.com.br",
			"id_nota_fiscal": "0",
			"pagamentos_integrados": [
				{
					"pagamento_integrado": {
						"valor": 10,
						"tipo_pagamento": 1,
						"cnpj_intermediador": "49525029000186",
						"codigo_autorizacao": "JFAUTH0000020820250904130544849357542",
						"codigo_bandeira": 1
					}
				}
			]
		}
	}
}
```

### Envio manual para Tiny (NS -> observacoes do pedido)

O plugin envia os numeros de serie para o Tiny somente quando voce clicar nos botoes manuais de envio na tela de bipagem:

- `Enviar NS ao Tiny KDT`
- `Enviar NS ao Tiny TEKE`
- `Enviar NS ao Tiny TECH`

Configure os tokens no painel admin do plugin, na secao `Integracao Tiny (tokens por sistema)`, preenchendo:

- Token Tiny KDT
- Token Tiny TEKE
- Token Tiny TECH

Configuracoes opcionais no `wp-config.php` continuam disponiveis:

```php
// Opcional: usar NF como id do pedido no Tiny. Padrao: pedido.
// define('NSR_TINY_ORDER_ID_SOURCE', 'nota_fiscal');
// Opcional: tambem escrever no campo publico obs (alem de obs_interna).
// define('NSR_TINY_WRITE_OBS_PUBLIC', true);
// Opcional: URL customizada do endpoint.
// define('NSR_TINY_PEDIDO_ALTERAR_URL', 'https://api.tiny.com.br/api2/pedido.alterar.php');
// Opcional: limite de caracteres da observacao enviada.
// define('NSR_TINY_OBS_MAX_LEN', 1800);
```

Opcionalmente, tambem e possivel definir tokens por constante (`NSR_TINY_TOKEN_KDT`, `NSR_TINY_TOKEN_TEKE`, `NSR_TINY_TOKEN_TECH`) como fallback.

Com isso habilitado, ao clicar em um dos botoes de envio Tiny:

- os NS bipados da sessao atual sao agrupados por SKU;
- o plugin chama `pedido.alterar.php` no Tiny do sistema escolhido (KDT/TEKE/TECH);
- os NS sao enviados para `dados_pedido.obs_interna` (e opcionalmente `obs`).

Ao clicar em `Finalizar e salvar NS`, o plugin apenas salva no banco local (sem envio automatico ao Tiny).

Em **Bipagens salvas**, use **Excluir bipagem** para remover os registros de NS daquele pedido e nota fiscal da base local. A confirmacao identifica ambos e permite cancelar antes da exclusao.

Na sessao aberta, use **Adicionar SKU**, **Alterar quantidade**, **Excluir SKU** e **Rebipar SKU**. Rebipar limpa os NS do item para nova leitura; tambem e possivel remover um NS individual pelo botao com seu numero. Clique em **Finalizar e salvar NS** para substituir a bipagem salva pelas alteracoes.
# Consulta de varios NS (1.5.6)

Na aba **Consultar NS** e no shortcode `[ns_rastreio_consulta]`, use **Consultar varios NS**.
Cole ate 500 numeros de serie ou bipe com o leitor enviando Enter ao final de cada NS.
A consulta aceita linhas, espacos, virgulas e ponto e virgula; remove repetidos e busca o NS exato.
Mostra encontrados e nao encontrados, todas as notas fiscais/pedidos associados, SKU, descricao,
data da compra/nota fiscal registrada em `Data Venda` e expiracao da garantia apos **2 anos**.
Datas ausentes ou invalidas ficam identificadas sem calcular expiracao. Para compras em 29/02,
o aniversario em ano sem esse dia e 28/02. A consulta apenas le os registros existentes.
