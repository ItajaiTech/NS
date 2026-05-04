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
