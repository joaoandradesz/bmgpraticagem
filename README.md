# BMG Serviços de Praticagem · BPO Financeiro (Exact BR)

Formato padrão de entrega mensal da BMG, no mesmo layout da ALS, gerado a partir do Omie, do Webpilot e dos documentos do banco:

- **`index.html`** — relatório de posição (leitura e PDF).
- **`painel.html`** — painel interativo: caixa por conta, fluxo clicável dia a dia, faturamento filtrável, cobrança, conciliação e calendário de recorrentes.

## O que o relatório traz

1. **Sumário executivo** — faturamento dos dois meses, caixa documentado, vencidos em cobrança.
2. **Faturamento por mês** — cada NFS-e com cliente, navio (Webpilot), vencimento, data de baixa e status (recebido, a vencer, vencido).
3. **Fluxo de caixa projetado** — clicável por dia; parte do saldo dos extratos e soma títulos a receber, contas a pagar e guias fora do Omie.
4. **Conciliação e cobrança** — saldo Omie × extrato por conta, vencidos sem pagamento e boletos já pagos que o banco ainda cobra.
5. **Histórico de pagamentos** — saídas por categoria nos últimos seis meses, base dos recorrentes do fluxo.

## Como gerar o mês

1. Coloque os documentos em `entradas/` (ver `entradas/README.md`): exportação "Faturas" do Webpilot e, se houver, o relatório de boletos vencidos do Itaú.
2. Atualize `config/bmg.json`:
   - `saldos_documentados` — saldo de cada conta **com o nome do extrato** que o sustenta. Sem documento, o saldo não entra.
   - `saidas_fora_do_omie` — guias (DARF, DCTFWeb) que ainda não estão cadastradas no Omie.
   - `estimativas` — títulos do Omie cujo valor é estimado (IRPJ/CSLL trimestral, ISS).
   - `recorrentes_variaveis` — itens do calendário do painel que não têm título fixo no Omie (IRRF sobre lucros, IRPJ/CSLL, distribuição).
   - `meses_faturamento` e `fluxo_ate`.
3. Rode:

```bash
cd src && python3 gerar_posicao.py --hoje AAAA-MM-DD
```

A saída vai para `relatorios/AAAA-MM/posicao-bmg-<meses>-<ano>.html` e `relatorios/AAAA-MM/painel-bmg-<meses>-<ano>.html`; a última versão de cada um é copiada para `index.html` e `painel.html` na raiz (GitHub Pages). Com `--offline` ele reaproveita o último `dados/omie.json` sem chamar a API.

## Credenciais

A chave da API do Omie fica em `config/credenciais.json` (fora do git) ou nas variáveis `OMIE_APP_KEY` e `OMIE_APP_SECRET`:

```json
{"app_key": "...", "app_secret": "..."}
```

## Regras que o gerador segue

- Saldo de banco só vem de documento do banco (`saldos_documentados`).
- Título vencido só aparece em cobrança se não houver baixa no Omie; boleto vencido no Itaú com título já recebido vira pedido de baixa.
- Distribuição de lucros fica fora do fluxo (decisão dos sócios).

Dependências: Python 3 e `openpyxl`.
