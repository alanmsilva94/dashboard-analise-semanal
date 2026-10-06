# Dashboard de Análise Semanal de Pagamentos

Painel web de análise financeira que transforma planilhas de pagamentos (saídas) em visões semanais por fornecedor, banco, departamento e natureza. Todo o processamento acontece no navegador: basta abrir o `index.html` e carregar uma planilha `.xlsx`.

> **Aviso:** os dados deste repositório são **100% fictícios** e gerados por script. O projeto não tem vínculo com dados reais de nenhuma empresa ou instituição; nomes como "Fornecedor 001", "Banco A" ou "Empresa Alfa" são inventados.

## Funcionalidades

- **Valores por dia**: fornecedores nas linhas, segunda a sexta nas colunas, total semanal e total geral por dia; filtro por banco e coluna de fim de semana quando houver.
- **Frequência**: quantidade de transações por fornecedor na semana, com selo *Único* / *Múltiplo* e detalhamento de cada pagamento.
- **Status e alertas**: limite configurável (padrão R$ 10.000,00), visões *por empresa* e *por pagamento*, e fluxo de **conferência** (Pendente / Autorizado / Verificar com o departamento) com responsável, data e hora.
- **Filtros universais**: ano, mês, semana do mês, agrupamento (Beneficiário ou Favorecido), departamentos e naturezas.
- Comparativo com a semana anterior (▲/▼ %), exportação para CSV (padrão Excel pt-BR), tema claro/escuro e layout responsivo.
- Leitura automática de colunas por sinônimos, com opção de ajuste manual; suporta vários arquivos, abas e subpastas por ano.
- Cópia local dos dados (IndexedDB/localStorage) e atualização automática ao selecionar uma pasta (Edge/Chrome).
- Alternativa em Python (`consolidar.py`) para consolidar planilhas e gerar um `dados.js` pré-carregado.

## Tecnologias

HTML, CSS e JavaScript puros (sem framework) · [SheetJS](https://sheetjs.com/) 0.18.5 via CDN · Python 3 com `openpyxl` (geração de dados) e `pandas` (consolidação opcional).

## Como executar

1. Gere (ou use) a planilha de exemplo em `dados_exemplo/`.
2. Abra o `index.html` no navegador, direto do arquivo ou por um servidor local:

```bash
python -m http.server 8811
# acesse http://localhost:8811/
```

3. Em **Fonte de dados**, selecione `dados_exemplo/Pagamentos_Exemplo_2026.xlsx` (ou arraste o arquivo). Recomendado: Edge ou Chrome.

> O leitor de planilhas é carregado do CDN `cdnjs.cloudflare.com`. Para uso totalmente offline, baixe o `xlsx.full.min.js` (versão 0.18.5) e altere o `src` da tag `<script>` no `index.html`.

## Estrutura de pastas

```
dashboard-analise-semanal/
├── index.html               # dashboard (HTML + CSS + JS)
├── consolidar.py            # consolidação opcional -> dados.js
├── gerar_dados_exemplo.py   # gera planilhas fictícias
├── dados_exemplo/           # planilha .xlsx sintética
├── docs/GUIA_DE_USO.md      # guia detalhado de uso
├── LICENSE
└── .gitignore
```

## Gerar dados de exemplo

```bash
pip install openpyxl
python gerar_dados_exemplo.py              # grava em dados_exemplo/
python gerar_dados_exemplo.py --linhas 3000 --saida bd_exemplo
```

A semente é fixa (`2026`), então a saída é reproduzível. Colunas geradas: `Data`, `Departamento`, `Descriçao das Naturezas`, `Favorecido`, `Beneficiario`, `Banco`, `Valor`.

## Consolidação opcional em Python

```bash
pip install pandas openpyxl xlrd
python consolidar.py dados_exemplo --saida .
```

Gera um `dados.js` ao lado do `index.html`, que o painel carrega sozinho ao abrir. Mais detalhes em [`docs/GUIA_DE_USO.md`](docs/GUIA_DE_USO.md).

## Licença

Distribuído sob a licença MIT. Veja [LICENSE](LICENSE).