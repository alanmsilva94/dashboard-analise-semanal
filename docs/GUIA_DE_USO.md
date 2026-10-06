# Guia de uso detalhado

Painel para acompanhar pagamentos por semana, empresa, banco e departamento.
Roda direto do navegador, sem backend (precisa de internet apenas para carregar o SheetJS via CDN).

## Arquivos

| Arquivo | Para que serve |
|---|---|
| `index.html` | O dashboard. Arquivo único; o leitor de XLSX (SheetJS) é carregado por CDN. |
| `consolidar.py` | Opcional. Varre as pastas da rede com pandas e gera o `dados.js`. |
| `dados.js` | Gerado pelo `consolidar.py`. Se estiver na mesma pasta do `index.html`, o painel carrega sozinho ao abrir. |
| `dados_exemplo/Pagamentos_Exemplo_2026.xlsx` | Planilha de exemplo para testar antes de apontar para os dados reais. |

## Instalação

Copie o `index.html` para a pasta da rede, por exemplo
`\\servidor-exemplo\financeiro\Pagamentos\Dashboard\index.html`, e crie um atalho na área de
trabalho de quem vai usar. Não precisa instalar nada na máquina do usuário.

Recomendado: Edge ou Chrome. No Firefox e no Safari o painel funciona, mas sem a
atualização automática (veja abaixo).

## A primeira leitura e as seguintes

Depois da primeira leitura, os dados ficam guardados no navegador daquele computador.
Nas próximas vezes que você abrir o `index.html`, o painel já aparece preenchido, com o
mês e a semana que estavam selecionados, sem precisar importar nada. A barra lateral
mostra a data e a hora da leitura que está na tela.

Para trazer lançamentos novos, clique em **Atualizar** — ou vá em *Fonte de dados* e
escolha outra pasta, se o caminho mudou. Em *Fonte de dados* há também **Esquecer dados
salvos**, que apaga a cópia local e o caminho gravado.

A cópia é compacta (cerca de 36 bytes por pagamento, ou 700 KB para 20 mil lançamentos)
e fica no IndexedDB; quando o navegador não o disponibiliza, o painel usa o
armazenamento local comum. Se a base for grande demais para caber ali, o painel avisa
em *Fonte de dados* e o caminho recomendado passa a ser o `consolidar.py`.

## Três formas de alimentar o painel

**1. Selecionar pasta da rede** (Edge/Chrome) — em *Fonte de dados*, clique em
**Selecionar pasta da rede** e escolha a pasta com as planilhas. O painel lê todas as
subpastas, une os anos e passa a conferir a cada minuto se algum arquivo mudou; quando
a planilha é alimentada, os números se atualizam sozinhos enquanto a aba está aberta.
Ao reabrir o painel, se a permissão da pasta ainda estiver válida ele relê tudo sozinho;
caso contrário mostra a cópia salva e o botão **Reconectar à pasta**.

**2. Selecionar arquivos ou arrastar** — funciona em qualquer navegador. Depois de
alimentar a planilha, clique em **Atualizar**.

**3. Consolidação agendada pelo Python** — para não depender de ninguém abrir a pasta:

```bat
python consolidar.py "\\servidor-exemplo\financeiro\Pagamentos" --saida "\\servidor-exemplo\financeiro\Pagamentos\Dashboard"
```

Isso grava o `dados.js` ao lado do `index.html`, e todo mundo que abrir o painel já vê
os dados prontos — sem cópia local, sem permissão de pasta. Agende no Agendador de
Tarefas do Windows (por exemplo, a cada hora):

- Programa: `python`
- Argumentos: `C:\scripts\consolidar.py "\\servidor-exemplo\financeiro\Pagamentos" --saida "\\servidor-exemplo\financeiro\Pagamentos\Dashboard"`

Requisitos do script: `pip install pandas openpyxl xlrd`.

## Estrutura esperada da planilha

Uma linha por pagamento. Formato esperado:

| Coluna da sua planilha | Como é usada | Obrigatória |
|---|---|---|
| `Data` | ano, mês, semana do mês e dia da semana | sim |
| `Beneficiario` | agrupador padrão das três páginas | sim |
| `Valor` | montante somado nas células e totais | sim |
| `Favorecido` | segundo agrupador, no seletor **Agrupar por** | não |
| `Banco` | filtro da página de valores | não |
| `Departamento` | filtro universal | não |
| `Descriçao das Naturezas` | filtro universal **Naturezas** e detalhe de cada pagamento | não |

Os nomes são reconhecidos automaticamente, com ou sem acento, e o mesmo vale para
variações comuns (`DT PGTO`, `Empresa`, `Montante`, `Setor`, `Centro de custo`,
`Natureza da despesa`…). Se algo não for reconhecido, abra *Fonte de dados ›
**Ajustar colunas*** e escolha a coluna certa; a escolha fica salva no navegador e
vale para toda planilha com o mesmo cabeçalho.

Detalhes que o leitor resolve sozinho:

- Cabeçalho fora da primeira linha (títulos e linhas em branco no topo).
- Várias abas por arquivo; abas de instrução são ignoradas.
- Um arquivo por ano (`Pagamentos_2025.xlsx`, `Pagamentos_2026.xlsx`) ou subpastas por ano.
- Valores como texto (`R$ 1.234,56`), datas em `dd/mm/aaaa` ou data do Excel.
- Linhas de `TOTAL` no fim da planilha são descartadas.

## Beneficiário ou Favorecido

Quando a base traz as duas colunas, o seletor **Agrupar por** na barra superior troca o
eixo das três páginas de uma vez. Por `Beneficiario`, a linha é o fornecedor que
originou a despesa. Por `Favorecido`, pagamentos de fornecedores diferentes que caem
na mesma conta — factoring, cooperativa, escritório de cobrança — somam na mesma linha,
o que muda quem aparece acima do limite de alerta. Ao abrir uma linha na página de
frequência, cada pagamento mostra o nome do outro lado, e o nome só é repetido quando
os dois campos diferem.

## As três páginas

**Valores por dia** — empresas nas linhas, segunda a sexta nas colunas, total da semana
por empresa na última coluna e total geral por dia no rodapé. A coluna da empresa fica
travada durante a rolagem horizontal. Filtro por banco. Uma coluna *Fim de semana*
aparece só quando existem pagamentos em sábado ou domingo, para nenhum valor ficar de
fora da soma.

**Frequência** — quantidade de transações por empresa, com selo *Único* ou *Múltiplo*.
Clique na linha para ver todos os pagamentos daquela empresa na semana. Filtros:
*Apenas 1* / *Mais de 1*.

**Status** — tem duas visões, no seletor à esquerda do limite:

- *Por empresa*: quantidade de pagamentos e total da semana por empresa, com chip
  ⬤ Normal abaixo do limite e ⚠ Alerta acima. Botão *Ver pagamentos* leva à página de
  frequência já filtrada naquela empresa.
- *Por pagamento*: um lançamento por linha, com status, data (marcada como *Hoje*
  quando for o dia corrente), natureza, banco e valor. Linhas em alerta recebem fundo
  avermelhado, e a lista é paginada de 50 em 50.

O limite vem em R$ 10.000,00 e pode ser alterado no campo ao lado. Os botões
*Ver todos* / *Normal* / *Alerta* filtram por faixa — na visão por empresa a
comparação é com o total da semana, na visão por pagamento é com o valor de cada
lançamento.

### Conferência do que está em alerta

Toda linha marcada como **Alerta** — seja uma empresa, seja um pagamento — ganha uma
lista suspensa na coluna **Conferência**, com três situações:

- **Pendente** — ainda não olhado. É como tudo começa.
- **Autorizado** — conferido e em ordem. A linha fica esverdeada.
- **Verificar com o departamento** — é preciso checar a autorização e a solicitação com
  quem pediu o pagamento. A linha fica amarelada.

Na visão *Por empresa*, a marca é sempre guardada na **semana** a que os pagamentos
pertencem, nunca no recorte que está na tela. Isso significa que:

- marcar numa semana continua valendo quando você abre o mês inteiro ou o ano;
- marcar com o filtro em *Todos* aplica a marca a todas as semanas daquele recorte de
  uma vez — a tela avisa quantas foram;
- a linha só aparece como **Autorizado** quando todas as semanas que ela representa
  estão autorizadas. Se faltar alguma, a célula mostra "3 de 5 semanas conferidas";
  se qualquer semana estiver como *a verificar*, a linha inteira pede verificação.

A mesma empresa pode, portanto, estar autorizada em março e pendente em abril, sem que
você precise remarcar ao trocar de filtro. Quando a empresa também tem pagamentos que
passam do limite sozinhos, a célula mostra abaixo quantos são e como estão
("3 pagamentos acima do limite · 1 autorizado").

Na visão *Por pagamento*, a lista aparece em cada lançamento acima do limite; os demais
mostram "dentro do limite".

Cada marca registra data, hora e o nome preenchido em *Fonte de dados › Conferência
por*. O filtro **Conferência** na barra da página isola pendentes, a verificar ou já
autorizadas, na visão que estiver aberta. O indicador **Conferência pendente** conta
quantas empresas (ou pagamentos, conforme a visão) ainda não foram tratadas, e o rodapé
mostra o andamento: "2 de 3 em alerta conferida(s)".

As marcas ficam guardadas por computador, separadas da cópia dos dados: sobrevivem a
uma nova leitura da planilha, à troca de pasta e ao botão *Esquecer dados salvos*. Cada
lançamento é identificado por data, empresa, valor, banco e natureza — não pela linha —
então a planilha pode ser reordenada ou receber lançamentos novos sem perder o que já
foi conferido. Para zerar tudo existe *Limpar conferências* em *Fonte de dados*.

O registro é local: é uma lista de trabalho da equipe que usa aquele computador, não
uma trilha de auditoria compartilhada. Quem abrir o painel em outra máquina começa com
tudo pendente. O CSV exportado inclui as colunas de conferência, responsável e data,
e é por ele que a informação circula.

Filtros universais na barra fixa: ano, mês, semana do mês, agrupamento, departamentos
e naturezas. Quando algum filtro está ativo aparece **Limpar filtros**, que zera tudo
menos o ano. Cada página tem o botão **Exportar**, que gera um CSV já no padrão do
Excel em português.

## Aparência

Menu lateral e barra superior em gradiente azul-escuro (#0f172a → #1d3a80), área de
trabalho em cinza claro, cards translúcidos com borda fina. O primeiro indicador de
cada página é um card escuro em destaque; quando ano, mês e semana estão selecionados,
ele mostra a variação contra a semana anterior (▲/▼ em %), calculada pela data — a
comparação atravessa a virada de mês corretamente.

O botão de lua na barra superior alterna para o tema escuro completo, e a escolha fica
salva no navegador de cada pessoa. Em telas de até 1000 px o menu lateral vira uma
faixa de abas no topo, e a barra de filtros continua fixa ao rolar.

## Como a semana do mês é contada

As semanas são blocos de segunda a domingo, e o dia 1º está sempre na primeira semana.
Quando o mês começa num sábado ou domingo, esses dias soltos entram na primeira semana
em vez de formarem uma semana própria — por isso todo mês tem no máximo cinco semanas,
da *Primeira* à *Quinta*.

Exemplos de 2026:

| Mês | 1º cai em | Primeira semana | Última semana |
|---|---|---|---|
| Março | domingo | 1 a 8 | Quinta: 30 e 31 |
| Agosto | sábado | 1 a 9 | Quinta: 31 |
| Junho | segunda | 1 a 7 | Quinta: 29 e 30 |
| Fevereiro | domingo | 1 a 8 | Quarta: 23 a 28 |

O filtro de semana lista apenas as semanas que têm pagamento no mês selecionado, então
uma semana sem lançamento não aparece na lista.

## Sobre a arquitetura

O pedido original previa Flask/Django com React. Como o requisito de execução é abrir
um `index.html` salvo na rede, sem servidor, o processamento que seria do pandas roda
no próprio navegador (leitura do XLSX, normalização, agregação por semana). O `consolidar.py`
mantém o caminho Python para quem quiser a consolidação agendada no servidor. Os dados da planilha nunca saem da máquina;
apenas a biblioteca de leitura (SheetJS) é baixada de um CDN.

## Problemas comuns

**"Nenhum dado carregado"** — o painel abre em *Fonte de dados*. Escolha a pasta ou os arquivos.

**Uma aba aparece como "sem dados"** — o cabeçalho não foi reconhecido. Use *Ajustar colunas*.

**O botão "Selecionar pasta da rede" está desabilitado** — navegador sem suporte
(Firefox/Safari). Use *Selecionar pasta (todos navegadores)* ou o `consolidar.py`.

**Abri o painel e os dados são de ontem** — é a cópia salva, e é o comportamento
esperado. Clique em **Atualizar** para reler a planilha. A barra lateral sempre mostra
a data e a hora da leitura que está na tela.

**Trocamos a planilha de lugar** — vá em *Fonte de dados* e escolha a nova pasta. O
caminho antigo é substituído na hora.

**"Falha ao ler a pasta: The request is not allowed…"** — a permissão de leitura da
pasta caducou. O navegador guarda o caminho, mas por segurança só devolve o acesso
depois de um clique seu; isso acontece sempre que a página é recarregada ou o navegador
é reaberto. Vá em *Fonte de dados* e clique em **Reconectar à pasta**, confirmando em
"Ver arquivos". Se a permissão for negada, o painel recarrega pelos arquivos já abertos
na sessão, quando houver.

**A permissão expira todo dia e isso incomoda** — nesse caso o caminho mais tranquilo
é o `consolidar.py` agendado no servidor: o `dados.js` fica pronto na pasta e o painel
abre preenchido sem pedir permissão nenhuma.

**"Por pagamento" com o filtro Alerta não traz nada** — provavelmente nenhum pagamento
*isolado* passou do limite no período, embora o *total da semana* de alguma empresa
tenha passado. São contas diferentes: a visão por empresa soma a semana, a visão por
pagamento olha lançamento a lançamento. A tela informa o maior pagamento do recorte
nesse caso.

**A lista some depois que clico em "Ver pagamentos"** — essa ação aplica um filtro pelo
nome da empresa, que vale para o painel inteiro e fica visível no campo de nome, na
barra de filtros. Use **Limpar filtros** para voltar a ver todos.

**Os valores não batem com a planilha** — confira em *Ajustar colunas* se a coluna de
valor é a certa (planilhas com `Valor Bruto` e `Valor Pago` juntas podem enganar a
detecção automática), verifique se há pagamentos fora do mês filtrado e se o
agrupamento está em `Beneficiario` ou `Favorecido`.

**Uma empresa aparece duas vezes com grafias diferentes** — o agrupamento usa o texto
exato da planilha. `Empresa Alfa ME` e `Empresa Álfa ME` são duas linhas. Vale
padronizar a coluna na origem ou usar validação de dados na planilha.
