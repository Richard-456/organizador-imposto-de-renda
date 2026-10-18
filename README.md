# IR em Dia — organizador de imposto de renda em Excel

Ferramenta em Excel para reunir, ao longo do ano, as informações que vão para a declaração do imposto de renda: dados do titular, saldos bancários, rendimentos recebidos e despesas dedutíveis. A ideia é chegar em março com tudo lançado e conferido, em vez de caçar comprovante na última semana.

Projeto do desafio de Excel da DIO (Expert Felipe Aguiar).

**Arquivo:** [`organizador-imposto-de-renda.xlsx`](./organizador-imposto-de-renda.xlsx)

> Todos os dados do arquivo são **fictícios**, usados só para demonstrar o funcionamento. Nome, CPF, endereço e valores são de exemplo.

---

## O problema que a ferramenta resolve

A declaração pede informação que está espalhada: o informe de rendimentos de cada banco, o holerite de cada mês, o recibo do dentista, a nota da escola. Quem não organiza durante o ano faz isso tudo de uma vez, com pressa, e é aí que aparece omissão e malha fina.

A planilha funciona como um caderno estruturado: cada tipo de informação tem seu lugar, as listas são validadas para não entrar texto livre, e um painel mostra o que já foi preenchido e o que falta.

---

## Como está organizada

| Aba | Para que serve |
|-----|----------------|
| **INICIO** | Painel com os totais consolidados e a conferência do preenchimento. Tudo calculado das outras abas. |
| **TITULAR** | Dados pessoais, endereço e situação da declaração. |
| **INFORMES** | Saldo de cada conta bancária em 31/12 e o arquivo do informe que comprova. |
| **NOTAS** | Rendimentos recebidos mês a mês, por categoria. |
| **DESPESAS** | Gastos dedutíveis (saúde, educação, previdência) que reduzem a base. |
| **TABELAS** | Bancos e categorias que alimentam as listas de validação. |

Em todas as abas de uso: linhas de grade e cabeçalhos de linha/coluna escondidos, **menu lateral fixo** à esquerda e o mesmo padrão de cores — **amarelo é campo de digitação, cinza é campo calculado**.

---

## Recursos usados

### Menu de navegação com HIPERLINK

O menu lateral existe em todas as abas e marca com `▸` em qual você está. Cada item é uma fórmula apontando para o destino:

```excel
=HIPERLINK("#INFORMES!C3"; "  Informes")
```

O `#` no começo do endereço é o que diz ao Excel que o destino é um local **dentro do próprio arquivo**, e não um site.

### Links rápidos externos

Logo abaixo do menu, a mesma função abre os portais oficiais em três cliques a menos:

```excel
=HIPERLINK("https://cav.receita.fazenda.gov.br/"; "   Portal e-CAC")
```

São links para o e-CAC, a página Meu Imposto de Renda e o portal da Receita Federal. Endereço de órgão público muda de tempos em tempos; se algum sair do ar, é só editar a fórmula.

### Validação de dados com intervalos nomeados

Nenhuma lista está escrita direto na validação. Todas apontam para um intervalo nomeado da aba `TABELAS`:

| Intervalo nomeado | Alimenta | Itens |
|---|---|---|
| `bancos` | campo BANCO dos quatro informes | 50 |
| `categorias_rendimento` | coluna CATEGORIA da aba NOTAS | 9 |
| `categorias_despesa` | coluna CATEGORIA da aba DESPESAS | 8 |
| `sim_nao` | perguntas de situação da declaração | 2 |

A validação fica escrita assim, com o nome no lugar do endereço:

```excel
Permitir: Lista
Fonte: =bancos
```

Isso resolve um detalhe chato do Excel: validação de dados apontando direto para outra aba (`=TABELAS!$A$4:$A$53`) só funciona em versões recentes. Apontando para um **nome**, funciona em qualquer versão e ainda fica legível. Para incluir um banco novo, basta acrescentar na aba `TABELAS` e esticar o intervalo — nenhuma validação precisa ser mexida.

Cada validação tem mensagem de erro própria, então quem digitar errado é avisado do que fazer.

### Endereço abreviado automático

O programa da Receita corta o endereço em 60 caracteres. Em vez de abreviar na mão, a planilha faz:

```excel
D15 = SE(NÚM.CARACT(D14)>60; ESQUERDA(D14;60); D14)
D16 = 60-NÚM.CARACT(D15)
```

Você digita o endereço completo em `D14` e a planilha mostra como ele vai ficar, com o contador de caracteres restantes ao lado.

### Painel consolidado

A aba `INICIO` não tem digitação nenhuma além do ano de referência. Ela puxa os totais das outras abas e faz a conferência:

```excel
D11 = INFORMES!D6                      ' saldo total em contas
D12 = NOTAS!D6                         ' rendimentos lançados
D13 = DESPESAS!D6                      ' despesas dedutíveis
D14 = D12-D13                          ' rendimentos menos deduções
D20 = CONT.VALORES(bancos informados) - CONT.VALORES(arquivos anexados)
```

A última é a que mais ajuda na prática: ela conta quantos bancos foram informados **sem** o arquivo do informe correspondente. Enquanto não estiver zero, falta documento.

---

## O que mudei em relação à ferramenta do Expert

A ferramenta da aula (LION APP) tem três abas — titular, informes e notas — com o menu montado em **formas com hiperlink** e a aba de tabelas oculta. Parti dessa estrutura e mudei:

- **Menu em fórmula `HIPERLINK`, não em formas.** Forma é desenho solto sobre a planilha: desalinha quando a linha muda de altura e some se alguém copiar o intervalo. Em fórmula, o menu é conteúdo de célula, acompanha a formatação e dá para replicar nas abas novas copiando e colando.
- **Aba `INICIO`**, que não existe na original: painel com totais consolidados e conferência de pendências.
- **Aba `DESPESAS`**, também nova. A original só registra entrada; sem as deduções a base não fecha.
- **Validação por intervalo nomeado** em vez de referência direta à aba de tabelas, pela compatibilidade explicada acima.
- **Quatro bancos** nos informes em vez de três, e o total passou para o topo da aba.
- **Endereço abreviado calculado**, em vez de digitado duas vezes.
- **CPF, CEP e telefone como texto**, não como número. Como número o Excel come o zero à esquerda e vira `123456789` no lugar de `012.345.678-90`.
- **Categorias de rendimento ampliadas** de três (holerite, CNPJ, freelance) para nove, incluindo 13º, férias, aluguel e dividendos.
- **Links rápidos** apontando para os portais da Receita, não para redes sociais.

---

## Como usar

1. Apague os dados de exemplo das abas `TITULAR`, `INFORMES`, `NOTAS` e `DESPESAS`.
2. Preencha o titular uma vez por ano.
3. Lance os rendimentos e as despesas conforme forem acontecendo — o painel atualiza sozinho.
4. Guarde os PDFs dos informes na mesma pasta da planilha e escreva o nome do arquivo no campo correspondente.
5. Antes de declarar, confira na aba `INICIO` se "informes sem arquivo anexado" está zerado.

Para deixar com cara de aplicativo fechado, dá para ocultar a aba `TABELAS` (botão direito na guia → Ocultar). Deixei visível de propósito, para quem abrir o arquivo conseguir ver como as listas funcionam.

---

## Prints

![Painel da declaração](./images/inicio.png)

![Informes bancários com a lista de validação](./images/informes.png)

---

## Tecnologias

Microsoft Excel — função `HIPERLINK`, `SE`, `ESQUERDA`, `NÚM.CARACT`, `SOMA`, `CONT.VALORES`, `CONT.NÚM`; intervalos nomeados; validação de dados com mensagem de erro personalizada; congelamento de painéis; interface com grade e cabeçalhos ocultos.
