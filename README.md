# Calendário de Provas — gerador automático por otimização

Gera o calendário de provas (AP1, AP2, AS) de vários cursos ao mesmo tempo,
distribuindo as disciplinas pelas datas disponíveis de modo a **minimizar o
número de alunos com mais de uma prova no mesmo dia** e **nunca colocar duas
provas da mesma turma no mesmo horário**.

A saída é uma planilha de trabalho com o calendário completo e um arquivo
`.xlsx` por curso no layout institucional, pronto para distribuição.

---

## Índice

1. [O que o método faz](#1-o-que-o-método-faz)
2. [Instalação](#2-instalação)
3. [Passo 1 — Transformar seu PDF ou planilha em CSV (com uma LLM)](#3-passo-1--transformar-seu-pdf-ou-planilha-em-csv-com-uma-llm)
4. [Passo 2 — Conferir o CSV](#4-passo-2--conferir-o-csv)
5. [Passo 3 — Configurar o script](#5-passo-3--configurar-o-script)
6. [Passo 4 — Rodar](#6-passo-4--rodar)
7. [O que sai](#7-o-que-sai)
8. [Regras de negócio que você pode ajustar](#8-regras-de-negócio-que-você-pode-ajustar)
9. [Como ler os avisos](#9-como-ler-os-avisos)
10. [Problemas comuns](#10-problemas-comuns)

---

## 1. O que o método faz

### O problema

Montar calendário de provas na mão é tentativa e erro: você encaixa uma
disciplina, descobre que criou choque em outro curso, desfaz, tenta de novo.
Com ~200 disciplinas, 6 cursos e 5 datas, o número de combinações possíveis
passa de qualquer coisa que dê para conferir manualmente.

### A modelagem

O script trata isso como um problema de **otimização combinatória**, resolvido
com o **CP-SAT**, o solver de programação por restrições do
[Google OR-Tools](https://developers.google.com/optimization).

**Oferta.** Cada disciplina realmente distinta vira uma *oferta*, identificada
por `código + turma + agenda semanal + professor`. Uma disciplina que aparece
em três cursos com nomes diferentes, mas mesma turma, mesmo horário e mesmo
professor, é **uma oferta só** — e portanto **uma prova só**, que conta
simultaneamente para os alunos dos três cursos.

**Grupo de alunos.** Cada `curso + período + turma` é um grupo. É o grupo que
sofre choque: duas provas no mesmo dia para a turma AA e para a turma AB do
mesmo período **não** é conflito, porque nenhum aluno faz as duas.

**Variáveis de decisão.** Para cada oferta e cada data/horário em que ela
*poderia* ter prova, existe uma variável binária: essa prova acontece aqui,
sim ou não. A restrição básica é que exatamente uma dessas variáveis seja
verdadeira por oferta.

**Opções elegíveis.** Uma disciplina só pode fazer prova num dia em que ela
tem aula. Se ela tem aula na quarta e na sexta, a prova é na quarta *ou* na
sexta — não em qualquer outro dia. (A exceção está na
[seção 8](#dias-substitutos-quando-falta-um-dia-da-semana).)

### O que ele otimiza, em ordem de prioridade

1. **Duas provas da mesma turma no mesmo dia E no mesmo horário.** Peso
   esmagador (`1_000_000_000`): é fisicamente impossível para o aluno, então o
   solver só aceita se não existir absolutamente nenhuma alternativa.
2. **Provas da mesma turma no mesmo dia**, em horários diferentes. Penalidade
   crescente: a 2ª prova do dia custa 100, a 3ª custa 10.000, da 4ª em diante
   100.000 cada. O efeito é que o solver espalha as provas e só concentra
   quando não tem escolha.
3. **Concentração por professor.** *Sem piorar em nada o resultado dos
   alunos*, agrupa as provas de cada professor no menor número de dias
   possível e aproxima esses dias. Isso é resolvido em duas etapas: o solver
   primeiro acha o melhor calendário para os alunos, **congela** esse valor, e
   só então otimiza os professores dentro do que sobrou de liberdade. Um
   critério nunca atropela o outro.

### Regras rígidas adicionais

- **Turmas conjuntas.** Se duas disciplinas com códigos diferentes têm o mesmo
  professor, a mesma agenda semanal **e a mesma sala**, elas são a mesma aula
  ofertada em cursos diferentes com nomes diferentes. O script amarra as duas
  à mesma data e ao mesmo horário — o professor não pode estar em duas salas
  ao mesmo tempo.
- **Disciplinas sem data compatível** não são inventadas em lugar nenhum: vão
  para a aba `Não Alocadas` com o motivo, para tratamento manual.

---

## 2. Instalação

Requer **Python 3.10 ou superior**.

```bash
git clone <url-do-seu-repositorio>
cd calendario-provas-ibmec
pip install -r requirements.txt
```

Se preferir conda:

```bash
conda create -n calendario python=3.12
conda activate calendario
pip install -r requirements.txt
```

Para conferir se está tudo certo, rode o exemplo que acompanha o repositório
(ver [seção 6](#6-passo-4--rodar)).

---

## 3. Passo 1 — Transformar seu PDF ou planilha em CSV (com uma LLM)

O script **não lê o quadro de horários da sua unidade diretamente**. Cada
unidade tem o seu formato — PDF da grade, planilha do acadêmico, exportação do
sistema — e não há como prever todos.

A ponte é uma LLM (Claude, ChatGPT, Gemini). Você entrega o arquivo que tem e
pede o CSV no formato abaixo. É um trabalho de transcrição, que a LLM faz bem,
mas **o resultado precisa ser conferido** ([seção 4](#4-passo-2--conferir-o-csv)).

### O formato de destino

Uma linha por **disciplina × turma × curso**. Se a mesma disciplina é ofertada
para dois cursos, são duas linhas.

| Coluna | Obrigatória | O que é |
|---|---|---|
| `codigo` | sim | Código da disciplina (`IBM0048`, `MAT101`...) |
| `materia` | sim | Nome da disciplina |
| `turma` | sim | Código numérico da turma na grade (`8001`) |
| `cod_turma` | recomendada | Código da **turma real de alunos** (`AA`, `AB`, `AC`) |
| `sala` | recomendada | Sala da prova; `xx` se não souber |
| `dia_1` | sim | Dia da 1ª aula semanal: `Segunda`…`Sexta` |
| `hora_1` | sim | Hora da 1ª aula, `HH:MM` |
| `dia_2` | sim (pode ficar vazia) | Dia da 2ª aula semanal |
| `hora_2` | sim (pode ficar vazia) | Hora da 2ª aula |
| `periodo` | sim | Período/semestre, número inteiro |
| `professor` | sim | Um nome por linha |
| `curso` | sim | Nome do curso |

> **Por que `cod_turma` importa tanto.** É ele que separa os grupos de alunos
> dentro de um mesmo período. Sem ele, o script trata o período inteiro como
> uma turma só e passa a acusar choques que não existem. A `turma` numérica
> **não** substitui: a mesma turma de alunos pode cursar disciplinas sob
> números diferentes quando a oferta é compartilhada.

### O prompt

Copie o texto abaixo, anexe o seu PDF ou planilha e envie. Ele também está em
[`PROMPT_LLM.md`](PROMPT_LLM.md), para copiar sem a formatação do README.

````text
Você vai converter um quadro de horários acadêmico em um arquivo CSV com um
formato exato. Vou anexar o arquivo original (PDF ou planilha).

## Formato de saída

CSV, codificação UTF-8, separado por vírgula, com esta linha de cabeçalho
exatamente, nesta ordem:

codigo,materia,turma,cod_turma,sala,dia_1,hora_1,dia_2,hora_2,periodo,professor,curso

## Granularidade

Uma linha por disciplina × turma × curso. Se a mesma disciplina aparece em
dois cursos, gere DUAS linhas — iguais em tudo, menos na coluna `curso` (e no
`periodo`, se o período for diferente em cada curso).

## Regras por coluna

- `codigo`: código da disciplina, como está no documento. Não invente.
- `materia`: nome da disciplina. Preserve marcadores entre parênteses como
  "(ed)" (estudo dirigido) e "(e)" (eletiva) — eles são significativos.
- `turma`: o código numérico da turma na grade (ex.: 8001). Se não existir,
  deixe vazio.
- `cod_turma`: o código da TURMA REAL DE ALUNOS, normalmente duas letras
  (AA, AB, AC). É o que distingue duas turmas do mesmo período. Se o
  documento não tiver essa informação, deixe a coluna vazia em TODAS as
  linhas — nunca preencha só algumas.
- `sala`: sala/laboratório da aula. Se não houver, escreva xx. Para aulas
  remotas, use o rótulo do documento (ex.: Teams).
- `dia_1` / `dia_2`: dias da semana em que a disciplina tem aula. Use
  exatamente uma destas palavras, com acento: Segunda, Terça, Quarta,
  Quinta, Sexta, Sábado. Se a disciplina tem só um encontro semanal, deixe
  `dia_2` e `hora_2` vazios. Se tem três ou mais encontros, use os dois mais
  representativos e me avise no fim.
- `hora_1` / `hora_2`: horário de INÍCIO da aula, formato 24h HH:MM
  (07:30, 09:50, 13:30, 15:50, 18:40, 20:40). Converta "7h30", "7:30-9:20"
  ou "07h30min" para 07:30. Nunca coloque o intervalo, só o início.
- `periodo`: número inteiro do período/semestre (1, 2, 3...). Sem "º".
- `professor`: UM nome por linha. Se houver dois professores para a mesma
  aula, use apenas o responsável principal e me avise no fim. Mantenha a
  grafia consistente para o mesmo professor em todas as linhas.
- `curso`: nome do curso, escrito de forma idêntica em todas as linhas
  daquele curso.

## Regras gerais

- NÃO invente nenhum dado. Se um campo não existe no documento, deixe-o
  vazio (ou xx, no caso da sala) e me avise no fim.
- NÃO acrescente, remova ou agrupe disciplinas. Cada oferta do documento
  vira uma linha.
- Se um valor tiver vírgula (ex.: "Hackathons, Innovation & Challenges"),
  coloque o campo entre aspas duplas.
- Não use ponto e vírgula como separador.

## O que me entregar

1. O CSV completo, em um bloco de código, pronto para eu salvar como
   `grade.csv`.
2. Depois do CSV, uma lista curta do que ficou duvidoso: campos que você
   deixou vazios, disciplinas com mais de dois encontros semanais, aulas com
   dois professores, horários que não se encaixaram nos padrões, e qualquer
   linha que você tenha achado inconsistente no documento original.
````

### Se o seu arquivo for muito grande

Divida por curso e peça um CSV por vez, depois junte os arquivos mantendo uma
única linha de cabeçalho no topo. É mais confiável do que pedir 200 linhas de
uma vez.

---

## 4. Passo 2 — Conferir o CSV

A LLM erra, principalmente em documentos com células mescladas. Antes de rodar
o script, confira:

- [ ] O número de linhas bate com o número de ofertas do documento original.
- [ ] Os dias da semana estão escritos com acento (`Terça`, não `Terca`) —
      embora o script aceite as duas formas e mais alguns sinônimos.
- [ ] Os horários estão no formato `HH:MM` e correspondem aos horários reais
      da unidade.
- [ ] O nome de cada curso está escrito **exatamente igual** em todas as suas
      linhas. `ENGENHARIA DE PRODUÇÃO` e `Engenharia de Produção` viram dois
      cursos diferentes.
- [ ] O nome de cada professor está grafado de forma consistente.
- [ ] `cod_turma` está preenchido em todas as linhas, ou em nenhuma.
- [ ] Disciplinas compartilhadas entre cursos aparecem uma vez por curso.

O próprio script roda uma validação e avisa sobre professor vazio, `dia_1`
igual a `dia_2`, linhas duplicadas e `cod_turma` preenchido pela metade. Esses
avisos vão para a aba `Validação` da planilha de saída.

---

## 5. Passo 3 — Configurar o script

Abra `gerar_calendario_provas_multicurso_v4_2.py` e vá até o final do arquivo,
no bloco marcado:

```python
# ============================================================
# MAIN — EDITE SOMENTE ESTA PARTE
# ============================================================
```

O que ajustar:

```python
ARQUIVO_GRADE   = "grade.csv"              # o CSV do passo 1
ARQUIVO_SAIDA   = "AP1_MINHA_UNIDADE.xlsx" # planilha de trabalho
NOME_AVALIACAO  = "AP1"                    # AP1, AP2 ou AS
PERIODO_LETIVO  = "2026.2"
PASTA_OFICIAL   = "OFICIAL/MINHA_UNIDADE"  # onde saem os .xlsx por curso
EXPORTAR_OFICIAL = True

DATAS_PROVAS = [
    "2026-09-24",
    "2026-09-25",
    "2026-09-28",
    "2026-09-29",
    "2026-09-30",
]
```

`DATAS_PROVAS` são as datas reais em que haverá prova, no formato
`AAAA-MM-DD`. Cada data vira uma coluna no calendário oficial. Datas repetidas
do mesmo dia da semana (duas quartas, por exemplo) são tratadas como colunas
independentes.

As colunas da tabela mantêm a ordem cronológica, mas giradas para começar na
segunda-feira: um calendário que começa numa quinta é apresentado a partir da
segunda, com a quinta e a sexta no fim.

---

## 6. Passo 4 — Rodar

```bash
python gerar_calendario_provas_multicurso_v4_2.py
```

Para testar a instalação com os dados fictícios que acompanham o repositório,
aponte `ARQUIVO_GRADE` para `exemplo/grade_exemplo.csv` e use as datas
`2026-05-11` a `2026-05-15`. O esperado é terminar com
`Nenhum conflito encontrado` e um aviso de turma conjunta
(`EXE0401 = EXE0402 — sala 201`), que está lá de propósito para demonstrar a
detecção.

A execução leva de alguns segundos a poucos minutos, conforme o tamanho da
grade. No fim, o resumo:

```
Calendário gerado com sucesso!
Objetivo dos alunos (mesmo critério da v3): 0
Professores considerados na concentração (com 2 ou mais provas): 46
Dias adicionais usados pelos professores: 49
Nenhum conflito encontrado entre as ofertas alocadas.
```

**Objetivo dos alunos = 0** significa calendário perfeito: nenhum aluno com
duas provas no mesmo dia. Valores maiores indicam quantos e quão graves são os
acúmulos — cada 100 é uma turma com duas provas num mesmo dia.

---

## 7. O que sai

### A planilha de trabalho (`ARQUIVO_SAIDA`)

| Aba | Conteúdo |
|---|---|
| `Calendário Visual` | Grade por curso/período/turma, com as datas em colunas |
| `Dados do Calendário` | Uma linha por prova alocada — é a fonte para conferências |
| `Conflitos` | Turmas com mais de uma prova no mesmo dia, com gravidade |
| `Diagnóstico` | Métricas gerais da solução |
| `Concentração Professores` | Quantos dias cada professor ficou usando |
| `Ofertas` | As ofertas consolidadas e os grupos que cada uma atende |
| `Não Alocadas` | Disciplinas sem data compatível — tratar manualmente |
| `Excluídas da Alocação` | (ED)/(E), se você tiver ligado a exclusão |
| `Validação` / `Validação Ofertas` | Avisos sobre o CSV de entrada |
| `Parâmetros` | Tudo que foi usado nesta execução |

### Os arquivos oficiais (`PASTA_OFICIAL`)

Um `.xlsx` por curso, no layout institucional, com todas as turmas empilhadas
numa única aba e uma página final de observações. É o que vai para
distribuição.

---

## 8. Regras de negócio que você pode ajustar

Todas ficam no topo do arquivo, na seção `CONFIGURAÇÃO`.

### Dias substitutos: quando falta um dia da semana

Se o calendário de provas não tem sexta-feira, as disciplinas que só têm aula
na sexta ficariam sem data. Em vez de mandá-las para tratamento manual, elas
são realocadas:

```python
DIAS_SUBSTITUTOS_PADRAO = {
    "Sexta": ("Terça", "Quarta", "Quinta"),
}
```

A regra vale **por oferta, não por dia de aula**: só entra em ação quando
*nenhum* dos dias de aula da disciplina tem data no calendário. Uma disciplina
de quarta e sexta simplesmente faz a prova na quarta. O solver escolhe qual dos
dias substitutos usar, pelo mesmo critério de conflito das demais provas.

Ajuste o dicionário para a realidade da sua unidade, ou use `{}` para
desligar.

### Horário das provas noturnas

```python
HORA_NOTURNA_PADRAO  = "18:40"   # noturna no seu próprio dia de aula
HORA_DIA_SUBSTITUTO  = "20:40"   # noturna realocada para outro dia
HORARIOS_NOTURNOS    = {"18:40", "20:40"}
```

Prova de disciplina noturna é sempre às 18:40, mesmo que a aula seja às 20:40.
A exceção é a disciplina realocada para um dia que não é o dela: aí vai para
20:40, porque o 18:40 daquele dia já pertence às disciplinas que têm aula
nele. **Disciplinas diurnas mantêm o horário da aula em qualquer situação** —
a turma delas não tem como fazer prova à noite.

### Estudo Dirigido e Eletivas

```python
EXCLUIR_ED_ELETIVAS = False
```

Com `False` (padrão), disciplinas marcadas com `(ed)` ou `(e)` no nome entram
na alocação como qualquer outra. Com `True`, ficam de fora do solver e vão
para a aba `Excluídas da Alocação`.

### Turmas conjuntas

```python
SALAS_SEM_IDENTIDADE = {"", "x", "xx", "-", "teams", "online", ...}
```

A união automática de disciplinas que são a mesma aula exige **mesma sala**, e
salas desta lista não contam como identificação. Acrescente os rótulos que a
sua unidade usa para "sala indefinida".

### Pesos da otimização

```python
PESO_SEGUNDA_PROVA  = 100
PESO_TERCEIRA_PROVA = 10_000
PESO_QUARTA_OU_MAIS = 100_000
PESO_MESMO_HORARIO  = 1_000_000_000
```

Só mexa se souber o que está fazendo. A escala é propositalmente separada por
ordens de grandeza para que uma prioridade nunca compense a outra.

---

## 9. Como ler os avisos

Durante a execução, o passo 7 imprime:

```
Sem data de Sexta: essas provas podem cair em Terça, Quarta, Quinta
  (as noturnas às 20:40 em vez de 18:40, as diurnas no horário da aula).
```
Informativo: o calendário não tem sexta e a regra de realocação foi acionada.

```
Turma conjunta (mesma aula, uma prova só): IBM1741 = IBM4023 — sala 43.
```
Informativo: os dois códigos foram identificados como a mesma aula e ficarão
sempre na mesma data e horário. **Confira se faz sentido.**

```
ATENÇÃO: Fulano tem IBM0108 (sala 47) e IBM0033 (sala 35)
  no mesmo horário, em salas diferentes. Conferir a grade.
```
O mesmo professor tem duas aulas simultâneas em salas diferentes na **grade de
origem**. Ou é erro de digitação no CSV (e as salas deveriam ser iguais — aí o
script passa a uni-las como turma conjunta sozinho), ou é um problema real da
grade. Não impede a geração do calendário.

---

## 10. Problemas comuns

**`Colunas obrigatórias ausentes: ...`**
O CSV não tem alguma coluna do formato. Confira o cabeçalho, inclusive
maiúsculas e espaços (o script normaliza para minúsculas, mas não adivinha
nomes diferentes).

**`Dia inválido: ...`**
Algum `dia_1`/`dia_2` está fora da lista aceita. São aceitos `Segunda`,
`Terça`, `Quarta`, `Quinta`, `Sexta`, `Sábado`, `Domingo`, com ou sem acento e
com ou sem o sufixo `-feira`.

**`O calendário contém datas sem coluna no modelo oficial`**
Alguma prova caiu numa data que não está em `DATAS_PROVAS`. Normalmente
significa que `DATAS_PROVAS` foi alterada depois de gerar. Rode de novo.

**Muitas disciplinas em `Não Alocadas`**
Os dias de aula delas não coincidem com nenhuma data de `DATAS_PROVAS`.
Acrescente datas, ou configure `DIAS_SUBSTITUTOS` para o dia que está
faltando.

**O objetivo dos alunos ficou alto**
Há mais provas do que a janela comporta. As saídas: mais datas em
`DATAS_PROVAS`, ou aceitar os acúmulos — a aba `Conflitos` mostra exatamente
quais turmas e em que dias.

**Choques que não existem, entre turmas diferentes do mesmo período**
`cod_turma` está vazio ou preenchido pela metade. Veja a aba `Validação`.
