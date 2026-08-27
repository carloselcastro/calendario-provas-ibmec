# Prompt para converter o quadro de horários em CSV

Copie o texto abaixo (tudo dentro do bloco), anexe o seu PDF ou planilha e
envie para a LLM (Claude, ChatGPT, Gemini).

Depois, confira o resultado com a checklist da seção 4 do
[README](README.md) antes de rodar o script.

---

```
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
```

---

## Se o arquivo for muito grande

Divida por curso e peça um CSV por vez. Depois junte tudo num arquivo só,
mantendo **uma única linha de cabeçalho** no topo. É mais confiável do que
pedir 200 linhas de uma vez.

## Prompt de conferência (opcional)

Depois de montar o CSV, vale uma segunda passada com a LLM:

```
Confira este CSV contra o documento original que te enviei antes e me diga:
1. Alguma disciplina do documento ficou de fora?
2. Alguma linha tem dia ou horário diferente do documento?
3. Algum nome de curso ou de professor está grafado de forma inconsistente
   entre linhas (maiúsculas/minúsculas, acentos, abreviações)?
Responda só com as divergências encontradas.
```
