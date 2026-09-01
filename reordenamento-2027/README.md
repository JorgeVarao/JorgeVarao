# Reordenamento 2027 — SEDUC-PI / SUPEX / UGERF / GDI

Três versões da aba **Reordenamento 2027** no mesmo arquivo, para comparar antes de decidir.

| Arquivo | O que é |
|---|---|
| `Reordenamento_2027_v1_v2_v3.xlsx` | A planilha pronta, com as três versões e todas as abas de apoio |
| `Reordenamento_2027.gs` | O script para colar em **Extensões → Apps Script** |
| `gerador/` | O código Python que montou o `.xlsx` e os testes que conferem o `.gs` |

---

## As três versões

**`Reordenamento 2027` (V1)** — retrato exato da aba que existe hoje na *Proposta
Reordenamento 2027 SEDUC - PI*. Nada foi recalculado. Serve de ponto de comparação.

**`Reordenamento 2027 V2`** — as mesmas 38 colunas da V1, com o que estava sem
funcionar já funcionando.

**`Reordenamento 2027 V3`** — a V2 mais os pedidos novos: anexos, fusão entre
escolas, UETEP, UEJA, fundamental e 2ª/3ª série editáveis. 48 colunas.

---

## O que estava quebrado na V1 e foi corrigido na V2

**Ainda faltam no município** devolvia `—` em todas as 576 linhas. O município
chegava acentuado da *Base GDI* (`BOM PRINCÍPIO DO PIAUÍ`) e o *Panorama
Municipal* guarda o nome sem acento (`BOM PRINCIPIO DO PIAUI`), então o PROCV
nunca encontrava. Agora as duas pontas usam o nome sem acento, e a coluna
responde ao que for digitado na 1ª série.

**Panorama Municipal** tinha `#ERROR!` nas colunas *Turmas decididas*, *Saldo* e
*Cobertura*. As fórmulas usavam coluna inteira (`$C:$C` e `$M:$M`), o que fecha
um ciclo de dependência no Google Sheets. Agora o intervalo é limitado às linhas
que realmente existem.

**Salas necessárias 2027** estava preenchida em 1 das 576 linhas. Como *Situação
da sala* e *Resumo 2027* dependem dela, as duas saíam vazias
(`Necessidade:  salas | `). Agora todas vêm calculadas pelo motor e continuam
editáveis.

**Possíveis fusões de turmas** devolvia `—` na coluna inteira. Voltou a ser
calculada, e a coluna ao lado conta quantas turmas saem e quantas chegam segundo
a aba *Fusão Turmas*.

**IDEB, ANEXOS e CURSOS** existiam no cabeçalho e nunca foram preenchidas. Agora
saem das abas *IDEB Municípios*, *Base Anexos* e *Base UETEP*.

---

## O que a V3 acrescenta

### Anexos
Três perguntas — *manter? encerrar? remanejar a oferta?* — na V3 por escola e na
aba **Base Anexos** uma linha por anexo, que é onde a decisão realmente cabe:
96 escolas têm anexo e algumas têm até 5. Cada anexo também ganhou a coluna
**Natureza do anexo**: prédio matriz, anexo, sala externa, comunidade/povoado,
sala cedida em outra escola, privação de liberdade, socioeducativo.

### Fusão entre escolas
Aba nova **Fusão Entre Escolas**. Cruza o mesmo curso de 1ª série entre escolas
do mesmo município e, quando os alunos somados cabem em menos turmas do que
existem hoje, sugere a fusão. Traz para a análise as matrículas dos dois lados,
as turmas de fundamental, o 9º ano e o IDEB do município. O destino é a escola
com mais alunos no curso; no empate, a que tem mais salas. **138 sugestões**
envolvendo 87 escolas.

A aba **Fusão Turmas**, que já existia (fusão turma a turma), continua intacta.

### UETEP
A base **Matriculas por etapa** entrou no arquivo. Turma sem enturmados e com
pré-matrícula é oferta nova: **466 ofertas novas, 1.276 matrículas**. Essas
matrículas viram turmas previstas (teto de `matrículas ÷ 40`, curso a curso) e
entram na conta de salas de 2027.

### UEJA
Aba nova **Base UEJA**. Separa a oferta de EJA por onde ela acontece, lendo o
curso e o nome do anexo. 535 linhas, uma por escola × local.

### Fundamental
O **Panorama Municipal** agora soma o 9º ano da própria rede estadual com o que
vem das outras redes do município, e recalcula as turmas necessárias a partir
dessa demanda total:

```
9º ano estadual (aba Turmas)  +  9º ano outras redes (Censo − estadual)
        = demanda total de 1ª série
        → turmas necessárias = teto(demanda ÷ 40)
```

A coluna **Ainda faltam no município** lê o saldo e responde a cada turma que for
digitada, em tempo real.

### Geral
2ª e 3ª série viraram colunas ★ editáveis. Cada turma que você muda mexe na hora
na coluna **Salas necessárias 2027 (calculado)**:

```
salas = base calculada pelo motor
        + (1ª decidida − 1ª projetada)
        + (2ª decidida − 2ª projetada)
        + (3ª decidida − 3ª projetada)
```

A base já respeita a regra física de turno — integral ocupa a sala o dia inteiro,
manhã e tarde dividem a mesma sala (vale o maior dos dois), noturno regular fica
à parte, EJA não entra.

---

## Como ler as cores

**Amarelo com ★** é seu: preencha à mão. Todo o resto é calculado e volta a ser
reescrito quando o script roda.

| Cor | Bloco |
|---|---|
| Cinza escuro | Identificação |
| Azul | Retrato de 2026 |
| Ciano | Fundamental e 9º ano |
| Verde | Ensino médio e séries |
| Verde-azulado | UETEP |
| Roxo | Fusões |
| Âmbar | Salas |
| Marrom | EJA |
| Violeta | Anexos |

---

## Abas do arquivo

| Aba | Linhas | O que traz |
|---|---:|---|
| `LEIA-ME` | — | Resumo do que mudou |
| `Reordenamento 2027` | 576 | V1 — como está hoje |
| `Reordenamento 2027 V2` | 576 | V2 — corrigida |
| `Reordenamento 2027 V3` | 576 | V3 — corrigida e ampliada |
| `Base Tratada` | 576 | Uma linha por prédio matriz — fonte de quase tudo |
| `Base Anexos` | 146 | Uma linha por anexo, com as colunas de decisão |
| `Panorama Municipal` | 227 | Demanda de 1ª série por município e o contador de turmas |
| `Fusão Turmas` | 62 | Fusão turma a turma (aba que já existia) |
| `Fusão Entre Escolas` | 138 | Sugestão nova, por município |
| `Base UETEP` | 497 | Oferta 2027 por escola: 1ª, 2ª, 3ª, subsequente e salas |
| `UETEP · Oferta Integral 2027` | 1.352 | 1ª série de 2027, escola × curso |
| `UETEP · Continuidade 2027` | 2.496 | 2ª e 3ª série, uma linha por turma |
| `UETEP · Subsequente 2027` | 48 | Oferta subsequente |
| `Base UEJA` | 535 | Onde cada oferta de EJA acontece |
| `IDEB Municípios` | 227 | **Para preencher à mão** |
| `Cursos` | 44 | Lista oficial de 2027 |
| `Validações` | 11 | Listas dos menus suspensos |
| `Turmas` · `Base GDI` · `Matriculas por etapa` | — | Dados brutos |

---

## Instalando o script

1. Abra a planilha no Google Sheets
2. **Extensões → Apps Script**
3. Apague o que estiver lá e cole `Reordenamento_2027.gs` inteiro
4. Salve, volte para a planilha e recarregue a página (F5)
5. Use o menu **📊 GDI / UGERF**

Na primeira execução o Google pede autorização — é o seu próprio script rodando
na sua própria planilha.

| Item do menu | Quando usar |
|---|---|
| ⟳ **Atualizar tudo** | Sempre que a aba *Turmas* mudar |
| ① Atualizar só as bases | Para conferir as bases sem mexer nas versões |
| ② Reescrever só as fórmulas da V2 e V3 | Se alguma fórmula for apagada sem querer |
| ③ Criar/atualizar Validações | Se as listas dos menus mudarem |
| ↩ Desfazer último curso da célula | Tira o último item do menu acumulativo de cursos |

**Regra de ouro: o script nunca apaga coluna marcada com ★.** A única exceção é a
primeira carga: coluna ★ numérica que estiver vazia recebe a projeção como ponto
de partida. As decisões de anexo, de fusão e de EJA são lidas antes da
atualização e devolvidas depois.

### Abas de entrada que precisam existir

| Aba | Colunas |
|---|---|
| `Turmas` | A INEP · B Escola · C Anexo · D Curso · E Etapa · F Organização · G Período · H Turno · I Enturmados · J Turmas |
| `Base GDI` | A INEP · C GRE · D Município · E Escola · F Nº salas · N Projeção 2026 · O Projeção 2027 · R Escola mais próxima |
| `Matriculas por etapa` | E INEP · H Turno · J Etapa · M Turma · O Curso · Q Enturmados · S Pré-matrícula |
| `Panorama Municipal` | A Município · B 9º ano todas as redes · C Turmas necessárias · G Fonte |

---

## O que ainda precisa ser preenchido à mão

**IDEB Municípios** — os 227 municípios já estão listados, as notas estão em
branco. Enquanto não forem preenchidas, a coluna IDEB mostra
`— sem IDEB cadastrado` e a aba de fusão mostra `— sem IDEB`. Assim que as notas
entrarem, as duas passam a mostrar o valor sozinhas.

---

## A base real da UETEP (segunda entrega)

A primeira versão trazia uma aba `Base UETEP` que eu **inferi** da aba
`Matriculas por etapa`, com a regra *"turma sem enturmados e com pré-matrícula é
oferta nova"*. A base de verdade não estava em nenhum dos arquivos de origem, e
eu devia ter dito isso na entrega em vez de só chamar a aba de `Base UETEP`.

A regra pegava lixo: 6 dos 30 cursos que ela classificava como oferta nova de
2027 não existem na oferta real — entre eles `AGRICULTOR FAMILIAR` no CETI Darcy
Ribeiro, uma linha sem nome de turma, com 26 pré-matrículas que nunca viraram
turma. Resíduo de 2026.

A base correta é a `OFERTA_2027_BASE.xlsx`, e ela não é "matrículas novas para
prever sala" — é **a oferta de 2027 inteira, já decidida**:

| Aba | Linhas | O que é |
|---|---:|---|
| `UETEP · Oferta Integral 2027` | 1.352 | 1ª série de 2027 por escola × curso — 1.454 turmas, 496 escolas |
| `UETEP · Continuidade 2027` | 2.496 | Uma linha por turma de 2ª e 3ª série — 1.254 + 1.242 turmas, 72.713 alunos |
| `UETEP · Subsequente 2027` | 48 | 58 turmas, 2.030 alunos, 41 escolas |

Isso substitui a regra de progressão (1ª→2ª, 2ª→3ª):

| | hoje 2026 | projeção antiga | oferta 2027 | dif |
|---|---:|---:|---:|---:|
| 1ª série | 1.156 | 1.156 | **1.454** | +298 |
| 2ª série | 1.170 | 1.156 | **1.254** | +98 |
| 3ª série | 1.208 | 1.170 | **1.242** | +72 |
| Total | 3.534 | 3.482 | **3.950** | +468 |

A projeção só acertava em 255 das 576 escolas.

### Salas com a oferta real

```
integral = 1ª + 2ª + 3ª + subsequente (a oferta da UETEP é integral)
           + fundamental integral + outras turmas de EM integrais
manhã e tarde dividem a mesma sala → vale o maior dos dois
noite fica à parte · EJA não entra
```

A coluna `Salas 2027 · composição`, na Base Tratada e na Base UETEP, mostra as
quatro parcelas para conferência.

### Três coisas na base que valem conferência

- **Linha de total no rodapé** de duas abas, sem INEP (`Oferta Integral` linha
  1354 = 1.454/58.160; `Subsequente` linha 50 = 58). O script descarta.
- **`PREVISÃO DE ALUNOS` é sempre turmas × 40** — é capacidade, não demanda.
  Serve para dimensionar sala, não para analisar fusão. Os alunos reais só
  existem na aba de continuidade.
- **Dois casos de escola + curso repetidos**, os dois no CETI Governador Freitas
  Neto (Desenvolvimento de Sistemas 2+2, Administração 1+2). Turnos diferentes
  ou duplicação?

E o **ECOESCOLA THOMAS A KEMPIS** tem 1 turma de 1ª série hoje e não aparece na
oferta de 2027. Das 80 escolas sem oferta integral, é a única com ensino médio
hoje — as outras 79 são unidades de EJA e fundamental.

---

## Preservação do trabalho manual

A segunda entrega foi aplicada **sobre** a planilha em uso, não regerada. O que
mudou, e só isso:

| Aba | Colunas tocadas |
|---|---|
| `Base Tratada` | S, T, U (1ª série de 2027) e AB (salas) mudaram de fonte; 9 colunas novas ao fim |
| `Reordenamento 2027 V2` | ★M e ★X re-semeadas; 4 colunas de referência ao fim |
| `Reordenamento 2027 V3` | ★O, ★S, ★T, ★AE re-semeadas; U/V/W repontadas; 2 colunas ao fim |

Regra da re-semeadura: a célula ★ só é atualizada se ainda for **exatamente** o
valor que a entrega anterior semeou. Qualquer célula digitada é preservada e
listada. Na conferência, uma decisão real foi encontrada e preservada — V3,
★ 1ª SÉRIE, INEP 22136703 = 4.

Ficaram intocadas: `IDEB - MUNICIPIO`, `IDEB - ESCOLAS`, `Base Anexos (1)` e as
fórmulas que vocês escreveram nela, `antiga base anexos`, `Panorama Municipal`
coluna O, `Fusão Turmas`, `Fusão Entre Escolas`, `Base UEJA`, `Turmas`,
`Matriculas por etapa`, `Cursos`, `Validações` e a aba V1.

O script também **parou de gerenciar a Base Anexos**: se `Base Anexos (1)`
existir, ele não encosta nela.

### Backup e reversão

A aba `★ Backup antes da UETEP` guarda o valor anterior das seis colunas ★.
Para desfazer, é copiar de volta.

### Base GDI

A aba era um `IMPORTRANGE` que o download em .xlsx congelou em
`=IFERROR(__xludf.DUMMYFUNCTION("IMPORTRANGE(...)"), valor)`. Essa fórmula não
existe no Sheets — o que aparece é sempre o argumento de reserva, e em parte das
células a reserva está vazia: 21 observações longas só viviam no cache do
arquivo e sumiriam em qualquer conversão, com ou sem mim.

A aba foi gravada como valor, e a aba `Como restaurar o IMPORTRANGE` traz a
fórmula original para religar a origem quando quiserem.

---

## Como isto foi conferido

O `.xlsx` foi montado por um gerador Python e o `.gs` foi escrito depois, com a
mesma lógica. Para garantir que os dois concordam, o `.gs` roda em Node com um
`SpreadsheetApp` de mentira e o resultado é comparado com o do gerador:

```
cd gerador
python3 extrai.py          # lê os dois arquivos de origem
python3 gera.py            # monta o .xlsx
python3 valida.py          # confere abas, intervalos e índices de PROCV
node    teste_gs.js        # roda a lógica do .gs sobre os mesmos dados
node    teste_form2.js     # captura as fórmulas que o .gs escreveria
```

Resultado da última conferência:

- 28.345 fórmulas verificadas estruturalmente — nenhum intervalo fora da aba,
  nenhum índice de PROCV maior que a largura do intervalo
- 6 abas de apoio geradas pelo `.gs` **idênticas** às do `.xlsx`, linha a linha
- 27.072 fórmulas da V2 e da V3 + 1.273 do Panorama e da Fusão **idênticas**
  entre `.gs` e `.xlsx`
- os valores foram calculados de fato (biblioteca `formulas`) numa amostra de
  120 escolas, conferindo o texto de cada coluna

---

## Parâmetros

| Constante | Valor | Onde pesa |
|---|---:|---|
| `MAX_T` | 40 | Capacidade máxima de uma turma — usada nas fusões e na UETEP |
| `CAP_1SERIE` | 40 | Alunos por turma ao converter 9º ano em 1ª série |

Os dois estão no topo do `.gs` e em `gerador/motor.py`.
