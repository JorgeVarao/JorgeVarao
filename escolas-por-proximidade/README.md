# Escolas por Proximidade — Teresina

Ferramenta da **GDI/UGERF · SEDUC-PI** que responde, em segundos, a duas perguntas
que chegam toda semana:

- *quais escolas da rede estadual estão a até X km do bairro Y?*
- *quais escolas estão a até X km da escola Y?*

Roda inteiramente no Google Apps Script sobre a planilha de infraestrutura escolar.
Sem API de mapas paga, sem chave, sem custo: a coordenada de cada unidade sai do
**Plus Code** que já existe na base, decodificado localmente.

> A base de dados **não** faz parte deste repositório (veja `.gitignore`).

---

## Arquivos

| Arquivo | Onde vai | O que faz |
|---|---|---|
| `AppsScript_GeoEscolas.gs` | Apps Script (arquivo de script) | Decodifica Plus Code e publica as funções de planilha |
| `AppsScript_WebApp.gs` | Apps Script (arquivo de script) | Publica a página web e monta o pacote de dados |
| `Index.html` | Apps Script (arquivo **HTML**, nomeado `Index`) | A página em si |

## Instalação / atualização

1. **Extensões > Apps Script** na planilha.
2. Cole o conteúdo dos dois `.gs` nos arquivos de script correspondentes.
3. Em **Arquivos > +  > HTML**, crie/abra o arquivo chamado exatamente `Index`
   (sem `.html`) e cole o conteúdo de `Index.html`.
4. **Implantar > Gerenciar implantações > editar > Nova versão**.
5. Na planilha, menu **Geo-Escolas > Atualizar página web agora** (limpa o cache
   de 10 minutos e faz a página refletir a planilha na hora).

O `Index.html` deste repositório lê os dados ao vivo pelo scriptlet
`<?!= dadosApp() ?>`. Ele só funciona dentro do Apps Script — abrir o arquivo
direto no navegador não carrega dado nenhum.

---

## O que mudou nesta versão (v5)

### 1. Nova aparência
Layout institucional: faixa de cores do Governo, cabeçalho com a marca da SEDUC,
painel de filtros à esquerda em cartões (Localização, Situação, Etapa de Ensino,
Turno) e resultados em cartões à direita, com a “Resposta Pronta” ao lado.

O brasão é um SVG desenhado no próprio arquivo. Para usar a arte oficial, preencha
a constante `BRASAO_URL` no início do script do `Index.html` — com uma URL pública
ou um `data:image/png;base64,...` — que o SVG é substituído automaticamente.

O mapa é um painel de 400x250 no canto superior direito. Ele serve de referência
visual; quem decide é a lista. Cada pino continua clicável e abre a ficha.

**Ver rota** usa o formato de caminho do Google Maps (`/maps/dir/lat,lon/lat,lon`),
sem query string. O formato `?api=1&origin=...&destination=...` caía numa página
de erro do Drive quando aberto de dentro do iframe do Apps Script — enquanto o
link de **Ver no mapa**, que tem um único parâmetro, sempre funcionou. Sem query
string não há o que ser reescrito no caminho.

### 2. Gestor e contato (colunas AN e AM)
`NOME DO GESTOR` e `CONTATO` agora aparecem:

- no cartão de cada escola na lista de resultados;
- na ficha que abre ao clicar no pino do mapa;
- na ficha do ponto de partida, quando a busca parte de uma escola;
- no texto da “Resposta Pronta” e no CSV exportado.

O telefone vira link de ligação (`tel:`) e, quando é celular, ganha um link de
WhatsApp. Há também um filtro novo — **Com gestor cadastrado** — para separar o
que já foi levantado do que falta levantar.

As colunas são localizadas **pelo nome do cabeçalho**, com uma lista de apelidos
aceitos (`CONTATO`, `TELEFONE`, `NOME DO GESTOR`, `DIRETOR`…) e, se nenhum casar,
pela letra da coluna (`AM` e `AN`). Ou seja: renomear o cabeçalho não quebra a
leitura. O menu **Geo-Escolas > Conferir colunas de gestor/contato** mostra de qual
coluna o script está lendo e quantas linhas estão preenchidas.

### 3. Distância a partir de uma escola específica
A pergunta “escolas a X km da escola Y” virou um modo próprio.

**Na página:** aba **Por Escola** no painel de Localização. Digite parte do nome
(ou o código INEP) e a lista sai ordenada pela distância até aquela unidade. A
escola de partida ganha um cartão “Ponto de partida” — com gestor e contato — e
fica fora da lista, para não aparecer como 0 m. Também dá para partir de qualquer
pino do mapa pelo link **“Escolas perto daqui”** na ficha.

**Na planilha:**

```
=ESCOLAS_PROXIMAS_DA_ESCOLA("CETI FIRMINA SOBREIRA"; 2)
=DISTANCIA_ENTRE_ESCOLAS("CETI FIRMINA SOBREIRA"; "CEJA GAYOSO E ALMENDRA")
=GESTOR_DA_ESCOLA("CETI FIRMINA SOBREIRA")
=CONTATO_DA_ESCOLA("CETI FIRMINA SOBREIRA")
```

---

## Funções de planilha

| Função | Para quê |
|---|---|
| `PLUSCODE_LAT(codigo)` / `PLUSCODE_LON(codigo)` | Geolocaliza a base a partir do Plus Code |
| `DISTKM(lat1; lon1; lat2; lon2)` | Distância em linha reta, em km |
| `ESCOLAS_PROXIMAS(origem; raioKm; somenteAtivas; limite)` | Unidades perto de um bairro, coordenada, Plus Code ou escola |
| `ESCOLAS_PROXIMAS_DA_ESCOLA(escola; raioKm; somenteAtivas; limite)` | Unidades perto de uma escola específica |
| `DISTANCIA_ENTRE_ESCOLAS(escolaA; escolaB)` | Distância entre duas unidades da base |
| `GESTOR_DA_ESCOLA(escola)` / `CONTATO_DA_ESCOLA(escola)` | Consulta rápida das colunas novas |

`ESCOLAS_PROXIMAS` e `ESCOLAS_PROXIMAS_DA_ESCOLA` devolvem, além do que já
devolviam, duas colunas ao final: **Nome do gestor** e **Contato**.

---

## Duas ressalvas

**Distância é em linha reta** (Haversine). Em Teresina isso pesa: dois pontos
separados pelo Parnaíba ou pelo Poti podem estar a 2 km em linha reta e a 8 km de
percurso real, dependendo da ponte. A lista serve de triagem — confirme o trajeto
quando a decisão depender disso.

**Contato de gestor é dado de trabalho.** A página passou a exibir nome e telefone
de servidor. Se a implantação estiver como *“Qualquer pessoa”*, vale restringir ao
domínio da SEDUC — ou zerar `CFG_EXTRA.CONTATO` em `AppsScript_GeoEscolas.gs`.
