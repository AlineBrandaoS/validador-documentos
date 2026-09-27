# Validador de Legibilidade de Documentos

Classifica cada documento (PDF ou foto) como **LEGÍVEL**, **ILEGÍVEL** (com o motivo, para pedir novo envio) ou **REVISAR** (caso duvidoso, para uma pessoa conferir) e gera uma planilha Excel com o resultado.

## Instalação (uma única vez)

1. **Python 3.10 ou mais recente**: baixe em python.org. Na instalação, marque **"Add Python to PATH"**.
2. **Tesseract OCR**: baixe o instalador do Windows em github.com/UB-Mannheim/tesseract/wiki. Na instalação, em *Additional language data*, marque **Portuguese**.
3. **Bibliotecas**: abra o terminal dentro desta pasta e rode:
   ```
   pip install -r requirements.txt
   ```
4. **Chave da IA**: crie uma conta em console.anthropic.com, adicione créditos e gere uma *API key*.
5. Renomeie `.env.exemplo` para **`.env`** e cole a chave em `ANTHROPIC_API_KEY=` (ou deixe a janela do programa pedir a chave na primeira vez).

## Levar para outro computador (ex.: o do trabalho)

1. Instale o Python (python.org, marcando **"Add Python to PATH"**) e, se possível, o Tesseract.
2. Copie a pasta inteira do programa para o outro computador (ex.: `D:\validador_documentos`).
3. Dê dois cliques em **`INSTALAR.bat`**.

O arquivo `.env` guarda a sua chave da IA. Se não copiar, a janela pede a chave na primeira vez.

## Como organizar os documentos

Crie uma subpasta por colaborador (o nome da subpasta vira a coluna "Colaborador" da planilha):

```
Mobilizacao\
    JOAO DA SILVA\   rg.pdf, cnh.jpg, aso.pdf
    MARIA SOUZA\     ctps.pdf, comprovante.png
```

Formatos aceitos: PDF, JPG, JPEG, PNG, TIF, BMP e WEBP.

## Como usar (sem VS Code)

1. **Uma única vez em cada computador:** dê dois cliques em **`INSTALAR.bat`**. Ele confere o Python, instala as bibliotecas e cria o ícone **"Validador de Documentos"** na Área de Trabalho.
2. **No dia a dia:** abra o ícone, clique em **Escolher pasta…**, depois em **▶ Analisar**. A janela mostra o andamento e o placar (legíveis / ilegíveis / revisar) e, no final, abre a planilha no Excel.
   - Na primeira vez, se a chave da IA não estiver configurada, a janela pede para colar a chave e salva sozinha (botão **Chave da IA** para trocar depois).
   - A janela lembra a última pasta usada.

Alternativas: arrastar a pasta em cima de `rodar_validador.bat`, ou pelo terminal:
```
python validador_legibilidade.py "C:\Documentos\Mobilizacao"
```

A planilha `resultado_legibilidade_DATA.xlsx` é salva dentro da própria pasta, com três abas:

| Aba | Conteúdo |
|---|---|
| Resumo | Totais de legíveis, ilegíveis e revisar |
| Documentos | Uma linha por arquivo: colaborador, tipo, status, motivo, certeza da IA |
| Detalhe por página | Métricas técnicas de cada página (nitidez, brilho, OCR…) |

Se você rodar de novo na mesma pasta, os arquivos que já foram analisados **não são cobrados de novo**: o programa guarda o resultado em `_cache_validador.json`.

## Checklist de mobilização (admissão + NRs + RACs)

Além da legibilidade, o programa confere **o que está faltando** para cada colaborador.

1. Monte uma planilha com as colunas **NOME** e **FUNÇÃO**. A coluna **APTO NR35** (SIM/NÃO) é opcional. O botão **Modelo** da janela cria uma pronta.
2. Na janela, escolha essa planilha no campo **2** e clique em **Analisar**.
3. A planilha de resultado ganha três abas novas no começo:
   - **Pendências**: uma linha por documento faltando, ilegível ou a conferir, com "o que fazer".
   - **Checklist**: um colaborador por linha e um documento por coluna (OK / FALTANDO / ILEGÍVEL / REVISAR / — não se aplica).
   - **Regras usadas**: quais NRs e RACs foram exigidos para cada função.

**O que é exigido:**

| Para quem | Documentos |
|---|---|
| Todos | Identificação (RG, CPF ou CNH), CTPS/registro/contrato, comprovante de residência, comprovante de escolaridade, foto, ASO, ficha de EPI, ordem de serviço, NR06 e NR18 |
| Conforme a função | NR11, NR12 (máquinas/equipamentos), NR35 (trabalho em altura ou APTO NR35 = SIM), treinamento do equipamento (operadores) |
| Quem tem RAC | Certificado de cada RAC + PRO de cada RAC + controle de frequência ART + frequência PST |

**Regras por função (`regras_funcoes.xlsx`):** é uma tabela com uma função por linha, onde você marca **X** nas NRs e RACs que ela exige. O botão **Regras das funções** da janela abre essa tabela. Quando aparece uma função nova, a IA sugere as regras e marca a linha em amarelo como **IA - REVISAR**. Confira e troque a ORIGEM para **CONFIRMADA**. Enquanto não for confirmada, a pendência aparece com o aviso "regra sugerida pela IA — confirmar".

**Nome dos arquivos:** siga o padrão do SGC, com uma pasta por colaborador e o tipo no começo do nome (`NR18_FILIPE MATTO.pdf`, `RAC01_...`, `PRO RAC 01_...`, `ASO_...`, `RG_...`, `CTPS_...`, `FICHA EPI_...`, `OS_...`, `ART_...`, `PST_...`, `FOTO_...`, `RESIDENCIA_...`, `ESCOLARIDADE_...`). Se o nome não seguir o padrão, a IA tenta reconhecer o documento pelo conteúdo.

Os documentos da empresa (PCMSO, PGR, ART do PGR) não entram nesse checklist, porque são por empresa e não por colaborador.

## Como a decisão é tomada

| Situação | Resultado |
|---|---|
| Imagem minúscula (menos de 300 px) | ILEGÍVEL direto |
| Página em branco dentro de um PDF | Ignorada (se todas estiverem em branco: ILEGÍVEL) |
| Arquivo corrompido ou PDF com senha | REVISAR |
| IA com certeza menor que 85% | REVISAR |
| IA diz "legível", mas o filtro técnico viu desfoque ou imagem escura | REVISAR |
| IA diz "não é documento" (ex.: selfie, print aleatório) | REVISAR |
| Falha de internet ou da IA | REVISAR (rode de novo depois) |
| IA diz "ilegível" com certeza alta | ILEGÍVEL, com o motivo e os campos que não dá para ler |
| IA diz "legível" com certeza alta e o filtro técnico concorda | LEGÍVEL |

Em PDFs com várias páginas, o arquivo recebe o **pior** status entre as páginas.

## Calibração (faça antes de usar para valer)

1. Separe de 30 a 50 documentos que a equipe já conferiu, em duas subpastas: `legivel\` e `ilegivel\`.
2. Rode:
   ```
   python validador_legibilidade.py --calibrar "C:\teste_calibracao"
   ```
3. O programa mostra:
   - **ERROS GRAVES** (documento ilegível aprovado). A meta é **zero**. Se aparecer algum, aumente `CONFIANCA_MINIMA` no `.env` (ex.: 0.90) ou troque o modelo para `claude-opus-5-5`.
   - **Automação**: a porcentagem decidida sem precisar de revisão humana.
   - **Nitidez por grupo**: ajuda a ajustar o limite de desfoque (`LIMITES` no início do script).

## Custo estimado

Com `claude-sonnet-5`, fica em torno de **US$ 0,01 por página** analisada. É uma estimativa: o valor real depende do tamanho das imagens e dos preços vigentes da Anthropic. Com `--sem-ia`, o programa roda só o filtro técnico e não tem custo, mas quase tudo vai para REVISAR.

## Privacidade

As imagens são enviadas à API da Anthropic para análise. O prompt pede que a IA **não transcreva** dados pessoais na resposta, e a planilha traz apenas o status e o motivo. Não compartilhe o arquivo `.env`.
