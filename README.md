# Automação de Briefing para Visitas em Campo

## Sobre o projeto

Este projeto foi desenvolvido para apoiar a gestão durante visitas em campo, automatizando a elaboração de briefings a partir de dados previamente consolidados e validados.

A solução solicitao município que deseja analisar. A partir dessa informação, o código realiza a leitura dos dados, aplica filtros, realiza cálculos definidos e preenche automaticamente um documento previamente estruturado.

## Problema

A elaboração dos briefings era realizada de forma manual, o que gerava muito tempo gasto com análise e variações no resultado dependendo do contexto da análise e do operador responsável pelo preenchimento.

A aplicação de diferentes filtros e critérios por diferentes operadores aumentava o risco de inconsistências nas informações apresentadas nos briefings.

## Solução

Foi desenvolvida uma automação em Python para padronizar o processo de elaboração dos briefings.

O código utiliza como fonte um arquivo Excel contendo dados previamente consolidados e validados. A partir do município informado pelo usuário, a aplicação:

1. Identifica os dados correspondentes ao município;
2. Realiza os filtros necessários;
3. Executa os cálculos definidos para a análise;
4. Insere os resultados nos campos variáveis do documento;
5. Gera o briefing preenchido em formato `.docx`.

Dessa forma, a mesma lógica de análise é aplicada independentemente do operador responsável pela geração do documento.

## Resultado

A automação padroniza o processo de elaboração dos briefings, reduzindo a variação causada pelo preenchimento manual e pela aplicação de diferentes filtros ou critérios de análise.

A partir de um mesmo conjunto de dados consolidados e validados, o código aplica uma lógica única para seleção, tratamento e cálculo das informações. Isso aumenta a consistência e a confiabilidade dos dados apresentados nos briefings utilizados durante as visitas em campo.

Além de reduzir a necessidade de intervenção manual, a solução mantém a estrutura previamente definida do documento e preenche automaticamente os campos variáveis.

## Tecnologias utilizadas

- Python
- Excel
- Pandas
- Python-docx

## Fluxo da solução

```text
Arquivo Excel
     ↓
Dados consolidados e validados
     ↓
Seleção do município
     ↓
Filtros e cálculos
     ↓
Preenchimento automático do template
     ↓
Briefing em formato .docx
