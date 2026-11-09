# Sistema de Sorteio de Salões de Condomínio

**Sistema de automação de inscrições e sorteios para espaços de uso coletivo em condomínios residenciais.**

## Sobre o projeto

Este projeto foi desenvolvido para automatizar o processo de inscrições e sorteios dos espaços de uso coletivo de um condomínio durante as festividades de Natal e Ano Novo.

A solução utiliza ferramentas do Google para receber inscrições, organizar participantes, realizar sorteios, registrar desistências e manter um histórico dos resultados.

O objetivo é reduzir atividades manuais, facilitar a administração e proporcionar maior organização e rastreabilidade ao processo.

## Tecnologias utilizadas

- **Google Apps Script:** linguagem JavaScript e automação dos processos.
- **Google Forms:** formulário para recebimento das inscrições.
- **Google Sheets:** armazenamento e gerenciamento das informações.
- **GitHub:** versionamento e documentação do código-fonte.

## Funcionalidades

1. Criação automática do formulário de inscrição.
2. Criação das abas de gerenciamento.
3. Recebimento automático das respostas.
4. Padronização dos textos das respostas em letras maiúsculas.
5. Identificação de inscrições duplicadas por apartamento.
6. Consideração somente da última inscrição enviada por apartamento.
7. Organização dos participantes em quatro listas independentes.
8. Sorteio individual por ambiente.
9. Sorteio de todos os ambientes.
10. Registro de desistências e realização de novo sorteio.
11. Exclusão do desistente somente do ambiente correspondente.
12. Atualização dos resultados.
13. Preservação do histórico de sorteios.

## Ambientes contemplados

| Ambiente | Capacidade |
|---|---:|
| Salão Gourmet – Natal | 25 pessoas |
| Salão de Festas – Natal | 50 pessoas |
| Salão Gourmet – Ano Novo | 25 pessoas |
| Salão de Festas – Ano Novo | 50 pessoas |

## Regras de negócio

- Cada apartamento possui uma inscrição válida por processo.
- Quando houver mais de uma resposta, será considerada a última enviada, substituindo integralmente as escolhas anteriores.
- Cada ambiente possui seu próprio sorteio.
- Um ganhador ativo impede a realização de outro sorteio naquele mesmo ambiente.
- Em caso de desistência, o apartamento é excluído somente do ambiente correspondente.
- O novo sorteio utiliza a lista do mesmo ambiente.
- Os registros anteriores permanecem no histórico.

## Estrutura do projeto

- `Code.gs`: código principal do Google Apps Script.
- `docs/funcionamento.md`: explicação das funcionalidades e regras.
- `docs/instalacao.md`: instruções para configuração e utilização.

## Como funciona

O morador preenche o formulário e seleciona os ambientes dos quais deseja participar.

As respostas são armazenadas no Google Sheets. O Apps Script processa os registros, identifica a inscrição mais recente de cada apartamento e atualiza as listas de participantes.

A administração utiliza o menu personalizado da planilha para executar os sorteios e registrar eventuais desistências.

## Autoria

Projeto desenvolvido por mim utilizando Google Apps Script, Google Forms e Google Sheets, com foco na automação de processos administrativos condominiais.
