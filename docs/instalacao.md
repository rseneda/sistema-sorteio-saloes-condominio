# Instalação e Configuração

## Requisitos

- Conta Google;
- Google Sheets;
- Google Apps Script.

## 1. Criar a planilha

Crie uma planilha no Google Sheets com o nome desejado.

## 2. Abrir o Apps Script

Na planilha, acesse:

**Extensões → Apps Script**

## 3. Inserir o código

Abra o arquivo principal do projeto e substitua seu conteúdo pelo código disponível em `Code.gs`.

## 4. Conferir a configuração

No início do código, confira:

- Nome do formulário;
- Nome da aba oficial de respostas;
- Opções de sorteio;
- Nomes das abas do sistema.

A configuração `ABA_RESPOSTAS` deve corresponder exatamente ao nome da aba que recebe as respostas do Google Forms.

## 5. Executar a configuração

Salve o código e execute a função `criarFormulario()` pela primeira vez.

O Google solicitará as permissões necessárias para acessar o Forms, Sheets e Apps Script.

Após a autorização, o sistema criará o formulário, as abas e o gatilho automático de recebimento das respostas.

## 6. Utilizar o sistema

Após recarregar a planilha, estará disponível o menu:

**SORTEIO DOS SALÕES**

Por meio dele, será possível:

- Criar o formulário;
- Atualizar as listas;
- Sortear todos os ambientes;
- Realizar sorteios individuais;
- Registrar desistências e realizar novos sorteios.

## Observações

- O código deve ser configurado na planilha que receberá as inscrições.
- Não altere os nomes das abas sem atualizar as configurações correspondentes.
- O histórico deve ser preservado para manter os registros dos sorteios.
- A publicação do código no GitHub não publica automaticamente a planilha nem o formulário.
- Para utilizar o sistema em outra planilha, será necessário realizar a instalação e autorizar as permissões novamente.
