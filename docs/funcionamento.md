# Funcionamento do Sistema

## 1. Recebimento das inscrições

O Google Forms permite que os moradores informem:

- Nome completo;
- Número do apartamento;
- Ambientes em que desejam participar.

O formulário permite selecionar um ou mais ambientes.

## 2. Tratamento das respostas

A função `aoEnviarFormulario(e)` é executada automaticamente quando uma nova resposta é recebida.

Ela padroniza os textos digitados em letras maiúsculas e solicita a atualização das listas.

## 3. Identificação da última inscrição

A função `atualizarListas_()` percorre as respostas e utiliza um `Map` para armazenar a inscrição mais recente de cada apartamento.

Quando encontra uma nova resposta para o mesmo apartamento, substitui os dados anteriores.

Assim, as escolhas antigas não são acumuladas.

## 4. Organização dos participantes

As inscrições válidas são distribuídas em quatro abas:

- Gourmet – Natal;
- Festas – Natal;
- Gourmet – Ano Novo;
- Festas – Ano Novo.

Cada lista contém nome, apartamento e data da inscrição.

O processamento considera o ano atual e o prazo de inscrição até 10 de dezembro.

## 5. Realização dos sorteios

O sistema permite realizar sorteios individuais ou de todos os ambientes.

A função `sortearAmbiente_()`:

1. Atualiza as listas;
2. Verifica se já existe ganhador ativo;
3. Identifica candidatos disponíveis;
4. Realiza a seleção aleatória;
5. Registra o resultado no histórico;
6. Atualiza a aba de resultados.

A seleção utiliza `Math.random()`.

## 6. Registro de desistência

A função `registrarDesistencia()` permite informar a desistência de um ganhador.

Após a confirmação:

1. O registro anterior é atualizado para a situação "Desistiu";
2. O apartamento é excluído somente daquele ambiente;
3. Um novo sorteio é realizado;
4. O resultado atualizado é registrado.

## 7. Histórico e resultados

O sistema mantém duas abas específicas:

**Resultados dos Sorteios:** apresenta o ganhador atual de cada ambiente ou informa que o sorteio está pendente.

**Histórico dos Sorteios:** mantém os registros dos sorteios e das desistências, permitindo acompanhar as alterações realizadas.

## 8. Principais funções do código

| Função | Responsabilidade |
|---|---|
| `onOpen()` | Cria o menu personalizado |
| `criarAbas()` | Cria as abas do sistema |
| `criarFormulario()` | Cria e configura o formulário |
| `aoEnviarFormulario(e)` | Processa o envio de respostas |
| `normalizar_()` | Padroniza textos para comparação |
| `localizarAbaRespostas_()` | Localiza a aba oficial de respostas |
| `atualizarListas()` | Executa a atualização manual |
| `atualizarListas_()` | Processa e distribui as inscrições |
| `sortearTodos()` | Executa os sorteios dos ambientes |
| `sortearAmbiente_()` | Realiza o sorteio de um ambiente |
| `registrarDesistencia()` | Registra desistência e solicita novo sorteio |
| `atualizarResultados_()` | Atualiza a tabela de resultados |
