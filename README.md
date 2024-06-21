# Desafio Técnico AAWZ - Automação
Este projeto foi desenvolvido como parte do desafio técnico proposto pela AAWZ para avaliar habilidades em automação de processos utilizando Python. O projeto consiste em duas etapas:

- Desafio Técnico - Etapa 1: Cálculo de Comissões e Validação de Pagamentos
- Desafio Técnico - Etapa 2: Análise de Contrato de Partnership

## Desafio Técnico - Etapa 1: Cálculo de Comissões e Validação de Pagamentos
Nesta etapa, o objetivo é calcular a comissão de cada vendedor e validar os pagamentos realizados, seguindo as regras estabelecidas pela empresa.

### Tarefa 1: Cálculo de Comissões
- Entrada: Primeira aba da planilha “Vendas”.
- Regras de Comissão:
  - Cada vendedor recebe 10% de cada venda.
  - Se a venda foi realizada por um canal online, 20% da comissão do vendedor vai para a equipe de marketing.
  - Se o valor total da comissão do vendedor for maior ou igual a R$ 1.500,00, 10% dessa comissão vai para o gerente de vendas.
- Saída Esperada: Uma tabela com o nome do vendedor, o valor da comissão e o valor que será pago a ele, considerando as regras estabelecidas.

### Tarefa 2: Validação de Pagamentos
- Entrada: Segunda aba da planilha “Pagamentos”.
- Regras de Validação:
  - Comparar os valores pagos aos vendedores com os valores calculados na Tarefa 1.
  - Identificar os pagamentos feitos incorretamente e o valor incorreto transferido.
- Saída Esperada: Uma lista dos pagamentos feitos incorretamente, indicando o vendedor, o valor pago erroneamente e o valor correto que deveria ter sido pago.
## Desafio Técnico - Etapa 2: Análise de Contrato de Partnership
Nesta etapa, o objetivo é extrair o nome de cada sócio e a quantidade de cotas que cada um possui a partir de um contrato de partnership.

- Entrada: [Link]() para o contrato de partnership
- Saída Esperada: Uma tabela ou lista contendo o nome de cada sócio e a quantidade de cotas que cada um possui.
## Conclusão
- Este README fornece uma visão geral do projeto, descrevendo cada etapa do desafio técnico e suas respectivas tarefas. Também inclui links para as entradas necessárias e especifica a saída esperada para cada etapa. Isso facilita o entendimento do projeto e como executar as tarefas propostas.

- Cada tarefa deve ser implementada em Python, utilizando bibliotecas como Pandas, OpenPyXL e outras, conforme necessário.
- Para executar o projeto:
  - Clone este repositório em sua máquina local.
  - Instale as dependências necessárias.
  - Execute os scripts Python correspondentes a cada etapa do desafio, fornecendo as entradas necessárias conforme especificado no README.
  - Verifique a saída gerada para cada tarefa e certifique-se de que atende aos requisitos estabelecidos.
