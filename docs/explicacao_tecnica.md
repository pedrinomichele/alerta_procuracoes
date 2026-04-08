## Visão Geral

O sistema foi desenvolvido para automatizar o controle de vencimento de procurações, utilizando uma base em Excel e envio automatizado de alertas por e-mail.


## Estrutura dos Dados

A base contém as seguintes colunas:

- outorgante
- tipo
- criticidade
- data_vencimento

A coluna de datas é convertida para o formato datetime, permitindo o cálculo preciso dos prazos.


## Lógica de Processamento e Fluxo do Sistema

O sistema executa as seguintes etapas:

1. Leitura da planilha em Excel com pandas  
2. Tratamento e padronização dos dados (datas e criticidade)  
3. Cálculo de dias restantes até o vencimento  
4. Aplicação das regras de alerta  
5. Geração do e-mail em HTML  
6. Envio automático via Outlook  
7. Execução diária via agendador do Windows  



## Regras de Alerta

- ALTA criticidade → alertas até 45 dias antes do vencimento  
- MÉDIA e BAIXA criticidade → alertas até 30 dias antes  

A lógica também considera documentos já vencidos, permitindo acompanhamento de pendências.



## Geração do E-mail

O e-mail é gerado em HTML com:

- tabela estruturada  
- destaque visual por nível de urgência  
- ordenação por proximidade do vencimento  



## Automação

A execução é realizada via Windows Task Scheduler, permitindo o envio automático diário sem intervenção manual.



## Decisões Técnicas

- Uso de Excel como fonte inicial, devido à facilidade de uso pelo usuário final  
- Utilização da biblioteca pandas para manipulação e tratamento dos dados  
- Integração com Outlook, evitando dependência de servidores SMTP e facilitando a execução em ambiente corporativo  