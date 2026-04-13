Título: Automação de Conferência de Notas de Empenho (NC)

Descrição: Ferramenta desenvolvida para otimizar o fluxo de conferência e cadastro de NCs no sistema SILOMS. O script valida as informações e automatiza a entrada de dados, reduzindo o erro humano e o tempo de execução em 50%.

📂 Estrutura de Dados (CSV)
O script processa as notas de empenho a partir de um arquivo chamado Cred recebido (relatorio para cadastro no SILOMS).csv.
* Importante: Este arquivo deve conter colunas específicas para o processamento de valores, datas e tipos de nota.
* Modelo: Um arquivo de exemplo (exemplo_nc.csv) foi incluído neste repositório com dados fictícios para demonstrar a estrutura necessária.
Destaque Técnico:

    Lógica de verificação para evitar duplicidade de cadastros.

    Tratamento de exceções para lidar com instabilidades do sistema governamental.

    Segurança: Uso de variáveis de ambiente para gestão de credenciais.
