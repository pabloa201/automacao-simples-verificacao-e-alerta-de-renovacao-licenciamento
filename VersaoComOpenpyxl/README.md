Esta versão foi projetada para executar em segundo plano, utilizando a biblioteca Openpyxl para leitura direta do arquivo .xlsx sem a necessidade de instanciar a interface gráfica do Excel. É a solução ideal para servidores onde o recurso de GUI não é necessário ou disponível.

Funcionamento:
Ingestão de Dados: O script realiza a leitura dos dados das colunas e linhas pré-definidas na planilha de licenciamento.

Processamento em Memória: Os dados são mapeados em variáveis para análise lógica de prazos.

Saída em Terminal: Exibição imediata dos resultados com temporizador de persistência (5 segundos) para conferência rápida.

LOG: Geração automática de arquivo .txt contendo o status atualizado, permitindo auditoria posterior do monitoramento.

Trigger de Alerta: Verificação condicional baseada em uma régua de alertas críticos (40, 30, 20, 15, 10, 5, 3 e 1 dia). Caso o critério seja atendido, o módulo de mensageria dispara o e-mail preventivo.

Requisitos:
Python 3.x

Bibliotecas: openpyxl

Acesso ao servidor SMTP para envio de alertas.
