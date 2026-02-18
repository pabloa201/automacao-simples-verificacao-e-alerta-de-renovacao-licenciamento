Sistema de Monitoramento e Alerta de Licenciamento (Python)
Este projeto automatiza a gestão de prazos de licenciamentos, certificados e contratos, eliminando conferências manuais e mitigando riscos de interrupção de serviços por expiração.

A aplicação foi desenhada para ser resiliente e flexível, oferecendo duas abordagens técnicas distintas conforme a necessidade da infraestrutura.

Estrutura do Projeto
O repositório está dividido em duas versões principais:

/versao-openpyxl (Recomendado para Servidores): Funciona em modo headless (sem interface gráfica). Ideal para rodar em servidores Windows Core ou em instâncias onde o Microsoft Excel não está instalado.

/versao-xlwings (Interação com Usuário): Ideal para máquinas de trabalho (workstations) onde o operador precisa que o Excel seja aberto e manipulado visualmente em tempo real.

Principais Diferenciais
Automação via Task Scheduler: Ambas as versões são compatíveis com o Agendador de Tarefas do Windows.

Régua de Alertas Críticos: Sistema de notificação por e-mail nos marcos de 40, 30, 20, 15, 10, 5, 3 e 1 dia(s) antes do vencimento.

Auditoria via Logs: Geração de arquivos .txt para histórico de execuções.

Portabilidade: Instruções inclusas para converter os scripts em executáveis .exe via PyInstaller.

Recursos Utilizadas:
Python 3.x

Openpyxl / Xlwings (Manipulação de dados)

SMTPLib (Mensageria e Alertas)

PyInstaller (Distribuição/Binários)
