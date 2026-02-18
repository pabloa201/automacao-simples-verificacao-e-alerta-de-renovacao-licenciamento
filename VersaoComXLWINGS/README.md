Esta versão foi desenvolvida utilizando a biblioteca Xlwings, ideal para cenários onde há necessidade de interação direta com a interface do Microsoft Excel. Diferente da versão Openpyxl, este método permite que o script manipule a planilha em tempo real, aproveitando recursos nativos do aplicativo instalado.

ATENÇÃO: É necessário ter o Microsoft Excel instalado na máquina para o funcionamento desta versão.

Funcionamento:
Observação de Execução: A aplicação é projetada para rodar de forma automática via Agendador de Tarefas do Windows.

Integração com Office: O script realiza a abertura do Excel e da planilha de licenciamento de forma automatizada.

Coleta de Dados: Realiza a leitura das colunas e linhas específicas, mapeando as informações em variáveis Python.

Persistência de Dados: Salva automaticamente o arquivo após a conferência para garantir a integridade dos dados.

Saída em Terminal: Exibe os resultados e o prazo de vencimento com um temporizador de 5 segundos para validação visual rápida.

Geração de LOG: Cria um arquivo .txt com os dados processados para fins de auditoria e histórico de monitoramento.

Trigger de Alerta: Realiza a verificação da régua de prazos (40, 30, 20, 15, 10, 5, 3 e 1 dia). Se identificado o critério, dispara um e-mail preventivo aos responsáveis.

Requisitos e Preparação:
Ambiente: Python 3.x e Microsoft Excel instalado.

Bibliotecas: xlwings (Instale com: pip install xlwings).

SMTP: Configuração de servidor de e-mail para alertas.

Transformando em Executável (.exe):
Para facilitar a implementação em servidores, utilize a biblioteca PyInstaller:

Instale: pip install pyinstaller

Gere o executável: pyinstaller --onefile main.py

O arquivo final estará disponível na pasta dist/ do seu projeto e pode ser renomeado conforme a necessidade.
