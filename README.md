Web Scraper de Diário Oficial
Este projeto é um web scraper automatizado desenvolvido em Python, utilizando Selenium para interação com páginas web e Pandas/Openpyxl para manipulação e armazenamento de dados em arquivos Excel.

Descrição
O script realiza a extração de informações de uma página específica do Diário Oficial do Estado de Alagoas. Ele busca informações relacionadas a clientes com base em um código CACEAL e um intervalo de datas especificado pelo usuário.

Funcionalidades
Busca automatizada: O script realiza buscas no site do Diário Oficial com base no código CACEAL de cada cliente.
Manipulação de dados: Os resultados das buscas são armazenados em uma lista e, posteriormente, salvos em arquivos Excel.
Relatórios detalhados: O script salva os resultados tanto utilizando Pandas quanto Openpyxl, gerando arquivos Excel com os dados extraídos.
Interface de linha de comando: Permite que o usuário insira o período de busca diretamente pela linha de comando.
Tecnologias Utilizadas
Python: Linguagem principal do projeto.
Selenium: Utilizado para automatizar a navegação e extração de informações da web.
Pandas: Utilizado para manipulação e salvamento de dados em arquivos Excel.
Openpyxl: Utilizado para manipulação de planilhas Excel.
Tqdm: Utilizado para exibir uma barra de progresso durante a execução do script.
Como Usar
Instale as dependências:

Utilize pip para instalar as bibliotecas necessárias:
bash
Copiar código
pip install selenium pandas openpyxl tqdm webdriver_manager
Execute o script:

Certifique-se de ter o arquivo clientes_caceal.xlsx no mesmo diretório que o script. Esse arquivo deve conter uma coluna chamada caceal e outra chamada Razao_social.
Execute o script e insira o período inicial e final quando solicitado:
bash
Copiar código
python nome_do_script.py
Resultados:

Os resultados serão salvos automaticamente em arquivos Excel, com o nome dados_encontrados_dd-mm-YYYY.xlsx, onde dd-mm-YYYY representa a data de execução do script.

Observações
Certifique-se de ter o Chrome instalado e atualizado, pois o webdriver_manager será utilizado para gerenciar o driver do Chrome automaticamente.
O script foi desenvolvido para ser executado em uma máquina local com Python instalado.
