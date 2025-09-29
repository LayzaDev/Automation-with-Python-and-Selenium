# Automação com Python e Selenium  

## Descrição  
Este projeto automatiza o acesso a um site específico (ex.: SINAN) para realizar tarefas como login, navegação e extração de informações. Ele foi desenvolvido para otimizar processos repetitivos e reduzir erros manuais.  

## Funcionalidades:  

### Funcionalidades implementadas:  
- [x] Login automático usando credenciais específicas.  
- [x] Abertura de um perfil de usuário do Chrome previamente configurado.  
- [x] Navegação automatizada entre páginas do site.  
- [x] Extração e manipulação de dados de arquivos .dbf.
- [x] Criação de uma nova planilha Excel e inserção dos dados extraídos nela.
- [x] Abertura da planilha 'Daily Reports' já existente no computador.   

## 🛠 Tecnologias  
As seguintes ferramentas foram utilizadas no desenvolvimento do projeto:  

**Linguagem:** Python  
**Bibliotecas e Frameworks:** Selenium
**Ambiente:** Jupyter Notebook, Anaconda  
**Navegador:** Google Chrome  

## Estrutura do Projeto:  

### Automação de Login:  
A automação inicia carregando um perfil específico do Chrome para autenticação e acesso ao site.  

### Navegação:  
O script realiza cliques automatizados e insere informações nos campos necessários do site.  

### Extração de Dados:  
Os dados são extraídos diretamente das páginas do site e processados usando a biblioteca Pandas.  

### Relatórios:  
Com os dados processados, o programa gera relatórios automatizados que podem ser salvos localmente ou enviados para outros sistemas.  

## Como Executar  
1. Certifique-se de que o Python e o Anaconda estão instalados no sistema.  
2. Instale as dependências necessárias:  
   ```bash  
   pip install selenium pandas  
   ```  
3. Configure o perfil do Chrome que será utilizado:  
   - Abra o Chrome.  
   - Digite `chrome://version` na barra de endereços.  
   - Copie o caminho do diretório do perfil de usuário e atualize no script.  
4. Execute o script principal:  
   ```bash  
   python main.py  
   ```
## Veja a automação funcionando na prática:
[Acesse o link](https://drive.google.com/file/d/1ohVG0MiZkgCebL4HyFB8aL8ue2DtJO9S/view?usp=drive_link)

