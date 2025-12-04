# 🤖 PROJETO_TERMOS: Automação de Reenvio de Termos (VCX)

Este projeto automatiza a navegação no sistema VCX para abrir links de ativos e reenviar e-mails de movimentação, garantindo que o clique seja realizado mesmo em situações de carregamento complexo da página.

## 🚀 Como Executar

Este projeto utiliza o modo de depuração (debug) do Google Chrome, permitindo que a automação se conecte a uma instância do navegador já aberta e configurada (por exemplo, já logada).

### 1. Pré-requisitos

* **Python:** Versão 3.x instalada.
* **Selenium:** Instale as bibliotecas necessárias:
    ```bash
    pip install selenium openpyxl
    ```
    *(Nota: `openpyxl` é assumido para a leitura do seu arquivo Excel, se necessário.)*
* **ChromeDriver:** O driver do Chrome (geralmente gerenciado automaticamente pelo Selenium, mas precisa corresponder à versão do seu Chrome).
* **Dependências de Arquivo:** Você precisa ter o script `abrir_chrome_debug.bat` na raiz do seu projeto.

### 2. Iniciar o Chrome em Modo Debug

Antes de rodar o script Python, o Chrome deve ser iniciado na porta 9222.

* **Arquivo `abrir_chrome_debug.bat` (Conteúdo esperado):**
    ```bat
    "C:\Program Files\Google\Chrome\Application\chrome.exe" --remote-debugging-port=9222 --user-data-dir="C:\ChromeTemp"
    ```
    *(Ajuste o caminho do Chrome se necessário.)*

* **Ação:** O script `main.py` executa este `.bat` **automaticamente** na inicialização da classe `Automator`.

### 3. Execução Principal

Execute o script principal (`main.py` ou o arquivo que contém a função `main`):

```bash
python src/main.py
