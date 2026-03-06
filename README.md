# Fila de Impressão - PDF v1.0 🚀

<p align="center">
  <img src="img/preview.png" alt="Interface do Fila de Impressão PDF" width="600">
</p>

Uma ferramenta desktop intuitiva desenvolvida em Python para gerenciar e automatizar a impressão de arquivos PDF em massa. O programa oferece uma interface moderna com suporte a "Drag and Drop" (arrastar e soltar) e um sistema inteligente de monitoramento de hardware da impressora.

## ✨ Funcionalidades

* **Interface Moderna:** Modo escuro com visual limpo e barra de título personalizada.
* **Arraste e Solte:** Adicione múltiplos arquivos PDF simplesmente arrastando-os para dentro da janela (Drag and Drop).
* **Escuta Ativa de Hardware:** O programa detecta e interrompe o processo automaticamente caso a impressora apresente problemas como:
    * Offline ou Desconectada.
    * Sem papel (Aviso específico).
    * Atolamento de papel.
    * Erro geral de hardware.
* **Feedback Visual:** As linhas da tabela mudam de cor em tempo real (Verde para sucesso, Vermelho para erro).
* **Contador Dinâmico:** Exibição em tempo real de quantos arquivos estão na fila.
* **Relatórios de Fluxo:** O sistema só confirma a conclusão se todos os arquivos chegarem à fila de impressão sem erros.

## 🛠️ Tecnologias Utilizadas

* **Python 3.x**
* **Tkinter:** Interface gráfica principal.
* **TkinterDnD2:** Suporte para arrastar e soltar.
* **PyMuPDF (fitz):** Manipulação e renderização de PDFs.
* **PyWin32:** Comunicação direta com a API de impressão do Windows.
* **Pillow (PIL):** Processamento de imagens para impressão.
* **Screeninfo:** Centralização inteligente da janela no monitor ativo.

## 🚀 Como Executar o Código

1.  **Instale as dependências:**
    Certifique-se de ter o arquivo `requirements.txt` na mesma pasta e rode:
    ```bash
    pip install -r requirements.txt
    ```

2.  **Execute o script:**
    ```bash
    python "Fila de Impressão - PDF v1.0.py"
    ```

---

## 📦 Como Criar o Executável (.exe)

Para gerar um arquivo único que rode em qualquer Windows sem precisar de Python instalado:

1.  **Instale o PyInstaller:**
    ```bash
    pip install pyinstaller
    ```

2.  **Gere o executável:**
    pyinstaller --noconsole --onefile --collect-all tkinterdnd2 --icon=Icon.ico "Fila de Impressão - PDF v1.0.py"
    ```

---

## ⚖️ Licença (License)

Este projeto está sob a **Licença MIT**.

### Resumo em Português
O software é fornecido "como está", sem garantias de qualquer tipo. Você tem permissão para usar, copiar, modificar e distribuir o código, desde que mantenha os créditos do autor original e o aviso de licença em todas as cópias. O autor não se responsabiliza por quaisquer danos decorrentes do uso do software.

Consulte o arquivo `LICENSE` para o texto integral em inglês (versão oficial).

---

## 👨‍💻 Desenvolvedor

**Marcos Castro** - [GitHub](https://github.com/marcosocastro)