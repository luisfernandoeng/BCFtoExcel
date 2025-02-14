## BCF para Excel Converter

Este é o meu primeiro projeto de código aberto! É uma ferramenta gráfica que converte arquivos BCF (BIM Collaboration Format) em planilhas Excel. Aqui estão alguns detalhes sobre o programa:

- **Interface Gráfica:** Utiliza `tkinter` e `ttkbootstrap` para criar uma interface amigável e intuitiva.
- **Funcionalidades Principais:**
  - Seleção de arquivos BCF, logo da empresa e pasta de destino.
  - Inserção de informações do projeto (nome, responsável, cliente, etapa, cidade e data).
  - Conversão dos tópicos do BCF para uma planilha Excel com formatação personalizada.
  - Inclusão de snapshots (imagens) associadas aos tópicos no Excel.
    
- **Tecnologias Utilizadas:** `tkinter`, `ttkbootstrap`, `PIL`, `openpyxl`, `bcf.bcfxml`, `zipfile`.

Você pode encontrar o código completo do programa no arquivo `bcf_to_excel_converter.py` deste repositório.

## Como Usar

1. **Instalação:** Clone este repositório e instale as dependências listadas no arquivo `requirements.txt`.
2. **Execução:** Execute o script `bcf_to_excel_converter.py` para iniciar a interface gráfica.
3. **Conversão:** Siga as instruções na interface para selecionar o arquivo BCF, logo da empresa e pasta de destino, preencha as informações do projeto e clique em "Converter".
