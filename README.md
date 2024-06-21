# Mesclar Abas do Excel

Este script em Python permite **mesclar múltiplas abas** de um arquivo Excel em uma única aba e **remover duplicatas**, utilizando uma **interface gráfica** criada com **Tkinter**.

## Requisitos

- **Python 3.x** instalado (recomenda-se Python 3.6 ou superior).
- Bibliotecas necessárias:
  - **pandas**
  - **tkinter** (incluída na biblioteca padrão do Python)

## Como Usar

1. **Executar o Script:**
   - Execute o script `merge_excel_sheets.py` em um ambiente Python compatível.
   
2. **Interface Gráfica:**
   - Ao abrir o programa, uma janela será exibida com um botão "Selecionar Arquivo e Mesclar Abas".

3. **Selecionar Arquivo Excel:**
   - Clique no botão para selecionar o arquivo Excel contendo as abas que deseja mesclar.

4. **Configurar a Mesclagem:**
   - O programa solicitará o número de abas que você deseja mesclar e os nomes dessas abas.

5. **Processamento:**
   - O script carregará as abas selecionadas, mesclará os dados e removerá quaisquer duplicatas.

6. **Salvar Resultado:**
   - Você será solicitado a fornecer um nome para a nova aba que será criada no arquivo Excel original.
   - O resultado da mesclagem será salvo como uma nova aba no mesmo arquivo Excel.

7. **Feedback de Sucesso ou Erro:**
   - Você receberá mensagens de sucesso ou erro conforme o resultado do processo.

## Observações

- Certifique-se de que o arquivo Excel original não esteja aberto em outro programa ao executar este script, para evitar conflitos de acesso ao arquivo.
- O script usa o módulo `openpyxl` para escrever no arquivo Excel, garantindo compatibilidade com formatos `.xlsx`.
