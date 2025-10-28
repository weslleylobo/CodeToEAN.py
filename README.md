# CodeToEAN.py – Gerenciador de Arquivos por Código e Código de Barras

Este script em Python automatiza a **localização e renomeação de arquivos** com base em códigos do fabricante ou códigos de barras (EAN/UPC). Ideal para padronizar arquivos recebidos de fornecedores de acordo com uma planilha de referência.

---

## 📂 Estrutura de pastas

- **Desktop**
  - `entrada/` → Pasta onde você coloca os arquivos recebidos do fornecedor.  
    - Pode conter arquivos em subpastas.  
  - `saida/` → Pasta onde os arquivos processados serão copiados e renomeados automaticamente.  
  - `planilha.xlsx` → Planilha contendo a relação entre código do fabricante, descrição e código de barras.  

> A pasta `saida` será criada automaticamente se não existir.

---

## 📄 Estrutura do arquivo Excel (`planilha.xlsx`)

A planilha deve conter as seguintes colunas (a primeira linha é o cabeçalho):

| Código | Descrição | Código de barras |
|--------|-----------|-----------------|
| 123    | Shampoo X | 7894561230012   |
| 124    | Condicionador Y | 7894561230013 |

- **Código** → Código do fabricante ou identificador interno.  
- **Descrição** → Opcional, para referência.  
- **Código de barras** → Código EAN/UPC correspondente.  

---

## ⚙️ Funcionamento do script

1. Lê todos os arquivos da pasta `entrada/`, incluindo subpastas.  
2. Lê a planilha `planilha.xlsx` e valida as colunas configuradas pelo usuário (`COLUNA_BUSCA` e `COLUNA_RENOME`).  
3. Normaliza os nomes dos arquivos e os valores da planilha (remove caracteres especiais, mantém apenas letras e números).  
4. Procura arquivos cujo nome contenha o valor definido em **COLUNA_BUSCA**.  
5. Dentre os arquivos encontrados, prioriza arquivos PNG; se houver múltiplos, escolhe o maior em tamanho.  
6. Copia os arquivos para `saida/` e renomeia com o valor definido em **COLUNA_RENOME**.  
7. Mantém os arquivos originais intactos.

> Você pode configurar o fluxo desejado alterando as colunas:

```python
# Para Código → Código de Barras
COLUNA_BUSCA = "Código"
COLUNA_RENOME = "Código de barras"

# Para Código de Barras → Código
COLUNA_BUSCA = "Código de barras"
COLUNA_RENOME = "Código"
```

---

## 🏗️ Como usar

1. Coloque os arquivos recebidos na pasta `entrada/`.  
2. Certifique-se de que a planilha `planilha.xlsx` esteja no Desktop, com as colunas corretas.  
3. Instale as dependências:

```bash
pip install -r requirements.txt
```

4. Configure as colunas `COLUNA_BUSCA` e `COLUNA_RENOME` no script.  
5. Execute o script:

```bash
python ~/Desktop/CodeToEAN.py
```

6. Os arquivos serão copiados para `saida/` e renomeados conforme a planilha.

---

## ⚡ Observações

- Arquivos originais **não são modificados**; apenas são copiados e renomeados na pasta `saida/`.  
- Funciona em **Windows, Linux e macOS**, usando caminhos relativos ao Desktop.  
- A normalização ignora caracteres especiais (`-`, `_`, `.`, espaços, etc.), considerando apenas letras e números.

---

## 📌 Requisitos

- Python 3.x  
- Pandas (`pip install pandas`)

---

## 📅 Informações

- Autor: Weslley Lobo  
- Data de criação: 28/10/2025  
- Versão: 1.0

---

## 📜 Licença

Este projeto é distribuído sob a licença MIT. Sinta-se à vontade para usar, modificar e compartilhar.

