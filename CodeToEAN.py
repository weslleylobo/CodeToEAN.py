"""=========================================================
SCRIPT: GerenciadorDeArquivosEAN.py
AUTOR: Weslley Lobo
DATA DE CRIAÇÃO: 28/10/2025
VERSÃO: 1.0

DESCRIÇÃO:
Este script automatiza o processo de localização e renomeação de arquivos
com base em códigos do fabricante ou códigos de barras (EAN/UPC).

Ele faz o seguinte:

1. Lê todos os arquivos da pasta "entrada" localizada na área de trabalho,
   incluindo subpastas.
2. Lê a planilha "planilha.xlsx" que deve conter, no mínimo, as colunas
   especificadas pelo usuário (ex: Código e Código de barras).
3. Normaliza os nomes dos arquivos e valores da planilha (removendo caracteres
   especiais e mantendo apenas letras e números) para facilitar a busca.
4. Procura arquivos cujo nome contenha o valor da coluna configurada como
   COLUNA_BUSCA.
5. Dentre os arquivos encontrados, prioriza arquivos PNG; se houver múltiplos,
   seleciona o maior em tamanho.
6. Copia os arquivos encontrados para a pasta "saida" e os renomeia com o valor
   da coluna configurada como COLUNA_RENOME.
7. Mantém os arquivos originais intactos.

SUPORTE A FORMATOS:
Todos os formatos de arquivo, priorizando PNG quando houver múltiplas opções.

OBSERVAÇÃO:
Se as pastas "entrada" ou "saida" não existirem, a pasta de saída será criada
automaticamente. A pasta de entrada deve ser criada pelo usuário com os arquivos
a serem processados.


Estrutura do arquivo Excel (`planilha.xlsx`)

A planilha deve conter as seguintes colunas (a primeira linha é o cabeçalho):

| Código | Descrição         | Código de barras |
|--------|-------------------|------------------|
| 123    | Shampoo X         | 7894561230012    |
| 124    | Condicionador Y   | 7894561230013    |

=========================================================""" 


from pathlib import Path
import re
import shutil
import pandas as pd

# === CONFIGURAÇÕES DO USUÁRIO ===
NOME_PASTA_ENTRADA = "entrada"
NOME_PASTA_SAIDA = "saida"
NOME_PLANILHA = "planilha.xlsx"

# Defina as colunas a serem usadas:
# Para Código → Código de Barras:
#     COLUNA_BUSCA = "Código"
#     COLUNA_RENOME = "Código de barras"
# Para Código de Barras → Código:
#     COLUNA_BUSCA = "Código de barras"
#     COLUNA_RENOME = "Código"

COLUNA_BUSCA = "Código"           # coluna usada para localizar no nome dos arquivos
COLUNA_RENOME = "Código de barras"                    # coluna usada para renomear os arquivos

# === CONSTRUÇÃO DOS CAMINHOS ===
desktop = Path.home() / "Desktop"
pasta_entrada = desktop / NOME_PASTA_ENTRADA
pasta_saida = desktop / NOME_PASTA_SAIDA
arquivo_excel = desktop / NOME_PLANILHA

# Criar pasta de saída, se não existir
pasta_saida.mkdir(exist_ok=True)

# === FUNÇÕES AUXILIARES ===
def normalizar_texto(texto):
    """Remove tudo que não for letra ou número e deixa em maiúsculo."""
    return re.sub(r'[^A-Za-z0-9]', '', str(texto)).upper()

def listar_arquivos_recursivo(pasta):
    """Retorna todos os arquivos da pasta e subpastas."""
    return [f for f in pasta.rglob("*") if f.is_file()]

def escolher_arquivo_preferido(lista_caminhos):
    """Dentre os arquivos encontrados, escolhe o PNG; se houver vários, o maior arquivo."""
    if not lista_caminhos:
        return None
    pngs = [f for f in lista_caminhos if f.suffix.lower() == ".png"]
    if pngs:
        lista_caminhos = pngs
    return max(lista_caminhos, key=lambda f: f.stat().st_size)

def nome_valido_windows(nome):
    """Remove caracteres inválidos em nomes de arquivo no Windows."""
    return re.sub(r'[<>:"/\\|?*]', '_', nome)

# === LER PLANILHA ===
if not arquivo_excel.exists():
    raise FileNotFoundError(f"❌ Planilha não encontrada: {arquivo_excel}")

df = pd.read_excel(arquivo_excel)

colunas_esperadas = {COLUNA_BUSCA, COLUNA_RENOME}
if not colunas_esperadas.issubset(df.columns):
    raise ValueError(f"A planilha deve conter as colunas: {colunas_esperadas}")

# === PROCESSAMENTO ===
arquivos = listar_arquivos_recursivo(pasta_entrada)
print(f"Foram encontrados {len(arquivos)} arquivos na pasta de entrada.\n")

for _, linha in df.iterrows():
    valor_busca = str(linha[COLUNA_BUSCA])
    valor_renome = str(linha[COLUNA_RENOME])

    if pd.isna(valor_busca) or not valor_busca.strip():
        continue

    busca_norm = normalizar_texto(valor_busca)
    candidatos = []

    # Procurar arquivos cujo nome contenha o valor de busca normalizado
    for caminho in arquivos:
        nome_normalizado = normalizar_texto(caminho.name)
        if busca_norm in nome_normalizado:
            candidatos.append(caminho)

    if candidatos:
        escolhido = escolher_arquivo_preferido(candidatos)
        extensao = escolhido.suffix
        destino = pasta_saida / f"{nome_valido_windows(valor_renome)}{extensao}"

        # ✅ cópia segura, preserva original
        if not destino.exists():
            shutil.copy2(escolhido, destino)
        else:
            print(f"⚠️ Arquivo já existe na saída: {destino.name}")

        print(f"✅ {valor_busca} encontrado → renomeado como {valor_renome}{extensao}")
    else:
        print(f"🚫 {valor_busca} não encontrado em nenhum arquivo.")

print("\n✅ Processo concluído — todos os arquivos originais foram mantidos intactos.")
