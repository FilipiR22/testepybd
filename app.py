import pandas as pd
import json
from collections import defaultdict

# Caminho do arquivo Excel
excel_path = "planilhas_materias.xlsx"

def importar_para_objetos(excel_path):
    # Ler todas as abas da planilha
    xls = pd.ExcelFile(excel_path)

    # Criar uma estrutura de dicionário aninhado
    dados = {}

    # Iterar sobre as abas
    for sheet_name in xls.sheet_names:
        df = pd.read_excel(xls, sheet_name=sheet_name)
        df["Matéria"] = sheet_name  # Adicionar a matéria como chave

        # Garantir que todas as colunas esperadas existem
        colunas_esperadas = ["Assunto", "Conteúdo", "horas"]
        for col in colunas_esperadas:
            if col not in df.columns:
                df[col] = ""  # Preencher colunas ausentes com string vazia

        # Substituir valores NaN por valores padrão
        df.fillna({"horas": 0, "Conteúdo": "Sem conteúdo"}, inplace=True)

        # Converter a coluna de horas para inteiro
        df["horas"] = df["horas"].astype(int)

        # Criar estrutura agrupada por assunto
        assuntos_dict = defaultdict(lambda: {"conteudo": [], "horas": 0})

        for _, row in df.iterrows():
            assunto = row["Assunto"]
            conteudo = row["Conteúdo"]
            horas = row["horas"]

            # Adicionar conteúdo à lista do assunto
            if conteudo not in assuntos_dict[assunto]["conteudo"]:
                assuntos_dict[assunto]["conteudo"].append(conteudo)
            
            # Somar horas
            assuntos_dict[assunto]["horas"] += horas

        # Aplicar a redução nas horas
        materias = []
        for key, value in assuntos_dict.items():
            horas_originais = value["horas"]
            if horas_originais > 60:
                horas_reduzidas = int(horas_originais * 0.6)  # Reduz 40%
            else:
                horas_reduzidas = int(horas_originais * 0.8)  # Reduz 20%
            
            # Garante que o mínimo de horas seja 1
            materias.append({
                "assunto": key,
                "conteudo": value["conteudo"],
                "horas": max(1, horas_reduzidas)
            })

        dados[sheet_name] = materias

    return dados

# Executar a importação e armazenar os objetos
dados_estudos = importar_para_objetos(excel_path)

# Salvar os dados em um arquivo TXT
output_path = "dados_estudos_reduzidos.txt"
with open(output_path, "w", encoding="utf-8") as f:
    json.dump(dados_estudos, f, indent=4, ensure_ascii=False)

print(f"Dados reduzidos salvos em {output_path}")
