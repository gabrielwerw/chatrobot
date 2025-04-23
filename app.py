
from flask import Flask, render_template, request, jsonify
import pandas as pd
from datetime import datetime
import unidecode
import os

app = Flask(__name__)

# Leitura dos arquivos
base_path = os.path.join("data")
faturamento = pd.read_excel(os.path.join(base_path, "Faturamento.xlsx"))
custo = pd.read_excel(os.path.join(base_path, "Custo.xlsx"))
compras = pd.read_excel(os.path.join(base_path, "Compras.xlsx"))
aso = pd.read_excel(os.path.join(base_path, "ASO.xlsx"))
funcionarios = pd.read_excel(os.path.join(base_path, "Funcionarios.xlsx"))
contratos = pd.read_excel(os.path.join(base_path, "Contratos.xlsx"))

# Normalizar texto
def normalizar(texto):
    return unidecode.unidecode(texto.lower())

# Extrair mês da pergunta
def extrair_mes(texto):
    meses = {
        "janeiro": "Janeiro", "fevereiro": "Fevereiro", "marco": "Março",
        "abril": "Abril", "maio": "Maio", "junho": "Junho",
        "julho": "Julho", "agosto": "Agosto", "setembro": "Setembro",
        "outubro": "Outubro", "novembro": "Novembro", "dezembro": "Dezembro"
    }
    texto = normalizar(texto)
    for chave, valor in meses.items():
        if chave in texto:
            return valor
    return None

# Função de resposta
def responder(pergunta):
    pergunta_norm = normalizar(pergunta)
    mes = extrair_mes(pergunta)

    if "faturamento" in pergunta_norm:
        if mes:
            valor = faturamento.loc[faturamento['Data'] == mes, 'Valor'].sum()
            historico = faturamento.to_dict(orient="records")
            return f"Faturamento de {mes}: R$ {valor:,.2f}" + "<br><br><b>Histórico:</b><br>" + "<br>".join([f"{x['Data']}: R$ {x['Valor']:,.2f}" for x in historico])
        else:
            total = faturamento['Valor'].sum()
            return f"Faturamento total: R$ {total:,.2f}"

    if "compras" in pergunta_norm:
        if mes:
            valor = compras.loc[compras['Data'].str.contains(mes, case=False, na=False), 'Valor'].sum()
            historico = compras.to_dict(orient="records")
            return f"Compras de {mes}: R$ {valor:,.2f}" + "<br><br><b>Histórico:</b><br>" + "<br>".join([f"{x['Data']}: R$ {x['Valor']:,.2f}" for x in historico])
        else:
            total = compras['Valor'].sum()
            return f"Total de compras: R$ {total:,.2f}"

    if "funcionario" in pergunta_norm:
        total = len(funcionarios)
        return f"Total de funcionários: {total}"

    if "aso" in pergunta_norm:
        vencendo = aso[aso['Data de Vencimento'] < datetime.now().strftime('%Y-%m-%d')]
        return f"Total de ASOs vencidos: {len(vencendo)}"

    if "custo" in pergunta_norm:
        if mes:
            valor = custo.loc[custo['Mes'].str.contains(mes, case=False, na=False), 'Custo'].sum()
            return f"Custo de {mes}: R$ {valor:,.2f}"
        else:
            total = custo['Custo'].sum()
            return f"Custo total: R$ {total:,.2f}"

    if "contrato" in pergunta_norm:
        if mes:
            contratos['Mes'] = pd.to_datetime(contratos['Valor do Contrato']).dt.strftime('%B')
            contratos['Mes'] = contratos['Mes'].str.capitalize()
            valor = contratos.loc[contratos['Mes'] == mes, 'Valor Contrato'].sum()
            return f"Total de contratos em {mes}: R$ {valor:,.2f}"
        else:
            total = contratos['Valor Contrato'].sum()
            return f"Valor total dos contratos: R$ {total:,.2f}"

    return "Não entendi sua pergunta. Tente novamente com termos como 'faturamento', 'compras', 'ASOs', 'funcionários', 'contratos' ou 'custo'."

@app.route("/")
def index():
    return render_template("index.html")

@app.route("/perguntar", methods=["POST"])
def perguntar():
    pergunta = request.form.get("mensagem")
    resposta = responder(pergunta)
    return jsonify({"resposta": resposta})


if __name__ == "__main__":
    port = int(os.environ.get("PORT", 5000))
    app.run(host="0.0.0.0", port=port)