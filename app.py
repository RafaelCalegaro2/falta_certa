from flask import Flask, render_template, request, session, redirect
import pandas as pd
from datetime import datetime, timedelta
import os

app = Flask(__name__)
app.secret_key = "sua_chave_secreta_aqui"  # Substitua por uma chave segura

# Arquivos
ARQUIVO_EXCEL = "faltas.xlsx"
ARQUIVO_RESPOSTAS = "respostas.csv"
ARQUIVO_ENCARREGADOS = "encarregados.txt"

COLUNAS_RESPOSTAS = [
    "MATRICULA",
    "FUNCIONARIO",
    "DATA_FALTA",
    "ACAO",
    "JUSTIFICATIVA",
    "DATA_RESPOSTA",
    "ENCARREGADO"
]

def ler_encarregados():
    if not os.path.exists(ARQUIVO_ENCARREGADOS):
        return []
    with open(ARQUIVO_ENCARREGADOS, "r", encoding="utf-8") as f:
        return [linha.strip().upper() for linha in f if linha.strip()]

def carrega_dados():
    df = pd.read_excel(ARQUIVO_EXCEL, header=1)
    df = df[["Matrícula", "Funcionário", "Encarregado", "Data", "Dia da Semana"]].dropna()
    df["Data"] = pd.to_datetime(df["Data"], dayfirst=True).dt.strftime("%d/%m/%Y")
    df["Encarregado"] = df["Encarregado"].str.upper()
    return df

def carrega_respostas():
    if os.path.exists(ARQUIVO_RESPOSTAS):
        return pd.read_csv(ARQUIVO_RESPOSTAS, encoding='utf-8')
    else:
        return pd.DataFrame(columns=COLUNAS_RESPOSTAS)

def salvar_resposta(matricula, funcionario, data, acao, justificativa, data_resposta, encarregado):
    nova = pd.DataFrame([[matricula, funcionario, data, acao, justificativa, data_resposta, encarregado]],
                        columns=COLUNAS_RESPOSTAS)
    escrever_cabecalho = not os.path.exists(ARQUIVO_RESPOSTAS) or os.path.getsize(ARQUIVO_RESPOSTAS) == 0
    try:
        nova.to_csv(ARQUIVO_RESPOSTAS, mode='a', header=escrever_cabecalho, index=False, encoding='utf-8-sig')
        print(f"Salvamento bem-sucedido: {matricula}, {data}, {acao}")  # Depuração
    except Exception as e:
        print(f"Erro ao salvar: {e}")  # Depuração

def gerar_csv_semanal():
    respostas = carrega_respostas()
    if respostas.empty:
        return

    hoje = datetime.now()
    inicio_semana = hoje - timedelta(days=hoje.weekday())
    inicio_semana = inicio_semana.replace(hour=0, minute=0, second=0, microsecond=0)
    fim_semana = inicio_semana + timedelta(days=6, hours=23, minutes=59, seconds=0)

    respostas["DATA_RESPOSTA"] = pd.to_datetime(respostas["DATA_RESPOSTA"], format="%d/%m/%Y %H:%M")
    respostas_semana = respostas[
        (respostas["DATA_RESPOSTA"] >= inicio_semana) &
        (respostas["DATA_RESPOSTA"] <= fim_semana)
    ]

    nome_arquivo = f"respostas_{inicio_semana.strftime('%Y-%m-%d')}_a_{fim_semana.strftime('%Y-%m-%d')}.csv"
    caminho_arquivo = nome_arquivo
    respostas_semana.to_csv(caminho_arquivo, index=False, encoding='utf-8-sig')

@app.route('/', methods=['GET', 'POST'])
def index():
    dados = carrega_dados()
    respostas = carrega_respostas()
    encarregados = ler_encarregados()
    faltas_encarregado = pd.DataFrame()
    idx_atual = 0

    if request.method == 'POST' and 'encarregado' in request.form:
        encarregado_nome = request.form.get('encarregado', '').strip().upper()
        if encarregado_nome and encarregado_nome in encarregados:
            faltas_encarregado = dados[dados["Encarregado"] == encarregado_nome].copy()
            if not respostas.empty:
                respondidas = respostas[respostas["ENCARREGADO"] == encarregado_nome]
                respondidas_set = set(zip(respondidas["MATRICULA"], respondidas["DATA_FALTA"]))
                faltas_encarregado = faltas_encarregado[
                    ~faltas_encarregado.apply(lambda r: (r["Matrícula"], r["Data"]) in respondidas_set, axis=1)
                ]
            faltas_encarregado.reset_index(drop=True, inplace=True)
            if not faltas_encarregado.empty:
                session['faltas_encarregado'] = faltas_encarregado.to_dict('records')
                session['encarregado_nome'] = encarregado_nome
                session['idx_atual'] = 0
                return render_template('index.html', encarregados=encarregados, faltas=[faltas_encarregado.loc[0].to_dict()], idx_atual=0, total=len(faltas_encarregado))
            else:
                return render_template('index.html', encarregados=encarregados, mensagem="Nenhuma falta pendente para este encarregado.")
        else:
            return render_template('index.html', encarregados=encarregados, mensagem="Encarregado inválido ou não encontrado. Por favor, selecione na lista.")

    return render_template('index.html', encarregados=encarregados)

@app.route('/acao', methods=['POST'])
def acao():
    if 'faltas_encarregado' not in session or 'encarregado_nome' not in session:
        return render_template('index.html', mensagem="Sessão inválida. Carregue as faltas novamente.")

    faltas_encarregado = pd.DataFrame(session['faltas_encarregado'])
    encarregado_nome = session['encarregado_nome']
    idx_atual = session.get('idx_atual', 0)
    total = len(faltas_encarregado)

    if idx_atual >= total:
        return redirect('/resumo')

    acao = request.form.get('acao')
    justificativa = request.form.get('justificativa', '').strip()
    print(f"Form data: {request.form}")  # Depuração

    falta = faltas_encarregado.loc[idx_atual]
    if acao == "justificar":
        if len(justificativa) < 5:
            return render_template('index.html', encarregados=[encarregado_nome], faltas=[falta.to_dict()], idx_atual=idx_atual, total=total, erro="Justificativa muito curta! Por favor, explique melhor.")
        salvar_resposta(
            falta["Matrícula"],
            falta["Funcionário"],
            falta["Data"],
            "Justificada",
            justificativa,
            datetime.now().strftime("%d/%m/%Y %H:%M"),
            encarregado_nome
        )
        idx_atual += 1
    elif acao == "confirmar":
        salvar_resposta(
            falta["Matrícula"],
            falta["Funcionário"],
            falta["Data"],
            "Confirmada",
            "",
            datetime.now().strftime("%d/%m/%Y %H:%M"),
            encarregado_nome
        )
        idx_atual += 1
    session['idx_atual'] = idx_atual
    if idx_atual < total:
        return render_template('index.html', encarregados=[encarregado_nome], faltas=[faltas_encarregado.loc[idx_atual].to_dict()], idx_atual=idx_atual, total=total)
    else:
        return redirect('/resumo')

@app.route('/resumo')
def resumo():
    if 'encarregado_nome' not in session:
        return render_template('index.html', mensagem="Por favor, carregue as faltas primeiro.")
    
    encarregado_nome = session['encarregado_nome']
    respostas = carrega_respostas()
    faltas_confirmadas = respostas[(respostas["ENCARREGADO"] == encarregado_nome) & (respostas["ACAO"] == "Confirmada")]
    
    if faltas_confirmadas.empty:
        return render_template('resumo.html', encarregado=encarregado_nome, faltas=[], mensagem="Nenhuma falta confirmada por este encarregado.")
    
    return render_template('resumo.html', encarregado=encarregado_nome, faltas=faltas_confirmadas.to_dict('records'))
if __name__ == "__main__":
    app.run(debug=True, host='0.0.0.0', port=5000)