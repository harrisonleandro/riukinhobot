import os
import requests
import shutil
import pandas as pd
from dotenv import load_dotenv
from telegram import Update, Bot
from telegram.ext import Application, CommandHandler, CallbackContext, Dispatcher
import logging
from flask import Flask, request

# Carregar variáveis de ambiente
load_dotenv()

# Configuração do bot
TOKEN = os.getenv("TELEGRAM_BOT_TOKEN")
WEBHOOK_URL = os.getenv("WEBHOOK_URL")
EXCEL_GITHUB_URL = "https://github.com/harrisonleandro/riukinhobot/blob/main/AprovaçõesdeOPs.xlsm"
SHEET_NAME = "Registros"

# Configuração do log
logging.basicConfig(format='%(asctime)s - %(name)s - %(levelname)s - %(message)s', level=logging.INFO)
logger = logging.getLogger(__name__)

# Configuração do Flask
app = Flask(__name__)

# Inicializando o bot
bot = Bot(token=TOKEN)

# Inicializando o dispatcher
dispatcher = Dispatcher(bot, None)

# Função para baixar o arquivo Excel do GitHub
def download_excel():
    try:
        response = requests.get(EXCEL_GITHUB_URL, stream=True)
        response.raise_for_status()
        with open("AprovaçõesdeOPs.xlsm", "wb") as file:
            shutil.copyfileobj(response.raw, file)
        print("Arquivo Excel baixado com sucesso!")
    except Exception as e:
        print(f"Erro ao baixar o arquivo: {e}")

# Função start
async def start(update: Update, context: CallbackContext) -> None:
    mensagem = (
        "Olá! Sou Riukinho, o Bot da Logística. Se precisar de informações das Listas de Picking, é só chamar!\n"
        "📌 Envie /status <OP> para ver o status de uma OP.\n"
        "📌 Envie /lista <Linha> para saber quais OPs de determinada linha estão registradas.\n"
        "📌 Envie /pendente <Linha> para saber qual OP da linha está Pendente (aberta)."
    )
    await update.message.reply_text(mensagem)

# Função status
async def status(update: Update, context: CallbackContext) -> None:
    if len(context.args) == 0:
        await update.message.reply_text("Uso: /status <número da OP>")
        return

    op = context.args[0].lstrip("0")  # Remove zeros à esquerda
    try:
        download_excel()  # Baixar o arquivo Excel do GitHub
        df = pd.read_excel("AprovaçõesdeOPs.xlsm", sheet_name=SHEET_NAME, engine='openpyxl')

        if 'OP' not in df.columns or 'Status' not in df.columns:
            await update.message.reply_text("Erro: A planilha não contém as colunas esperadas ('OP' e 'Status').")
            return

        resultado = df[df['OP'].astype(str).str.lstrip("0") == op]

        if resultado.empty:
            await update.message.reply_text(f"OP {op} não encontrada.")
        else:
            status_op = resultado.iloc[0]['Status']
            await update.message.reply_text(f"A OP {op} está com status: {status_op}")

    except Exception as e:
        logger.error(f"Erro ao buscar OP: {str(e)}")
        await update.message.reply_text(f"Erro ao processar a OP: {str(e)}")

# Função para listar todas as OPs de uma linha
async def lista(update: Update, context: CallbackContext) -> None:
    if len(context.args) == 0:
        await update.message.reply_text("Uso: /lista <número da Linha>")
        return

    linha = context.args[0].lstrip("0")  # Remove zeros à esquerda
    try:
        download_excel()  # Baixar o arquivo Excel do GitHub
        df = pd.read_excel("AprovaçõesdeOPs.xlsm", sheet_name=SHEET_NAME, engine='openpyxl')

        if 'Linha' not in df.columns or 'OP' not in df.columns or 'Status' not in df.columns:
            await update.message.reply_text("Erro: A planilha não contém as colunas esperadas ('Linha', 'OP' e 'Status').")
            return

        resultado = df[df['Linha'].astype(str).str.lstrip("0") == linha]

        if resultado.empty:
            await update.message.reply_text(f"Nenhuma OP encontrada para a linha {linha}.")
        else:
            lista_ops = "\n".join([f"OP {row['OP']} - {row['Status']}" for _, row in resultado.iterrows()])
            await update.message.reply_text(f"Lista de OPs da linha {linha}:\n{lista_ops}")

    except Exception as e:
        logger.error(f"Erro ao buscar OPs da linha {linha}: {str(e)}")
        await update.message.reply_text(f"Erro ao processar a linha {linha}: {str(e)}")

# Função para configurar o webhook
async def set_webhook():
    await bot.set_webhook(url=WEBHOOK_URL)

# Rota do Flask para responder ao webhook
@app.route('/' + os.getenv("WEBHOOK_SECRET"), methods=['POST'])
def webhook():
    json_str = request.get_data().decode('UTF-8')
    update = Update.de_json(json_str, bot)
    dispatcher.process_update(update)
    return "OK"

# Função principal para iniciar o bot
def main():
    import asyncio
    loop = asyncio.get_event_loop()
    loop.run_until_complete(set_webhook())  # Configurar o webhook antes de iniciar o servidor

    app.run(host='0.0.0.0', port=int(os.getenv("PORT", 10000)))  # Iniciar o Flask

# Execução do script
if __name__ == "__main__":
    main()

