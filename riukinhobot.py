import shutil
import os
import pandas as pd
import openpyxl
import logging
from telegram import Update
from telegram.ext import ApplicationBuilder, CommandHandler, CallbackContext
from dotenv import load_dotenv

# Carregar variáveis de ambiente
load_dotenv()

# Configuração do bot
TOKEN = os.getenv("TELEGRAM_BOT_TOKEN")
EXCEL_FILE = os.getenv("EXCEL_FILE_PATH")
SHEET_NAME = "Registros"

# Configuração do log
logging.basicConfig(format='%(asctime)s - %(name)s - %(levelname)s - %(message)s', level=logging.INFO)
logger = logging.getLogger(__name__)

# Função start
async def start(update: Update, context: CallbackContext) -> None:
    mensagem = (
        "Olá! Sou Riukinho, o Bot da Logística. Se precisar de informações das Listas de Picking, é só chamar!\n"
        "📌 Envie /status <OP> para ver o status de uma OP.\n"
        "📌 Envie /lista <Linha> para saber quais OPs de determinada linha estão registradas.\n"
        "📌 Envie /pendente <Linha> para saber qual OP da linha está Pendente (aberta)."
    )
    await update.message.reply_text(mensagem)

# Função de status
async def status(update: Update, context: CallbackContext) -> None:
    if len(context.args) == 0:
        await update.message.reply_text("Uso: /status <número da OP>")
        return

    op = context.args[0].lstrip("0")  # Remove zeros à esquerda
    try:
        TEMP_FILE = EXCEL_FILE.replace(".xlsm", "_temp.xlsm")
        shutil.copy(EXCEL_FILE, TEMP_FILE)
        df = pd.read_excel(TEMP_FILE, sheet_name=SHEET_NAME, engine='openpyxl')

        if 'OP' not in df.columns or 'Status' not in df.columns:
            await update.message.reply_text("Erro: A planilha não contém as colunas esperadas ('OP' e 'Status').")
            return

        resultado = df[df['OP'].astype(str).str.lstrip("0") == op]

        if resultado.empty:
            await update.message.reply_text(f"OP {op} não encontrada.")
        else:
            status_op = resultado.iloc[0]['Status']
            await update.message.reply_text(f"A OP {op} está com status: {status_op}")

        os.remove(TEMP_FILE)

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
        TEMP_FILE = EXCEL_FILE.replace(".xlsm", "_temp.xlsm")
        shutil.copy(EXCEL_FILE, TEMP_FILE)
        df = pd.read_excel(TEMP_FILE, sheet_name=SHEET_NAME, engine='openpyxl')

        if 'Linha' not in df.columns or 'OP' not in df.columns or 'Status' not in df.columns:
            await update.message.reply_text("Erro: A planilha não contém as colunas esperadas ('Linha', 'OP' e 'Status').")
            return

        resultado = df[df['Linha'].astype(str).str.lstrip("0") == linha]

        if resultado.empty:
            await update.message.reply_text(f"Nenhuma OP encontrada para a linha {linha}.")
        else:
            lista_ops = "\n".join([f"OP {row['OP']} - {row['Status']}" for _, row in resultado.iterrows()])
            await update.message.reply_text(f"Lista de OPs da linha {linha}:\n{lista_ops}")

        os.remove(TEMP_FILE)

    except Exception as e:
        logger.error(f"Erro ao buscar OPs da linha {linha}: {str(e)}")
        await update.message.reply_text(f"Erro ao processar a linha {linha}: {str(e)}")

# Função para buscar o último registro pendente de uma linha
async def pendente(update: Update, context: CallbackContext) -> None:
    if len(context.args) == 0:
        await update.message.reply_text("Uso: /pendente <número da Linha>")
        return

    linha = context.args[0].lstrip("0")  # Remove zeros à esquerda
    try:
        TEMP_FILE = EXCEL_FILE.replace(".xlsm", "_temp.xlsm")
        shutil.copy(EXCEL_FILE, TEMP_FILE)
        df = pd.read_excel(TEMP_FILE, sheet_name=SHEET_NAME, engine='openpyxl')

        if 'Linha' not in df.columns or 'OP' not in df.columns or 'Status' not in df.columns:
            await update.message.reply_text("Erro: A planilha não contém as colunas esperadas ('Linha', 'OP' e 'Status').")
            return

        # Filtra registros da linha especificada e com status "Pendente"
        resultado = df[(df['Linha'].astype(str).str.lstrip("0") == linha) & (df['Status'] == 'Pendente')]

        if resultado.empty:
            await update.message.reply_text(f"Não há registros pendentes para a linha {linha}.")
        else:
            # Encontrar o último registro pendente baseado na data
            resultado = resultado.sort_values(by="Data de Registro", ascending=False)  # Ordenar pela data de registro
            ultimo_pendente = resultado.iloc[0]  # Pega o último (mais recente) registro
            await update.message.reply_text(
                f"Último registro pendente da linha {linha}:\n"
                f"OP: {ultimo_pendente['OP']}\n"
                f"Status: {ultimo_pendente['Status']}\n"
                f"Data de Registro: {ultimo_pendente['Data de Registro']}"
            )

        os.remove(TEMP_FILE)

    except Exception as e:
        logger.error(f"Erro ao buscar registro pendente da linha {linha}: {str(e)}")
        await update.message.reply_text(f"Erro ao processar a linha {linha}: {str(e)}")

# Função principal
def main():
    app = ApplicationBuilder().token(TOKEN).build()

    app.add_handler(CommandHandler("start", start))
    app.add_handler(CommandHandler("status", status))
    app.add_handler(CommandHandler("lista", lista))  # Comando lista já existente
    app.add_handler(CommandHandler("pendente", pendente))  # Novo comando pendente

    # Ativando Webhook
    app.run_webhook(
        listen="0.0.0.0",
        port=8443,
        url_path=TOKEN,
        webhook_url=f"https://SEU_DOMINIO_AQUI/{TOKEN}",
    )

# Execução do script
if __name__ == "__main__":
    main()
