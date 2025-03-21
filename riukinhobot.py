import os
import requests
import shutil
import xlwings as xw
import pandas as pd  # Importando o pandas
from dotenv import load_dotenv
from telegram import Update
from telegram.ext import Application, CommandHandler, CallbackContext
import logging

# Carregar variáveis de ambiente
load_dotenv()

# Configuração do bot
TOKEN = os.getenv("TELEGRAM_BOT_TOKEN")
EXCEL_GITHUB_URL = "https://raw.githubusercontent.com/harrisonleandro/riukinhobot/main/AprovaçõesdeOPs.xlsm"
SHEET_NAME = "Registros"

# Configuração do log
logging.basicConfig(format='%(asctime)s - %(name)s - %(levelname)s - %(message)s', level=logging.INFO)
logger = logging.getLogger(__name__)

# Função para baixar o arquivo Excel do GitHub
def download_excel():
    try:
        response = requests.get(EXCEL_GITHUB_URL, stream=True)
        response.raise_for_status()  # Garante que o código de status seja 200
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
        
        # Abrir o arquivo usando xlwings
        wb = xw.Book("AprovaçõesdeOPs.xlsm")
        sheet = wb.sheets[SHEET_NAME]

        # Ler os dados da planilha
        df = sheet.range('A1').expand().options(pd.DataFrame, header=1, index=False).value

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
        
        # Abrir o arquivo usando xlwings
        wb = xw.Book("AprovaçõesdeOPs.xlsm")
        sheet = wb.sheets[SHEET_NAME]

        # Ler os dados da planilha
        df = sheet.range('A1').expand().options(pd.DataFrame, header=1, index=False).value

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

# Função principal
def main():
    app = Application.builder().token(TOKEN).build()

    app.add_handler(CommandHandler("start", start))
    app.add_handler(CommandHandler("status", status))
    app.add_handler(CommandHandler("lista", lista))

    # Run polling, usando um timeout maior para evitar conflitos
    try:
        app.run_polling(timeout=30, poll_interval=5)
    except Exception as e:
        logger.error(f"Erro ao iniciar o polling: {e}")

# Execução do script
if __name__ == "__main__":
    main()
