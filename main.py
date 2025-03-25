import os, shutil
import re
import pandas as pd
import seaborn as sns
import arabic_reshaper
import matplotlib.pyplot as plt
from bidi.algorithm import get_display

from docx import Document
from docx.oxml.ns import qn
from docx.shared import Cm
from docx.oxml import OxmlElement
from docx.enum.text import WD_ALIGN_PARAGRAPH
from excelfile import fillnotes
from headinformation import HeadInformation
from filldoc import FillDoc
from analysvisualizationfile import AnalyseVisualizationFile

from telegram import Update
from makefolderuser import MakeFolderUser
from telegram.ext import ApplicationBuilder, CommandHandler, ContextTypes, MessageHandler, filters

with open("token.txt") as f:
    TOKEN = f.read()


async def helped(update: Update, context: ContextTypes.DEFAULT_TYPE):
    await context.bot.send_message(chat_id=update.effective_chat.id, text="Hi, Im the Analyse ClassSchool excel bot")


async def start(update: Update, context: ContextTypes.DEFAULT_TYPE):
    await context.bot.send_message(chat_id=update.effective_chat.id, text="To create your ClassSchool analyse,"
                                                                          " please send the file")


async def hand_message(update: Update, context: ContextTypes.DEFAULT_TYPE):
    # Information about file
    name_id = update.message.chat.id
    file_id = update.message.document.file_id
    f_name, f_ext = os.path.split(update.message.document.file_name)
    f_unique_id = update.message.document.file_unique_id
    f_id_name = f"{f_ext}"

    # Create Path For Client
    user_name = update.message.chat.username
    create_folder_client = MakeFolderUser(name_id)
    path_file_receive = create_folder_client.pathreceive + f"/{f_id_name}"
    path_file_send = create_folder_client.pathsend




    # Download the file
    file = await context.bot.get_file(file_id)
    await file.download_to_drive(custom_path=path_file_receive)
    await context.bot.send_message(chat_id=update.effective_chat.id, text="We are processing your file. Please wait...")

    # processed_file = await processfile(f_id_name)
    # Get Sheet Names List
    filexcel = fillnotes(path_file_receive)
    sheet_names_list = pd.ExcelFile(path_file_receive).sheet_names
    path_fiel_list = []
    for sheet in sheet_names_list:
        print(sheet)
        if not pd.read_excel(path_file_receive, sheet_name=sheet).empty:
            # Information About Head File
            headinfo = HeadInformation(path_file_receive, sheet)
            dirclass = create_folder_client.getpathclass(headinfo.nameclass+sheet)
            pathclass = create_folder_client.pathsend + headinfo.nameclass+sheet
            path_file_send_class = path_file_send + f"/{create_folder_client.nameclass}" +f"/{headinfo.nameclass}.docx"

            data = headinfo.sentence_di
            # Analyse Visualization File
            # Calcule Statistic
            d = AnalyseVisualizationFile(path_file_receive, sheet)
            d.numberstudent()
            d.overallaverage()
            d.percentage()
            d.standarddeviation()
            d.studenthighscore()
            t_d = d.percentagedistribution()
            data.update(d.analysedata)
            # Visualization
            print(create_folder_client.path1)
            create_folder_client.getpathpicture()
            d.histogram(create_folder_client.path2)
            d.boxplot(create_folder_client.path3)
            d.barchar(create_folder_client.path4)
            d.scatterplot(create_folder_client.path5)
            d.piechart(create_folder_client.path1)
            d.corrolation(create_folder_client.path6)
            filexcel.notes(sheet)

            # Fill Document Doc2.docx
            fildoc = FillDoc(path_file_send_class, data, d.analysetable, create_folder_client.path_picture)
            path_fiel_list.append(fildoc.output_path)
        else:
            continue
    filexcel.savefile()
    for document in path_fiel_list:
        with open(document, 'rb') as file:
            await context.bot.send_document(chat_id=update.effective_chat.id, document=file)
    await context.bot.send_document(chat_id=update.effective_chat.id, document=path_file_receive)
    create_folder_client.deletefolderusersned()
    create_folder_client.deletefolderreceive()

# async def main():
#     bot = telegram.Bot(TOKEN)
#
#     async with bot:
#         update = (await bot.get_updates())[-1]
#         chat_id = update.message.chat.id
#         user_name = update.message.chat.username
#         await bot.send_message(text=f"Hi {user_name}", chat_id=chat_id)


# execute the program
if __name__ == '__main__':
    # asyncio.run(main())
    application = ApplicationBuilder().token(TOKEN).build()

    # Command Handler
    help_handler = CommandHandler("help", helped)
    star_handler = CommandHandler("start", start)
    message_handler = MessageHandler(filters.Document.ALL, hand_message)

    # register command
    application.add_handler(help_handler)
    application.add_handler(star_handler)
    application.add_handler(message_handler)

    # run bot
    application.run_polling()
