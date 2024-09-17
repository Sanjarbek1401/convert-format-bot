import os
import asyncio
import logging
from aiogram import Bot, Dispatcher, types
from aiogram.filters import CommandStart
from aiogram.types import FSInputFile
from aiogram.enums import ParseMode
from aiogram.client.default import DefaultBotProperties
from pdf2docx import Converter as PdfToDocxConverter
from docx2pdf import convert as DocxToPdfConverter
from dotenv import load_dotenv

# Load environment variables from .env file
load_dotenv()

# Bot token from BotFather
BOT_TOKEN = os.getenv("BOT_TOKEN")

# Initialize bot with default properties
bot = Bot(token=BOT_TOKEN, default=DefaultBotProperties(parse_mode=ParseMode.HTML))
dp = Dispatcher()

# Set up logging
logging.basicConfig(level=logging.INFO)


# Function to convert PDF to DOCX
def convert_pdf_to_doc(pdf_path, docx_path):
    cv = PdfToDocxConverter(pdf_path)
    cv.convert(docx_path)
    cv.close()


# Function to convert DOCX to PDF
def convert_doc_to_pdf(docx_path, pdf_path):
    DocxToPdfConverter(docx_path, pdf_path)


# Handle the /start command
@dp.message(CommandStart())
async def start_command(message: types.Message):
    await message.answer(
        "Hello! Send me a PDF file to convert it to DOC format or a DOC file to convert it to PDF format."
    )


# Handle incoming PDF files
@dp.message()
async def handle_document(message: types.Message):
    # Check if the message contains a document
    if message.document:
        file_id = message.document.file_id
        file_name = message.document.file_name
        file_extension = file_name.split('.')[-1].lower()

        # Download the file
        file_path = f'{file_id}.{file_extension}'
        await bot.download(message.document, file_path)

        if file_extension == 'pdf':
            # Convert PDF to DOCX
            docx_file_path = f'{file_id}.docx'
            convert_pdf_to_doc(file_path, docx_file_path)

            # Send the converted DOCX file
            docx_file = FSInputFile(docx_file_path)
            await message.answer_document(docx_file)

            # Clean up temporary files
            os.remove(file_path)
            os.remove(docx_file_path)

        elif file_extension == 'docx':
            # Convert DOCX to PDF
            pdf_file_path = f'{file_id}.pdf'
            convert_doc_to_pdf(file_path, pdf_file_path)

            # Send the converted PDF file
            pdf_file = FSInputFile(pdf_file_path)
            await message.answer_document(pdf_file)

            # Clean up temporary files
            os.remove(file_path)
            os.remove(pdf_file_path)

        else:
            await message.answer("Please upload a PDF or DOCX file.")


# Main function to start the bot
async def main():
    await dp.start_polling(bot)


if __name__ == "__main__":
    asyncio.run(main())
