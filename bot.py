import asyncio
import os
import logging
import shutil
import tempfile
import subprocess
from pathlib import Path

from dotenv import load_dotenv
from telegram import Update
from telegram.ext import Application, CommandHandler, MessageHandler, filters, ContextTypes
from pdf2docx import Converter

load_dotenv()

logging.basicConfig(
    format="%(asctime)s - %(name)s - %(levelname)s - %(message)s",
    level=logging.INFO
)
logger = logging.getLogger(__name__)

BOT_TOKEN = os.getenv("BOT_TOKEN")
MAX_FILE_BYTES = 20 * 1024 * 1024  # Telegram Bot API download limit


LIBREOFFICE_PATHS = [
    "libreoffice",
    "soffice",
    "/Applications/LibreOffice.app/Contents/MacOS/soffice",
    "/usr/bin/libreoffice",
    "/usr/bin/soffice",
    "/usr/lib/libreoffice/program/soffice",
]


def _find_libreoffice() -> str | None:
    for candidate in LIBREOFFICE_PATHS:
        if shutil.which(candidate) or (os.path.isfile(candidate) and os.access(candidate, os.X_OK)):
            return candidate
    for name in ("libreoffice", "soffice"):
        try:
            result = subprocess.run(["which", name], capture_output=True, text=True)
            if result.returncode == 0 and result.stdout.strip():
                return result.stdout.strip()
        except Exception:
            pass
    return None


def _check_libreoffice() -> None:
    if _find_libreoffice() is None:
        logger.warning(
            "LibreOffice not found — DOC/DOCX→PDF conversion will fail. "
            "Install: brew install --cask libreoffice  (macOS) | apt install libreoffice  (Linux)"
        )


async def start(update: Update, context: ContextTypes.DEFAULT_TYPE) -> None:
    await update.message.reply_text(
        "👋 Hello! I'm a file converter bot.\n\n"
        "📄 *PDF → DOCX*: Send any PDF file\n"
        "📝 *DOCX / DOC → PDF*: Send any Word document\n\n"
        "Just drop a file and I'll convert it automatically.\n"
        "Use /help for more details.",
        parse_mode="Markdown",
    )


async def help_command(update: Update, context: ContextTypes.DEFAULT_TYPE) -> None:
    await update.message.reply_text(
        "📋 *How to use:*\n\n"
        "1. Send a *PDF* → receive a *DOCX*\n"
        "2. Send a *DOCX* or *DOC* → receive a *PDF*\n\n"
        "⚠️ *Limits:*\n"
        "• Max file size: 20 MB\n"
        "• Supported formats: PDF, DOCX, DOC\n\n"
        "💡 *Tips:*\n"
        "• PDF→DOCX works best on text-based PDFs\n"
        "• Scanned / image-only PDFs may not convert well",
        parse_mode="Markdown",
    )


def convert_pdf_to_docx(pdf_path: str, docx_path: str) -> None:
    cv = Converter(pdf_path)
    cv.convert(docx_path)
    cv.close()


def convert_docx_to_pdf(docx_path: str, output_dir: str) -> str:
    lo = _find_libreoffice()
    if lo is None:
        raise RuntimeError(
            "LibreOffice is not installed.\n"
            "Install: brew install --cask libreoffice  (macOS) | apt install libreoffice  (Linux)"
        )

    result = subprocess.run(
        [lo, "--headless", "--convert-to", "pdf", "--outdir", output_dir, docx_path],
        capture_output=True,
        text=True,
        timeout=120,
    )

    if result.returncode != 0:
        raise RuntimeError(result.stderr.strip() or "LibreOffice returned a non-zero exit code")

    pdf_path = Path(output_dir) / f"{Path(docx_path).stem}.pdf"
    if not pdf_path.exists():
        raise RuntimeError("LibreOffice ran but no output file was produced")

    return str(pdf_path)


async def handle_document(update: Update, context: ContextTypes.DEFAULT_TYPE) -> None:
    doc = update.message.document
    if not doc:
        return

    file_name = doc.file_name or "file"
    ext = Path(file_name).suffix.lower()

    if ext not in (".pdf", ".docx", ".doc"):
        await update.message.reply_text(
            "⚠️ Unsupported format. Please send a *PDF*, *DOCX*, or *DOC* file.",
            parse_mode="Markdown",
        )
        return

    if doc.file_size and doc.file_size > MAX_FILE_BYTES:
        size_mb = doc.file_size / 1024 / 1024
        await update.message.reply_text(
            f"⚠️ File too large ({size_mb:.1f} MB). Maximum allowed size is 20 MB."
        )
        return

    status = await update.message.reply_text("⏳ Converting, please wait…")

    with tempfile.TemporaryDirectory() as tmpdir:
        input_path = os.path.join(tmpdir, file_name)

        tg_file = await context.bot.get_file(doc.file_id)
        await tg_file.download_to_drive(input_path)

        try:
            if ext == ".pdf":
                output_name = Path(file_name).stem + ".docx"
                output_path = os.path.join(tmpdir, output_name)
                convert_pdf_to_docx(input_path, output_path)
                caption = "✅ Here is your DOCX file!"
            else:
                output_path = convert_docx_to_pdf(input_path, tmpdir)
                output_name = Path(output_path).name
                caption = "✅ Here is your PDF file!"

            with open(output_path, "rb") as f:
                await update.message.reply_document(
                    document=f,
                    filename=output_name,
                    caption=caption,
                )

            await status.delete()

        except Exception as e:
            logger.error("Conversion failed for %s: %s", file_name, e)
            await status.edit_text(f"❌ Conversion failed:\n{e}")


async def handle_text(update: Update, context: ContextTypes.DEFAULT_TYPE) -> None:
    await update.message.reply_text(
        "Please send me a *PDF*, *DOCX*, or *DOC* file to convert.\n"
        "Use /help for instructions.",
        parse_mode="Markdown",
    )


def main() -> None:
    if not BOT_TOKEN:
        raise ValueError("BOT_TOKEN is not set. Add it to a .env file: BOT_TOKEN=<your_token>")

    _check_libreoffice()

    app = Application.builder().token(BOT_TOKEN).build()
    app.add_handler(CommandHandler("start", start))
    app.add_handler(CommandHandler("help", help_command))
    app.add_handler(MessageHandler(filters.Document.ALL, handle_document))
    app.add_handler(MessageHandler(filters.TEXT & ~filters.COMMAND, handle_text))

    loop = asyncio.new_event_loop()
    asyncio.set_event_loop(loop)

    logger.info("Bot is running…")
    app.run_polling()


if __name__ == "__main__":
    main()
