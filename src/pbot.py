import pandas as pd
import matplotlib
import matplotlib.pyplot as plt
from telegram import Update
from telegram.ext import ApplicationBuilder, MessageHandler, ContextTypes, filters
from dotenv import load_dotenv
import nest_asyncio
import asyncio
import os
import io

nest_asyncio.apply()

# Load environment variables from the .env file
load_dotenv()
BOT_TOKEN = os.getenv("BOT_TOKEN")

matplotlib.rcParams['font.family'] = 'Times New Roman'

# Name mapping
name_list = [
    "first", "second", "third", "fourth", "fifth",
    "sixth", "seventh", "eighth", "ninth", "tenth"
]

# Create spreadsheet image from data
def create_spreadsheet_image(rows):
    try:
        df = pd.DataFrame(rows, columns=["Name", "Num1", "Num2", "Num3", "Product"])
        print(f"Creating image for {len(df)} rows...")

        fig, ax = plt.subplots(figsize=(7, 0.5 + 0.4 * len(df)))
        ax.axis('off')
        table = ax.table(cellText=df.values, colLabels=df.columns, loc='center', cellLoc='center')

        table.auto_set_font_size(False)
        table.set_fontsize(12)
        table.scale(1, 1.5)

        img_stream = io.BytesIO()
        plt.savefig(img_stream, format='png', bbox_inches='tight')
        plt.close(fig)

        img_stream.seek(0)
        print("✅ Image created successfully.")
        return img_stream

    except Exception as e:
        print(f"❌ Error creating image: {e}")
        return None

# Telegram message handler
async def handle_message(update: Update, context: ContextTypes.DEFAULT_TYPE):
    try:
        text = update.message.text.strip()
        print(f"Received message:\n{text}")

        rows = []
        lines = text.splitlines()

        for i, line in enumerate(lines):
            parts = line.strip().split()
            if len(parts) != 3:
                await update.message.reply_text("❌ Each line must contain exactly 3 numbers.")
                return

            try:
                num1, num2, num3 = map(int, parts)
            except ValueError:
                await update.message.reply_text("❌ All values must be integers.")
                return

            name = name_list[i] if i < len(name_list) else f"row{i+1}"
            product = num1 + num2 + num3
            rows.append([name, num1, num2, num3, product])

        img = create_spreadsheet_image(rows)

        if img:
            await update.message.reply_photo(photo=img, caption="✅ Spreadsheet image created!")
        else:
            await update.message.reply_text("❌ Failed to create image.")

    except Exception as e:
        print(f"❌ Error: {e}")
        await update.message.reply_text(f"Error: {e}")

# Start the bot
async def main():
    app = ApplicationBuilder().token(BOT_TOKEN).build()
    app.add_handler(MessageHandler(filters.TEXT & ~filters.COMMAND, handle_message))
    print("🤖 Bot is running...")
    await app.run_polling()

# To ensure we can gracefully stop the bot (e.g., on Ctrl+C)
if __name__ == '__main__':
    asyncio.run(main())
