import pandas as pd
import matplotlib
import matplotlib.pyplot as plt
from telegram import Update
from telegram.ext import ApplicationBuilder, MessageHandler, CommandHandler, ContextTypes, filters
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

# TELDA names are fixed
telda_names = [
    'BINJAI', 'INNER SUMUT', 'KABANJAHE', 'KISARAN', 'LUBUK PAKAM',
    'PADANGSIDEMPUAN', 'RANTAU PRAPAT', 'SIANTAR', 'SIBOLGA', 'TOBA'
]

# Store MTD and TGT values
mtd_values = {name: {"Maret": None, "April": None} for name in telda_names}
tgt_value = None


# Function to create spreadsheet image
def create_spreadsheet_image(df):
    try:
        print(f"Creating image for {len(df)} rows...")

        fig, ax = plt.subplots(figsize=(12, 0.6 + 0.4 * len(df)))
        ax.axis('off')
        table = ax.table(cellText=df.values, colLabels=df.columns, loc='center', cellLoc='center')

        table.auto_set_font_size(False)
        table.set_fontsize(10)
        table.scale(1.2, 1.5)

        img_stream = io.BytesIO()
        plt.savefig(img_stream, format='png', bbox_inches='tight')
        plt.close(fig)

        img_stream.seek(0)
        print("✅ Image created successfully.")
        return img_stream

    except Exception as e:
        print(f"❌ Error creating image: {e}")
        return None


# Command to set MTD values
async def set_mtd(update: Update, context: ContextTypes.DEFAULT_TYPE):
    try:
        if len(context.args) != 4 or context.args[0].lower() != "maret" or context.args[2].lower() != "april":
            await update.message.reply_text("❌ Usage: /set_mtd Maret <value> April <value>")
            return
        
        maret_value = int(context.args[1])
        april_value = int(context.args[3])

        for name in telda_names:
            mtd_values[name] = {"Maret": maret_value, "April": april_value}

        await update.message.reply_text(f"✅ MTD values set successfully:\nMaret = {maret_value}, April = {april_value}")

    except ValueError:
        await update.message.reply_text("❌ MTD values must be integers.")


# Command to set TGT value
async def set_tgt(update: Update, context: ContextTypes.DEFAULT_TYPE):
    global tgt_value
    try:
        if len(context.args) != 1:
            await update.message.reply_text("❌ Usage: /set_tgt <value>")
            return
        
        tgt_value = int(context.args[0])
        await update.message.reply_text(f"✅ TGT value set successfully: {tgt_value}")

    except ValueError:
        await update.message.reply_text("❌ TGT value must be an integer.")


# Command to list MTD values
async def list_mtd(update: Update, context: ContextTypes.DEFAULT_TYPE):
    msg = "📋 *MTD Values Status:*\n\n"
    for name, values in mtd_values.items():
        status = f"Maret: {values['Maret']}, April: {values['April']}"
        msg += f"{name}: {status}\n"
    
    await update.message.reply_text(msg, parse_mode="Markdown")


# Message handler for REAL inputs
async def handle_real_input(update: Update, context: ContextTypes.DEFAULT_TYPE):
    global tgt_value
    try:
        if not tgt_value:
            await update.message.reply_text("❌ Please set the TGT value using /set_tgt <value> first.")
            return

        text = update.message.text.strip()
        print(f"Received REAL values:\n{text}")

        real_values = text.split()
        
        if len(real_values) != len(telda_names):
            await update.message.reply_text(f"❌ Please provide exactly {len(telda_names)} values, one for each TELDA.")
            return

        rows = []
        for idx, value in enumerate(real_values):
            name = telda_names[idx]
            try:
                real_value = int(value)
            except ValueError:
                await update.message.reply_text(f"❌ '{value}' is not a valid integer.")
                return

            if not mtd_values[name]["Maret"] or not mtd_values[name]["April"]:
                await update.message.reply_text(f"❌ MTD values for '{name}' are not set. Use /set_mtd to configure.")
                return

            # Calculations
            mtd_maret = mtd_values[name]["Maret"]
            mtd_april = mtd_values[name]["April"]

            ach = real_value / tgt_value
            ytd_real = mtd_maret + mtd_april
            ytd_ach = ytd_real / tgt_value
            ytd_gap = ytd_real - tgt_value
            mom = (mtd_april - mtd_maret) / mtd_maret if mtd_maret != 0 else 0

            rows.append([name, tgt_value, real_value, ach, ytd_real, ytd_ach, ytd_gap, mom])

        # Create DataFrame and generate image
        df = pd.DataFrame(rows, columns=["Name", "TGT", "REAL", "ACH", "YtD Real", "YtD ACH", "YtD GAP", "MoM"])
        img = create_spreadsheet_image(df)

        if img:
            await update.message.reply_photo(photo=img, caption="✅ Report generated!")
        else:
            await update.message.reply_text("❌ Failed to create image.")

    except Exception as e:
        print(f"❌ Error: {e}")
        await update.message.reply_text(f"Error: {e}")

# Start the bot
async def main():
    app = ApplicationBuilder().token(BOT_TOKEN).build()
    app.add_handler(CommandHandler("set_mtd", set_mtd))
    app.add_handler(CommandHandler("set_tgt", set_tgt))
    app.add_handler(CommandHandler("list_mtd", list_mtd))
    app.add_handler(MessageHandler(filters.TEXT & ~filters.COMMAND, handle_real_input))
    print("🤖 Bot is running...")
    await app.run_polling()

# To ensure we can gracefully stop the bot (e.g., on Ctrl+C)
if __name__ == '__main__':
    asyncio.run(main())
