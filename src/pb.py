import pandas as pd
import matplotlib
import matplotlib.pyplot as plt
from telegram import Update
from telegram.ext import ApplicationBuilder, MessageHandler, CommandHandler, ContextTypes, filters
from dotenv import load_dotenv
import nest_asyncio
import asyncio
import json
import os
import io

nest_asyncio.apply()

# Load environment variables from the .env file
load_dotenv()
BOT_TOKEN = os.getenv("BOT_TOKEN")

matplotlib.rcParams['font.family'] = 'Roboto'

# TELDA names are fixed
telda_names = [
    'BINJAI', 'INNER SUMUT', 'KABANJAHE', 'KISARAN', 'LUBUK PAKAM',
    'PADANGSIDEMPUAN', 'RANTAU PRAPAT', 'SIANTAR', 'SIBOLGA', 'TOBA'
]

# Store TGT and Real values per TELDA
tgt_values = {name: None for name in telda_names}
real_values = {name: None for name in telda_names}

# Save data
def save_data():
    data = {
        "tgt": tgt_values,
        "real": real_values,
        "month1": month1_value,
        "year1": year1_value
    }
    with open("data.json", "w") as f:
        json.dump(data, f)

# Load data
def load_data():
    global tgt_values, real_values, month1_value, year1_value
    try:
        with open("data.json", "r") as f:
            data = json.load(f)
            tgt_values = data.get("tgt", {})
            real_values = data.get("real", {})
            month1_value = data.get("month1")
            year1_value = data.get("year1")

        print("Data loaded successfully")
        
    except FileNotFoundError:
        print("File not found")
        pass

# Function to create spreadsheet image
def create_spreadsheet_image(df):
    try:
        print(f"Creating image for {len(df)} rows...")

        fig, ax = plt.subplots(figsize=(6, 0.6 + 0.4 * len(df)))
        ax.axis('off')
        table = ax.table(cellText=df.values, colLabels=df.columns, loc='center', cellLoc='center')

        table.auto_set_font_size(False)
        table.set_fontsize(10)
        table.scale(1.2, 1.5)
        fig.suptitle(f"MTD {month1_value.upper()} {year1_value}", fontsize=14, fontweight='bold')

        # Color the first column (TELDA) — column index 0
        for row in range(len(df) + 1):  # +1 to include header row
            cell = table[(row, 0)]
            cell.set_facecolor("#00FFFF")  # Cyan

        # Color the first row (header) — row index 0
        for col in range(len(df.columns)):
            cell = table[(0, col)]
            cell.set_facecolor("#00FFFF")  # Cyan
            cell.set_text_props(weight='bold')

        img_stream = io.BytesIO()
        plt.savefig(img_stream, format='png', bbox_inches='tight')
        plt.close(fig)

        img_stream.seek(0)
        print("✅ Image created successfully.")
        return img_stream

    except Exception as e:
        print(f"❌ Error creating image: {e}")
        return None

async def save(Update: Update, context:ContextTypes.DEFAULT_TYPE):
    try:
        save_data()
        await Update.message.reply_text("Successfully saved data")
    
    except Exception as e:
        print(f"❌ Error: {e}")
        return None

# Command to set MTD values
async def set_mtd(update: Update, context: ContextTypes.DEFAULT_TYPE):
    global month1_value
    global year1_value
    # global month2_value
    # global year2_value
    try:
        if len(context.args) != 2:
            await update.message.reply_text("❌ Usage: /set_mtd <month1> <year1>")
            return
        
        month1_value = str(context.args[0])
        year1_value = int(context.args[1])
        # month2_value = str(context.args[2])
        # year2_value = int(context.args[3])

        await update.message.reply_text(f"✅ MTD values set successfully:\n{month1_value} = {year1_value}")

    except ValueError:
        await update.message.reply_text("❌ Please enter again.")

# Command to set TGT values per TELDA
async def set_tgt(update: Update, context: ContextTypes.DEFAULT_TYPE):
    global tgt_values
    try:
        if len(context.args) != len(telda_names):
            await update.message.reply_text(f"❌ Usage: /set_tgt <{len(telda_names)} values, one for each TELDA")
            return
        
        for idx, value in enumerate(context.args):
            try:
                tgt_values[telda_names[idx]] = int(value)
            except ValueError:
                await update.message.reply_text(f"❌ '{value}' is not a valid integer for {telda_names[idx]}.")
                return

        await update.message.reply_text("✅ TGT values set successfully for each TELDA!")

    except Exception as e:
        print(f"❌ Error: {e}")
        await update.message.reply_text(f"❌ Error: {e}")

# Command to list MTD values
async def list_tgt(update: Update):
    msg = "📋 *TGT Values Status:*\n\n"
    for name, value in tgt_values.items():
        msg += f"• {name}: {value}\n"
    
    await update.message.reply_text(msg, parse_mode="Markdown")

async def set_real(update: Update, context: ContextTypes.DEFAULT_TYPE):
    global real_values
    try:
        if len(context.args) != len(telda_names):
            await update.message.reply_text(f"❌ Usage: /set_real <{len(telda_names)} values, one for each TELDA")
            return
        
        for idx, value in enumerate(context.args):
            try:
                real_values[telda_names[idx]] = int(value)
            except ValueError:
                await update.message.reply_text(f"❌ '{value}' is not a valid integer for {telda_names[idx]}.")
                return

        await update.message.reply_text("✅ Real values set successfully for each TELDA!")

    except Exception as e:
        print(f"❌ Error: {e}")
        await update.message.reply_text(f"❌ Error: {e}")

async def list_tgt(update: Update):
    msg = "📋 *Real Values Status:*\n\n"
    for name, value in real_values.items():
        msg += f"• {name}: {value}\n"
    
    await update.message.reply_text(msg, parse_mode="Markdown")

async def print_table(update: Update, context: ContextTypes.DEFAULT_TYPE):
    try:
        if not month1_value or not year1_value:
            await update.message.reply_text("❌ MTD month and year not set. Use /set_mtd to set them.")
            return

        rows = []
        total_tgt = 0
        total_real = 0

        for name in telda_names:
            if tgt_values[name] is None:
                await update.message.reply_text(f"❌ TGT value for '{name}' is not set. Use /set_tgt to set it.")
                return

            if real_values[name] is None:
                await update.message.reply_text(f"❌ REAL value for '{name}' is not set. Use /set_real to set it.")
                return

            real_value = real_values[name]
            tgt = tgt_values[name]
            ach = (real_value / tgt * 100) if tgt != 0 else 0

            rows.append([name, tgt, real_value, f"{ach:.2f}%"])
            total_tgt += tgt
            total_real += real_value

        # Calculate average ACH for the bottom row
        avg_ach = (total_real / total_tgt * 100) if total_tgt != 0 else 0

        # Add total row
        rows.append(["TOTAL", total_tgt, total_real, f"{avg_ach:.2f}%"])

        # Columns title
        columns = ["TELDA", "TGT", "REAL", "ACH"]

        # Create DataFrame
        df = pd.DataFrame(rows, columns=columns)

        # Create and send the table image
        img = create_spreadsheet_image(df)
        if img:
            await update.message.reply_photo(photo=img, caption="✅ Report generated!")
        else:
            await update.message.reply_text("❌ Failed to create image.")

    except Exception as e:
        print(f"❌ Error: {e}")
        await update.message.reply_text(f"❌ Error: {e}")

# Start the bot
async def main():
    app = ApplicationBuilder().token(BOT_TOKEN).build()
    app.add_handler(CommandHandler("set_mtd", set_mtd))
    app.add_handler(CommandHandler("set_tgt", set_tgt))
    app.add_handler(CommandHandler("list_tgt", list_tgt))
    app.add_handler(CommandHandler("set_real", set_real))
    app.add_handler(CommandHandler("print_table", print_table))
    app.add_handler(CommandHandler("save", save))
    print("🤖 Bot is running...")
    load_data()
    await app.run_polling()

# To ensure we can gracefully stop the bot (e.g., on Ctrl+C)
if __name__ == '__main__':
    try:
        asyncio.run(main())
    except KeyboardInterrupt:
        print("\nBot stopped by user.")
