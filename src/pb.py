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

matplotlib.rcParams['font.family'] = 'Segoe UI'

# TELDA names
telda_names = [
    'BINJAI', 'INNER SUMUT', 'KABANJAHE', 'KISARAN', 'LUBUK PAKAM',
    'PADANGSIDEMPUAN', 'RANTAU PRAPAT', 'SIANTAR', 'SIBOLGA', 'TOBA'
]

# Store PS HI, TGT Real, MTD Date, MTD values
pshi_values = {name: None for name in telda_names}
tgt_values = {name: None for name in telda_names}
real_values = {name: None for name in telda_names}
mtd_values = {name: None for name in telda_names}
tgt_values_ytd = {name: None for name in telda_names}
real_values_ytd = {name: None for name in telda_names}

# Save data to json file
def save_data():
    data = {
        "month1": month1_value,
        "year1": year1_value,
        "month2": month2_value,
        "year2": year2_value,
        "pshi": pshi_values,
        "tgt": tgt_values,
        "real": real_values,
        "mtd": mtd_values,
        "tgt_ytd": tgt_values_ytd,
        "real_ytd": real_values_ytd
    }
    with open("data.json", "w") as f:
        json.dump(data, f, indent=4)

# Load data from json file
def load_data():
    global pshi_values, tgt_values, real_values, month1_value, year1_value, month2_value, year2_value, tgt_values_ytd, real_values_ytd, mtd_values
    try:
        with open("data.json", "r") as f:
            data = json.load(f)
            month1_value = data.get("month1")
            year1_value = data.get("year1")
            month2_value = data.get("month2")
            year2_value = data.get("year2")
            pshi_values = data.get("pshi", {})
            tgt_values = data.get("tgt", {})
            real_values = data.get("real", {})
            mtd_values = data.get("mtd", {})
            tgt_values_ytd = data.get("tgt_ytd", {})
            real_values_ytd = data.get("real_ytd", {})

        print("Data loaded successfully")
        
    except FileNotFoundError:
        print("File not found")
        pass

# Function to create spreadsheet image
def create_spreadsheet_image(df):
    try:
        print(f"Creating table image...")

        fig, ax = plt.subplots(figsize=(10, 0.6 + 0.4 * len(df)))
        ax.axis('off')
        table = ax.table(cellText=df.values, colLabels=df.columns, loc='center', cellLoc='center')

        table.auto_set_font_size(False)
        table.set_fontsize(10)
        table.scale(1.2, 1.5)
        
        fig.text(0.33, 0.82, f"MTD {month1_value.upper()} {year1_value}", fontsize=14, fontweight='bold', ha='left')
        fig.text(0.815, 0.82, f"YTD {month2_value.upper()} {year2_value}", fontsize=14, fontweight='bold', ha='right')

        # Color the first column (TELDA)
        for row in range(len(df)):
            cell = table[(row, 0)]
            cell.set_facecolor("navy")
            cell.set_text_props(weight='bold', color='white')

        # Color the first row (header) — row MTD
        for col in range(6):
            cell = table[(0, col+1)]
            cell.set_facecolor("navy")
            cell.set_text_props(weight='bold', color='white')

        # Row TGT - ACH | YTD
        for col in range(3):
            cell = table[(0, col+6)]
            cell.set_facecolor("steelblue")
            cell.set_text_props(weight='bold', color='white')

        # Row GAP
        cell2 = table[(0, 9)]
        cell2.set_facecolor('red')
        cell2.set_text_props(weight='bold', color='white')

        # Row MoM
        cell3 = table[(0, 10)]
        cell3.set_facecolor("navy")
        cell3.set_text_props(weight='bold', color='white')

        # Row total
        for col in range(len(df.columns)):
            cell = table[(11, col)]
            cell.set_facecolor("navy")
            cell.set_text_props(weight='bold', color='white')

        def get_named_color(value):
            value = max(0, min(value, 100))
            if value < 10:
                return 'red'
            elif value < 20:
                return 'orangered'
            elif value < 30:
                return 'orange'
            elif value < 40:
                return 'gold'
            elif value < 50:
                return 'yellow'
            elif value < 60:
                return 'yellowgreen'
            elif value < 70:
                return 'limegreen'
            elif value < 80:
                return 'green'
            elif value < 90:
                return 'forestgreen'
            else:
                return 'forestgreen'

        for row in range(1, len(df)):
            for col in [4, 8]:
                cell = table[(row, col)]
                try:
                    value_str = str(df.iloc[row-1, col]).strip()
                    if value_str.endswith('%'):
                        value = float(value_str.strip('%'))
                    else:
                        value = float(value_str)
                    cell.set_facecolor(get_named_color(value))
                    cell.set_text_props(color='black')
                except Exception as e:
                    print(f"❌ Could not parse value at row {row}, col {col}: {e}")

        def get_signal_color(value):
            if value < 0:
                return 'firebrick'
            else:
                return 'forestgreen'

        for row in range(1, len(df)):
            for col in [10]:
                cell = table[(row, col)]
                try:
                    value_str = str(df.iloc[row-1, col]).strip()
                    if value_str.endswith('%'):
                        value = float(value_str.strip('%'))
                    else:
                        value = float(value_str)
                    cell.set_facecolor(get_signal_color(value))
                    cell.set_text_props(color='black')
                except Exception as e:
                    print(f"❌ Could not parse value at row {row}, col {col}: {e}")

        # Column widths
        for i, key in enumerate(df.columns):
            for j in range(len(df) + 1):
                cell = table[(j, i)]
                if i == 0:
                    cell.set_width(0.2)
                else:
                    cell.set_width(0.1)

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
        print(f"❌ Error saving data: {e}")
        return None
    
async def set_pshi(update: Update, context: ContextTypes.DEFAULT_TYPE):
    global pshi_values
    try:
        if len(context.args) != len(telda_names):
            await update.message.reply_text(f"❌ Usage: /set_pshi <{len(telda_names)} values")
            return
        
        for idx, value in enumerate(context.args):
            try:
                pshi_values[telda_names[idx]] = int(value)
            except ValueError:
                await update.message.reply_text(f"❌ '{value}' is not a valid integer for {telda_names[idx]}.")
                return

        await update.message.reply_text("✅ PS HI values set successfully")

    except Exception as e:
        print(f"❌ Error set_pshi: {e}")
        await update.message.reply_text(f"❌ Error: {e}")

async def set_mtd_date(update: Update, context: ContextTypes.DEFAULT_TYPE):
    global month1_value
    global year1_value
    global month2_value
    global year2_value
    try:
        if len(context.args) != 4:
            await update.message.reply_text("❌ Usage: /set_mtd_date <month1> <year1> <month2> <year2>")
            return
        
        month1_value = str(context.args[0])
        year1_value = int(context.args[1])
        month2_value = str(context.args[2])
        year2_value = int(context.args[3])

        await update.message.reply_text(f"✅ MTD Date set successfully:\n{month1_value}, {year1_value} | {month2_value}, {year2_value}")

    except ValueError:
        await update.message.reply_text("❌ Please enter again.")

async def set_mtd(update: Update, context: ContextTypes.DEFAULT_TYPE):
    global mtd_values
    try:
        if len(context.args) != len(telda_names):
            await update.message.reply_text(f"❌ Usage: /set_mtd <{len(telda_names)} values")
            return
        
        for idx, value in enumerate(context.args):
            try:
                mtd_values[telda_names[idx]] = int(value)
            except ValueError:
                await update.message.reply_text(f"❌ '{value}' is not a valid integer for {telda_names[idx]}.")
                return

        await update.message.reply_text("✅ MTD values set successfully")

    except Exception as e:
        print(f"❌ Error set_mtd: {e}")
        await update.message.reply_text(f"❌ Error: {e}")

async def set_tgt(update: Update, context: ContextTypes.DEFAULT_TYPE):
    global tgt_values
    try:
        if len(context.args) != len(telda_names):
            await update.message.reply_text(f"❌ Usage: /set_tgt <{len(telda_names)} values")
            return
        
        for idx, value in enumerate(context.args):
            try:
                tgt_values[telda_names[idx]] = int(value)
            except ValueError:
                await update.message.reply_text(f"❌ '{value}' is not a valid integer for {telda_names[idx]}.")
                return

        await update.message.reply_text("✅ TGT values set successfully")

    except Exception as e:
        print(f"❌ Error set_tgt: {e}")
        await update.message.reply_text(f"❌ Error: {e}")

async def set_tgt_ytd(update: Update, context: ContextTypes.DEFAULT_TYPE):
    global tgt_values_ytd
    try:
        if len(context.args) != len(telda_names):
            await update.message.reply_text(f"❌ Usage: /set_tgt_ytd <{len(telda_names)} values")
            return
        
        for idx, value in enumerate(context.args):
            try:
                tgt_values_ytd[telda_names[idx]] = int(value)
            except ValueError:
                await update.message.reply_text(f"❌ '{value}' is not a valid integer for {telda_names[idx]}.")
                return

        await update.message.reply_text("✅ TGT YTD values set successfully")

    except Exception as e:
        print(f"❌ Error set_tgt_ytd: {e}")
        await update.message.reply_text(f"❌ Error: {e}")

async def list_tgt(update: Update, context: ContextTypes.DEFAULT_TYPE):
    msg = "📋 *TGT Values Status:*\n\n"
    for name, value in tgt_values.items():
        msg += f"• {name}: {value}\n"
    
    await update.message.reply_text(msg, parse_mode="Markdown")

async def list_tgt_ytd(update: Update, context: ContextTypes.DEFAULT_TYPE):
    msg = "📋 *TGT YTD Values Status:*\n\n"
    for name, value in tgt_values_ytd.items():
        msg += f"• {name}: {value}\n"
    
    await update.message.reply_text(msg, parse_mode="Markdown")

async def set_real(update: Update, context: ContextTypes.DEFAULT_TYPE):
    global real_values
    try:
        if len(context.args) != len(telda_names):
            await update.message.reply_text(f"❌ Usage: /set_real <{len(telda_names)} values")
            return
        
        for idx, value in enumerate(context.args):
            try:
                real_values[telda_names[idx]] = int(value)
            except ValueError:
                await update.message.reply_text(f"❌ '{value}' is not a valid integer for {telda_names[idx]}.")
                return

        await update.message.reply_text("✅ Real values set successfully")

    except Exception as e:
        print(f"❌ Error set_real: {e}")
        await update.message.reply_text(f"❌ Error: {e}")

async def set_real_ytd(update: Update, context: ContextTypes.DEFAULT_TYPE):
    global real_values_ytd
    try:
        if len(context.args) != len(telda_names):
            await update.message.reply_text(f"❌ Usage: /set_real_ytd <{len(telda_names)} values")
            return
        
        for idx, value in enumerate(context.args):
            try:
                real_values_ytd[telda_names[idx]] = int(value)
            except ValueError:
                await update.message.reply_text(f"❌ '{value}' is not a valid integer for {telda_names[idx]}.")
                return

        await update.message.reply_text("✅ Real YTD values set successfully")

    except Exception as e:
        print(f"❌ Error set_real_ytd: {e}")
        await update.message.reply_text(f"❌ Error: {e}")

async def list_real(update: Update, context: ContextTypes.DEFAULT_TYPE):
    msg = "📋 *Real Values Status:*\n\n"
    for name, value in real_values.items():
        msg += f"• {name}: {value}\n"
    
    await update.message.reply_text(msg, parse_mode="Markdown")

async def list_real_ytd(update: Update, context: ContextTypes.DEFAULT_TYPE):
    msg = "📋 *Real YTD Values Status:*\n\n"
    for name, value in real_values_ytd.items():
        msg += f"• {name}: {value}\n"
    
    await update.message.reply_text(msg, parse_mode="Markdown")

async def print_table(update: Update, context: ContextTypes.DEFAULT_TYPE):
    try:
        if not month1_value or not year1_value or not month2_value or not year2_value:
            await update.message.reply_text("❌ MTD month and year not set. Use /set_mtd_date to set them.")
            return

        rows = []
        total_pshi = 0
        total_tgt = 0
        total_real = 0
        total_tgt_ytd = 0
        total_real_ytd = 0
        total_mtd = 0
        total_gap = 0
        pshi = 0
        gap = 0
        mom = 0

        for name in telda_names:
            if tgt_values[name] is None:
                await update.message.reply_text(f"❌ TGT value for '{name}' is not set. Use /set_tgt to set it.")
                return

            if real_values[name] is None:
                await update.message.reply_text(f"❌ REAL value for '{name}' is not set. Use /set_real to set it.")
                return
            
            if tgt_values_ytd[name] is None:
                await update.message.reply_text(f"❌ TGT YTD value for '{name}' is not set. Use /set_tgt_ytd to set it.")
                return

            if real_values_ytd[name] is None:
                await update.message.reply_text(f"❌ REAL YTD value for '{name}' is not set. Use /set_real_ytd to set it.")
                return

            tgt = tgt_values[name]
            real= real_values[name]
            mtd_value = mtd_values[name]
            tgt_ytd = tgt_values_ytd[name]
            real_ytd = real_values_ytd[name]
            ach = (real / tgt * 100) if tgt != 0 else 0
            ach_ytd = (real_ytd / tgt_ytd * 100) if tgt_ytd != 0 else 0
            gap = tgt_ytd - real_ytd
            mom = ((real - mtd_value) / mtd_value * 100) if mtd_value != 0 else 0

            rows.append([name, pshi, tgt, real, f"{ach:.2f}%", mtd_value, tgt_ytd, real_ytd, f"{ach_ytd:.2f}%", gap, f"{mom:.2f}%"])
            total_tgt += tgt
            total_real += real
            total_tgt_ytd += tgt_ytd
            total_real_ytd += real_ytd
            total_mtd += mtd_value
            total_gap += gap

        # Calculate average ACH for the bottom row
        avg_ach = (total_real / total_tgt * 100) if total_tgt != 0 else 0
        avg_ach_ytd = (total_real_ytd / total_tgt_ytd * 100) if total_tgt_ytd != 0 else 0
        avg_mom = (total_real / total_mtd * 100) if total_mtd != 0 else 0

        # Add total row
        rows.append(["TOTAL", total_pshi, total_tgt, total_real, f"{avg_ach:.2f}%", total_mtd, total_tgt_ytd, total_real_ytd, f"{avg_ach_ytd:.2f}%", total_gap, f"{avg_mom:.2f}%"])

        # Columns title
        columns = ["TELDA", "PS HI", "TGT", "REAL", "ACH", "MTD", "TGT", "REAL", "ACH", "GAP", "MoM"]

        # Create DataFrame
        df = pd.DataFrame(rows, columns=columns)

        # Create and send the table image
        img = create_spreadsheet_image(df)
        if img:
            await update.message.reply_photo(photo=img, caption="✅ Report generated!")
        else:
            await update.message.reply_text("❌ Failed to create image.")

    except Exception as e:
        print(f"❌ Error print_table: {e}")
        await update.message.reply_text(f"❌ Error: {e}")

# Start the bot
async def main():
    app = ApplicationBuilder().token(BOT_TOKEN).build()

    commands = [
        ("set_mtd_date", set_mtd_date),
        ("set_mtd", set_mtd),
        ("set_pshi", set_pshi),
        ("set_tgt", set_tgt),
        ("set_tgt_ytd", set_tgt_ytd),
        ("list_tgt", list_tgt),
        ("list_tgt_ytd", list_tgt_ytd),
        ("set_real", set_real),
        ("set_real_ytd", set_real_ytd),
        ("list_real", list_real),
        ("list_real_ytd", list_real_ytd),
        ("table", print_table),
        ("save", save),
    ]

    for cmd, handler in commands:
        app.add_handler(CommandHandler(cmd, handler))

    print("🤖 Bot is running...")
    load_data()
    await app.run_polling()

if __name__ == '__main__':
    try:
        asyncio.run(main())
    except KeyboardInterrupt:
        print("\nBot stopped by user.")
