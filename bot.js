import "dotenv/config";
import TelegramBot from "node-telegram-bot-api";
import { GoogleSpreadsheet } from "google-spreadsheet";
import { JWT } from "google-auth-library";
import puppeteer from "puppeteer";

const bot = new TelegramBot(process.env.BOT_TOKEN, { polling: true });

// Authenticate with Google Sheets
async function accessSheet() {
  try {
    const auth = new JWT({
      email: process.env.GOOGLE_SERVICE_ACCOUNT_EMAIL,
      key: process.env.GOOGLE_PRIVATE_KEY,
      scopes: ["https://www.googleapis.com/auth/spreadsheets"],
    });
    const doc = new GoogleSpreadsheet(process.env.SPREADSHEET_ID, auth);
    await doc.loadInfo(); // Load spreadsheet details
    console.log("✅ Google Sheets authenticated successfully.");
    const sheet = doc.sheetsByIndex[0]; // Get first sheet
    if (!sheet) {
      throw new Error("❌ No sheets found in the document!");
    }
    return sheet;
  } catch (error) {
    console.error("❌ Error authenticating Google Sheets:", error);
    return null;
  }
}

// Handle messages
bot.on("message", async (msg) => {
  const chatId = msg.chat.id;
  const text = msg.text.trim();
  
  // Validate input (should be 3 numbers)
  const numbers = text.split(" ").map(Number).filter(n => !isNaN(n));
  if (numbers.length !== 3) {
    bot.sendMessage(chatId, "❌ Please send exactly 3 numbers separated by spaces.");
    return;
  }

  const [num1, num2, num3] = numbers;
  const total = num1 + num2 + num3;

  // Store in Google Sheets
  const sheet = await accessSheet();
  await sheet.addRow({ 
    Num1: num1,
    Num2: num2,
    Num3: num3,
    Product: total });

  bot.sendMessage(chatId, `✅ Data saved! Product: ${total}`);

  await sendSpreadsheetImage(chatId);
  await clearSheet()
});

async function sendSpreadsheetImage(chatId) {
  try {
    const sheetURL = `https://docs.google.com/spreadsheets/d/${process.env.SPREADSHEET_ID}`;
    const browser = await puppeteer.launch();
    const page = await browser.newPage();
    await page.goto(sheetURL, { waitUntil: "networkidle2" });
    await page.waitForSelector("canvas", { timeout: 60000 });
    const tableElement = await page.$("canvas")
    const screenshotBuffer = Buffer.from(await tableElement.screenshot())
    await browser.close();
    bot.sendPhoto(chatId, screenshotBuffer);
  } catch (error) {
    console.log(error.message)
  }
}

async function clearSheet() {
  try {
    const sheet = await accessSheet();
    const rows = await sheet.getRows(); // Get all rows (excluding header)

    if (rows.length === 0) {
      console.log("No data to clear.");
      return;
    }
    // Delete each row
    for (const row of rows) {
      await row.delete();
    }
    console.log("✅ Sheet cleared, headers preserved.");
  } catch (error) {
    console.error("❌ Error clearing sheet:", error.message);
  }
}
