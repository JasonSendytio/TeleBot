require("dotenv").config();
import TelegramBot from "node-telegram-bot-api";
import { GoogleSpreadsheet } from "google-spreadsheet";
import { launch } from "puppeteer";

// Initialize bot & Google Sheets
const bot = new TelegramBot(process.env.BOT_TOKEN, { polling: true });
const doc = new GoogleSpreadsheet(process.env.SPREADSHEET_ID);

// Authenticate with Google Sheets
async function accessSheet() {
  await doc.useServiceAccountAuth(JSON.parse(process.env.GOOGLE_SERVICE_ACCOUNT_JSON));
  await doc.loadInfo();
  return doc.sheetsByIndex[0]; // First sheet
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
  const multiplication = num1 * num2 * num3;

  // Store in Google Sheets
  const sheet = await accessSheet();
  await sheet.addRow({ Num1: num1, Num2: num2, Num3: num3, Product: multiplication });

  bot.sendMessage(chatId, `✅ Data saved! Multiplication: ${multiplication}`);

  // Generate and send spreadsheet image
  await sendSpreadsheetImage(chatId);
});

// Function to generate image from spreadsheet
async function sendSpreadsheetImage(chatId) {
  const sheetURL = `https://docs.google.com/spreadsheets/d/${process.env.SPREADSHEET_ID}`;

  const browser = await launch();
  const page = await browser.newPage();
  await page.goto(sheetURL, { waitUntil: "networkidle2" });
  await page.screenshot({ path: "spreadsheet.png", fullPage: true });
  await browser.close();

  bot.sendPhoto(chatId, "spreadsheet.png");
}
