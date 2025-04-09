import fs from 'fs';
import fetch from 'node-fetch';
import ExcelJS from 'exceljs';
import nodemailer from 'nodemailer';
import dayjs from 'dayjs';
import utc from 'dayjs/plugin/utc.js';
import timezone from 'dayjs/plugin/timezone.js';

import dotenv from 'dotenv';
dotenv.config();


dayjs.extend(utc);
dayjs.extend(timezone);

// GitHub Secrets via process.env
const SHOP = process.env.SHOP;
const ACCESS_TOKEN = process.env.ACCESS_TOKEN;
const CITY_FILTERS = ["Chennai", "Kanchipuram", "Tiruvallur"];
const EMAIL_USER = process.env.EMAIL_USER;
const EMAIL_PASS = process.env.EMAIL_PASS;
const EMAIL_TO = process.env.RECEIVER_EMAILS.split(',');

const TIMEZONE = "Asia/Kolkata";
const nowIST = dayjs().tz(TIMEZONE);
const today830IST = nowIST.hour(20).minute(30).second(0).millisecond(0);

const endTime = nowIST.isBefore(today830IST)
  ? today830IST.subtract(1, 'day')
  : today830IST;

const startTime = endTime.subtract(1, 'day');

const formattedStart = startTime.toISOString();
const formattedEnd = endTime.toISOString();

console.log(`📦 Fetching orders from ${formattedStart} to ${formattedEnd} for cities: ${CITY_FILTERS.join(", ")}`);

const ordersUrl = `https://${SHOP}.myshopify.com/admin/api/2023-10/orders.json?status=any&created_at_min=${formattedStart}&created_at_max=${formattedEnd}`;

async function fetchOrders() {
  const response = await fetch(ordersUrl, {
    headers: {
      'X-Shopify-Access-Token': ACCESS_TOKEN,
      'Content-Type': 'application/json'
    }
  });

  if (!response.ok) {
    const errorText = await response.text();
    throw new Error(`Failed to fetch orders: ${response.status} - ${errorText}`);
  }

  const data = await response.json();
  return data.orders;
}

function filterOrdersByCities(orders) {
  return orders.filter(order => {
    const city = order.shipping_address?.city?.toLowerCase();
    return CITY_FILTERS.map(c => c.toLowerCase()).includes(city);
  });
}

async function generateExcel(orders) {
  const workbook = new ExcelJS.Workbook();
  const sheet = workbook.addWorksheet("Orders");

  sheet.columns = [
    { header: "Order Number", key: "order_number", width: 15 },
    { header: "Product Title", key: "title", width: 30 },
    { header: "Quantity", key: "quantity", width: 10 },
    { header: "City", key: "city", width: 15 }
  ];

  orders.forEach(order => {
    const city = order.shipping_address?.city || '';
    order.line_items.forEach(item => {
      sheet.addRow({
        order_number: order.name,
        title: item.title,
        quantity: item.quantity,
        city: city
      });
    });
  });

  const timestamp = startTime.format("YYYY-MM-DD-HH-mm");
  const filename = `order-report-${timestamp}.xlsx`;
  await workbook.xlsx.writeFile(filename);
  console.log(`✅ Report generated: ${filename}`);
  return filename;
}

async function sendEmailWithAttachment(filePath) {
  const transporter = nodemailer.createTransport({
    service: "gmail",
    auth: {
      user: EMAIL_USER,
      pass: EMAIL_PASS
    }
  });

  const info = await transporter.sendMail({
    from: `"Order Bot" <${EMAIL_USER}>`,
    to: EMAIL_TO,
    subject: "📦 Shopify Order Report",
    text: "Please find the attached Excel report for the latest filtered orders.",
    attachments: [
      {
        filename: filePath,
        path: `./${filePath}`
      }
    ]
  });

  console.log(`📧 Email sent: ${info.messageId}`);
}

async function run() {
  try {
    const allOrders = await fetchOrders();
    const filteredOrders = filterOrdersByCities(allOrders);
    const filePath = await generateExcel(filteredOrders);
    await sendEmailWithAttachment(filePath);
  } catch (err) {
    console.error("❌ Error:", err.message);
    process.exit(1);
  }
}

run();
