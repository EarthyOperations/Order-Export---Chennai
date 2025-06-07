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

// ENV config
const SHOP = process.env.SHOP;
const ACCESS_TOKEN = process.env.ACCESS_TOKEN;
const CITY_FILTERS = ["Bangalore","Bengaluru"];
const EMAIL_USER = process.env.EMAIL_USER;
const EMAIL_PASS = process.env.EMAIL_PASS;
const EMAIL_TO = process.env.RECEIVER_EMAILS.split(',');

// Time range: yesterday 00:00 IST to today 00:00 IST
const TIMEZONE = "Asia/Kolkata";
const nowIST = dayjs().tz(TIMEZONE);
const todayStartIST = nowIST.startOf('day'); // Today 00:00
const yesterdayStartIST = todayStartIST.subtract(1, 'day'); // Yesterday 00:00

const formattedStart = yesterdayStartIST.toISOString(); // UTC
const formattedEnd = todayStartIST.toISOString(); // UTC

console.log(`📦 Fetching orders from ${formattedStart} to ${formattedEnd} for city: ${CITY_FILTERS[0]}`);

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

function filterOrdersByCity(orders) {
  return orders.filter(order => {
    const city = order.shipping_address?.city?.trim().toLowerCase();
    const isCityMatch = CITY_FILTERS.map(c => c.toLowerCase()).includes(city);
    const isNotCancelled = order.cancelled_at === null;
    return isCityMatch && isNotCancelled;
  });
}

function formatFullAddress(address) {
  if (!address) return "";
  const parts = [
    address.name,
    address.address1,
    address.address2,
    address.city,
    address.province,
    address.zip,
    address.country
  ];
  return parts.filter(Boolean).join(", ");
}

async function generateExcel(orders) {
  const workbook = new ExcelJS.Workbook();
  const sheet = workbook.addWorksheet("Orders");

  sheet.columns = [
    { header: "Order ID", key: "order_id", width: 20 },
    { header: "Order Number", key: "order_number", width: 15 },
    { header: "Product Title", key: "title", width: 30 },
    { header: "Quantity", key: "quantity", width: 10 },
    { header: "City", key: "city", width: 15 },
    { header: "Full Address", key: "address", width: 50 }
  ];

  orders.forEach(order => {
    const city = order.shipping_address?.city || '';
    const fullAddress = formatFullAddress(order.shipping_address);
    order.line_items.forEach(item => {
      sheet.addRow({
        order_id: order.id,
        order_number: order.name,
        title: item.title,
        quantity: item.quantity,
        city: city,
        address: fullAddress
      });
    });
  });

  const timestamp = yesterdayStartIST.format("YYYY-MM-DD");
  const filename = `bangalore-orders-${timestamp}.xlsx`;
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
    subject: "📦 Bangalore Orders Report",
    text: "Please find the attached Excel report for Bangalore orders.",
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
    const filteredOrders = filterOrdersByCity(allOrders);
    const filePath = await generateExcel(filteredOrders);
    await sendEmailWithAttachment(filePath);
  } catch (err) {
    console.error("❌ Error:", err.message);
    process.exit(1);
  }
}

run();
