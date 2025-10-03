// shopify-unfulfilled-bangalore.js
import fs from 'fs';
import path from 'path';
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

/* ───────────────────────── Env & constants ───────────────────────── */

const requiredEnv = [
  'SHOP',
  'ACCESS_TOKEN',
  'EMAIL_USER',
  'EMAIL_PASS',
  'RECEIVER_EMAILS',
];

const missing = requiredEnv.filter((k) => !process.env[k]);
if (missing.length) {
  console.error(`Missing required env vars: ${missing.join(', ')}`);
  process.exit(1);
}

const SHOP = process.env.SHOP;
const ACCESS_TOKEN = process.env.ACCESS_TOKEN;
const EMAIL_USER = process.env.EMAIL_USER;
const EMAIL_PASS = process.env.EMAIL_PASS;
const EMAIL_TO = process.env.RECEIVER_EMAILS.split(',').map((s) => s.trim()).filter(Boolean);

// Accept override via env, else default
const CITY_FILTERS = (process.env.CITY_FILTERS || 'Bangalore,Bengaluru')
  .split(',')
  .map((s) => s.trim())
  .filter(Boolean);

const TIMEZONE = 'Asia/Kolkata';
const API_VERSION = process.env.SHOPIFY_API_VERSION || '2023-10';

/* ───────────────────────── Time window (IST) ───────────────────────── */

const nowIST = dayjs().tz(TIMEZONE);
const todayStartIST = nowIST.startOf('day');
const yesterdayStartIST = todayStartIST.subtract(1, 'day');

const formattedStart = yesterdayStartIST.toDate().toISOString(); // UTC ISO
const formattedEnd = todayStartIST.toDate().toISOString();       // UTC ISO

console.log(`📦 Fetching UNFULFILLED orders created from ${formattedStart} to ${formattedEnd} (IST window)`);

/* ───────────────────────── Helpers ───────────────────────── */

const sleep = (ms) => new Promise((r) => setTimeout(r, ms));

function normalizeCity(s) {
  return (s || '')
    .normalize('NFKD')
    .replace(/\p{Diacritic}/gu, '')
    .trim()
    .toLowerCase();
}

const citySet = new Set(CITY_FILTERS.map(normalizeCity));
// Add common aliases to be safe
['bangalore', 'bengaluru'].forEach((c) => citySet.add(c));

function isCityAllowed(city) {
  return citySet.has(normalizeCity(city));
}

function formatFullAddress(address) {
  if (!address) return '';
  const parts = [
    address.name,
    address.address1,
    address.address2,
    address.city,
    address.province,
    address.zip,
    address.country,
  ];
  return parts.filter(Boolean).join(', ');
}

/* ───────────────── Shopify fetch with pagination & retries ───────────────── */

async function fetchWithRetries(url, { headers, attempt = 1, maxAttempts = 5 } = {}) {
  const res = await fetch(url, { headers });

  if (res.status === 429 || (res.status >= 500 && res.status < 600)) {
    if (attempt >= maxAttempts) {
      const txt = await res.text().catch(() => '');
      throw new Error(`Shopify error ${res.status} after ${attempt} attempts: ${txt}`);
    }
    const retryAfter = Number(res.headers.get('Retry-After')) || Math.min(2 ** attempt, 30);
    console.warn(`⚠️  Got ${res.status}. Retrying in ${retryAfter}s (attempt ${attempt}/${maxAttempts})…`);
    await sleep(retryAfter * 1000);
    return fetchWithRetries(url, { headers, attempt: attempt + 1, maxAttempts });
  }

  if (!res.ok) {
    const msg = await res.text().catch(() => '');
    throw new Error(`Failed request ${res.status}: ${msg}`);
  }
  return res;
}

function parseLinkHeader(link) {
  if (!link) return {};
  const parts = link.split(',').map((s) => s.trim());
  const out = {};
  for (const p of parts) {
    const m = p.match(/<([^>]+)>;\s*rel="([^"]+)"/);
    if (m) {
      const [, url, rel] = m;
      out[rel] = url;
    }
  }
  return out;
}

async function fetchOrdersAllPages({ createdAtMinISO, createdAtMaxISO }) {
  const base = `https://${SHOP}.myshopify.com/admin/api/${API_VERSION}/orders.json`;
  // Keep Shopify's unfulfilled filter to reduce payload,
  // but we'll strictly verify locally using fulfillable_quantity.
  const common = `status=any&limit=250&created_at_min=${encodeURIComponent(createdAtMinISO)}&created_at_max=${encodeURIComponent(createdAtMaxISO)}&fulfillment_status=unfulfilled&order=created_at%20asc`;

  let url = `${base}?${common}`;
  const headers = {
    'X-Shopify-Access-Token': ACCESS_TOKEN,
    'Content-Type': 'application/json',
  };

  const all = [];
  let page = 1;

  while (url) {
    console.log(`🔎 Fetching page ${page}…`);
    const res = await fetchWithRetries(url, { headers });
    const data = await res.json();
    const chunk = data?.orders || [];
    all.push(...chunk);

    const link = parseLinkHeader(res.headers.get('Link'));
    url = link.next || null;
    page += 1;
  }
  console.log(`✅ Pulled ${all.length} orders total across ${page - 1} page(s).`);
  return all;
}

/* ───────────────────────── Stricter unfulfillment filter ───────────────────────── */

function isActuallyUnfulfilled(order) {
  // Exclude cancelled outright
  if (order.cancelled_at) return false;

  // Deterministic check: any line with fulfillable_quantity > 0 means not fully fulfilled
  const hasFulfillable =
    Array.isArray(order.line_items) &&
    order.line_items.some((li) => Number(li?.fulfillable_quantity || 0) > 0);

  // Also exclude orders that Shopify already marks as fully fulfilled
  const notFullyFulfilled = order.fulfillment_status !== 'fulfilled';

  return hasFulfillable && notFullyFulfilled;
}

function filterOrdersByCityAndUnfulfilled(orders) {
  const filtered = orders.filter((order) => {
    const city =
      order.shipping_address?.city ||
      order.customer?.default_address?.city ||
      '';
    return isCityAllowed(city) && isActuallyUnfulfilled(order);
  });
  console.log(`🏙️ City-matched & strictly unfulfilled orders: ${filtered.length}`);
  return filtered;
}

/* ───────────────────────── Excel ───────────────────────── */

async function generateExcel(orders) {
  const workbook = new ExcelJS.Workbook();
  const sheet = workbook.addWorksheet('Orders', {
    views: [{ state: 'frozen', xSplit: 0, ySplit: 1 }],
  });

  sheet.columns = [
    { header: 'Order Number', key: 'order_number', width: 18 },
    { header: 'Product Title', key: 'title', width: 36 },
    { header: 'Quantity', key: 'quantity', width: 10 },
    { header: 'City', key: 'city', width: 16 },
    { header: 'Phone', key: 'phone', width: 16 },
    { header: 'Full Address', key: 'address', width: 60 },
    { header: 'Financial Status', key: 'financial_status', width: 18 },
    { header: 'Total Price (₹)', key: 'total_price', width: 16 },
  ];

  sheet.getRow(1).font = { bold: true };
  sheet.autoFilter = {
    from: { row: 1, column: 1 },
    to: { row: 1, column: sheet.columnCount },
  };

  for (const order of orders) {
    const city =
      order.shipping_address?.city ||
      order.customer?.default_address?.city ||
      '';
    const fullAddress = formatFullAddress(order.shipping_address || order.customer?.default_address);
    const phone =
      order.shipping_address?.phone ||
      order.phone ||
      order.customer?.phone ||
      '';
    const financialStatus = order.financial_status || '';
    const totalPrice = Number(order.total_price || 0);

    for (const item of order.line_items || []) {
      sheet.addRow({
        order_number: order.name,
        title: item.title,
        quantity: item.quantity,
        city,
        phone,
        address: fullAddress,
        financial_status: financialStatus,
        total_price: totalPrice,
      });
    }
  }

  const totalCol = sheet.getColumn('total_price');
  totalCol.numFmt = '#,##0.00';

  const stamp = yesterdayStartIST.format('YYYY-MM-DD');
  const filename = `unfulfilled-bangalore-orders-${stamp}.xlsx`;
  const outPath = path.resolve(process.cwd(), filename);
  await workbook.xlsx.writeFile(outPath);
  console.log(`📄 Excel written: ${outPath}`);
  return outPath;
}

/* ───────────────────────── Email ───────────────────────── */

async function sendEmailWithAttachment(filePath) {
  const transporter = nodemailer.createTransport({
    host: 'smtp.gmail.com',
    port: 465,
    secure: true,
    auth: {
      user: EMAIL_USER,
      pass: EMAIL_PASS,
    },
  });

  const dateLabel = yesterdayStartIST.format('YYYY-MM-DD');
  const info = await transporter.sendMail({
    from: `"Order Bot" <${EMAIL_USER}>`,
    to: EMAIL_TO,
    subject: `📦 Unfulfilled Bangalore Orders Report — ${dateLabel}`,
    text: `Attached: Excel report for unfulfilled Bangalore/Bengaluru orders created ${dateLabel}.`,
    attachments: [
      {
        filename: path.basename(filePath),
        path: filePath,
      },
    ],
  });

  console.log(`📧 Email sent: ${info.messageId}`);
}

/* ───────────────────────── Main ───────────────────────── */

async function run() {
  console.log(`🕒 Script start (IST): ${dayjs().tz(TIMEZONE).format('YYYY-MM-DD HH:mm:ss')}`);
  try {
    const allOrders = await fetchOrdersAllPages({
      createdAtMinISO: formattedStart,
      createdAtMaxISO: formattedEnd,
    });

    const filteredOrders = filterOrdersByCityAndUnfulfilled(allOrders);

    if (filteredOrders.length === 0) {
      console.log('ℹ️ No unfulfilled orders found for the specified cities and window.');
      return;
    }

    const filePath = await generateExcel(filteredOrders);
    await sendEmailWithAttachment(filePath);
    console.log('✅ Done.');
  } catch (err) {
    console.error('❌ Error:', err?.message || err);
    process.exit(1);
  }
}

run();
