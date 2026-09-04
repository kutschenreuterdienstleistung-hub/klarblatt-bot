#!/usr/bin/env node
// ============================================================
//  Klarblatt — Vollautomatisierter E-Mail-Service
//  Bestellformular → Stripe → Claude API → E-Mail-Zustellung
// ============================================================

require('dotenv').config();
const express = require('express');
const Stripe = require('stripe');
const Anthropic = require('@anthropic-ai/sdk');
const PptxGenJS = require('pptxgenjs');
const { Document, Packer, Paragraph, TextRun, HeadingLevel } = require('docx');
const { Resend } = require('resend');
const fs = require('fs');
const path = require('path');

// ── Config ──────────────────────────────────────────────────
const STRIPE_KEY = process.env.STRIPE_SECRET_KEY;
const STRIPE_WH  = process.env.STRIPE_WEBHOOK_SECRET;
const CLAUDE_KEY = process.env.ANTHROPIC_API_KEY;
const RESEND_KEY = process.env.RESEND_API_KEY;
const FROM_EMAIL = process.env.FROM_EMAIL || 'Klarblatt <onboarding@resend.dev>';
const NOTIFY_EMAIL = process.env.NOTIFY_EMAIL || 'kutschenreuter.dienstleistung@gmail.com';
const PORT = process.env.PORT || 3000;
const BASE_URL = process.env.BASE_URL || `http://localhost:${PORT}`;

const stripe    = new Stripe(STRIPE_KEY);
const anthropic = new Anthropic({ apiKey: CLAUDE_KEY });
const resend    = new Resend(RESEND_KEY);
const app       = express();

// ── Klarblatt Brand ─────────────────────────────────────────
const COLORS = {
  DARK: '06393E', TEAL: '028090', MINT: '02C39A', WHITE: 'FFFFFF',
  LIGHT_BG: 'F0F7F4', TEXT: '1A1A2E',
};

// ── Pakete ──────────────────────────────────────────────────
const PACKAGES = {
  auffrischen: {
    name: 'Auffrischen', desc: 'Bestehendes Deck aufpolieren — bis 15 Folien',
    price: 39000, priceLabel: '390 €', slides: 15, type: 'pptx',
  },
  neuaufbau: {
    name: 'Neuaufbau', desc: 'Komplett neues Deck — bis 20 Folien',
    price: 79000, priceLabel: '790 €', slides: 20, type: 'pptx',
  },
  bericht_deck: {
    name: 'Bericht & Deck', desc: 'Report + Präsentation — bis 30 Folien',
    price: 149000, priceLabel: '1.490 €', slides: 30, type: 'both',
  },
  retainer: {
    name: 'Retainer', desc: '4 Auffrischungen pro Monat',
    price: 60000, priceLabel: '600 €/Mo', slides: 15, type: 'pptx',
  },
};

// ── Order Store (file-backed for Render free-plan restarts) ─
const ORDERS_FILE = path.join('/tmp', 'klarblatt-orders.json');
const orders = new Map();

function loadOrders() {
  try {
    if (fs.existsSync(ORDERS_FILE)) {
      const data = JSON.parse(fs.readFileSync(ORDERS_FILE, 'utf8'));
      for (const [k, v] of Object.entries(data)) orders.set(k, v);
      console.log(`📂 ${orders.size} Bestellungen aus Datei geladen`);
    }
  } catch (e) { console.error('Order-Datei Ladefehler:', e.message); }
}

function saveOrders() {
  try {
    const obj = Object.fromEntries(orders);
    fs.writeFileSync(ORDERS_FILE, JSON.stringify(obj));
  } catch (e) { console.error('Order-Datei Speicherfehler:', e.message); }
}

loadOrders();

// ── Safe JSON parse from Claude ─────────────────────────────
function safeParseJSON(text) {
  // Strip markdown code fences if present
  let cleaned = text.trim();
  if (cleaned.startsWith('```')) {
    cleaned = cleaned.replace(/^```(?:json)?\s*\n?/, '').replace(/\n?```\s*$/, '');
  }
  return JSON.parse(cleaned);
}

// ── Bestellformular (HTML) ──────────────────────────────────
app.get('/', (req, res) => {
  res.send(`<!DOCTYPE html>
<html lang="de">
<head>
<meta charset="utf-8"><meta name="viewport" content="width=device-width,initial-scale=1">
<title>Klarblatt — Bestellung</title>
<style>
  :root { --dark:#06393E; --teal:#028090; --mint:#02C39A; --bg:#F0F7F4; --text:#1A1A2E; }
  * { box-sizing:border-box; margin:0; padding:0; }
  body { font-family:'Segoe UI',system-ui,sans-serif; background:var(--bg); color:var(--text); }
  .header { background:var(--dark); color:#fff; padding:2rem; text-align:center; }
  .header h1 { font-size:2rem; font-weight:700; }
  .header p { color:var(--mint); margin-top:.5rem; }
  .container { max-width:640px; margin:2rem auto; padding:0 1rem; }
  .card { background:#fff; border-radius:12px; padding:2rem; box-shadow:0 2px 12px rgba(0,0,0,.08); margin-bottom:1.5rem; }
  .card h2 { color:var(--dark); margin-bottom:1rem; font-size:1.3rem; }
  label { display:block; font-weight:600; margin:1rem 0 .3rem; }
  select, input, textarea { width:100%; padding:.75rem; border:2px solid #ddd; border-radius:8px; font-size:1rem; transition:border .2s; }
  select:focus, input:focus, textarea:focus { border-color:var(--teal); outline:none; }
  textarea { min-height:150px; resize:vertical; }
  .pkg-info { background:var(--bg); padding:1rem; border-radius:8px; margin:.5rem 0; display:none; }
  .pkg-info.active { display:block; }
  .addons { display:flex; flex-wrap:wrap; gap:.5rem; margin-top:.5rem; }
  .addons label { display:flex; align-items:center; gap:.4rem; font-weight:400; background:#fff; border:2px solid #ddd; padding:.5rem 1rem; border-radius:8px; cursor:pointer; }
  .addons input:checked + span { color:var(--teal); font-weight:600; }
  .price-display { text-align:center; font-size:1.8rem; font-weight:700; color:var(--teal); padding:1rem; }
  .btn { display:block; width:100%; padding:1rem; background:var(--mint); color:var(--dark); border:none; border-radius:8px; font-size:1.1rem; font-weight:700; cursor:pointer; transition:background .2s; }
  .btn:hover { background:var(--teal); color:#fff; }
  .footer { text-align:center; padding:2rem; color:#888; font-size:.85rem; }
  .steps { display:flex; justify-content:center; gap:2rem; margin:1.5rem 0; }
  .step { text-align:center; }
  .step-num { width:36px; height:36px; border-radius:50%; background:var(--mint); color:var(--dark); display:flex; align-items:center; justify-content:center; font-weight:700; margin:0 auto .3rem; }
  .step-label { font-size:.85rem; color:#666; }
</style>
</head>
<body>
<div class="header">
  <h1>Klarblatt</h1>
  <p>Professionelle Präsentationen & Reports — KI-gestützt</p>
</div>

<div class="container">
  <div class="steps">
    <div class="step"><div class="step-num">1</div><div class="step-label">Bestellen</div></div>
    <div class="step"><div class="step-num">2</div><div class="step-label">Bezahlen</div></div>
    <div class="step"><div class="step-num">3</div><div class="step-label">Erhalten</div></div>
  </div>

  <form id="orderForm" action="/checkout" method="POST">
    <div class="card">
      <h2>1. Paket wählen</h2>
      <select name="package" id="pkgSelect" required>
        <option value="">— Bitte wählen —</option>
        <option value="auffrischen">Auffrischen — 390 € (bis 15 Folien)</option>
        <option value="neuaufbau">Neuaufbau — 790 € (bis 20 Folien)</option>
        <option value="bericht_deck">Bericht & Deck — 1.490 € (bis 30 Folien)</option>
        <option value="retainer">Retainer — 600 €/Mo (4× Auffrischung)</option>
      </select>
    </div>

    <div class="card">
      <h2>2. Add-Ons (optional)</h2>
      <div class="addons">
        <label><input type="checkbox" name="addon_express" value="1"><span>⚡ Express 24h (+35%)</span></label>
        <label><input type="checkbox" name="addon_bilingual" value="1"><span>🌍 Zweite Sprache (+18%)</span></label>
        <label><input type="checkbox" name="addon_template" value="1"><span>🎨 Master-Vorlage (+450 €)</span></label>
      </div>
    </div>

    <div class="card">
      <h2>3. Dein Briefing</h2>
      <label for="email">E-Mail-Adresse *</label>
      <input type="email" name="email" id="email" placeholder="deine@email.de" required>

      <label for="name">Name</label>
      <input type="text" name="name" id="name" placeholder="Dein Name">

      <label for="briefing">Was soll erstellt werden? *</label>
      <textarea name="briefing" id="briefing" placeholder="Beschreibe Thema, Zielgruppe, Kernaussagen, Stil-Wünsche…" required></textarea>
    </div>

    <div class="price-display" id="totalPrice">390,00 €</div>
    <button type="submit" class="btn">💳 Jetzt bezahlen & bestellen</button>
  </form>

  <div class="footer">
    Klarblatt · Kleinunternehmer §19 UStG — keine MwSt. · Alle Preise netto
  </div>
</div>

<script>
const prices = {auffrischen:39000,neuaufbau:79000,bericht_deck:149000,retainer:60000};
function calc(){
  const pkg = document.getElementById('pkgSelect').value;
  if(!pkg) return;
  let total = prices[pkg];
  if(document.querySelector('[name=addon_express]').checked) total = Math.round(total*1.35);
  if(document.querySelector('[name=addon_bilingual]').checked) total = Math.round(total*1.18);
  if(document.querySelector('[name=addon_template]').checked) total += 45000;
  document.getElementById('totalPrice').textContent = (total/100).toFixed(2).replace('.',',')+' €';
}
document.getElementById('pkgSelect').addEventListener('change',calc);
document.querySelectorAll('.addons input').forEach(i=>i.addEventListener('change',calc));
</script>
</body>
</html>`);
});

// ── Checkout → Stripe ───────────────────────────────────────
app.use('/checkout', express.urlencoded({ extended: true }));
app.post('/checkout', async (req, res) => {
  const { package: pkgKey, email, name, briefing, addon_express, addon_bilingual, addon_template } = req.body;
  const pkg = PACKAGES[pkgKey];
  if (!pkg || !email || !briefing) return res.status(400).send('Fehlende Daten');

  const addons = [];
  let totalCents = pkg.price;

  if (addon_express)  { addons.push('express');  totalCents = Math.round(totalCents * 1.35); }
  if (addon_bilingual){ addons.push('bilingual'); totalCents = Math.round(totalCents * 1.18); }
  if (addon_template) { addons.push('template');  totalCents += 45000; }

  try {
    const session = await stripe.checkout.sessions.create({
      payment_method_types: ['card'],
      customer_email: email,
      line_items: [{
        price_data: {
          currency: 'eur',
          product_data: {
            name: `Klarblatt ${pkg.name}`,
            description: addons.length ? `Add-Ons: ${addons.join(', ')}` : pkg.desc,
          },
          unit_amount: totalCents,
        },
        quantity: 1,
      }],
      mode: 'payment',
      success_url: `${BASE_URL}/danke`,
      cancel_url: BASE_URL,
      metadata: {
        package: pkgKey,
        addons: JSON.stringify(addons),
        // Store briefing in Stripe metadata as fallback
        briefing: briefing.substring(0, 490),
        email,
        name: name || '',
      },
    });

    // Store order data (in memory + file)
    orders.set(session.id, { email, name, pkgKey, addons, briefing, created: Date.now() });
    saveOrders();
    res.redirect(303, session.url);
  } catch (err) {
    console.error('Stripe error:', err.message);
    res.status(500).send('Zahlungsfehler — bitte versuche es erneut.');
  }
});

// ── Danke-Seite ─────────────────────────────────────────────
app.get('/danke', (req, res) => {
  res.send(`<!DOCTYPE html><html><head><meta charset="utf-8"><meta name="viewport" content="width=device-width,initial-scale=1">
<title>Danke — Klarblatt</title>
<style>body{font-family:system-ui;display:flex;align-items:center;justify-content:center;min-height:100vh;background:#F0F7F4;margin:0}
.box{text-align:center;background:#fff;padding:3rem;border-radius:16px;box-shadow:0 4px 20px rgba(0,0,0,.08);max-width:480px}
h1{color:#06393E;font-size:2rem}p{color:#555;margin-top:1rem;line-height:1.6}
.check{font-size:4rem;margin-bottom:1rem}</style></head>
<body><div class="box"><div class="check">✅</div><h1>Danke für deine Bestellung!</h1>
<p>Dein Deliverable wird jetzt erstellt und in wenigen Minuten an deine E-Mail geschickt.</p>
<p style="color:#028090;font-weight:600">Prüfe auch deinen Spam-Ordner.</p></div></body></html>`);
});

// ── Stripe Webhook ──────────────────────────────────────────
app.post('/webhook/stripe', express.raw({ type: 'application/json' }), async (req, res) => {
  let event;
  try {
    event = stripe.webhooks.constructEvent(req.body, req.headers['stripe-signature'], STRIPE_WH);
  } catch (err) {
    console.error('Webhook error:', err.message);
    return res.status(400).send('Webhook Error');
  }

  if (event.type === 'checkout.session.completed') {
    const session = event.data.object;
    let order = orders.get(session.id);

    // Fallback: reconstruct order from Stripe metadata if server restarted
    if (!order && session.metadata) {
      console.log('⚠️ Bestellung nicht im Speicher — rekonstruiere aus Stripe-Metadaten');
      order = {
        email: session.metadata.email || session.customer_email,
        name: session.metadata.name || '',
        pkgKey: session.metadata.package,
        addons: JSON.parse(session.metadata.addons || '[]'),
        briefing: session.metadata.briefing || 'Bitte Briefing per E-Mail nachliefern.',
      };
    }

    if (order) {
      console.log(`✅ Zahlung von ${order.email} — Paket: ${order.pkgKey}`);
      // Respond to Stripe immediately, then process
      res.json({ received: true });
      try {
        await produceAndDeliver(order);
        console.log(`📧 Deliverable an ${order.email} verschickt`);
      } catch (err) {
        console.error('Production error:', err);
        // Notify business owner
        try {
          await resend.emails.send({
            from: FROM_EMAIL, to: NOTIFY_EMAIL,
            subject: `⚠️ Klarblatt Fehler — Bestellung von ${order.email}`,
            text: `Fehler bei Produktion:\n${err.message}\n\nBriefing:\n${order.briefing}\n\nPaket: ${order.pkgKey}`,
          });
        } catch (mailErr) {
          console.error('Fehlerbenachrichtigung fehlgeschlagen:', mailErr.message);
        }
      }
      orders.delete(session.id);
      saveOrders();
      return; // Already responded
    }
  }
  res.json({ received: true });
});

// ── Produce & Deliver ───────────────────────────────────────
async function produceAndDeliver(order) {
  const pkg = PACKAGES[order.pkgKey];
  if (!pkg) throw new Error(`Unbekanntes Paket: ${order.pkgKey}`);
  const attachments = [];

  // Generate PPTX
  if (pkg.type === 'pptx' || pkg.type === 'both') {
    const slides = await generateSlideContent(order.briefing, pkg.slides);
    const filePath = await buildPptx(slides);
    attachments.push({ filename: 'Klarblatt-Praesentation.pptx', path: filePath });
  }

  // Generate Report
  if (pkg.type === 'both') {
    const report = await generateReportContent(order.briefing);
    const filePath = await buildDocx(report);
    attachments.push({ filename: 'Klarblatt-Report.docx', path: filePath });
  }

  // Send email via Resend
  const emailAttachments = attachments.map(a => ({
    filename: a.filename,
    content: fs.readFileSync(a.path),
  }));

  await resend.emails.send({
    from: FROM_EMAIL,
    to: order.email,
    subject: `Dein Klarblatt ${pkg.name} ist fertig!`,
    html: `
      <div style="font-family:system-ui;max-width:520px;margin:0 auto">
        <div style="background:#06393E;color:#fff;padding:1.5rem;text-align:center;border-radius:12px 12px 0 0">
          <h1 style="margin:0;font-size:1.5rem">Klarblatt</h1>
        </div>
        <div style="background:#fff;padding:2rem;border:1px solid #eee;border-radius:0 0 12px 12px">
          <p>Hallo ${order.name || 'dort'},</p>
          <p>dein <strong>${pkg.name}</strong> ist fertig! Die Datei${emailAttachments.length > 1 ? 'en sind' : ' ist'} angehängt.</p>
          <p>Zufrieden? Empfiehl uns gerne weiter oder bestelle erneut:</p>
          <p style="text-align:center">
            <a href="${BASE_URL}" style="background:#02C39A;color:#06393E;padding:.75rem 2rem;border-radius:8px;text-decoration:none;font-weight:700;display:inline-block">Neue Bestellung</a>
          </p>
          <p style="color:#888;font-size:.85rem;margin-top:2rem">Klarblatt · Professionelle Präsentationen & Reports</p>
        </div>
      </div>
    `,
    attachments: emailAttachments,
  });

  // Notify Michael about new order
  try {
    await resend.emails.send({
      from: FROM_EMAIL, to: NOTIFY_EMAIL,
      subject: `🎉 Neue Klarblatt-Bestellung: ${pkg.name} von ${order.name || order.email}`,
      text: `Neue Bestellung eingegangen!\n\nKunde: ${order.name || '(kein Name)'}\nE-Mail: ${order.email}\nPaket: ${pkg.name} (${pkg.priceLabel})\nAdd-Ons: ${order.addons.length ? order.addons.join(', ') : 'keine'}\n\nBriefing:\n${order.briefing}\n\nDeliverable wurde automatisch erstellt und zugestellt.`,
    });
  } catch (e) {
    console.error('Benachrichtigung an Michael fehlgeschlagen:', e.message);
  }

  // Cleanup temp files
  for (const a of attachments) {
    try { fs.unlinkSync(a.path); } catch (e) { /* ignore */ }
  }
}

// ── Claude API: Slides ──────────────────────────────────────
async function generateSlideContent(briefing, maxSlides) {
  const response = await anthropic.messages.create({
    model: 'claude-sonnet-4-20250514',
    max_tokens: 4000,
    messages: [{
      role: 'user',
      content: `Du bist ein professioneller Präsentationsdesigner für "Klarblatt".
Erstelle den Inhalt für eine Präsentation mit maximal ${maxSlides} Folien.

Kundenbriefing:
${briefing}

Antworte als JSON-Array. Jede Folie hat:
- "title": Folientitel
- "bullets": Array mit 3-5 Stichpunkten
- "notes": Sprechernotizen (1-2 Sätze)

Antworte NUR mit dem JSON-Array, kein Markdown.`
    }],
  });

  try {
    return safeParseJSON(response.content[0].text);
  } catch (err) {
    console.error('Claude Slide-JSON Parse-Fehler:', err.message);
    console.error('Rohtext:', response.content[0].text.substring(0, 500));
    throw new Error('Claude-Antwort konnte nicht als JSON gelesen werden. Bitte erneut versuchen.');
  }
}

// ── Claude API: Report ──────────────────────────────────────
async function generateReportContent(briefing) {
  const response = await anthropic.messages.create({
    model: 'claude-sonnet-4-20250514',
    max_tokens: 6000,
    messages: [{
      role: 'user',
      content: `Du bist ein professioneller Business-Autor für "Klarblatt".
Erstelle einen ausführlichen Report.

Kundenbriefing:
${briefing}

Antworte als JSON:
{"title":"Report-Titel","sections":[{"heading":"Titel","body":"Fließtext..."}]}

5-8 Abschnitte. NUR JSON, kein Markdown.`
    }],
  });

  try {
    return safeParseJSON(response.content[0].text);
  } catch (err) {
    console.error('Claude Report-JSON Parse-Fehler:', err.message);
    console.error('Rohtext:', response.content[0].text.substring(0, 500));
    throw new Error('Claude-Antwort konnte nicht als JSON gelesen werden. Bitte erneut versuchen.');
  }
}

// ── PPTX Builder ────────────────────────────────────────────
async function buildPptx(slides) {
  const pptx = new PptxGenJS();
  pptx.defineLayout({ name: 'KB', width: 13.33, height: 7.5 });
  pptx.layout = 'KB';

  const titleSlide = pptx.addSlide();
  titleSlide.background = { color: COLORS.DARK };
  titleSlide.addText(slides[0]?.title || 'Präsentation', {
    x: 0.8, y: 2.0, w: 11.7, h: 1.5,
    fontSize: 36, fontFace: 'Cambria', color: COLORS.WHITE, bold: true,
  });
  titleSlide.addText('Erstellt von Klarblatt', {
    x: 0.8, y: 4.0, w: 11.7, h: 0.6,
    fontSize: 14, fontFace: 'Calibri', color: COLORS.MINT,
  });

  for (let i = 1; i < slides.length; i++) {
    const s = slides[i];
    const slide = pptx.addSlide();
    slide.background = { color: COLORS.WHITE };
    slide.addShape(pptx.ShapeType.rect, { x: 0, y: 0, w: 0.15, h: 7.5, fill: { color: COLORS.TEAL } });
    slide.addText(s.title, { x: 0.6, y: 0.4, w: 12, h: 0.8, fontSize: 28, fontFace: 'Cambria', color: COLORS.DARK, bold: true });
    if (s.bullets?.length) {
      slide.addText(
        s.bullets.map(b => ({ text: b, options: { bullet: { code: '25CF', color: COLORS.MINT } } })),
        { x: 0.8, y: 1.6, w: 11, h: 4.5, fontSize: 18, fontFace: 'Calibri', color: COLORS.TEXT, lineSpacingMultiple: 1.5 }
      );
    }
    if (s.notes) slide.addNotes(s.notes);
  }

  const endSlide = pptx.addSlide();
  endSlide.background = { color: COLORS.TEAL };
  endSlide.addText('Vielen Dank', { x: 0.8, y: 2.5, w: 11.7, h: 1.2, fontSize: 40, fontFace: 'Cambria', color: COLORS.WHITE, bold: true, align: 'center' });
  endSlide.addText('klarblatt.de', { x: 0.8, y: 4.2, w: 11.7, h: 0.6, fontSize: 16, fontFace: 'Calibri', color: COLORS.WHITE, align: 'center' });

  const out = path.join('/tmp', `kb-${Date.now()}.pptx`);
  await pptx.writeFile({ fileName: out });
  return out;
}

// ── DOCX Builder ────────────────────────────────────────────
async function buildDocx(report) {
  const children = [
    new Paragraph({ text: report.title, heading: HeadingLevel.TITLE, spacing: { after: 400 } }),
    new Paragraph({ children: [new TextRun({ text: 'Erstellt von Klarblatt', italics: true, color: COLORS.TEAL, size: 24 })], spacing: { after: 600 } }),
  ];
  for (const sec of report.sections) {
    children.push(new Paragraph({ text: sec.heading, heading: HeadingLevel.HEADING_1, spacing: { before: 400, after: 200 } }));
    for (const p of sec.body.split('\n').filter(x => x.trim())) {
      children.push(new Paragraph({ children: [new TextRun({ text: p, font: 'Calibri', size: 22 })], spacing: { after: 200 } }));
    }
  }
  const doc = new Document({ sections: [{ properties: {}, children }] });
  const out = path.join('/tmp', `kb-report-${Date.now()}.docx`);
  fs.writeFileSync(out, await Packer.toBuffer(doc));
  return out;
}

// ── Health ──────────────────────────────────────────────────
app.get('/health', (req, res) => res.json({ status: 'ok', service: 'klarblatt', orders: orders.size }));

// ── Robots.txt (allow SEO crawling) ─────────────────────────
app.get('/robots.txt', (req, res) => {
  res.type('text/plain').send('User-agent: *\nAllow: /\n');
});

// ── Start ───────────────────────────────────────────────────
app.listen(PORT, () => {
  console.log(`✅ Klarblatt läuft auf http://localhost:${PORT}`);
  console.log(`📋 Bestellformular: http://localhost:${PORT}`);
  console.log(`💳 Stripe Webhook:  ${BASE_URL}/webhook/stripe`);
});
