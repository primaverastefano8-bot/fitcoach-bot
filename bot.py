import os
import json
import re
import tempfile
import anthropic
from telegram import Update
from telegram.ext import Application, CommandHandler, MessageHandler, filters, ContextTypes
from openpyxl import Workbook
from openpyxl.styles import PatternFill, Font, Alignment, Border, Side
from openpyxl.utils import get_column_letter

TELEGRAM_TOKEN = os.environ.get("TELEGRAM_TOKEN")
ANTHROPIC_API_KEY = os.environ.get("ANTHROPIC_API_KEY")

client = anthropic.Anthropic(api_key=ANTHROPIC_API_KEY)

SYSTEM_PROMPT = """Sei un coach fitness esperto che si basa sulla metodologia scientifica di Jeff Nippard.

Il tuo processo per creare la scheda:
1. Cerca su internet schede di allenamento esistenti per gli obiettivi dell'utente
2. Cerca su YouTube "Jeff Nippard [gruppo muscolare] exercise" per trovare gli esercizi che lui raccomanda
3. Cerca "Jeff Nippard program [obiettivo]" per la sua metodologia di programmazione
4. Combina le schede esistenti con gli esercizi di Jeff Nippard per creare un programma ottimale
5. Per ogni esercizio includi il link al video YouTube di Jeff Nippard più pertinente

Principi di Jeff Nippard da seguire sempre:
- Esercizi basati sull'evidenza scientifica
- Enfasi sulla meccanica muscolare e range of motion completo
- Progressione di carico sistematica
- Volume e intensità bilanciati
- Preferenza per esercizi compound + isolamento strategico

Rispondi SOLO con un oggetto JSON valido, nessun testo prima o dopo, nessun markdown.

Il JSON deve avere ESATTAMENTE questa struttura:
{
  "nome_utente": "Nome",
  "obiettivo": "Obiettivo principale",
  "livello": "Principiante/Intermedio/Avanzato",
  "giorni_settimana": 4,
  "metodologia": "Breve descrizione dell'approccio scientifico usato (stile Jeff Nippard)",
  "note_generali": "Consigli generali su nutrizione e recupero secondo la scienza",
  "progressione": "Piano dettagliato progressione settimane 1-8 con carichi e volume",
  "giorni": [
    {
      "giorno": "Lunedì",
      "focus": "Petto e Tricipiti",
      "esercizi": [
        {
          "nome": "Panca Piana con Bilanciere",
          "serie": 4,
          "ripetizioni": "6-8",
          "recupero": "2-3 min",
          "note": "Schiena arcuata, busto fermo, discesa controllata 2-3 secondi",
          "perche": "Esercizio compound primario per petto, raccomandato da Jeff Nippard per massima attivazione delle fibre",
          "video_jeff": "https://www.youtube.com/results?search_query=Jeff+Nippard+bench+press"
        }
      ]
    }
  ]
}"""


def crea_excel(dati: dict, nome_file: str) -> str:
    wb = Workbook()

    NERO      = "1A1A2E"
    ROSSO     = "E63946"
    GRIGIO_SC = "2D2D2D"
    GRIGIO_CH = "F5F5F5"
    BIANCO    = "FFFFFF"
    GIALLO    = "FFD700"
    COLORI_GIORNI = ["16213E","0F3460","533483","2C2C54","1B4332","3D0C02"]

    def fill(h):   return PatternFill("solid", fgColor=h)
    def fnt(bold=False, color=BIANCO, size=11):
        return Font(bold=bold, color=color, name="Calibri", size=size)
    def brd():
        s = Side(style="thin", color="CCCCCC")
        return Border(left=s, right=s, top=s, bottom=s)
    def ctr(wrap=False): return Alignment(horizontal="center", vertical="center", wrap_text=wrap)
    def lft(wrap=False): return Alignment(horizontal="left",   vertical="center", wrap_text=wrap)

    # ══════════════════════════════════════════
    # FOGLIO 1 – SCHEDA SETTIMANALE
    # ══════════════════════════════════════════
    ws1 = wb.active
    ws1.title = "📅 Scheda Settimanale"
    ws1.sheet_view.showGridLines = False

    # Header brand
    ws1.merge_cells("A1:I1")
    ws1.row_dimensions[1].height = 50
    c = ws1["A1"]
    c.value     = "🏋️  FITCOACH PRO  |  Metodo Jeff Nippard  |  Il tuo Personal Trainer AI"
    c.fill      = fill(NERO)
    c.font      = Font(bold=True, color=GIALLO, name="Calibri", size=16)
    c.alignment = ctr()

    # Info utente
    ws1.merge_cells("A2:I2")
    ws1.row_dimensions[2].height = 28
    c = ws1["A2"]
    c.value     = (f"👤 {dati.get('nome_utente','Utente')}   |   "
                   f"🎯 {dati.get('obiettivo','—')}   |   "
                   f"📊 {dati.get('livello','—')}   |   "
                   f"📅 {dati.get('giorni_settimana','—')} giorni/settimana")
    c.fill      = fill(ROSSO)
    c.font      = fnt(bold=True, size=11)
    c.alignment = ctr()

    # Metodologia
    ws1.merge_cells("A3:I3")
    ws1.row_dimensions[3].height = 36
    c = ws1["A3"]
    c.value     = f"🔬 METODOLOGIA: {dati.get('metodologia','—')}"
    c.fill      = fill(GRIGIO_SC)
    c.font      = fnt(bold=False, size=10)
    c.alignment = lft(wrap=True)

    # Intestazioni colonne
    headers = ["GIORNO","FOCUS","ESERCIZIO","SERIE","REPS","RECUPERO","NOTE TECNICHE","PERCHÉ","🎥 VIDEO JEFF NIPPARD"]
    widths  = [13, 18, 26, 7, 10, 10, 34, 34, 36]
    ws1.row_dimensions[4].height = 30
    for col, (h, w) in enumerate(zip(headers, widths), 1):
        ws1.column_dimensions[get_column_letter(col)].width = w
        c = ws1.cell(row=4, column=col, value=h)
        c.fill      = fill(GRIGIO_SC)
        c.font      = fnt(bold=True, size=10)
        c.alignment = ctr()
        c.border    = brd()

    # Dati esercizi
    row = 5
    for idx, giorno in enumerate(dati.get("giorni", [])):
        col_g = COLORI_GIORNI[idx % len(COLORI_GIORNI)]
        esercizi = giorno.get("esercizi", [])
        prima = row

        for i, ex in enumerate(esercizi):
            ws1.row_dimensions[row].height = 42

            if i == 0:
                if len(esercizi) > 1:
                    ws1.merge_cells(f"A{prima}:A{prima+len(esercizi)-1}")
                    ws1.merge_cells(f"B{prima}:B{prima+len(esercizi)-1}")
                for col_idx, val in enumerate([giorno.get("giorno",""), giorno.get("focus","")], 1):
                    c = ws1.cell(row=prima, column=col_idx, value=val)
                    c.fill = fill(col_g); c.font = fnt(bold=True, size=10)
                    c.alignment = ctr(wrap=True); c.border = brd()

            bg = GRIGIO_CH if i % 2 == 0 else BIANCO
            valori = [
                ex.get("nome",""),
                ex.get("serie",""),
                ex.get("ripetizioni",""),
                ex.get("recupero",""),
                ex.get("note",""),
                ex.get("perche",""),
                ex.get("video_jeff","")
            ]
            for col_off, val in enumerate(valori):
                c = ws1.cell(row=row, column=3+col_off, value=val)
                c.fill      = fill(bg)
                c.font      = Font(color=GRIGIO_SC, name="Calibri", size=10,
                                   bold=(col_off == 0))
                c.alignment = ctr() if col_off < 4 else lft(wrap=True)
                c.border    = brd()
                # Link video cliccabile
                if col_off == 6 and val and val.startswith("http"):
                    ws1.cell(row=row, column=9).hyperlink = val
                    ws1.cell(row=row, column=9).font = Font(
                        color="0563C1", name="Calibri", size=10, underline="single")
            row += 1

        # Separatore
        for col in range(1, 10):
            c = ws1.cell(row=row, column=col, value="")
            c.fill = fill(NERO)
        ws1.row_dimensions[row].height = 5
        row += 1

    # ══════════════════════════════════════════
    # FOGLIO 2 – SPIEGAZIONI
    # ══════════════════════════════════════════
    ws2 = wb.create_sheet("📖 Spiegazioni")
    ws2.sheet_view.showGridLines = False

    ws2.merge_cells("A1:E1")
    ws2.row_dimensions[1].height = 50
    c = ws2["A1"]
    c.value     = "🏋️  FITCOACH PRO  |  Guida Scientifica agli Esercizi  (Metodo Jeff Nippard)"
    c.fill      = fill(NERO)
    c.font      = Font(bold=True, color=GIALLO, name="Calibri", size=15)
    c.alignment = ctr()

    hdrs2   = ["ESERCIZIO","GIORNO / FOCUS","NOTE TECNICHE","PERCHÉ QUESTO ESERCIZIO","🎥 VIDEO"]
    widths2 = [26, 20, 44, 44, 36]
    ws2.row_dimensions[2].height = 30
    for col, (h, w) in enumerate(zip(hdrs2, widths2), 1):
        ws2.column_dimensions[get_column_letter(col)].width = w
        c = ws2.cell(row=2, column=col, value=h)
        c.fill = fill(GRIGIO_SC); c.font = fnt(bold=True, size=10)
        c.alignment = ctr(); c.border = brd()

    r2 = 3
    for idx, giorno in enumerate(dati.get("giorni", [])):
        col_g = COLORI_GIORNI[idx % len(COLORI_GIORNI)]
        for i, ex in enumerate(giorno.get("esercizi", [])):
            ws2.row_dimensions[r2].height = 52
            bg = GRIGIO_CH if i % 2 == 0 else BIANCO

            vals = [
                ex.get("nome",""),
                f"{giorno.get('giorno','')} – {giorno.get('focus','')}",
                ex.get("note",""),
                ex.get("perche",""),
                ex.get("video_jeff","")
            ]
            fills_row = [bg, col_g, bg, bg, bg]
            fonts_row = [
                Font(bold=True, color=GRIGIO_SC, name="Calibri", size=10),
                fnt(size=10),
                Font(color=GRIGIO_SC, name="Calibri", size=10),
                Font(color=GRIGIO_SC, name="Calibri", size=10),
                Font(color="0563C1", name="Calibri", size=10, underline="single")
            ]
            aligns_row = [lft(), ctr(wrap=True), lft(wrap=True), lft(wrap=True), lft(wrap=True)]

            for col_idx, (val, f_, fn_, al_) in enumerate(zip(vals, fills_row, fonts_row, aligns_row), 1):
                c = ws2.cell(row=r2, column=col_idx, value=val)
                c.fill = fill(f_); c.font = fn_; c.alignment = al_; c.border = brd()
                if col_idx == 5 and val and val.startswith("http"):
                    ws2.cell(row=r2, column=5).hyperlink = val
            r2 += 1

    # ══════════════════════════════════════════
    # FOGLIO 3 – PROGRESSIONE
    # ══════════════════════════════════════════
    ws3 = wb.create_sheet("📈 Progressione")
    ws3.sheet_view.showGridLines = False
    ws3.column_dimensions["A"].width = 28
    ws3.column_dimensions["B"].width = 72
    ws3.column_dimensions["C"].width = 4

    ws3.merge_cells("A1:B1")
    ws3.row_dimensions[1].height = 50
    c = ws3["A1"]
    c.value     = "🏋️  FITCOACH PRO  |  Piano di Progressione Scientifica"
    c.fill      = fill(NERO)
    c.font      = Font(bold=True, color=GIALLO, name="Calibri", size=16)
    c.alignment = ctr()

    sezioni = [
        ("🎯 Obiettivo",                    dati.get("obiettivo","—")),
        ("📊 Livello",                      dati.get("livello","—")),
        ("🔬 Metodologia Jeff Nippard",     dati.get("metodologia","—")),
        ("📈 Progressione Settimane 1-8",   dati.get("progressione","—")),
        ("🥗 Nutrizione & Recupero",        dati.get("note_generali","—")),
    ]
    r3 = 2
    for titolo, valore in sezioni:
        ws3.row_dimensions[r3].height = max(32, min(120, len(str(valore)) // 2))
        c = ws3.cell(row=r3, column=1, value=titolo)
        c.fill = fill(ROSSO); c.font = fnt(bold=True, size=11)
        c.alignment = lft(); c.border = brd()
        c = ws3.cell(row=r3, column=2, value=valore)
        c.fill = fill(GRIGIO_CH)
        c.font = Font(color=GRIGIO_SC, name="Calibri", size=10)
        c.alignment = lft(wrap=True); c.border = brd()
        r3 += 1

    # Footer
    footer = "Creato da FitCoach Pro 🤖  |  Basato sulla metodologia scientifica di Jeff Nippard  |  Personal Trainer AI su Telegram"
    for ws in [ws1, ws2, ws3]:
        last = ws.max_row + 2
        ws.merge_cells(f"A{last}:I{last}")
        c = ws.cell(row=last, column=1, value=footer)
        c.fill = fill(NERO)
        c.font = Font(color=GIALLO, name="Calibri", size=9, italic=True)
        c.alignment = ctr()

    wb.save(nome_file)
    return nome_file


async def start(update: Update, context: ContextTypes.DEFAULT_TYPE):
    await update.message.reply_text(
        "👋 *Benvenuto in FitCoach Pro!*\n\n"
        "Creo schede personalizzate basate sulla *metodologia scientifica di Jeff Nippard* 🔬\n\n"
        "Dimmi:\n"
        "✅ I tuoi *obiettivi* (massa, forza, dimagrimento…)\n"
        "✅ I tuoi *punti deboli* (es. tricipiti scarsi)\n"
        "✅ Quanti *giorni a settimana* ti alleni\n"
        "✅ La tua *esperienza* (principiante/intermedio/avanzato)\n"
        "✅ Eventuali *infortuni o limitazioni*\n\n"
        "Riceverai un file Excel professionale con scheda, spiegazioni e piano di progressione! 🔥",
        parse_mode="Markdown"
    )


async def help_command(update: Update, context: ContextTypes.DEFAULT_TYPE):
    await update.message.reply_text(
        "💡 *Come usarmi:*\n\n"
        "Descrivimi la tua situazione in un messaggio!\n\n"
        "_Esempio: 'Ho i tricipiti carenti, mi alleno 4 giorni, "
        "livello intermedio, voglio aumentare la massa.'_\n\n"
        "Riceverai un Excel con la tua scheda! 💪",
        parse_mode="Markdown"
    )


async def crea_scheda(update: Update, context: ContextTypes.DEFAULT_TYPE):
    user_message = update.message.text
    user_name    = update.message.from_user.first_name

    attesa = await update.message.reply_text(
        "⏳ Sto cercando gli esercizi di Jeff Nippard e creando la tua scheda...\n"
        "Circa 30-40 secondi 🔄"
    )

    try:
        response = client.messages.create(
            model="claude-sonnet-4-6",
            max_tokens=4000,
            system=SYSTEM_PROMPT,
            tools=[{"type": "web_search_20250305", "name": "web_search"}],
            messages=[{
                "role": "user",
                "content": (
                    f"Crea una scheda per {user_name}. "
                    f"Informazioni: {user_message}\n\n"
                    f"Cerca su YouTube i video di Jeff Nippard per ogni esercizio "
                    f"e cerca schede esistenti online come riferimento. "
                    f"Rispondi SOLO con il JSON."
                )
            }]
        )

        testo = "".join(b.text for b in response.content if b.type == "text")
        testo = re.sub(r"```json|```", "", testo).strip()
        match = re.search(r"\{.*\}", testo, re.DOTALL)
        if not match:
            raise ValueError("JSON non trovato")
        dati = json.loads(match.group())

        with tempfile.NamedTemporaryFile(suffix=".xlsx", delete=False) as tmp:
            nome_file = tmp.name
        crea_excel(dati, nome_file)

        await attesa.delete()

        with open(nome_file, "rb") as f:
            await update.message.reply_document(
                document=f,
                filename=f"FitCoachPro_{user_name}.xlsx",
                caption=(
                    f"💪 *Scheda pronta, {user_name}!*\n\n"
                    f"📊 3 fogli nel file:\n"
                    f"• *Scheda Settimanale* — esercizi + link video Jeff Nippard\n"
                    f"• *Spiegazioni* — note tecniche e motivazioni scientifiche\n"
                    f"• *Progressione* — piano 8 settimane\n\n"
                    f"Clicca sui link 🎥 per vedere i video di Jeff Nippard! 🔥"
                ),
                parse_mode="Markdown"
            )

        os.unlink(nome_file)

    except Exception as e:
        await attesa.delete()
        await update.message.reply_text(
            "❌ Si è verificato un errore. Riprova o scrivi /start"
        )
        print(f"Errore: {e}")


def main():
    app = Application.builder().token(TELEGRAM_TOKEN).build()
    app.add_handler(CommandHandler("start", start))
    app.add_handler(CommandHandler("help", help_command))
    app.add_handler(MessageHandler(filters.TEXT & ~filters.COMMAND, crea_scheda))
    print("Bot avviato! ✅")
    app.run_polling()

if __name__ == "__main__":
    main()
