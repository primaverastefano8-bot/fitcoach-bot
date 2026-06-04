import os
import anthropic
from telegram import Update
from telegram.ext import Application, CommandHandler, MessageHandler, filters, ContextTypes

TELEGRAM_TOKEN = os.environ.get("TELEGRAM_TOKEN")
ANTHROPIC_API_KEY = os.environ.get("ANTHROPIC_API_KEY")

client = anthropic.Anthropic(api_key=ANTHROPIC_API_KEY)

SYSTEM_PROMPT = """Sei un coach fitness esperto e personale trainer certificato con 15 anni di esperienza.
Il tuo compito è creare schede di allenamento personalizzate basandoti sulle informazioni dell'utente.

Quando crei una scheda:
1. Analizza attentamente i punti deboli e gli obiettivi dell'utente
2. Crea un programma settimanale dettagliato rispettando ESATTAMENTE il numero di giorni richiesto
3. Per ogni esercizio specifica: serie, ripetizioni, recupero, note tecniche
4. Includi progressione nel tempo (settimane 1-4, 5-8, ecc.)
5. Aggiungi consigli su nutrizione base e recupero
6. Spiega il PERCHE di ogni scelta

Formato la scheda in modo chiaro con:
- Obiettivo principale
- Struttura settimanale
- Esercizi dettagliati per ogni giorno
- Progressione consigliata
- Consigli nutrizionali base
- Note di sicurezza importanti

Basa le tue raccomandazioni sui principi scientifici dell'allenamento."""

async def start(update: Update, context: ContextTypes.DEFAULT_TYPE):
    messaggio = (
        "Benvenuto in FitCoach Pro!\n\n"
        "Sono il tuo personal trainer AI. Posso creare schede di allenamento "
        "completamente personalizzate per te.\n\n"
        "Per creare la tua scheda, dimmi:\n\n"
        "- I tuoi obiettivi (es. massa, forza, dimagrimento)\n"
        "- I tuoi punti deboli (es. tricipiti scarsi, poca forza nel petto)\n"
        "- Quanti giorni a settimana ti alleni\n"
        "- La tua esperienza (principiante/intermedio/avanzato)\n"
        "- Eventuali infortuni o limitazioni\n\n"
        "Scrivi tutto in un messaggio e creo la tua scheda personalizzata!"
    )
    await update.message.reply_text(messaggio)

async def help_command(update: Update, context: ContextTypes.DEFAULT_TYPE):
    await update.message.reply_text(
        "Come usarmi:\n\n"
        "Semplicemente descrivimi la tua situazione in un messaggio!\n\n"
        "Esempio:\n"
        "Ho i tricipiti molto carenti e voglio migliorarli. Mi alleno 4 giorni a settimana, "
        "livello intermedio, nessun infortunio. Voglio anche aumentare la forza dei pettorali.\n\n"
        "E io creo la tua scheda!"
    )

async def crea_scheda(update: Update, context: ContextTypes.DEFAULT_TYPE):
    user_message = update.message.text
    user_name = update.message.from_user.first_name

    attesa = await update.message.reply_text(
        "Sto analizzando la tua situazione e creando la scheda personalizzata...\n"
        "Ci vogliono circa 30 secondi!"
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
                    "Crea una scheda di allenamento personalizzata per " + user_name + ". "
                    "Ecco le sue informazioni: " + user_message + "\n\n"
                    "Prima cerca su internet esempi di schede per questi obiettivi specifici, "
                    "poi crea un programma personalizzato e ottimizzato. "
                    "Rispetta ESATTAMENTE il numero di giorni di allenamento richiesto dall utente."
                )
            }]
        )

        scheda_testo = ""
        for block in response.content:
            if hasattr(block, "text"):
                scheda_testo += block.text

        await attesa.delete()

        if len(scheda_testo) > 4000:
            parti = [scheda_testo[i:i+4000] for i in range(0, len(scheda_testo), 4000)]
            for i, parte in enumerate(parti):
                if i == 0:
                    await update.message.reply_text("LA TUA SCHEDA PERSONALIZZATA\n\n" + parte)
                else:
                    await update.message.reply_text(parte)
        else:
            await update.message.reply_text("LA TUA SCHEDA PERSONALIZZATA\n\n" + scheda_testo)

        await update.message.reply_text(
            "Scheda creata! Hai domande o vuoi modifiche? Scrivimi!\n\n"
            "Per una nuova scheda scrivi /start"
        )

    except Exception as e:
        await attesa.delete()
        await update.message.reply_text(
            "Si e verificato un errore. Riprova tra qualche secondo o scrivi /start"
        )
        print("Errore: " + str(e))

def main():
    app = Application.builder().token(TELEGRAM_TOKEN).build()
    app.add_handler(CommandHandler("start", start))
    app.add_handler(CommandHandler("help", help_command))
    app.add_handler(MessageHandler(filters.TEXT & ~filters.COMMAND, crea_scheda))
    print("Bot avviato!")
    app.run_polling()

if __name__ == "__main__":
    main()
