import os
import anthropic
from telegram import Update
from telegram.ext import Application, CommandHandler, MessageHandler, filters, ContextTypes

# --- CONFIGURAZIONE ---
TELEGRAM_TOKEN = os.environ.get("TELEGRAM_TOKEN")
ANTHROPIC_API_KEY = os.environ.get("ANTHROPIC_API_KEY")

client = anthropic.Anthropic(api_key=ANTHROPIC_API_KEY)

# --- SISTEMA PROMPT PER L'AI ---
SYSTEM_PROMPT = """Sei un coach fitness esperto e personale trainer certificato con 15 anni di esperienza.
Il tuo compito è creare schede di allenamento personalizzate basandoti sulle informazioni dell'utente.

Quando crei una scheda:
1. Analizza attentamente i punti deboli e gli obiettivi dell'utente
2. Crea un programma settimanale dettagliato (3-5 giorni)
3. Per ogni esercizio specifica: serie, ripetizioni, recupero, note tecniche
4. Includi progressione nel tempo (settimane 1-4, 5-8, ecc.)
5. Aggiungi consigli su nutrizione base e recupero
6. Spiega il PERCHÉ di ogni scelta (es. "ho inserito le dip per i tricipiti perché...")

Formato la scheda in modo chiaro con:
- 💪 Obiettivo principale
- 📅 Struttura settimanale  
- 🏋️ Esercizi dettagliati per ogni giorno
- 📈 Progressione consigliata
- 🥗 Consigli nutrizionali base
- ⚠️ Note di sicurezza importanti

Basa le tue raccomandazioni sui principi scientifici dell'allenamento (ipertrofia, forza, recupero muscolare)."""

# --- COMANDO /START ---
async def start(update: Update, context: ContextTypes.DEFAULT_TYPE):
    messaggio = """👋 *Benvenuto in FitCoach Pro!*

Sono il tuo personal trainer AI. Posso creare schede di allenamento *completamente personalizzate* per te.

Per creare la tua scheda, dimmi:

✅ I tuoi *obiettivi* (es. massa, forza, dimagrimento)
✅ I tuoi *punti deboli* (es. "tricipiti scarsi", "poca forza nel petto")
✅ Quanti *giorni a settimana* ti alleni
✅ La tua *esperienza* (principiante/intermedio/avanzato)
✅ Eventuali *infortuni o limitazioni*

Scrivi tutto in un messaggio e creo la tua scheda personalizzata! 🔥"""
    
    await update.message.reply_text(messaggio, parse_mode='Markdown')

# --- COMANDO /HELP ---
async def help_command(update: Update, context: ContextTypes.DEFAULT_TYPE):
    await update.message.reply_text(
        "💡 *Come usarmi:*\n\n"
        "Semplicemente *descrivimi la tua situazione* in un messaggio!\n\n"
        "Esempio:\n"
        "_'Ho i tricipiti molto carenti e voglio migliorarli. Mi alleno 4 giorni a settimana, "
        "livello intermedio, nessun infortunio. Voglio anche aumentare la forza dei pettorali.'_\n\n"
        "E io creo la tua scheda! 💪",
        parse_mode='Markdown'
    )

# --- GESTIONE MESSAGGI PRINCIPALI ---
async def crea_scheda(update: Update, context: ContextTypes.DEFAULT_TYPE):
    user_message = update.message.text
    user_name = update.message.from_user.first_name
    
    # Messaggio di attesa
    attesa = await update.message.reply_text(
        "⏳ Sto analizzando la tua situazione e creando la scheda personalizzata...\n"
        "Ci vogliono circa 30 secondi 🔄"
    )
    
    try:
        # Chiamata all'AI con web search
        response = client.messages.create(
            model="claude-sonnet-4-20250514",
            max_tokens=4000,
            system=SYSTEM_PROMPT,
            tools=[{"type": "web_search_20250305", "name": "web_search"}],
            messages=[
                {
                    "role": "user", 
                    "content": f"Crea una scheda di allenamento personalizzata per {user_name}. "
                               f"Ecco le sue informazioni: {user_message}\n\n"
                               f"Prima cerca su internet esempi di schede per questi obiettivi specifici, "
                               f"poi crea un programma personalizzato e ottimizzato."
                }
            ]
        )
        
        # Estrai il testo dalla risposta
        scheda_testo = ""
        for block in response.content:
            if block.type == "text":
                scheda_testo += block.text
        
        # Elimina messaggio di attesa
        await attesa.delete()
        
        # Invia la scheda (Telegram ha limite 4096 caratteri per messaggio)
        if len(scheda_testo) > 4000:
            # Dividi in più messaggi
            parti = [scheda_testo[i:i+4000] for i in range(0, len(scheda_testo), 4000)]
            for i, parte in enumerate(parti):
                if i == 0:
                    await update.message.reply_text(f"🏋️ *LA TUA SCHEDA PERSONALIZZATA*\n\n{parte}", parse_mode='Markdown')
                else:
                    await update.message.reply_text(parte, parse_mode='Markdown')
        else:
            await update.message.reply_text(
                f"🏋️ *LA TUA SCHEDA PERSONALIZZATA*\n\n{scheda_testo}", 
                parse_mode='Markdown'
            )
            
        # Messaggio finale
        await update.message.reply_text(
            "✅ Scheda creata! Hai domande o vuoi modifiche? Scrivimi!\n\n"
            "🔄 Per una nuova scheda scrivi /start"
        )
        
    except Exception as e:
        await attesa.delete()
        await update.message.reply_text(
            "❌ Si è verificato un errore. Riprova tra qualche secondo o scrivi /start"
        )
        print(f"Errore: {e}")

# --- AVVIO BOT ---
def main():
    app = Application.builder().token(TELEGRAM_TOKEN).build()
    
    app.add_handler(CommandHandler("start", start))
    app.add_handler(CommandHandler("help", help_command))
    app.add_handler(MessageHandler(filters.TEXT & ~filters.COMMAND, crea_scheda))
    
    print("Bot avviato! ✅")
    app.run_polling()

if __name__ == "__main__":
    main()