import os
import anthropic
from telegram import Update
from telegram.ext import Application, CommandHandler, MessageHandler, filters, ContextTypes

TELEGRAM_TOKEN = os.environ.get("TELEGRAM_TOKEN")
ANTHROPIC_API_KEY = os.environ.get("ANTHROPIC_API_KEY")

client = anthropic.Anthropic(api_key=ANTHROPIC_API_KEY)

SYSTEM_PROMPT = """Sei un coach fitness esperto e personal trainer certificato con 15 anni di esperienza.
Crea schede di allenamento personalizzate SEMPRE in questo formato esatto:

---
# SCHEDA DI ALLENAMENTO PERSONALIZZATA — [NOME UTENTE]
### Livello: [livello] | Focus: [obiettivo principale] | Giorni: [N]/settimana

---
### [EMOJI COLORE] GIORNO A — [GIORNO SETTIMANA] | [Gruppi muscolari]

> **Perche?** [Spiegazione scientifica del perche questi muscoli e questo approccio]

| # | Esercizio | Serie | Rip. | Recupero | Note Tecniche |
|---|-----------|-------|------|----------|---------------|
| 1 | **[Nome Esercizio]** | [N] | [N-N] | [N'N"] | [note tecniche dettagliate] |
| 2 | **[Nome Esercizio]** | [N] | [N-N] | [N'N"] | [note tecniche dettagliate] |
...

> Note aggiuntive scientifiche se necessario

---
[RIPETI PER OGNI GIORNO DI ALLENAMENTO]

---
## PROGRESSIONE NEL TEMPO

### SETTIMANE 1-4 | Fase di Accumulo
- Obiettivo e dettagli
- RIR consigliato
- Incrementi settimanali

### SETTIMANE 5-8 | Fase di Intensificazione
- Obiettivo e dettagli
- RIR consigliato
- Tecniche avanzate

### SETTIMANE 9-10 | Deload
- Come ridurre il volume
- Come recuperare

---
## NOTE DI SICUREZZA IMPORTANTI
1. [nota 1]
2. [nota 2]
3. [nota 3]
4. [nota 4]
5. [nota 5]

---
## OBIETTIVO PRINCIPALE
[Descrizione dettagliata con spiegazione scientifica]

---

REGOLE IMPORTANTI:
1. Rispetta ESATTAMENTE il numero di giorni richiesto dall utente
2. Rispetta tutti i vincoli (infortuni, esercizi da evitare, attrezzatura)
3. Spiega sempre il PERCHE scientifico di ogni scelta
4. Usa le emoji per rendere visivamente chiara la struttura
5. Le tabelle devono essere sempre complete con tutte le colonne
6. Cerca sempre le migliori schede scientifiche online come riferimento"""

async def start(update: Update, context: ContextTypes.DEFAULT_TYPE):
    await update.message.reply_text(
        "Benvenuto in FitCoach Pro!\n\n"
        "Sono il tuo personal trainer AI.\n\n"
        "Per creare la tua scheda dimmi:\n\n"
        "- I tuoi obiettivi (massa, forza, dimagrimento)\n"
        "- I tuoi punti deboli (es. tricipiti scarsi)\n"
        "- Quanti giorni a settimana ti alleni\n"
        "- La tua esperienza (principiante/intermedio/avanzato)\n"
        "- Eventuali infortuni o limitazioni\n"
        "- Attrezzatura disponibile\n\n"
        "Scrivi tutto in un messaggio e creo la tua scheda!"
    )

async def help_command(update: Update, context: ContextTypes.DEFAULT_TYPE):
    await update.message.reply_text(
        "Come usarmi:\n\n"
        "Descrivimi la tua situazione in un messaggio!\n\n"
        "Esempio:\n"
        "Ho i tricipiti molto carenti. Mi alleno 4 giorni a settimana, "
        "livello intermedio, nessun infortunio. Voglio aumentare la forza del petto.\n\n"
        "Creo la tua scheda personalizzata!"
    )

async def crea_scheda(update: Update, context: ContextTypes.DEFAULT_TYPE):
    user_message = update.message.text
    user_name = update.message.from_user.first_name

    attesa = await update.message.reply_text(
        "Sto analizzando la tua situazione e cercando le migliori schede...\n"
        "Circa 30-40 secondi!"
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
                    "Crea una scheda di allenamento personalizzata per " + user_name + ".\n"
                    "Informazioni: " + user_message + "\n\n"
                    "IMPORTANTE:\n"
                    "1. Rispetta ESATTAMENTE il numero di giorni richiesto\n"
                    "2. Usa il formato con tabelle e sezioni come da istruzioni\n"
                    "3. Cerca online schede scientifiche come riferimento\n"
                    "4. Spiega il perche scientifico di ogni scelta"
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
                    await update.message.reply_text(
                        "LA TUA SCHEDA PERSONALIZZATA\n\n" + parte,
                        parse_mode="Markdown"
                    )
                else:
                    await update.message.reply_text(parte, parse_mode="Markdown")
        else:
            await update.message.reply_text(
                "LA TUA SCHEDA PERSONALIZZATA\n\n" + scheda_testo,
                parse_mode="Markdown"
            )

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
