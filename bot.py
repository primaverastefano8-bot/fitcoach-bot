import os
import anthropic
from telegram import Update, LabeledPrice
from telegram.ext import Application, CommandHandler, MessageHandler, PreCheckoutQueryHandler, filters, ContextTypes

TELEGRAM_TOKEN = os.environ.get("TELEGRAM_TOKEN")
ANTHROPIC_API_KEY = os.environ.get("ANTHROPIC_API_KEY")

client = anthropic.Anthropic(api_key=ANTHROPIC_API_KEY)

PREZZO_STARS = 250

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

---
## OBIETTIVO PRINCIPALE
[Descrizione dettagliata con spiegazione scientifica]

REGOLE IMPORTANTI:
1. Rispetta ESATTAMENTE il numero di giorni richiesto dall utente
2. Rispetta tutti i vincoli (infortuni, esercizi da evitare, attrezzatura)
3. Spiega sempre il PERCHE scientifico di ogni scelta
4. Le tabelle devono essere sempre complete con tutte le colonne
5. Cerca sempre le migliori schede scientifiche online come riferimento"""


async def start(update: Update, context: ContextTypes.DEFAULT_TYPE):
    await update.message.reply_text(
        "Benvenuto in FitCoach Pro!\n\n"
        "Sono il tuo personal trainer AI.\n\n"
        "Ricevi una scheda di allenamento completamente personalizzata "
        "al prezzo di soli 5 euro!\n\n"
        "Per creare la tua scheda dimmi:\n\n"
        "- I tuoi obiettivi (massa, forza, dimagrimento)\n"
        "- I tuoi punti deboli (es. tricipiti scarsi)\n"
        "- Quanti giorni a settimana ti alleni\n"
        "- La tua esperienza (principiante/intermedio/avanzato)\n"
        "- Eventuali infortuni o limitazioni\n"
        "- Attrezzatura disponibile\n\n"
        "Scrivi tutto in un messaggio e ti mostro come acquistare!"
    )


async def help_command(update: Update, context: ContextTypes.DEFAULT_TYPE):
    await update.message.reply_text(
        "Come funziona FitCoach Pro:\n\n"
        "1. Descrivi la tua situazione in un messaggio\n"
        "2. Paga 250 Stars (circa 5 euro)\n"
        "3. Ricevi la tua scheda personalizzata!\n\n"
        "Esempio:\n"
        "Ho i tricipiti molto carenti. Mi alleno 4 giorni a settimana, "
        "livello intermedio, nessun infortunio.\n\n"
        "Scrivi /start per iniziare!"
    )


async def richiesta_pagamento(update: Update, context: ContextTypes.DEFAULT_TYPE):
    user_message = update.message.text
    context.user_data["richiesta"] = user_message

    await update.message.reply_text(
        "Ottimo! Ho preso nota della tua richiesta.\n\n"
        "Per ricevere la tua scheda personalizzata procedi con il pagamento di "
        "250 Stars (circa 5 euro).\n\n"
        "Clicca su /paga per procedere!"
    )


async def paga(update: Update, context: ContextTypes.DEFAULT_TYPE):
    if "richiesta" not in context.user_data:
        await update.message.reply_text(
            "Prima dimmi le tue informazioni!\n\n"
            "Scrivi /start e descrivi la tua situazione."
        )
        return

    await update.message.reply_invoice(
        title="Scheda di Allenamento Personalizzata",
        description="Scheda scientifica personalizzata creata da FitCoach Pro AI",
        payload="scheda_allenamento",
        currency="XTR",
        prices=[LabeledPrice("Scheda Personalizzata", PREZZO_STARS)],
    )


async def pre_checkout(update: Update, context: ContextTypes.DEFAULT_TYPE):
    await update.pre_checkout_query.answer(ok=True)


async def pagamento_completato(update: Update, context: ContextTypes.DEFAULT_TYPE):
    user_name = update.message.from_user.first_name
    user_message = context.user_data.get("richiesta", "scheda generica")

    attesa = await update.message.reply_text(
        "Pagamento ricevuto! Grazie!\n\n"
        "Sto creando la tua scheda personalizzata...\n"
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

        context.user_data.clear()

    except Exception as e:
        await attesa.delete()
        await update.message.reply_text(
            "Si e verificato un errore nella creazione della scheda.\n"
            "Ti rimborseremo le Stars. Contatta il supporto scrivendo /supporto"
        )
        print("Errore: " + str(e))


async def supporto(update: Update, context: ContextTypes.DEFAULT_TYPE):
    await update.message.reply_text(
        "Per assistenza o rimborsi contattaci direttamente.\n\n"
        "Descrivi il problema e ti risponderemo al piu presto!"
    )


def main():
    app = Application.builder().token(TELEGRAM_TOKEN).build()
    app.add_handler(CommandHandler("start", start))
    app.add_handler(CommandHandler("help", help_command))
    app.add_handler(CommandHandler("paga", paga))
    app.add_handler(CommandHandler("supporto", supporto))
    app.add_handler(PreCheckoutQueryHandler(pre_checkout))
    app.add_handler(MessageHandler(filters.SUCCESSFUL_PAYMENT, pagamento_completato))
    app.add_handler(MessageHandler(filters.TEXT & ~filters.COMMAND, richiesta_pagamento))
    print("Bot avviato!")
    app.run_polling()


if __name__ == "__main__":
    main()
