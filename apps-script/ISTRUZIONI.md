# SE Practice Partners · come funziona e come si aggiorna

L'app ha due parti:

| Parte | Dove | Cosa fa | Come si aggiorna |
|---|---|---|---|
| **Pagina** | `index.html` in questo repository → `se-practice.mnesti.it` (GitHub Pages) | Interfaccia: persone, filtri, prenotazioni, lingue | Modifica su GitHub: è online in 1–2 minuti, senza altro |
| **Motore** | `apps-script/Code.gs` → Apps Script legato al foglio delle risposte | Legge il foglio, salva le prenotazioni, manda le email | Incolla in Apps Script → **Gestisci deployment → matita → Nuova versione** |

## Installare il motore (una volta)
1. Apri il **foglio delle risposte** → **Estensioni → Apps Script**.
2. Incolla `Code.gs` (copialo dalla versione *Raw* su GitHub). Nel progetto basta questo file:
   se c'è ancora un file `Index`, puoi eliminarlo.
3. In `CONFIG` controlla `FORM_URL` (link del modulo) e `APP_URL` (`https://se-practice.mnesti.it`).
4. Esegui una volta `testEmail` e autorizza (Fogli, Drive, Gmail).
5. **Esegui il deployment** con queste impostazioni:
   - *Esegui come*: **Me**
   - *Chi ha accesso*: **Chiunque**

L'URL che termina con `/exec` è già scritto in `index.html` (costante `API_URL`). Se un giorno cambia, va aggiornato lì.
Aggiornando sempre lo **stesso** deployment (Nuova versione), l'URL non cambia.

## Modulo Google
- Titoli delle domande con queste parole chiave: *Name, Region, Session format, Available days, time slots, How often, setup, Language, Phone, Notes, Photo*.
- Email: attiva **Raccogli indirizzi email** nelle impostazioni del modulo (oppure una domanda con «email» nel titolo).
- Foto facoltativa: domanda **Caricamento file** intitolata «Photo (optional)», solo immagini, 1 file.

## Notifiche email
- proposta → all'invitato · conferma → a chi ha proposto · annullamento → all'altra persona.
- Partono dal tuo account come «SE Practice Partners»; rispondendo si scrive all'altro studente.
- Gli indirizzi non arrivano mai alla pagina. Per spegnerle: `SEND_EMAILS: false`. Limite Gmail: circa 100 al giorno.

## Prova senza dati veri
Apri `index.html` dal computer, oppure `se-practice.mnesti.it/?demo`: usa dati di esempio e non scrive nulla.

## Da sapere
- Non c'è accesso con password: chi ha il link può prenotare a nome di altri. Non diffondere il link fuori dal gruppo.
- Il numero di telefono compare solo se la persona l'ha inserito (link diretto a WhatsApp).
