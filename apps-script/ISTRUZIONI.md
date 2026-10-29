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
- Per «Modifica i tuoi dati» nelle **Impostazioni → Risposte** del modulo attiva:
  **Consenti la modifica delle risposte** e **Limita a 1 risposta** (richiede l'accesso con Google).
  Chi riapre il modulo vede così la propria risposta e la corregge, invece di crearne una nuova.
- Foto facoltativa: domanda **Caricamento file** intitolata «Photo (optional)», solo immagini, 1 file.

## Notifiche email
- proposta → all'invitato · conferma → a chi ha proposto · annullamento → all'altra persona.
- Partono dal tuo account come «SE Practice Partners»; rispondendo si scrive all'altro studente.
- Gli indirizzi non arrivano mai alla pagina. Per spegnerle: `SEND_EMAILS: false`. Limite Gmail: circa 100 al giorno.

## Triadi
Nel modulo di proposta si può scegliere un **terzo partecipante** (facoltativo). Ricevono l'email entrambi gli invitati;
ognuno conferma per sé e la sessione diventa «Confermata» quando hanno confermato tutti e due.
Nel foglio «Bookings» le colonne *Third* e *Confirmed by* si aggiungono da sole.

## Link personale e foto
- Ogni email dell'app contiene il **link personale** del destinatario (`…/?me=<codice>`): aprendolo, l'app lo riconosce
  e preimposta «Io sono». Il codice è segreto, diverso per ognuno, e non compare mai nella pagina.
- Chi non ha la foto vede una finestra per **scattarla o sceglierla dalla galleria**. Per caricarla serve il link personale:
  chi non ce l'ha lo chiede con «Inviami il link» e lo riceve via email (al massimo uno ogni 10 minuti).
- Le foto caricate dall'app (piccole, 320 px) stanno nel foglio «Photos» e hanno la precedenza su quelle del modulo.

## «Ora sono libero»
Nella propria card c'è uno switch con una durata (30 min, 1, 2 o 3 ore): finché è attivo la card è evidenziata in arancio,
compare tra i «Liberi ora» e gli altri hanno il pulsante «Proponi adesso» (oggi, al prossimo quarto d'ora).
Si spegne da solo allo scadere. La pagina aggiorna questo stato ogni minuto. I dati stanno nel foglio «Now».

## Prova senza dati veri
Apri `index.html` dal computer, oppure `se-practice.mnesti.it/?demo`: usa dati di esempio e non scrive nulla.

## Da sapere
- Non c'è accesso con password: chi ha il link può prenotare a nome di altri. Non diffondere il link fuori dal gruppo.
- Il numero di telefono compare solo se la persona l'ha inserito (link diretto a WhatsApp).
