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

## Aggiornamento automatico da GitHub
Ogni modifica a `apps-script/` su GitHub viene pubblicata da sola sullo **stesso** deployment (azione *Deploy Apps Script*).
Configurazione, una volta sola:
1. Attiva l'API: https://script.google.com/home/usersettings → **API di Google Apps Script: On**.
2. In Apps Script → **Impostazioni progetto** (ingranaggio) copia lo **ID script** e mettilo in `.clasp.json` al posto di `INCOLLA_QUI_LO_SCRIPT_ID`.
3. Sul tuo computer, nel Terminale: `npx @google/clasp@2.4.2 login` → si apre il browser, accedi con l'account dello script e consenti.
4. Apri il file creato: `cat ~/.clasprc.json`, copia tutto il contenuto.
5. Su GitHub → **Settings → Secrets and variables → Actions → New repository secret**: nome `CLASPRC_JSON`, valore = il contenuto copiato.
6. Su GitHub → **Actions → Deploy Apps Script → Run workflow** per la prima pubblicazione.

Da lì in poi non si copia più nulla a mano. Il file `appsscript.json` porta con sé fuso orario, servizio Calendar e
impostazioni del web app (*Esegui come: Me*, *Accesso: Chiunque*). Le autorizzazioni nuove (per esempio un servizio aggiunto)
vanno comunque confermate una volta nell'editor eseguendo `testEmail`.

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

## Accesso: Google o link personale
Non c'è più il menu «Io sono»: chi non ha fatto l'accesso può solo guardare.
- **Accedi con Google**: un clic con l'account usato nel modulo. Il motore verifica con Google la firma dell'accesso
  e confronta l'email con quella del modulo. Dalle visite successive l'accesso è automatico. ID client in `GOOGLE_CLIENT_ID`
  (`Code.gs`) e in `index.html`; l'origine autorizzata nella console Google è `https://se-practice.mnesti.it`.
- **Link personale via email** (per chi non usa Google): si inserisce l'email del modulo e arriva `…/?me=<codice>`.
  La risposta è la stessa anche per email sconosciute; al massimo un link ogni 10 minuti per persona.
- Ogni azione (proporre, confermare, annullare, «Ora sono libero», foto) richiede il codice personale:
  il nome lo ricava il motore, quindi nessuno può agire a nome di altri. «Esci» in alto toglie l'accesso dal dispositivo.
- Chi non ha fatto l'accesso vede persone e disponibilità abituali, ma **non** chi è libero adesso e non può proporre:
  il motore non invia proprio quell'informazione senza un codice personale valido.
- Le foto caricate dall'app (320 px) stanno nel foglio «Photos» e hanno la precedenza su quelle del modulo.

## «Ora sono libero»
Nella propria card c'è uno switch con una durata (30 min, 1, 2 o 3 ore): finché è attivo la card è evidenziata in arancio,
compare tra i «Liberi ora» e gli altri hanno il pulsante «Proponi adesso» (oggi, al prossimo quarto d'ora).
Si spegne da solo allo scadere. La pagina aggiorna questo stato ogni minuto. I dati stanno nel foglio «Now».

## Google Meet alla conferma
Quando **tutti** gli invitati hanno confermato, lo script crea nel tuo Google Calendar un evento (durata `SESSION_MINUTES`,
90 minuti) con i partecipanti come invitati e, per le sessioni online, un **link Google Meet**. Poi manda a tutti
un'email «Sessione confermata» con il pulsante «Entra su Google Meet». Se la sessione viene annullata, l'evento si cancella.
Il link Meet arriva solo via email: non compare nella pagina.

**Attivazione (una volta):** in Apps Script, a sinistra, **Servizi → +** → **Google Calendar API** → **Aggiungi**
(identificatore `Calendar`). Poi esegui una funzione qualsiasi (per esempio `testEmail`), autorizza il calendario
e pubblica una **Nuova versione**. Senza il servizio, la conferma funziona lo stesso ma senza link Meet.
Per spegnere: `CREATE_MEET: false`.

## Fusi orari (studenti all'estero)
- Le sessioni sono salvate **in ora italiana**; ognuno le vede nel proprio fuso, letto dal browser, con «(19:00 in Italia)» accanto se diverso.
- Chi propone dall'estero inserisce data e ora **locali**: l'app le converte. Il cambio dell'ora legale è gestito per ogni data.
- Il fuso di ciascuno viene salvato al primo accesso nel foglio «Zones»; le card di chi è all'estero mostrano città e differenza
  e ricordano che le sue fasce abituali sono in ora locale.
- Email: ora italiana più l'ora locale per chi è all'estero. L'evento Calendar/Meet si adatta da solo al fuso di ognuno.

## Sicurezza del foglio e backup
- Il foglio **non va condiviso** con gli studenti: l'app passa dal motore, che lavora con il tuo account.
  **Condividi → Accesso generale → Con restrizioni** e nessun editor oltre a te.
- **Backup automatico:** in Apps Script esegui una volta `installBackup`. Ogni notte verso le 3 salva una copia `.xlsx`
  nella cartella Drive «SE Practice · backup» e tiene le ultime 30 (`BACKUP_KEEP`). Backup al volo: esegui `backupNow`.
- Per tornare indietro di poche ore c'è anche **File → Cronologia versioni** del foglio.

## Prova senza dati veri
Apri `index.html` dal computer, oppure `se-practice.mnesti.it/?demo`: usa dati di esempio e non scrive nulla.

## Da sapere
- Si agisce solo dopo l'accesso (Google o link personale): nessuno può proporre o confermare a nome di altri.
- La schermata di consenso OAuth nella console Google va **pubblicata** («In produzione»): in modalità Test solo gli utenti di prova possono accedere con Google.
- Il numero di telefono compare solo se la persona l'ha inserito (link diretto a WhatsApp).
