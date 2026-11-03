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
- L'invitato può **confermare o rifiutare** (pulsante «Conferma o rifiuta» nelle email, «Rifiuta» nell'app): se rifiuta, la sessione è annullata e gli altri ricevono «… ha rifiutato l'invito».
- Partono dal tuo account come «SE Practice Partners»; rispondendo si scrive all'altro studente.
- **Lingua:** ognuno riceve le email nella lingua che usa sulla dashboard (IT/EN, salvata nel foglio «Zones»);
  chi non l'ha mai aperta le riceve nella lingua della pagina di chi propone. Anche l'evento in calendario segue la lingua.
- **Registro:** ogni email è una riga del foglio «Mail log» (quando, tipo, destinatario, oggetto, esito: `sent`,
  `skipped: no email`, `skipped: daily quota reached` o `error: …`). Le copie sono anche in Gmail → Posta inviata.
- Gli indirizzi non arrivano mai alla pagina. Per spegnerle: `SEND_EMAILS: false`. Limite Gmail: circa 100 al giorno.

## Triadi
Nel modulo di proposta si può scegliere un **terzo partecipante** (facoltativo). Ricevono l'email entrambi gli invitati;
ognuno conferma per sé e la sessione diventa «Confermata» quando hanno confermato tutti e due.
Nel foglio «Bookings» le colonne *Third* e *Confirmed by* si aggiungono da sole.

## Sessioni scadute
- Una sessione **non confermata** da tutti, passato l'orario d'inizio, esce dall'app e nel foglio «Bookings» diventa `expired`
  (resta come storico, non si cancella). Dopo l'inizio non si può più confermare.
- Una sessione **confermata** resta visibile fino alla fine dell'incontro (inizio + `SESSION_MINUTES`), poi esce dall'app.
- Non si possono proporre sessioni con data e ora già passate.

## Promemoria 2 ore prima (sessioni confermate)
Circa 2 ore prima (`SOON_HOURS`) di una sessione confermata, tutti i partecipanti ricevono «Tra circa 2 ore: sessione di
pratica con …», con il pulsante **Entra su Google Meet** (o «Apri l'app» se in presenza), l'ora locale per chi è all'estero e la nota.
Per le sessioni prima delle 9:00 il promemoria arriva **la sera prima** (dopo le 20), mai di notte. Una volta sola (colonna *Soon*).

**«Sta per iniziare»:** circa 15 minuti prima (`START_MINUTES`; arriva tra 15 e 30 minuti prima) tutti i partecipanti ricevono
un'email con il pulsante **Entra su Google Meet**. Nel tuo Google Calendar l'evento ha anche pop-up ed email 15 minuti prima;
i partecipanti ricevono le notifiche del loro calendario tramite l'invito di Google.

Il controllo automatico gira **ogni 15 minuti**. Se manca, si installa da solo alla prima apertura dell'app (controllato al massimo ogni 6 ore).

## Reminder 24 ore prima
Se mancano 24 ore (`REMIND_HOURS`) a una sessione non ancora confermata da tutti:
- ogni invitato che non ha confermato riceve «Conferma la sessione di domani alle 19:00 con Dimitri: senza conferma
  verrà annullata», con il pulsante che apre l'app già riconosciuti (rispondendo si scrive a chi ha proposto);
- chi ha proposto riceve l'avviso di chi manca all'appello.
Una volta sola per sessione (colonna *Reminded* in «Bookings»), tra le 8 e le 22, mai nelle prime 3 ore dalla proposta.
Gira nello stesso controllo automatico del promemoria qui sotto (`installTriggers`). Per spegnere: `REMIND_HOURS: 0`.

## Promemoria «hanno appena praticato»
Quando una sessione **confermata** finisce, tutti gli altri iscritti ricevono un'email: «Dimitri e Branham hanno appena
concluso una sessione di pratica. Che ne dici di programmarne una anche tu?», con il pulsante che apre l'app già
riconosciuti e un riquadro per condividere l'app su WhatsApp con chi non è ancora iscritto.
- Al massimo **una al giorno** per persona (più sessioni finite nello stesso giorno finiscono nella stessa email),
  solo tra le 8 e le 22, mai a chi ha appena praticato, solo nomi di battesimo.
- In fondo c'è «Non voglio più ricevere questi promemoria» (salvato nel foglio «Zones», colonna *Reminders*).
- Restano sempre libere almeno `DIGEST_RESERVE` email della quota Gmail giornaliera per proposte e conferme.
- **Attivazione:** in Apps Script esegui una volta `installTriggers`. Per spegnere: `DIGEST: false`.

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
- Le foto caricate dall'app (200 px) stanno nel foglio «Photos» e hanno la precedenza su quelle del modulo.

## «Ora sono libero»
Nella propria card c'è uno switch con una durata (30 min, 1, 2 o 3 ore): finché è attivo la card è evidenziata in arancio,
compare tra i «Liberi ora» e gli altri hanno il pulsante «Proponi adesso» (oggi, al prossimo quarto d'ora).
Si spegne da solo allo scadere. La pagina aggiorna questo stato ogni minuto. I dati stanno nel foglio «Now».

## Chi è connesso
In alto «● 3 connessi» (con l'elenco al clic) e un pallino verde sulla foto di chi ha la pagina aperta ed è entrato.
Lo vedono solo gli utenti che hanno fatto l'accesso. La presenza si aggiorna ogni minuto, sparisce subito alla chiusura
della pagina (o dopo 2,5 minuti) e sta solo nella memoria temporanea del motore: nulla viene scritto nel foglio.

## Google Meet alla conferma
Quando **tutti** gli invitati hanno confermato, lo script crea nel tuo Google Calendar un evento (durata `SESSION_MINUTES`,
90 minuti) con i partecipanti come invitati e, per le sessioni online, un **link Google Meet**. Poi manda a tutti
un'email «Sessione confermata» con il pulsante «Entra su Google Meet». Se la sessione viene annullata, l'evento si cancella.
Google Calendar manda anche il suo invito a ogni partecipante: l'evento con il Meet compare nel loro calendario,
con le loro notifiche abituali. Il link Meet non compare nella pagina.

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
- **Backup automatico:** in Apps Script esegui una volta `installTriggers` (attiva anche il promemoria qui sotto). Ogni notte verso le 3 salva una copia `.xlsx`
  nella cartella Drive «SE Practice · backup» e tiene le ultime 30 (`BACKUP_KEEP`). Backup al volo: esegui `backupNow`.
- Per tornare indietro di poche ore c'è anche **File → Cronologia versioni** del foglio.

## Avvio rapido
La pagina salva nel browser una copia dei dati dell'ultima visita (massimo 7 giorni) e la mostra subito, con
«Aggiorno i dati…» mentre arrivano quelli nuovi. Se il motore non risponde, resta visibile la copia.
«Liberi ora» e «connessi» della copia valgono solo per la stessa persona; «Esci» cancella la copia.

## Feedback post-sessione
Tab **Feedback**: per ogni sessione confermata e conclusa (ultimi `FEEDBACK_DAYS` giorni) c'è una bacheca di commenti.
- Finalità formativa: la **leggono tutti gli studenti che hanno fatto l'accesso**, ci **scrivono solo i partecipanti**
  di quella sessione (`FEEDBACK_PUBLIC: true`; con `false` la vedono solo i partecipanti). Ognuno può eliminare i propri commenti.
- A ogni nuovo commento gli altri partecipanti ricevono un'email con il testo e il pulsante «Leggi e rispondi»,
  che apre l'app direttamente sul tab Feedback. I commenti stanno nel foglio «Feedback».

## Triadi: conferma con due su tre
In una triade la sessione è confermata appena il primo invitato conferma (proponente + un invitato):
si crea subito l'evento Calendar con Meet, con tutti e tre invitati. Il terzo riceve un'email per unirsi o rifiutare.
Se poi conferma, gli altri ricevono un avviso (nessun secondo link Meet). Se rifiuta, esce dalla sessione e dall'evento,
e la sessione resta per gli altri due. Lo stesso vale prima di ogni conferma: se in una triade un invitato rifiuta,
la sessione resta in programma come coppia. Il contatore «sessioni fatte» e il riepilogo «hanno appena praticato» contano solo chi ha confermato.

## Knowledge base
Tab **Knowledge base**, visibile solo agli studenti che hanno fatto l'accesso: appunti, articoli, video, libri, esercizi
condivisi dal gruppo, con ricerca e filtro per tipo. Tutti leggono e aggiungono; ognuno può eliminare solo le proprie risorse.
I contenuti stanno nel foglio «Risorse» (creato da solo alla prima pubblicazione). Link accettati solo `http(s)://`.
Link diretto al tab: `?tab=kb` insieme al link personale.

## Prova senza dati veri
Apri `index.html` dal computer, oppure `se-practice.mnesti.it/?demo`: usa dati di esempio e non scrive nulla.

## Da sapere
- Si agisce solo dopo l'accesso (Google o link personale): nessuno può proporre o confermare a nome di altri.
- La schermata di consenso OAuth nella console Google va **pubblicata** («In produzione»): in modalità Test solo gli utenti di prova possono accedere con Google.
- Il numero di telefono compare solo se la persona l'ha inserito (link diretto a WhatsApp).
