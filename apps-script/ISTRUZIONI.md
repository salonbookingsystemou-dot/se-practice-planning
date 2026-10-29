# SE Practice Partners · installazione (10 minuti)

## 1. Incolla il codice
1. Apri il **foglio delle risposte** del modulo.
2. Menu **Estensioni → Apps Script**.
3. Nel file `Codice.gs` cancella tutto e incolla il contenuto di `Code.gs`.
4. Clicca **+ → HTML**, chiamalo esattamente `Index` e incolla il contenuto di `Index.html`.
5. Salva (icona del dischetto).

## 2. Pubblica
1. In alto a destra **Esegui il deployment → Nuovo deployment**.
2. Tipo: **App web**.
3. *Esegui come*: **Me**. *Chi ha accesso*: **Chiunque**.
4. **Esegui il deployment** → autorizza con il tuo account Google (avviso «app non verificata»: *Avanzate → Vai al progetto*).
5. Copia l'**URL dell'app web** e mandalo nel gruppo WhatsApp.

Alla prima prenotazione lo script crea da solo la scheda **Bookings** nel foglio.

## Come funziona
- **People**: una scheda per persona con la griglia delle disponibilità (giorni × fasce). Filtri per formato, giorno e fascia.
- **I am**: chi sceglie il proprio nome vede in ocra le fasce in comune con ogni persona e le schede ordinate per compatibilità.
- **Propose a session**: data, ora, formato e nota. Impedisce due sessioni della stessa persona a meno di 90 minuti.
- **Sessions**: elenco delle sessioni in programma. L'invitato può **confermare**, entrambi possono **annullare**.

## Foto (facoltativa)
1. Nel Google Form aggiungi una domanda **Caricamento file** con titolo **Photo (optional)**:
   solo immagini, massimo 1 file, 10 MB. Lasciala non obbligatoria.
2. Dopo aver aggiornato il codice, in Apps Script esegui una volta la funzione `getData` (menu in alto → **Esegui**)
   e autorizza l'accesso a Drive. Poi pubblica una **nuova versione** del deployment.

Le foto restano private nel tuo Drive: l'app ne mostra solo una miniatura. Chi non carica una foto vede le iniziali.
Il caricamento file nel modulo richiede che chi risponde abbia un account Google.

## Da sapere
- Non c'è un accesso con password: chiunque abbia il link può prenotare a nome di altri. Tra 45 compagni di corso basta, ma non diffondere il link fuori dal gruppo.
- Il numero di telefono compare solo se la persona l'ha inserito nel modulo (link diretto a WhatsApp).
- Se cambi le domande del modulo, tieni nei titoli le parole chiave: *Name, Region, Session format, Available days, time slots, How often, setup, Language, Phone, Notes*.
- Dopo una modifica al codice: **Esegui il deployment → Gestisci deployment → modifica → Nuova versione**, così l'URL resta lo stesso.
