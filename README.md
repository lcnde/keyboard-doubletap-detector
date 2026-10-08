# Double Tap Detector

Controlla se la tastiera registra due volte lo stesso tasto (double tapping / chattering).

## Soglie
Sono tarate sui tempi misurati dall'utente: al massimo della velocità, 70 ms tra due pressioni con un dito e 43 ms con due dita. Nelle prove a due dita nessuna ripressione è avvenuta entro 20 ms dal rilascio.

| Livello | Quando | Cosa succede |
|---|---|---|
| **Doppio tap** | stesso tasto ripremuto entro **35 ms** dalla pressione precedente, oppure entro **10 ms** dal rilascio (rimbalzo al rilascio) | notifica + log |
| **Sospetto** | ripremuto tra 35 e **60 ms** dopo la pressione precedente: più veloce del minimo con un dito, quindi impossibile scrivendo normalmente | solo log e statistiche |

Si possono cambiare all'avvio con `-SogliaMs`, `-SogliaRilascioMs` e `-SogliaSospettoMs`.

## Uso
- **Avvio:** doppio clic su `Avvia.cmd`. La finestra parte ridotta a icona e resta nella barra delle applicazioni. Se è già in esecuzione, non parte una seconda copia.
- **Dati in tempo reale:** apri la finestra dalla barra delle applicazioni oppure clicca l'icona rossa "2x" nell'area di notifica.
- **Notifiche:** arriva una notifica di Windows quando viene rilevato un doppio tap (al massimo una ogni 30 s; quelli nel frattempo vengono sommati).
- **Tasti nella finestra:** `Q` esci, `H` riduci a icona, `L` apri la cartella dei log, `R` azzera le statistiche.
- **Chiusura:** `Q`, menu dell'icona → Esci, oppure la X della finestra.
- **Avvio automatico a ogni accesso (facoltativo):** `powershell -ExecutionPolicy Bypass -File avvio-automatico.ps1` (per toglierlo aggiungi `-Rimuovi`).
- **Verifica di funzionamento:** `powershell -ExecutionPolicy Bypass -File double-tap-detector.ps1 -Autotest`.

## Log (cartella `log\`)
- `doppi-tap.csv`: una riga per ogni doppio tap o pressione sospetta (data, ora, tasto, intervalli in ms, tipo, tastiera). Si apre con Excel.
- `sessioni.log`: avvio e arresto, con un riepilogo (durata, pressioni, doppi tap per tasto).

I tasti premuti normalmente **non vengono salvati da nessuna parte**: nel file finiscono solo i doppi tap e le pressioni sospette.

## Perché è sicuro
- Legge la tastiera in sola lettura con la Raw Input API di Windows: non blocca, non modifica e non ritarda nessun tasto.
- Non serve essere amministratore, non tocca il registro e non installa niente.
- Ignora gli input simulati dai programmi (password manager, macro, tastiera su schermo), per evitare falsi allarmi.
- Se il PC era rallentato e la misura non è affidabile, scarta l'evento invece di segnalarlo.
- Quando la finestra è ridotta a icona non ridisegna lo schermo, quindi il consumo di CPU è praticamente nullo.
