# Keyboard Double-Tap Detector

[English](README.md) | **Italiano**

Controlla se la tastiera registra due volte lo stesso tasto (**double tapping** o *chattering*).
Lavora in background su Windows, ti avvisa appena succede e salva solo le pressioni difettose.

![Double Tap Detector in modalità demo](docs/screenshot-demo.png)
<sub>Modalità demo con dati simulati (`-Demo`).</sub>

## Perché

La mia tastiera ogni tanto scriveva una lettera due volte. Volevo capire se era un difetto hardware o il mio modo di scrivere, con dei numeri invece che a sensazione.

## Soglie, ricavate da misure

Prima ho misurato i miei limiti alla massima velocità: **70 ms** tra due pressioni dello stesso tasto con un dito, **43 ms** alternando due dita, e mai una ripressione entro 20 ms dal rilascio. Le soglie stanno sotto questi limiti:

| Livello | Quando | Cosa succede |
|---|---|---|
| **Doppio tap** | stesso tasto ripremuto entro **35 ms** dalla pressione precedente, oppure entro **10 ms** dal rilascio (rimbalzo al rilascio) | notifica + log |
| **Sospetto** | ripremuto tra **35 e 60 ms** dopo la pressione precedente: più veloce di quanto si possa fare con un dito | solo log e statistiche |

Si possono cambiare all'avvio con `-SogliaMs`, `-SogliaRilascioMs` e `-SogliaSospettoMs`.

## Come funziona

- Legge la tastiera **in sola lettura** con la **Raw Input API** di Windows (`RegisterRawInputDevices` con `RIDEV_INPUTSINK`). Non è un hook: non può bloccare, modificare o ritardare nessun tasto.
- Gli intervalli si misurano con il timer ad alta precisione, letto appena arriva il messaggio `WM_INPUT`.
- Ogni tastiera fisica è seguita separatamente (dispositivo + codice di scansione), quindi due tastiere non si confondono.
- **Pensato per non dare falsi allarmi:**
  - la ripetizione automatica (tasto tenuto premuto) viene riconosciuta e ignorata;
  - un rilascio perso (per esempio dopo la sospensione del PC) non crea un falso rilevamento;
  - gli input simulati dai programmi (password manager, macro, tastiera su schermo) vengono ignorati;
  - se il PC era rallentato e la misura non è affidabile (l'orologio dei messaggi di Windows non concorda con il timer preciso), l'evento viene scartato invece di essere segnalato.

### Privacy

I tasti premuti normalmente **non vengono salvati da nessuna parte**. Nel log finiscono solo doppi tap e pressioni sospette: nome del tasto, intervalli in millisecondi e identificativo della tastiera.

## Requisiti

Windows 10 o 11 con Windows PowerShell 5.1 (già incluso). Niente da installare, non servono permessi di amministratore e non tocca il registro.

## Avvio rapido

1. Scarica o clona la repository.
2. Doppio clic su **`Avvia.cmd`**. La finestra parte ridotta a icona nella barra delle applicazioni, con l'icona rossa **"2x"** nell'area di notifica.
3. Clicca l'icona (o la finestra nella barra delle applicazioni) per vedere i dati in tempo reale.

Vuoi prima vederlo in azione? Demo con dati simulati, non scrive niente su disco:

```
powershell -ExecutionPolicy Bypass -File double-tap-detector.ps1 -Demo
```

> **Avviso di Windows:** SmartScreen o l'antivirus possono segnalare uno script PowerShell scaricato da internet. Il programma è un unico file leggibile, `double-tap-detector.ps1`. Per sbloccarlo: tasto destro → Proprietà → Annulla blocco, oppure `Unblock-File .\double-tap-detector.ps1`.

Tasti nella finestra: `Q` esci, `H` riduci a icona, `L` apri la cartella dei log, `R` azzera le statistiche.

Avvio automatico a ogni accesso (facoltativo): `powershell -ExecutionPolicy Bypass -File avvio-automatico.ps1` (per toglierlo aggiungi `-Rimuovi`). Crea solo un collegamento nella cartella "Esecuzione automatica".

## Log (cartella `log\`)

- `doppi-tap.csv`: una riga per ogni doppio tap o pressione sospetta (data, ora, tasto, intervalli in ms, tipo, tastiera). Si apre direttamente con Excel.
- `sessioni.log`: avvio e arresto di ogni sessione, con un riepilogo.

## Test

```
powershell -ExecutionPolicy Bypass -File double-tap-detector.ps1 -Autotest
```

18 verifiche della logica con eventi simulati (limiti esatti di 35/36/60 ms, rimbalzo al rilascio, ripetizione automatica, PC rallentato, rilascio perso, tasti diversi ravvicinati…) più un test completo che preme F24 due volte tramite Windows e controlla che venga letto con la Raw Input API e rilevato.

![Risultato dell'autotest](docs/autotest.png)

## Tecnologie

PowerShell e C# compilato all'avvio con `Add-Type`; API Win32 tramite P/Invoke (`RegisterRawInputDevices`, `GetRawInputData`, `SendInput`); Windows Forms per l'icona e le notifiche. Un solo file, nessuna dipendenza.

## Licenza

[MIT](LICENSE)
