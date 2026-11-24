# Scadenzario SSL e Ambiente

Web app per gestire le scadenze di **sicurezza sul lavoro e ambiente** di decine di aziende clienti: valutazioni dei rischi, verifiche DPR 462, CPI, prove di evacuazione, AUA, MUD, F-gas, ecc. Ci lavora un team, anche in contemporanea. Si può **installare come app desktop**.

Non c'è server e non c'è database: i dati sono file JSON in una **cartella condivisa su Dropbox** (o su OneDrive o Google Drive per desktop).

```
Repository GitHub (solo il programma)     Cartella condivisa Dropbox (solo i dati)
├── index.html                            ├── config.json     ← catalogo, sezioni, soglie
├── manifest.webmanifest                  └── aziende/
├── sw.js                                     ├── alimec.json ← un file per azienda
├── icons/                                    └── ...
└── tools/import_excel.py
```

> **Nel repository va solo il codice.** I dati dei clienti stanno su Dropbox e non passano mai da GitHub.

---

## Novità della versione 0.2
- **Nuova grafica**: tema verde, menu a sinistra (diventa un menu a scomparsa su schermi piccoli), aziende mostrate come card.
- **Aziende con RSPP dello studio** ben riconoscibili: etichetta con scudo **RSPP**, bordo dorato sulle card, barra dorata a sinistra negli elenchi e 🛡 nel calendario. Ovunque si possono filtrare con "solo RSPP".
- **Sezioni separate**: Valutazioni dei rischi, Adempimenti SSL, Adempimenti ambientali, Altri. Nella scheda azienda sono tab distinte; in Scadenze, Calendario e Matrice sono filtri.
- **Filtro per stato nella scheda azienda**: In regola, Scaduti, Entro 30/180 gg, Da fare, Valutati nel DVR, Non in gestione, NP. Accanto a ogni pulsante c'è il numero di voci.
- **"Scadenza in gestione allo studio"** (spunta su ogni adempimento): se la togli, il dato resta registrato con la sua data e la scadenza calcolata, ma **non diventa mai critico** e non compare nei conteggi. Serve per gli adempimenti seguiti da altri, per esempio il manutentore, il frigorista o il professionista antincendio.
- **Calendario** mensile: elenco delle scadenze del giorno, filtri, ed **esportazione .ics per Outlook o Google Calendar** con promemoria 30 giorni prima.
- **Stampa / PDF** della scheda azienda.
- **Filtro per referente interno** in Scadenze.
- **Nuovi adempimenti ambientali** nel catalogo: autocontrolli emissioni, autocontrolli scarichi, MUD, F-gas, valutazione di impatto acustico.
- **App installabile** sul computer, con icona dedicata.

Un `config.json` creato con la versione 0.1 viene **aggiornato da solo** alla prima apertura: aggiunge le sezioni e i nuovi adempimenti ambientali, senza toccare i dati delle aziende.

---

## Installazione

### 1. Pubblica l'app su GitHub Pages
1. Crea un repository su GitHub e carica **tutti i file di questa cartella** (`index.html`, `manifest.webmanifest`, `sw.js`, la cartella `icons/`, ecc.).
2. Vai in *Settings → Pages → Branch: main / root → Save*.
3. Dopo un minuto l'app è raggiungibile a `https://<utente>.github.io/<repo>/`.

### 2. Installala come app desktop (ogni collega)
1. Apri l'indirizzo di GitHub Pages con **Chrome o Edge**.
2. Clicca **"⬇ Installa come app"** in basso nel menu a sinistra. In alternativa usa l'icona di installazione nella barra degli indirizzi, oppure il menu ⋮ → *Installa Scadenzario SSL e Ambiente*.
3. Da quel momento l'app si comporta come un programma:
   - si apre in una finestra sua;
   - ha l'icona nel menu Start (Windows) o nel Launchpad/Dock (Mac) e si può fissare sulla barra delle applicazioni;
   - **si aggiorna da sola** quando pubblicate una nuova versione su GitHub;
   - si apre anche senza internet; i dati vengono comunque letti dalla cartella Dropbox locale.

> L'installazione funziona solo aprendo l'app dall'indirizzo GitHub Pages (https). Se apri `index.html` con un doppio clic, l'app funziona ma non si può installare.

### 3. Importa l'Excel esistente
Serve Python 3 con `openpyxl` (`pip install openpyxl`).
```bash
python3 tools/import_excel.py "DVR_SCADENZARIO_CLIENTI.xlsx" "C:/Users/.../Dropbox/Scadenzario"
```
Se un file esiste già, lo script non lo sovrascrive. Per forzare, aggiungi `--force`.

### 4. Primo avvio
Clicca **Seleziona cartella dati**, scegli la cartella Dropbox condivisa e inserisci il tuo nome, che finisce nel registro modifiche. Dalla volta successiva basta cliccare **Riconnetti**.

---

## Lavoro in team
1. **Ogni azienda è un file separato**: chi lavora su aziende diverse non si tocca mai.
2. **Ogni salvataggio rilegge il file appena prima di scrivere** e applica solo la modifica fatta, così le modifiche dei colleghi sulla stessa azienda non si perdono.
3. **Ogni 30 secondi l'app ricarica** i file cambiati dai colleghi.
4. Se due persone salvano la stessa azienda **nello stesso istante**, Dropbox crea una "copia in conflitto". L'app la segnala con un avviso in alto.
5. Ogni azienda ha un **registro modifiche** (chi, quando, cosa).

## Come si calcola lo stato
- **Scadenza** = data dell'ultima revisione + periodicità. La periodicità è quella dell'azienda se impostata, altrimenti quella standard del catalogo. Se c'è una scadenza manuale, vale quella.
- **Stati in gestione**:
  - Scaduto;
  - Entro 30 gg e Entro 180 gg (le soglie si cambiano in Impostazioni);
  - In regola;
  - Da fare (applicabile ma senza date, oppure "da valutare");
  - Senza scadenza (per esempio il piano di emergenza).
- **Stati che non generano allarmi**: Valutato nel DVR, Non in gestione, NP, Non serve.

## Periodicità standard e fonti
| Sezione | Adempimento | Anni | Fonte | Note |
|---|---|---|---|---|
| Valutazioni rischi | DVR | 4 | D.Lgs. 81/08 artt. 28-29 | **Prassi**: la legge chiede di rielaborarlo entro 30 gg da modifiche significative (art. 29 c.3). |
| | Rumore, vibrazioni, ROA, CEM | 4 | D.Lgs. 81/08 art. 181 c.2 | Obbligo di legge: "almeno ogni quattro anni". |
| | ATEX, VDT, MMC, OCRA | 4 | Titoli XI, VII, VI | **Prassi** |
| | Rischio chimico | 3 | art. 223 c.7 | **Prassi** ("periodicamente") |
| | Stress lavoro-correlato | 2 | art. 28 c.1-bis | **Prassi** |
| | Amianto | 1 | DM 06/09/1994 | |
| Adempimenti SSL | Messa a terra, scariche atmosferiche | 5 (o 2) | DPR 462/2001 art. 4 | 2 anni per cantieri, locali a uso medico e ambienti a maggior rischio incendio. |
| | Prova di evacuazione | 1 | DM 02/09/2021 (GSA) | |
| | Riunione periodica | 1 | art. 35 | Aziende con più di 15 lavoratori. |
| | SCIA / CPI | 5 | DPR 151/2011 art. 5 | 10 anni per alcune attività. |
| | Apparecchi di sollevamento | 2 | art. 71 c.11, All. VII | 1 o 2 anni a seconda del tipo. |
| Ambiente | AUA | 15 | DPR 59/2013 art. 3 c.6 | |
| | Autocontrolli emissioni / scarichi | 1 | D.Lgs. 152/06 artt. 269 e 124 | La frequenza **reale** è quella scritta nell'autorizzazione: impostala azienda per azienda. |
| | MUD | 1 | L. 70/1994 | Entro il 30 aprile di ogni anno. |
| | F-gas (controllo perdite) | 1 | Reg. (UE) 2024/573 art. 5 | 12, 6 o 3 mesi in base alle t CO2 eq: imposta 0,5 o 0,25 anni dove serve. |
| | Impatto acustico | — | L. 447/1995 art. 8 | Da rifare quando cambiano impianti o attività. |
| Altri | HACCP | 3 | Reg. CE 852/2004 | **Prassi** |

Le voci segnate come **prassi** sono scelte dello studio e si modificano in *Impostazioni → Catalogo*. Lì si può anche spostare un adempimento in un'altra sezione o aggiungerne di nuovi.

## Limiti
- Serve **Chrome o Edge** su computer: niente telefono e niente Safari/Firefox.
- La sicurezza dei dati è quella di Dropbox: chi ha accesso alla cartella vede tutto.
