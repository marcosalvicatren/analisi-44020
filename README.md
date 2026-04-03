# Analisi Conto 44020 – Fatture da Ricevere

Applicazione web per la verifica delle scritture di assestamento sul conto **44020 – Fornitori fatture da ricevere beni e servizi** (GB Software, codice D0007A02).

## Funzionalità

1. **Check 1** – Verifica che il saldo vada a zero durante l'anno di bilancio e chiuda in avere al 31.12
2. **Check 2** – Matching cifra per cifra tra gli avere al 31.12 dell'anno di bilancio e i dare dell'anno successivo
3. **Check 3** – Verifica che i totali avere/dare corrispondano (anche senza matching puntuale)
4. **Check 4** – Elenca gli importi senza corrispondente con data, descrizione e segno
5. **Export Excel** – Report scaricabile in formato .xls

## Formati file accettati

- **TXT** (tab-separato, export diretto da GB Software) ← **consigliato**
- **XLS** (vecchio formato BIFF, export da GB Software)
- **PDF** (stampa da GB Software)

## Deploy su Streamlit Cloud (gratuito)

### 1. Crea il repository su GitHub

```bash
# Crea una cartella locale
mkdir analisi-44020
cd analisi-44020

# Copia i file
# app.py
# requirements.txt
# README.md

# Inizializza git
git init
git add .
git commit -m "Prima versione analisi 44020"

# Crea repo su GitHub (github.com → New repository → analisi-44020)
git remote add origin https://github.com/TUO_USERNAME/analisi-44020.git
git push -u origin main
```

### 2. Deploy su Streamlit Cloud

1. Vai su **https://share.streamlit.io**
2. Accedi con il tuo account GitHub
3. Clicca **"New app"**
4. Seleziona il repository `analisi-44020`
5. Branch: `main`
6. Main file path: `app.py`
7. Clicca **"Deploy"**

In pochi minuti avrai un link pubblico (es. `https://analisi-44020.streamlit.app`) da condividere con i collaboratori.

### Struttura del repository

```
analisi-44020/
├── app.py           ← l'applicazione
├── requirements.txt ← dipendenze Python
└── README.md        ← questo file
```

## Come si usa

1. Apri l'app dal link Streamlit
2. Nella **barra laterale**: inserisci gli anni (precedente, bilancio, successivo)
3. Carica il file mastrino esportato da GB Software
4. Leggi il report nelle 5 tab:
   - **① Saldo a zero** – verifica apertura/chiusura
   - **② Matching cifre** – abbinamento puntuale
   - **③ Totali** – confronto somme
   - **④ Non abbinati** – anomalie dettagliate
   - **📋 Mastrino** – tutti i movimenti filtrabili
5. Scarica il report Excel con il pulsante in fondo

## Note tecniche

- Tutto gira in cloud, nessuna installazione richiesta
- I file caricati non vengono salvati (sessione temporanea)
- Compatibile con GB Software versione corrente (export al 2026)
