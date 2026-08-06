# Offerte Vergelijker (Web)

Streamlit-app die PDF-offertes vergelijkt met NETTO-prijslijsten en het IB-budget.
De app heeft een startpagina met een tegel per leverancier (bijv. Van Keulen, Koeling).

## Lokaal draaien

```bash
pip install -r ../requirements.txt
python -m streamlit run Home.py
```

De app is dan bereikbaar op `http://localhost:8501`.

## Draaien op een computer

Deze stappen zijn voor iemand die de app lokaal wil draaien, zonder dat Python of dit project al geïnstalleerd is.

### 1. Python installeren

- Download Python 3.11 (of hoger) van [python.org/downloads](https://www.python.org/downloads/)
- Bij de installatie: vink **"Add python.exe to PATH"** aan onderaan het installatiescherm

Controleer daarna in een terminal (Opdrachtprompt / PowerShell):

```bash
python --version
```

### 2. Bestanden ontvangen

Kopieer de map `IB_met_OfferteVergelijker` en het bestand `requirements.txt` naar de computer (bijv. via een USB-stick, e-mail-bijlage, of gedeelde map), met deze structuur:

```
project-map/
├── requirements.txt
└── IB_met_OfferteVergelijker/
    ├── Home.py
    ├── common.py
    └── pages/
        ├── 1_🟡_Van_Keulen.py
        └── 2_🧊_Koeling.py
```

> Als het via GitHub gedeeld is: gebruik "Code" → "Download ZIP" op de repository-pagina en pak het uit.

### 3. Dependencies installeren

Open een terminal in de `project-map` en voer uit:

```bash
pip install -r requirements.txt
```

### 4. App starten

```bash
cd IB_met_OfferteVergelijker
python -m streamlit run Home.py
```

Streamlit opent automatisch een browservenster op `http://localhost:8501`. Gebeurt dit niet, open dan die link handmatig.

Om te stoppen: ga terug naar de terminal en druk op `Ctrl + C`.

## Deployen naar Azure App Service

### Vereisten
- Een Azure-abonnement
- [Azure CLI](https://learn.microsoft.com/cli/azure/install-azure-cli) geïnstalleerd, en ingelogd via `az login`

### 1. Resources aanmaken

```bash
az group create -n rg-offerte-app -l westeurope

az appservice plan create \
  -n plan-offerte \
  --resource-group rg-offerte-app \
  --sku B1 \
  --is-linux

az webapp create \
  -n offerte-vergelijker \
  --resource-group rg-offerte-app \
  --plan plan-offerte \
  --runtime "PYTHON:3.11"
```

> `offerte-vergelijker` moet een wereldwijd unieke naam zijn — pas aan indien nodig. De uiteindelijke URL wordt `https://<naam>.azurewebsites.net`.

### 2. Startup command instellen

App Service moet weten hoe de app te starten (`startup.sh` in deze map):

```bash
az webapp config set \
  -n offerte-vergelijker \
  --resource-group rg-offerte-app \
  --startup-file "IB_met_OfferteVergelijker/startup.sh"
```

### 3. Deployen

Vanuit de root van de repository (waar `requirements.txt` staat):

```bash
az webapp up \
  -n offerte-vergelijker \
  --resource-group rg-offerte-app \
  --plan plan-offerte \
  --runtime "PYTHON:3.11"
```

Dit zipt de lokale bestanden en deployt ze direct — geen GitHub nodig.

### 4. Controleren

```bash
az webapp browse -n offerte-vergelijker --resource-group rg-offerte-app
```

Logs bekijken bij problemen:

```bash
az webapp log tail -n offerte-vergelijker --resource-group rg-offerte-app
```

## Toegang beperken (optioneel, aanbevolen)

Om de app niet publiek toegankelijk te maken, kan Azure AD-authenticatie worden ingeschakeld via het **Authentication**-blad van de App Service in de Azure Portal, of via:

```bash
az webapp auth update \
  -n offerte-vergelijker \
  --resource-group rg-offerte-app \
  --enabled true \
  --action LoginWithAzureActiveDirectory
```

## Bestanden

| Bestand | Doel |
|---|---|
| `Home.py` | Startpagina — kies een leverancier |
| `common.py` | Gedeelde styling/header helpers |
| `pages/1_🟡_Van_Keulen.py` | Sectie: Van Keulen vergelijker |
| `pages/2_🧊_Koeling.py` | Sectie: Koeling vergelijker (in opbouw) |
| `startup.sh` | Startcommando voor Azure App Service |
| `../requirements.txt` | Python-dependencies (gedeeld met de andere scripts in dit project) |
