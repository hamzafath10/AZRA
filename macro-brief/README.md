# MacroBrief

Briefing macro / marchés du jour, directement dans le terminal Windows, avec un rapport HTML propre.

Le script scanne une vingtaine de sources, regroupe les articles qui parlent du même sujet
(« 3 sources ») et les classe par thème et par importance.

| Bloc | Source |
|---|---|
| Cotations (indices, taux US, dollar, pétrole, or, BTC) | Yahoo Finance (différé) |
| Agenda macro du jour / demain / semaine (prévisions, précédents, réels) | Calendrier ForexFactory |
| News | Google News FR/EN, CNBC, MarketWatch, WSJ, FT, Yahoo Finance, Investing, Les Echos, Le Monde, BFM |
| Banques centrales | Flux officiels Fed, BCE, Bank of England |

Thèmes : Banques centrales, Inflation, Emploi, Croissance & activité, Politique & géopolitique,
Énergie & matières premières, Marchés, Crypto.

## Utilisation

1. Télécharge le dossier `macro-brief` sur ton PC.
2. Double-clique sur **`MacroBrief.bat`**.

Ou bien, dans PowerShell :

```powershell
.\MacroBrief.ps1                         # briefing complet, dernières 24 h
.\MacroBrief.ps1 -Hours 12               # seulement les 12 dernières heures
.\MacroBrief.ps1 -Lang fr                # uniquement les sources françaises
.\MacroBrief.ps1 -PerCategory 10 -Top 15 # plus de sujets
.\MacroBrief.ps1 -NoBrowser              # sans ouvrir le rapport HTML
```

Les rapports sont enregistrés dans `Documents\MacroBrief\MacroBrief_AAAA-MM-JJ.html`.

> Si Windows bloque le script (fichier téléchargé), clic droit sur le `.ps1` → Propriétés → **Débloquer**,
> ou lance `Unblock-File .\*.ps1` dans le dossier.

## Tous les matins automatiquement

```powershell
.\Install-MacroBriefTask.ps1               # lundi→vendredi à 07:30
.\Install-MacroBriefTask.ps1 -Time 08:15 -Weekend
.\Install-MacroBriefTask.ps1 -Remove
```

## Personnaliser

Tout se règle en haut de `MacroBrief.ps1`, dans la section **CONFIGURATION** :

- `$Sources` : ajoute ou retire un flux RSS/Atom (`Name`, `Lang`, `Url`). La fonction
  `New-GoogleNewsUrl` permet de créer un flux sur n'importe quelle recherche.
- `$Tickers` : symboles Yahoo Finance (ex. `^FCHI`, `EURUSD=X`, `GC=F`).
- `$CalendarCurrencies` : devises suivies dans l'agenda.
- `$Categories` / `$HighImpact` : mots-clés des thèmes et des sujets prioritaires.

Compatible Windows PowerShell 5.1 (inclus dans Windows 10/11) et PowerShell 7. Aucune clé API.
Une source indisponible est simplement ignorée (elle apparaît barrée en bas du rapport).
