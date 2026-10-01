<#
.SYNOPSIS
    MacroBrief - Briefing macro / marchés quotidien dans le terminal Windows.

.DESCRIPTION
    Scanne plusieurs sources (RSS presse FR/EN, banques centrales, Google News),
    le calendrier économique (ForexFactory) et un snapshot des marchés (Yahoo Finance).
    Regroupe les articles qui parlent du même sujet, les classe par thème et par
    importance, puis affiche un résumé propre dans la console et génère un rapport
    HTML (ouvert automatiquement dans le navigateur).

    Compatible Windows PowerShell 5.1 et PowerShell 7+. Aucune clé API requise.

.EXAMPLE
    .\MacroBrief.ps1
.EXAMPLE
    .\MacroBrief.ps1 -Hours 12 -Lang fr -NoBrowser
#>
[CmdletBinding()]
param(
    # Fenêtre de temps des news (en heures)
    [int]$Hours = 24,
    # Langue des sources : all | fr | en
    [ValidateSet('all', 'fr', 'en')]
    [string]$Lang = 'all',
    # Nombre de sujets affichés par thème
    [int]$PerCategory = 6,
    # Nombre de sujets dans "A la une"
    [int]$Top = 10,
    # Dossier de sortie des rapports HTML
    [string]$OutDir = (Join-Path ([Environment]::GetFolderPath('MyDocuments')) 'MacroBrief'),
    # Ne pas ouvrir le rapport HTML
    [switch]$NoBrowser,
    # Ne rien afficher dans la console (utile en tâche planifiée)
    [switch]$Quiet,
    # Timeout réseau par requête (secondes)
    [int]$TimeoutSec = 15,
    # Fichier de sources local pour les tests (dossier contenant des .xml/.json) - usage interne
    [string]$OfflineDir
)

$ErrorActionPreference = 'Stop'
Set-StrictMode -Version 2

# =============================================================================
#  CONFIGURATION - modifie librement les sources, tickers et mots-clés
# =============================================================================

function New-GoogleNewsUrl([string]$Query, [string]$L) {
    $q = [uri]::EscapeDataString($Query)
    if ($L -eq 'fr') { return "https://news.google.com/rss/search?q=$q&hl=fr&gl=FR&ceid=FR:fr" }
    return "https://news.google.com/rss/search?q=$q&hl=en-US&gl=US&ceid=US:en"
}

$Sources = @(
    # --- Agrégateurs (Google News regroupe des centaines de médias) ---
    @{ Name = 'Google News FR - Macro';   Lang = 'fr'; Url = (New-GoogleNewsUrl '(BCE OR Fed OR inflation OR "taux directeur" OR PIB OR chômage OR récession) when:1d' 'fr') }
    @{ Name = 'Google News FR - Marchés'; Lang = 'fr'; Url = (New-GoogleNewsUrl '(Bourse OR "CAC 40" OR "Wall Street" OR obligataire OR pétrole) when:1d' 'fr') }
    @{ Name = 'Google News EN - Macro';   Lang = 'en'; Url = (New-GoogleNewsUrl '("Federal Reserve" OR ECB OR CPI OR payrolls OR GDP OR "rate cut" OR tariffs) when:1d' 'en') }
    @{ Name = 'Google News EN - Markets'; Lang = 'en'; Url = (New-GoogleNewsUrl '(stocks OR "Treasury yields" OR "oil prices" OR dollar OR "S&P 500") when:1d' 'en') }

    # --- Presse financière EN ---
    @{ Name = 'CNBC Economy';        Lang = 'en'; Url = 'https://www.cnbc.com/id/20910258/device/rss/rss.html' }
    @{ Name = 'CNBC Markets';        Lang = 'en'; Url = 'https://www.cnbc.com/id/100727362/device/rss/rss.html' }
    @{ Name = 'MarketWatch';         Lang = 'en'; Url = 'https://feeds.content.dowjones.io/public/rss/mw_topstories' }
    @{ Name = 'WSJ Markets';         Lang = 'en'; Url = 'https://feeds.a.dj.com/rss/RSSMarketsMain.xml' }
    @{ Name = 'FT Global Economy';   Lang = 'en'; Url = 'https://www.ft.com/global-economy?format=rss' }
    @{ Name = 'FT Markets';          Lang = 'en'; Url = 'https://www.ft.com/markets?format=rss' }
    @{ Name = 'Yahoo Finance';       Lang = 'en'; Url = 'https://finance.yahoo.com/news/rssindex' }
    @{ Name = 'Investing.com Eco';   Lang = 'en'; Url = 'https://www.investing.com/rss/news_14.rss' }

    # --- Presse financière FR ---
    @{ Name = 'Les Echos Economie';  Lang = 'fr'; Url = 'https://services.lesechos.fr/rss/les-echos-economie.xml' }
    @{ Name = 'Les Echos Marchés';   Lang = 'fr'; Url = 'https://services.lesechos.fr/rss/les-echos-finance-marches.xml' }
    @{ Name = 'Le Monde Economie';   Lang = 'fr'; Url = 'https://www.lemonde.fr/economie/rss_full.xml' }
    @{ Name = 'BFM Economie';        Lang = 'fr'; Url = 'https://www.bfmtv.com/rss/economie/' }
    @{ Name = 'Investing.com FR';    Lang = 'fr'; Url = 'https://fr.investing.com/rss/news_14.rss' }

    # --- Banques centrales (sources officielles) ---
    @{ Name = 'Federal Reserve';     Lang = 'en'; Url = 'https://www.federalreserve.gov/feeds/press_all.xml' }
    @{ Name = 'BCE / ECB';           Lang = 'en'; Url = 'https://www.ecb.europa.eu/rss/press.html' }
    @{ Name = 'Bank of England';     Lang = 'en'; Url = 'https://www.bankofengland.co.uk/rss/news' }
)

# Snapshot marchés (symboles Yahoo Finance). Kind: index | rate | fx | commo | crypto
$Tickers = @(
    @{ Sym = '^GSPC';     Name = 'S&P 500';        Kind = 'index' }
    @{ Sym = '^IXIC';     Name = 'Nasdaq';         Kind = 'index' }
    @{ Sym = '^DJI';      Name = 'Dow Jones';      Kind = 'index' }
    @{ Sym = '^FCHI';     Name = 'CAC 40';         Kind = 'index' }
    @{ Sym = '^STOXX50E'; Name = 'Euro Stoxx 50';  Kind = 'index' }
    @{ Sym = '^GDAXI';    Name = 'DAX';            Kind = 'index' }
    @{ Sym = '^N225';     Name = 'Nikkei 225';     Kind = 'index' }
    @{ Sym = '^VIX';      Name = 'VIX';            Kind = 'index' }
    @{ Sym = '^TNX';      Name = 'US 10 ans';      Kind = 'rate' }
    @{ Sym = '^IRX';      Name = 'US 3 mois';      Kind = 'rate' }
    @{ Sym = 'EURUSD=X';  Name = 'EUR/USD';        Kind = 'fx' }
    @{ Sym = 'DX-Y.NYB';  Name = 'Dollar Index';   Kind = 'fx' }
    @{ Sym = 'USDJPY=X';  Name = 'USD/JPY';        Kind = 'fx' }
    @{ Sym = 'BZ=F';      Name = 'Brent';          Kind = 'commo' }
    @{ Sym = 'CL=F';      Name = 'WTI';            Kind = 'commo' }
    @{ Sym = 'GC=F';      Name = 'Or';             Kind = 'commo' }
    @{ Sym = 'NG=F';      Name = 'Gaz naturel US'; Kind = 'commo' }
    @{ Sym = 'BTC-USD';   Name = 'Bitcoin';        Kind = 'crypto' }
)

# Calendrier économique (ForexFactory, gratuit, sans clé)
$CalendarUrl = 'https://nfs.faireconomy.media/ff_calendar_thisweek.json'
# Devises suivies dans le calendrier
$CalendarCurrencies = @('USD', 'EUR', 'GBP', 'JPY', 'CNY', 'CHF', 'CAD')

# Thèmes : regex (FR + EN, insensible à la casse, accents retirés avant comparaison)
$Categories = [ordered]@{
    'Banques centrales'   = '\b(fed|fomc|powell|bce|ecb|lagarde|boe|bank of england|boj|bank of japan|pboc|snb|banque centrale|central bank|taux directeurs?|rate (cut|hike)s?|baisse des taux|hausse des taux|monetary policy|politique monetaire|quantitative|bailey|ueda|waller|schnabel)\b'
    'Inflation'           = '\b(inflation|cpi|pce|ppi|ipc|prix a la consommation|consumer prices|desinflation|disinflation|deflation)\b'
    'Emploi'              = '\b(payrolls?|nfp|jobless|unemployment|chomage|emplois?|jobs report|jolts|labor market|marche du travail|salaires|wages)\b'
    'Croissance & activité' = '\b(gdp|pib|pmi|ism|recession|croissance|growth|retail sales|ventes au detail|industrial production|production industrielle|consumer confidence|confiance des menages|ifo|zew|housing)\b'
    'Politique & géopolitique' = '\b(tariffs?|droits de douane|trade war|guerre commerciale|sanctions?|elections?|trump|shutdown|budget|deficit|dette publique|debt ceiling|ukraine|russie|russia|chine|china|israel|iran|taiwan|g7|g20)\b'
    'Énergie & matières premières' = '\b(oil|petrole|brent|wti|opep|opec|gas|gaz|gold|or noir|copper|cuivre|commodities|matieres premieres|lng)\b'
    'Marchés'             = '\b(stocks?|actions?|wall street|s&p|nasdaq|dow|cac|dax|stoxx|bourse|bonds?|yields?|treasur(y|ies)|obligataires?|oat|bund|dollar|euro|forex|devises|earnings|resultats|ipo|volatilit[ey]|vix)\b'
    'Crypto'              = '\b(bitcoin|btc|ether(eum)?|crypto(s|monnaies?)?|stablecoins?)\b'
}

# Mots-clés "gros impact" : bonus de score
$HighImpact = '\b(fomc|powell|lagarde|rate (cut|hike)|baisse des taux|hausse des taux|cpi|pce|payrolls|nfp|gdp|pib|recession|tariffs?|droits de douane|default|defaut|shutdown|crash|krach|emergency|urgence|surprise)\b'

# Mots ignorés pour le regroupement des titres
$StopWords = @(
    'the','a','an','and','or','of','to','in','on','for','with','as','at','by','from','is','are','be','its','it','this','that','after','over','into','amid','says','say','will','new','how','why','what','who','than','more','up','down','out','about','but','not','has','have','was','were',
    'le','la','les','un','une','des','du','de','et','ou','en','au','aux','dans','sur','pour','par','avec','sans','est','sont','qui','que','quoi','ce','cette','ces','se','sa','son','ses','plus','moins','pas','apres','avant','face','selon','comme','vers','fait','leur','leurs','entre'
)

$UserAgent = 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/128.0 Safari/537.36'

# =============================================================================
#  OUTILS
# =============================================================================

try { [Console]::OutputEncoding = [System.Text.Encoding]::UTF8 } catch { }
[Net.ServicePointManager]::SecurityProtocol = [Net.ServicePointManager]::SecurityProtocol -bor [Net.SecurityProtocolType]::Tls12
$Invariant = [Globalization.CultureInfo]::InvariantCulture
$French = [Globalization.CultureInfo]::GetCultureInfo('fr-FR')

function Write-Log([string]$Msg) {
    if (-not $Quiet) { Write-Host "  $Msg" -ForegroundColor DarkGray }
}

function Remove-Diacritics([string]$s) {
    if (-not $s) { return '' }
    $n = $s.Normalize([Text.NormalizationForm]::FormD)
    return ([regex]::Replace($n, '\p{Mn}', '')).Normalize([Text.NormalizationForm]::FormC)
}

function ConvertFrom-Html([string]$s) {
    if (-not $s) { return '' }
    $s = [regex]::Replace($s, '<[^>]+>', ' ')
    $s = [System.Net.WebUtility]::HtmlDecode($s)
    return ([regex]::Replace($s, '\s+', ' ')).Trim()
}

function Get-Html([string]$s) { return [System.Net.WebUtility]::HtmlEncode($s) }

$ZoneMap = @{
    'GMT' = '+00:00'; 'UTC' = '+00:00'; 'UT' = '+00:00'; 'Z' = '+00:00'
    'EST' = '-05:00'; 'EDT' = '-04:00'; 'CST' = '-06:00'; 'CDT' = '-05:00'
    'MST' = '-07:00'; 'MDT' = '-06:00'; 'PST' = '-08:00'; 'PDT' = '-07:00'
    'CET' = '+01:00'; 'CEST' = '+02:00'; 'BST' = '+01:00'; 'JST' = '+09:00'
}

function ConvertTo-DateOffset([string]$s) {
    if ([string]::IsNullOrWhiteSpace($s)) { return $null }
    $s = $s.Trim()
    $dto = [DateTimeOffset]::MinValue
    $styles = [Globalization.DateTimeStyles]::AssumeUniversal
    if ([DateTimeOffset]::TryParse($s, $Invariant, $styles, [ref]$dto)) { return $dto }
    # RFC 822 : "Wed, 01 Oct 2026 10:00:00 GMT" / "+0000"
    $t = $s -replace '^[A-Za-z]+,\s*', ''
    $m = [regex]::Match($t, '\s([A-Za-z]{1,4})$')
    if ($m.Success -and $ZoneMap.ContainsKey($m.Groups[1].Value.ToUpper())) {
        $t = $t.Substring(0, $m.Index) + ' ' + $ZoneMap[$m.Groups[1].Value.ToUpper()]
    }
    $t = $t -replace '\s([+-]\d{2})(\d{2})$', ' $1:$2'
    if ([DateTimeOffset]::TryParse($t, $Invariant, $styles, [ref]$dto)) { return $dto }
    return $null
}

function Format-Ago([DateTimeOffset]$d) {
    $span = [DateTimeOffset]::Now - $d
    if ($span.TotalMinutes -lt 1) { return "à l'instant" }
    if ($span.TotalMinutes -lt 60) { return ('il y a {0} min' -f [int]$span.TotalMinutes) }
    if ($span.TotalHours -lt 24) { return ('il y a {0} h' -f [int][math]::Floor($span.TotalHours)) }
    return ('il y a {0} j' -f [int][math]::Floor($span.TotalDays))
}

# Téléchargement parallèle (runspaces) : compatible PowerShell 5.1
function Invoke-ParallelFetch {
    param([string[]]$Urls, [int]$Throttle = 12)

    $results = @{}
    if ($OfflineDir) {
        # Mode test : les fichiers sont nommés d'après le hash de l'URL
        foreach ($u in $Urls) {
            $f = Join-Path $OfflineDir (Get-UrlKey $u)
            if (Test-Path $f) { $results[$u] = [pscustomobject]@{ Ok = $true; Bytes = [IO.File]::ReadAllBytes($f); Error = $null } }
            else { $results[$u] = [pscustomobject]@{ Ok = $false; Bytes = $null; Error = 'absent (offline)' } }
        }
        return $results
    }

    $worker = {
        param($Url, $TimeoutMs, $UA)
        try {
            $req = [System.Net.HttpWebRequest]::Create($Url)
            $req.UserAgent = $UA
            $req.Accept = 'application/rss+xml, application/atom+xml, application/xml, application/json, text/xml, */*'
            $req.Headers.Add('Accept-Language', 'fr-FR,fr;q=0.9,en;q=0.8')
            $req.Timeout = $TimeoutMs
            $req.ReadWriteTimeout = $TimeoutMs
            $req.AutomaticDecompression = [System.Net.DecompressionMethods]::GZip -bor [System.Net.DecompressionMethods]::Deflate
            $resp = $req.GetResponse()
            try {
                $ms = New-Object System.IO.MemoryStream
                $resp.GetResponseStream().CopyTo($ms)
                [pscustomobject]@{ Ok = $true; Bytes = $ms.ToArray(); Error = $null }
            } finally { $resp.Close() }
        } catch {
            $msg = $_.Exception.Message
            if ($_.Exception.InnerException) { $msg = $_.Exception.InnerException.Message }
            [pscustomobject]@{ Ok = $false; Bytes = $null; Error = $msg }
        }
    }

    $pool = [runspacefactory]::CreateRunspacePool(1, $Throttle)
    $pool.Open()
    $jobs = @()
    foreach ($u in $Urls) {
        $ps = [powershell]::Create()
        $ps.RunspacePool = $pool
        [void]$ps.AddScript($worker).AddArgument($u).AddArgument($TimeoutSec * 1000).AddArgument($UserAgent)
        $jobs += [pscustomobject]@{ Url = $u; PS = $ps; Handle = $ps.BeginInvoke() }
    }
    foreach ($j in $jobs) {
        try {
            $out = $j.PS.EndInvoke($j.Handle)
            if ($out -and $out.Count -gt 0) { $results[$j.Url] = $out[0] }
            else { $results[$j.Url] = [pscustomobject]@{ Ok = $false; Bytes = $null; Error = 'réponse vide' } }
        } catch {
            $results[$j.Url] = [pscustomobject]@{ Ok = $false; Bytes = $null; Error = $_.Exception.Message }
        } finally { $j.PS.Dispose() }
    }
    $pool.Close(); $pool.Dispose()
    return $results
}

function Get-UrlKey([string]$Url) {
    $sha = [Security.Cryptography.SHA1]::Create()
    $h = $sha.ComputeHash([Text.Encoding]::UTF8.GetBytes($Url))
    return (-join ($h | ForEach-Object { $_.ToString('x2') })).Substring(0, 16)
}

function Get-ChildText($Node, [string[]]$Names) {
    foreach ($n in $Names) {
        foreach ($c in $Node.ChildNodes) {
            if ($c.LocalName -eq $n) {
                if ($n -eq 'link' -and $c.Attributes -and $c.Attributes['href']) {
                    $rel = $c.Attributes['rel']
                    if (-not $rel -or $rel.Value -eq 'alternate') { return $c.Attributes['href'].Value }
                    continue
                }
                $txt = $c.InnerText
                if (-not [string]::IsNullOrWhiteSpace($txt)) { return $txt.Trim() }
            }
        }
    }
    return ''
}

function ConvertFrom-Feed([byte[]]$Bytes, $Source) {
    $settings = New-Object System.Xml.XmlReaderSettings
    $settings.DtdProcessing = [System.Xml.DtdProcessing]::Ignore
    $settings.XmlResolver = $null
    $ms = New-Object System.IO.MemoryStream(, $Bytes)
    $reader = [System.Xml.XmlReader]::Create($ms, $settings)
    $xml = New-Object System.Xml.XmlDocument
    $xml.Load($reader)
    $reader.Close()

    $items = New-Object System.Collections.Generic.List[object]
    foreach ($node in $xml.SelectNodes("//*[local-name()='item' or local-name()='entry']")) {
        $title = ConvertFrom-Html (Get-ChildText $node @('title'))
        if (-not $title) { continue }
        $link = (Get-ChildText $node @('link', 'guid', 'id')).Trim()
        $date = ConvertTo-DateOffset (Get-ChildText $node @('pubDate', 'published', 'updated', 'date', 'issued'))
        $desc = ConvertFrom-Html (Get-ChildText $node @('description', 'summary', 'encoded', 'content'))
        if ($desc.Length -gt 280) { $desc = $desc.Substring(0, 277) + '...' }

        $publisher = $Source.Name
        # Google News : le média est en fin de titre "Titre - Média"
        if ($Source.Url -like '*news.google.com*') {
            $srcNode = Get-ChildText $node @('source')
            $m = [regex]::Match($title, '^(.*)\s+[-–]\s+([^-–]{2,60})$')
            if ($m.Success) { $title = $m.Groups[1].Value.Trim(); $publisher = $m.Groups[2].Value.Trim() }
            if ($srcNode) { $publisher = $srcNode }
            $desc = ''
        }

        $items.Add([pscustomobject]@{
            Title     = $title
            Link      = $link
            Date      = $date
            Desc      = $desc
            Publisher = $publisher
            Feed      = $Source.Name
            Lang      = $Source.Lang
        })
    }
    return $items
}

function Get-Tokens([string]$Title) {
    $t = (Remove-Diacritics $Title).ToLowerInvariant()
    $t = [regex]::Replace($t, "[^a-z0-9&%\.]+", ' ')
    $set = New-Object 'System.Collections.Generic.HashSet[string]'
    foreach ($w in $t.Split(' ', [StringSplitOptions]::RemoveEmptyEntries)) {
        $w = $w.Trim('.')
        if ($w.Length -lt 3 -and $w -notmatch '^\d') { continue }
        if ($StopWords -contains $w) { continue }
        [void]$set.Add($w)
    }
    return , $set
}

# =============================================================================
#  1. TELECHARGEMENT
# =============================================================================

$started = Get-Date
if (-not $Quiet) {
    Write-Host ''
    Write-Host '  MacroBrief - scan des sources en cours...' -ForegroundColor Cyan
}

$activeSources = @($Sources | Where-Object { $Lang -eq 'all' -or $_.Lang -eq $Lang })
$tickerUrls = @{}
foreach ($t in $Tickers) {
    $tickerUrls[$t.Sym] = 'https://query1.finance.yahoo.com/v8/finance/chart/' + [uri]::EscapeDataString($t.Sym) + '?range=1d&interval=15m'
}

$allUrls = @($activeSources | ForEach-Object { $_.Url }) + @($tickerUrls.Values) + @($CalendarUrl)
$fetched = Invoke-ParallelFetch -Urls $allUrls
Write-Log ('{0} requêtes en {1:N1} s' -f $allUrls.Count, ((Get-Date) - $started).TotalSeconds)

# =============================================================================
#  2. NEWS : parsing, filtrage, regroupement, scoring
# =============================================================================

$cutoff = [DateTimeOffset]::Now.AddHours(-$Hours)
$rawItems = New-Object System.Collections.Generic.List[object]
$sourceStatus = New-Object System.Collections.Generic.List[object]

foreach ($s in $activeSources) {
    $r = $fetched[$s.Url]
    if (-not $r.Ok) {
        $sourceStatus.Add([pscustomobject]@{ Name = $s.Name; Ok = $false; Count = 0; Error = $r.Error }); continue
    }
    try {
        $its = ConvertFrom-Feed $r.Bytes $s
        $kept = 0
        foreach ($i in $its) {
            # sans date : on garde (souvent pages officielles), sinon filtre fenêtre
            if ($i.Date -and $i.Date -lt $cutoff) { continue }
            if ($i.Date -and $i.Date -gt [DateTimeOffset]::Now.AddHours(1)) { continue }
            $rawItems.Add($i); $kept++
        }
        $sourceStatus.Add([pscustomobject]@{ Name = $s.Name; Ok = $true; Count = $kept; Error = $null })
    } catch {
        $sourceStatus.Add([pscustomobject]@{ Name = $s.Name; Ok = $false; Count = 0; Error = 'flux illisible: ' + $_.Exception.Message })
    }
}

# Regroupement des articles qui parlent du même sujet (similarité de Jaccard sur les mots du titre)
$sorted = $rawItems | Sort-Object -Property @{ Expression = { if ($_.Date) { $_.Date.UtcTicks } else { 0 } } } -Descending
$clusters = New-Object System.Collections.Generic.List[object]
$index = @{}
$Threshold = 0.42

foreach ($it in $sorted) {
    $tokens = Get-Tokens $it.Title
    $best = -1; $bestScore = 0.0
    if ($tokens.Count -ge 3) {
        $counts = @{}
        foreach ($tk in $tokens) {
            if ($index.ContainsKey($tk)) { foreach ($cid in $index[$tk]) { $counts[$cid] = 1 + [int]$counts[$cid] } }
        }
        foreach ($cid in $counts.Keys) {
            $inter = $counts[$cid]
            $union = $tokens.Count + $clusters[$cid].Tokens.Count - $inter
            $j = $inter / [double]$union
            if ($j -gt $bestScore) { $bestScore = $j; $best = $cid }
        }
    }
    if ($best -ge 0 -and $bestScore -ge $Threshold) {
        $c = $clusters[$best]
        # même lien / même média deux fois : on ignore le doublon
        if (-not ($c.Items | Where-Object { $_.Link -eq $it.Link -or ($_.Publisher -eq $it.Publisher -and $_.Title -eq $it.Title) })) {
            $c.Items.Add($it)
        }
    } else {
        $id = $clusters.Count
        $c = [pscustomobject]@{ Id = $id; Tokens = $tokens; Items = (New-Object System.Collections.Generic.List[object]) }
        $c.Items.Add($it)
        $clusters.Add($c)
        foreach ($tk in $tokens) {
            if (-not $index.ContainsKey($tk)) { $index[$tk] = New-Object System.Collections.Generic.List[int] }
            $index[$tk].Add($id)
        }
    }
}

# Thème + score de chaque sujet
$stories = foreach ($c in $clusters) {
    $lead = $c.Items[0]
    $text = (Remove-Diacritics ($lead.Title + ' ' + $lead.Desc)).ToLowerInvariant()
    $cats = @()
    foreach ($k in $Categories.Keys) { if ($text -match $Categories[$k]) { $cats += $k } }
    if ($cats.Count -eq 0) { continue }   # hors sujet macro/marchés : ignoré

    $publishers = @($c.Items | ForEach-Object { $_.Publisher } | Select-Object -Unique)
    $dates = @($c.Items | Where-Object { $_.Date } | ForEach-Object { $_.Date })
    $latest = $null
    if ($dates.Count) { $latest = ($dates | Sort-Object -Descending)[0] }

    $score = 3 * [math]::Min($publishers.Count, 8)
    $score += 2 * [math]::Min($cats.Count, 3)
    if ($text -match $HighImpact) { $score += 4 }
    if ($c.Items | Where-Object { $_.Feed -in @('Federal Reserve', 'BCE / ECB', 'Bank of England') }) { $score += 5 }
    if ($latest) {
        $age = ([DateTimeOffset]::Now - $latest).TotalHours
        if ($age -lt 3) { $score += 3 } elseif ($age -lt 8) { $score += 1 }
    }

    [pscustomobject]@{
        Title      = $lead.Title
        Link       = $lead.Link
        Desc       = (@($c.Items | Where-Object { $_.Desc } | Select-Object -First 1 | ForEach-Object { $_.Desc }) -join '')
        Category   = $cats[0]
        Categories = $cats
        Publishers = $publishers
        Items      = $c.Items
        Latest     = $latest
        Score      = $score
    }
}
$stories = @($stories | Sort-Object -Property @{ Expression = 'Score'; Descending = $true }, @{ Expression = { if ($_.Latest) { $_.Latest.UtcTicks } else { 0 } }; Descending = $true })
$topStories = @($stories | Select-Object -First $Top)
$topSet = @{}; foreach ($s in $topStories) { $topSet[$s.Link] = $true }

# =============================================================================
#  3. MARCHES
# =============================================================================

$markets = foreach ($t in $Tickers) {
    $r = $fetched[$tickerUrls[$t.Sym]]
    if (-not $r -or -not $r.Ok) { continue }
    try {
        $j = [Text.Encoding]::UTF8.GetString($r.Bytes) | ConvertFrom-Json
        $meta = $j.chart.result[0].meta
        $last = [double]$meta.regularMarketPrice
        $prev = 0.0
        foreach ($pn in @('chartPreviousClose', 'previousClose')) {
            if (-not $prev -and $meta.PSObject.Properties[$pn]) { $prev = [double]$meta.$pn }
        }
        if (-not $last -or -not $prev) { continue }
        $chg = $last - $prev
        [pscustomobject]@{
            Name   = $t.Name
            Kind   = $t.Kind
            Last   = $last
            Chg    = $chg
            Pct    = 100 * $chg / $prev
            Time   = [DateTimeOffset]::FromUnixTimeSeconds([long]$meta.regularMarketTime).ToLocalTime()
        }
    } catch { continue }
}
$markets = @($markets)

function Format-Price($m) {
    if ($m.Kind -eq 'rate') { return ($m.Last.ToString('N2', $French) + ' %') }
    if ($m.Kind -eq 'fx' -and $m.Last -lt 10) { return $m.Last.ToString('N4', $French) }
    return $m.Last.ToString('N2', $French)
}
function Format-Change($m) {
    if ($m.Kind -eq 'rate') { return ('{0:+0;-0;0} pb' -f [math]::Round($m.Chg * 100)) }
    return ($m.Pct.ToString('+0.00;-0.00;0.00', $French) + ' %')
}

# =============================================================================
#  4. CALENDRIER ECONOMIQUE
# =============================================================================

if ($env:LOCALAPPDATA) { $cacheDir = Join-Path $env:LOCALAPPDATA 'MacroBrief' }
else { $cacheDir = Join-Path $OutDir '.cache' }
$calCache = Join-Path $cacheDir 'calendar_thisweek.json'
$calJson = $null
$cr = $fetched[$CalendarUrl]
if ($cr -and $cr.Ok) {
    $calJson = [Text.Encoding]::UTF8.GetString($cr.Bytes)
    try { New-Item -ItemType Directory -Force -Path $cacheDir | Out-Null; [IO.File]::WriteAllText($calCache, $calJson) } catch { }
} elseif (Test-Path $calCache) {
    $calJson = [IO.File]::ReadAllText($calCache)   # ForexFactory limite les appels : on reprend le cache
}

$events = @()
if ($calJson) {
    try {
        $events = @($calJson | ConvertFrom-Json | ForEach-Object {
            $d = ConvertTo-DateOffset ([string]$_.date)
            if (-not $d) { return }
            if ($CalendarCurrencies -notcontains $_.country) { return }
            [pscustomobject]@{
                Date     = $d.ToLocalTime()
                Currency = [string]$_.country
                Title    = [string]$_.title
                Impact   = [string]$_.impact
                Forecast = [string]$_.forecast
                Previous = [string]$_.previous
                Actual   = $(if ($_.PSObject.Properties['actual']) { [string]$_.actual } else { '' })
            }
        } | Sort-Object Date)
    } catch { $events = @() }
}
$today = (Get-Date).Date
$todayEvents = @($events | Where-Object { $_.Date.Date -eq $today -and $_.Impact -in @('High', 'Medium', 'Holiday') })
$tomorrowEvents = @($events | Where-Object { $_.Date.Date -eq $today.AddDays(1) -and $_.Impact -in @('High', 'Medium') })
$weekHigh = @($events | Where-Object { $_.Date.Date -gt $today.AddDays(1) -and $_.Impact -eq 'High' })

# =============================================================================
#  5. AFFICHAGE CONSOLE
# =============================================================================

function Get-ConsoleWidth {
    try { $w = $Host.UI.RawUI.WindowSize.Width; if ($w -ge 60) { return [math]::Min($w, 160) } } catch { }
    return 110
}
function Limit-Text([string]$s, [int]$max) {
    if ($s.Length -le $max) { return $s }
    return $s.Substring(0, [math]::Max(1, $max - 1)) + '…'
}
function Write-Section([string]$Title) {
    $w = Get-ConsoleWidth
    Write-Host ''
    Write-Host ('  ' + $Title.ToUpper()) -ForegroundColor Yellow
    Write-Host ('  ' + ('─' * [math]::Min($w - 4, 100))) -ForegroundColor DarkYellow
}
function Get-ImpactColor([string]$i) {
    switch ($i) { 'High' { 'Red' } 'Medium' { 'DarkYellow' } 'Holiday' { 'Magenta' } default { 'Gray' } }
}
function Get-ImpactLabel([string]$i) {
    switch ($i) { 'High' { 'FORT ' } 'Medium' { 'MOYEN' } 'Holiday' { 'FÉRIÉ' } default { 'faible' } }
}

$okCount = @($sourceStatus | Where-Object Ok).Count
$failed = @($sourceStatus | Where-Object { -not $_.Ok })

if (-not $Quiet) {
    $w = Get-ConsoleWidth
    $dateLabel = (Get-Date).ToString('dddd d MMMM yyyy - HH:mm', $French)
    Write-Host ''
    Write-Host ('  ╔' + ('═' * ([math]::Min($w, 104) - 6)) + '╗') -ForegroundColor Cyan
    Write-Host ('  ║  MACRO BRIEF  ·  ' + $dateLabel).PadRight([math]::Min($w, 104) - 3) -ForegroundColor Cyan -NoNewline
    Write-Host '║' -ForegroundColor Cyan
    Write-Host ('  ╚' + ('═' * ([math]::Min($w, 104) - 6)) + '╝') -ForegroundColor Cyan
    Write-Host ('  {0}/{1} sources OK · {2} articles scannés · {3} sujets retenus (fenêtre {4} h)' -f $okCount, $activeSources.Count, $rawItems.Count, $stories.Count, $Hours) -ForegroundColor DarkGray

    # --- Marchés ---
    Write-Section 'Cotations'
    if ($markets.Count -eq 0) {
        Write-Host '  (données de marché indisponibles)' -ForegroundColor DarkGray
    } else {
        $col = 0
        $perRow = [math]::Max(1, [math]::Floor(($w - 2) / 38))
        foreach ($m in $markets) {
            $color = 'Gray'
            if ($m.Chg -gt 0) { $color = 'Green' } elseif ($m.Chg -lt 0) { $color = 'Red' }
            if ($m.Name -eq 'VIX') { if ($m.Chg -gt 0) { $color = 'Red' } elseif ($m.Chg -lt 0) { $color = 'Green' } }
            Write-Host ('  ' + $m.Name.PadRight(15)) -NoNewline -ForegroundColor White
            Write-Host ((Format-Price $m).PadLeft(11)) -NoNewline -ForegroundColor Gray
            Write-Host ((Format-Change $m).PadLeft(10)) -NoNewline -ForegroundColor $color
            $col++
            if ($col -ge $perRow) { Write-Host ''; $col = 0 }
        }
        if ($col -ne 0) { Write-Host '' }
    }

    # --- Agenda ---
    Write-Section ("Agenda macro aujourd'hui (heure locale)")
    if (-not $calJson) {
        Write-Host '  (calendrier indisponible)' -ForegroundColor DarkGray
    } elseif ($todayEvents.Count -eq 0) {
        Write-Host '  Aucune publication importante aujourd''hui.' -ForegroundColor DarkGray
    } else {
        foreach ($e in $todayEvents) {
            Write-Host ('  ' + $e.Date.ToString('HH:mm') + '  ') -NoNewline -ForegroundColor White
            Write-Host ((Get-ImpactLabel $e.Impact) + '  ') -NoNewline -ForegroundColor (Get-ImpactColor $e.Impact)
            Write-Host ($e.Currency + '  ') -NoNewline -ForegroundColor Cyan
            Write-Host (Limit-Text $e.Title 48).PadRight(49) -NoNewline -ForegroundColor White
            $detail = @()
            if ($e.Actual) { $detail += "réel $($e.Actual)" }
            if ($e.Forecast) { $detail += "prév. $($e.Forecast)" }
            if ($e.Previous) { $detail += "préc. $($e.Previous)" }
            Write-Host ($detail -join ' · ') -ForegroundColor DarkGray
        }
    }
    if ($tomorrowEvents.Count -gt 0) {
        Write-Host ''
        Write-Host '  Demain :' -ForegroundColor DarkCyan
        foreach ($e in ($tomorrowEvents | Where-Object Impact -eq 'High')) {
            Write-Host ('  ' + $e.Date.ToString('HH:mm') + '  ' + $e.Currency + '  ' + $e.Title) -ForegroundColor Gray
        }
    }
    if ($weekHigh.Count -gt 0) {
        Write-Host ''
        Write-Host '  Plus tard cette semaine (impact fort) :' -ForegroundColor DarkCyan
        foreach ($e in $weekHigh) {
            Write-Host ('  ' + $e.Date.ToString('ddd dd HH:mm', $French).PadRight(16) + $e.Currency + '  ' + $e.Title) -ForegroundColor Gray
        }
    }

    # --- News ---
    function Write-Story($s, [int]$n) {
        $w = Get-ConsoleWidth
        $badge = ''
        if ($s.Publishers.Count -gt 1) { $badge = " [$($s.Publishers.Count) sources]" }
        $ago = ''
        if ($s.Latest) { $ago = Format-Ago $s.Latest }
        $prefix = ('  {0,2}. ' -f $n)
        $max = $w - $prefix.Length - $badge.Length - 4
        Write-Host $prefix -NoNewline -ForegroundColor DarkGray
        Write-Host (Limit-Text $s.Title $max) -NoNewline -ForegroundColor White
        Write-Host $badge -ForegroundColor Green
        $pubs = Limit-Text (($s.Publishers | Select-Object -First 5) -join ', ') ($w - 30)
        Write-Host ('      ' + $pubs + ' · ' + $ago) -ForegroundColor DarkGray
    }

    Write-Section 'À la une'
    if ($topStories.Count -eq 0) { Write-Host '  (aucune news récupérée)' -ForegroundColor DarkGray }
    $n = 1
    foreach ($s in $topStories) { Write-Story $s $n; $n++ }

    foreach ($cat in $Categories.Keys) {
        $list = @($stories | Where-Object { $_.Category -eq $cat -and -not $topSet.ContainsKey($_.Link) } | Select-Object -First $PerCategory)
        if ($list.Count -eq 0) { continue }
        Write-Section $cat
        $n = 1
        foreach ($s in $list) { Write-Story $s $n; $n++ }
    }

    if ($failed.Count -gt 0) {
        Write-Host ''
        Write-Host ('  Sources indisponibles : ' + (($failed | ForEach-Object { $_.Name }) -join ', ')) -ForegroundColor DarkGray
    }
}

# =============================================================================
#  6. RAPPORT HTML
# =============================================================================

New-Item -ItemType Directory -Force -Path $OutDir | Out-Null
$reportPath = Join-Path $OutDir ('MacroBrief_{0}.html' -f (Get-Date -Format 'yyyy-MM-dd'))

$sb = New-Object System.Text.StringBuilder
function Add-Html([string]$s) { [void]$sb.AppendLine($s) }

function Get-StoryHtml($s) {
    $title = Get-Html $s.Title
    $href = Get-Html $s.Link
    $ago = ''
    if ($s.Latest) { $ago = Get-Html (Format-Ago $s.Latest) }
    $chips = ($s.Items | Group-Object Publisher | ForEach-Object {
        '<a class="chip" href="{0}" target="_blank" rel="noopener">{1}</a>' -f (Get-Html $_.Group[0].Link), (Get-Html $_.Name)
    }) -join ''
    $count = ''
    if ($s.Publishers.Count -gt 1) { $count = '<span class="count">{0} sources</span>' -f $s.Publishers.Count }
    $desc = ''
    if ($s.Desc) { $desc = '<p class="desc">{0}</p>' -f (Get-Html $s.Desc) }
    $tags = ($s.Categories | ForEach-Object { '<span class="tag">{0}</span>' -f (Get-Html $_) }) -join ''
    return @"
<article class="story">
  <div class="meta">$tags $count<span class="ago">$ago</span></div>
  <h3><a href="$href" target="_blank" rel="noopener">$title</a></h3>
  $desc
  <div class="chips">$chips</div>
</article>
"@
}

$dateTitle = (Get-Date).ToString('dddd d MMMM yyyy', $French)
$genTime = (Get-Date).ToString('HH:mm')

Add-Html @"
<!doctype html>
<html lang="fr"><head><meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<title>Macro Brief - $dateTitle</title>
<style>
:root{--bg:#f6f7f9;--card:#fff;--fg:#14171c;--muted:#5d6672;--line:#e3e6ea;--accent:#2457d6;--up:#0f8a4b;--down:#c8322d;--chip:#eef1f6;--high:#c8322d;--med:#c27c0e}
@media (prefers-color-scheme:dark){:root{--bg:#0f1216;--card:#171b21;--fg:#e8ebef;--muted:#949ca7;--line:#262c34;--accent:#7aa2ff;--up:#3ccf7f;--down:#ff6b63;--chip:#222831;--high:#ff6b63;--med:#f0b44a}}
*{box-sizing:border-box}body{margin:0;background:var(--bg);color:var(--fg);font:15px/1.5 -apple-system,Segoe UI,Roboto,Helvetica,Arial,sans-serif}
.wrap{max-width:1100px;margin:0 auto;padding:24px 16px 60px}
header h1{margin:0;font-size:26px;letter-spacing:-.02em}header .sub{color:var(--muted);font-size:13px;margin-top:4px}
h2{font-size:13px;text-transform:uppercase;letter-spacing:.08em;color:var(--muted);margin:32px 0 12px;border-bottom:1px solid var(--line);padding-bottom:6px}
.grid{display:grid;grid-template-columns:repeat(auto-fill,minmax(160px,1fr));gap:8px}
.tile{background:var(--card);border:1px solid var(--line);border-radius:10px;padding:10px 12px}
.tile .n{font-size:12px;color:var(--muted)}.tile .p{font-size:18px;font-weight:600;font-variant-numeric:tabular-nums}
.tile .c{font-size:13px;font-weight:600;font-variant-numeric:tabular-nums}.up{color:var(--up)}.down{color:var(--down)}.flat{color:var(--muted)}
table{width:100%;border-collapse:collapse;background:var(--card);border:1px solid var(--line);border-radius:10px;overflow:hidden}
td,th{padding:8px 10px;border-bottom:1px solid var(--line);text-align:left;font-size:14px;font-variant-numeric:tabular-nums}th{font-size:12px;color:var(--muted);font-weight:600}
tr:last-child td{border-bottom:0}.imp{font-size:11px;font-weight:700;padding:2px 6px;border-radius:4px;white-space:nowrap}
.imp.High{background:var(--high);color:#fff}.imp.Medium{background:var(--med);color:#fff}.imp.Holiday{background:var(--chip);color:var(--muted)}
.cols{display:grid;grid-template-columns:1fr 1fr;gap:12px}@media(max-width:760px){.cols{grid-template-columns:1fr}}
.story{background:var(--card);border:1px solid var(--line);border-radius:10px;padding:12px 14px;margin-bottom:8px}
.story h3{margin:4px 0;font-size:16px;line-height:1.35}.story h3 a{color:var(--fg);text-decoration:none}.story h3 a:hover{color:var(--accent)}
.desc{margin:4px 0 8px;color:var(--muted);font-size:14px}.meta{display:flex;flex-wrap:wrap;gap:6px;align-items:center;font-size:12px}
.tag{color:var(--accent);font-weight:600}.count{background:var(--up);color:#fff;border-radius:999px;padding:1px 8px;font-weight:600}.ago{color:var(--muted);margin-left:auto}
.chips{display:flex;flex-wrap:wrap;gap:4px}.chip{background:var(--chip);color:var(--muted);text-decoration:none;font-size:12px;padding:2px 8px;border-radius:999px}.chip:hover{color:var(--accent)}
.empty{color:var(--muted);font-style:italic}footer{margin-top:40px;color:var(--muted);font-size:12px}
nav{display:flex;flex-wrap:wrap;gap:6px;margin-top:14px}nav a{font-size:12px;color:var(--accent);text-decoration:none;background:var(--chip);padding:3px 10px;border-radius:999px}
</style></head><body><div class="wrap">
<header><h1>Macro Brief</h1><div class="sub">$(Get-Html $dateTitle) · généré à $genTime · $okCount/$($activeSources.Count) sources · $($rawItems.Count) articles scannés · $($stories.Count) sujets (dernières $Hours h)</div>
<nav><a href="#marches">Cotations</a><a href="#agenda">Agenda</a><a href="#une">À la une</a>
"@
$i = 0
foreach ($cat in $Categories.Keys) {
    $has = @($stories | Where-Object { $_.Category -eq $cat -and -not $topSet.ContainsKey($_.Link) }).Count
    if ($has) { Add-Html ('<a href="#cat{0}">{1}</a>' -f $i, (Get-Html $cat)) }
    $i++
}
Add-Html '</nav></header>'

# Marchés
Add-Html '<h2 id="marches">Cotations</h2>'
if ($markets.Count -eq 0) { Add-Html '<p class="empty">Données de marché indisponibles.</p>' }
else {
    Add-Html '<div class="grid">'
    foreach ($m in $markets) {
        $cls = 'flat'
        if ($m.Chg -gt 0) { $cls = 'up' } elseif ($m.Chg -lt 0) { $cls = 'down' }
        if ($m.Name -eq 'VIX' -and $cls -ne 'flat') { if ($cls -eq 'up') { $cls = 'down' } else { $cls = 'up' } }
        Add-Html ('<div class="tile"><div class="n">{0}</div><div class="p">{1}</div><div class="c {2}">{3}</div></div>' -f (Get-Html $m.Name), (Get-Html (Format-Price $m)), $cls, (Get-Html (Format-Change $m)))
    }
    Add-Html '</div>'
}

# Agenda
function Get-EventRows($list, [string]$fmt) {
    $rows = foreach ($e in $list) {
        $d = @(); if ($e.Actual) { $d += '<b>' + (Get-Html $e.Actual) + '</b>' }
        '<tr><td>{0}</td><td><span class="imp {1}">{2}</span></td><td><b>{3}</b></td><td>{4}</td><td>{5}</td><td>{6}</td><td>{7}</td></tr>' -f `
            (Get-Html $e.Date.ToString($fmt, $French)), $e.Impact, (Get-Html (Get-ImpactLabel $e.Impact).Trim()), (Get-Html $e.Currency), (Get-Html $e.Title), ($d -join ''), (Get-Html $e.Forecast), (Get-Html $e.Previous)
    }
    return ($rows -join "`n")
}
$thead = '<tr><th>Heure</th><th>Impact</th><th>Dev.</th><th>Publication</th><th>Réel</th><th>Prévision</th><th>Précédent</th></tr>'
Add-Html '<h2 id="agenda">Agenda macro du jour <small>(heure locale)</small></h2>'
if (-not $calJson) { Add-Html '<p class="empty">Calendrier indisponible.</p>' }
elseif ($todayEvents.Count -eq 0) { Add-Html '<p class="empty">Aucune publication importante aujourd''hui.</p>' }
else { Add-Html ("<table>$thead" + (Get-EventRows $todayEvents 'HH:mm') + '</table>') }
if ($tomorrowEvents.Count -gt 0) {
    Add-Html '<h2>Demain</h2>'
    Add-Html ("<table>$thead" + (Get-EventRows $tomorrowEvents 'HH:mm') + '</table>')
}
if ($weekHigh.Count -gt 0) {
    Add-Html '<h2>Plus tard cette semaine · impact fort</h2>'
    Add-Html ("<table>$thead" + (Get-EventRows $weekHigh 'ddd dd HH:mm') + '</table>')
}

# News
Add-Html '<h2 id="une">À la une</h2>'
if ($topStories.Count -eq 0) { Add-Html '<p class="empty">Aucune news récupérée.</p>' }
else { Add-Html '<div class="cols">'; foreach ($s in $topStories) { Add-Html (Get-StoryHtml $s) }; Add-Html '</div>' }

$i = 0
foreach ($cat in $Categories.Keys) {
    $list = @($stories | Where-Object { $_.Category -eq $cat -and -not $topSet.ContainsKey($_.Link) } | Select-Object -First ($PerCategory * 2))
    $i++
    if ($list.Count -eq 0) { continue }
    Add-Html ('<h2 id="cat{0}">{1}</h2>' -f ($i - 1), (Get-Html $cat))
    Add-Html '<div class="cols">'; foreach ($s in $list) { Add-Html (Get-StoryHtml $s) }; Add-Html '</div>'
}

# Pied de page : état des sources
$srcList = ($sourceStatus | ForEach-Object {
    if ($_.Ok) { '{0} ({1})' -f (Get-Html $_.Name), $_.Count } else { '<s title="{1}">{0}</s>' -f (Get-Html $_.Name), (Get-Html $_.Error) }
}) -join ' · '
Add-Html "<footer><p><b>Sources :</b> $srcList</p><p>Marchés : Yahoo Finance (différé). Calendrier : ForexFactory. Les regroupements de sujets sont automatiques.</p></footer>"
Add-Html '</div></body></html>'

[IO.File]::WriteAllText($reportPath, $sb.ToString(), (New-Object System.Text.UTF8Encoding($false)))

if (-not $Quiet) {
    Write-Host ''
    Write-Host ('  Rapport HTML : ' + $reportPath) -ForegroundColor Cyan
    Write-Host ('  Terminé en {0:N1} s' -f ((Get-Date) - $started).TotalSeconds) -ForegroundColor DarkGray
    Write-Host ''
}
if (-not $NoBrowser) {
    try { Start-Process $reportPath } catch { Write-Log "Impossible d'ouvrir le navigateur : $($_.Exception.Message)" }
}
