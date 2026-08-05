/**
 * =============================================================
 *  Počasí Tour – denní log předpovědi pro 13.–16.8.2026
 * =============================================================
 *
 *  INSTALACE:
 *  1. Otevři spreadsheet "Počasí Tour"
 *  2. Extensions  ▸  Apps Script
 *  3. Smaž vše a vlož tento kód
 *  4. Ulož (Ctrl+S)
 *  5. Spusť ručně: vyber fetchWeatherAndLog ▸ Run
 *     (při prvním spuštění potvrdíš oprávnění)
 *  6. Nastav triggery: vyber setupTriggers ▸ Run.
 *     Založí dva denní běhy — ráno (~6:00) a večer (~18:00) —
 *     a smaže případné staré triggery, ať se nehromadí.
 *     Ruční nastavení přes UI už není potřeba.
 *
 *  Apps Script u denního triggeru volí náhodnou minutu v dané hodině,
 *  takže "6:05" je jen kdy to zrovna vyšlo, ne přesný čas. Grafy vývoje
 *  tím dostanou dva body denně místo jednoho.
 *
 *  Zdroj dat: open-meteo.com (zdarma, bez API klíče)
 *    – Forecast API   ≤ 16 dní předem: deterministická předpověď
 *    – Ensemble API   ≤ 35 dní předem: GEFS 0,5° (31 členů)
 *                                    + ECMWF IFS 0,25° (51 členů)
 *    – Seasonal API   > 35 dní předem: ECMWF SEAS5
 *    – Climate API    poslední záchrana: CMIP6
 *
 *  Ensemble je zdroj pravděpodobností: každý člen je jeden možný vývoj
 *  počasí, takže "P(déšť) = 42 %" znamená, že déšť vyšel u 42 % běhů.
 *  Není to odhad ani heuristika.
 *
 *  POZOR NA MÍCHÁNÍ ZDROJŮ:
 *  Sloupce A–K jsou beze změny z původní verze a drží deterministickou
 *  (resp. sezónní) předpověď. Všechno z ensemblu má prefix ens_ a vlastní
 *  sloupec. Kdyby jeden sloupec střídal zdroje podle toho, který je zrovna
 *  dostupný, vznikl by v grafu vývoje skok způsobený přepnutím modelu,
 *  ne změnou počasí.
 *
 *  Sloupce v sheetu (script zapíše hlavičku automaticky):
 *    A  datum_zaznamu
 *    B  predikovany_den
 *    C  dny_predem
 *    D  teplota_C                 ← deterministický / sezónní model
 *    E  popis_pocasi
 *    F  srazky_mm
 *    G  srazky_pravdepodobnost_%  ← precipitation_probability_max
 *    H  spolehlivost_%            ← tabulková hodnota dle dní předem
 *    I  rozptyl_teploty_C         ← směrodatná odchylka (sezónní model)
 *    J  typ_udaje
 *    K  zdroj
 *    ── nové, vše z Ensemble API ──
 *    L  realfeel_max_C            ← deterministický model
 *    M  realfeel_min_C            ← deterministický model
 *    N  ens_teplota_med_C         ← medián členů
 *    O  ens_teplota_p10_C
 *    P  ens_teplota_p90_C
 *    Q  ens_realfeel_max_C
 *    R  ens_realfeel_min_C
 *    S  ens_srazky_med_mm
 *    T  ens_p_dest_05_%           ← P(déšť > 0,5 mm)
 *    U  ens_p_dest_5mm_%          ← P(déšť > 5 mm)
 *    V  ens_p_horko_28_%          ← P(max > 28 °C)
 *    W  ens_p_chladno_20_%        ← P(max < 20 °C)
 *    X  ens_p_bourka_%            ← P(CAPE > 300 + srážky)
 *    Y  ens_clenu
 *    Z  ens_jistota_%             ← spočtená z šířky p10–p90
 * =============================================================
 */

// ── Konfigurace ────────────────────────────────────────────────

var CONFIG = {
  sheetId:   '1TBvQe48oZyu2iFChZWax8RtgC_aXTZ8RW9IMUCFXYcY',
  sheetName: '2026',
  // Musí sedět s FORECAST_LAT/FORECAST_LON v pocasi.html, jinak
  // dashboard počítá pravděpodobnosti pro jiné místo než tenhle log.
  lat:       50.7052,       // Jablonec nad Nisou / Kokonín
  lon:       15.1863,
  targets:   ['2026-08-13', '2026-08-14', '2026-08-15', '2026-08-16']
};

// Prahy jevů — musí být stejné jako EVENTS v pocasi.html, jinak by
// tabulka na webu ukazovala jiná čísla než historie v sheetu.
var THRESHOLDS = {
  rainAny:   0.5,   // mm/den
  rainHeavy: 5,     // mm/den
  hot:       28,    // °C
  cold:      20,    // °C
  capeStorm: 300,   // J/kg
  capeRain:  1      // mm/den, musí platit současně s CAPE
};

var HEADERS = [
  'datum_zaznamu', 'predikovany_den', 'dny_predem',
  'teplota_C', 'popis_pocasi', 'srazky_mm',
  'srazky_pravdepodobnost_%', 'spolehlivost_%', 'rozptyl_teploty_C',
  'typ_udaje', 'zdroj',
  'realfeel_max_C', 'realfeel_min_C',
  'ens_teplota_med_C', 'ens_teplota_p10_C', 'ens_teplota_p90_C',
  'ens_realfeel_max_C', 'ens_realfeel_min_C', 'ens_srazky_med_mm',
  'ens_p_dest_05_%', 'ens_p_dest_5mm_%', 'ens_p_horko_28_%',
  'ens_p_chladno_20_%', 'ens_p_bourka_%',
  'ens_clenu', 'ens_jistota_%'
];

// ── Hlavní funkce ──────────────────────────────────────────────

function fetchWeatherAndLog() {
  var today    = new Date();
  var todayISO = _fmtISO(today);
  var nowCZ    = Utilities.formatDate(today, 'Europe/Prague', 'dd.MM.yyyy HH:mm');
  var todayMid = new Date(todayISO + 'T00:00:00');

  var lastTarget = new Date(CONFIG.targets[CONFIG.targets.length - 1] + 'T00:00:00');
  if (todayMid > lastTarget) {
    Logger.log('Všechny cílové dny už proběhly – ukončuji.');
    return;
  }

  // Deterministická předpověď pokrývá jen část dnů (dosah 16 dní).
  // Dřív se při jakémkoli nepokrytém dni zahodila celá odpověď a všechny
  // čtyři dny spadly na sezónní model — i ty, pro které už byla přesná
  // předpověď k dispozici. Teď se doplňuje jen to, co chybí.
  var byDate = {};
  var fc = _fetchForecast();
  if (fc) {
    for (var d in fc.byDate) byDate[d] = fc.byDate[d];
  }

  var missing = CONFIG.targets.filter(function(d) { return !byDate[d]; });
  if (missing.length > 0) {
    var fb = _fetchSeasonal() || _fetchClimate();
    if (fb) {
      missing.forEach(function(d) {
        if (fb.byDate[d]) byDate[d] = fb.byDate[d];
      });
    }
  }

  if (Object.keys(byDate).length === 0) {
    Logger.log('Nepodařilo se získat data z žádného API.');
    return;
  }

  var ens   = _fetchEnsemble()      || {};
  var storm = _fetchStormEnsemble() || {};

  var sheet = SpreadsheetApp.openById(CONFIG.sheetId)
                            .getSheetByName(CONFIG.sheetName);
  _ensureHeaders(sheet);

  var rows = [];

  CONFIG.targets.forEach(function(date) {
    var d = byDate[date];
    if (!d) return;

    var e         = ens[date] || null;
    var tgtDate   = new Date(date + 'T00:00:00');
    var dnyPredem = Math.round((tgtDate - todayMid) / 86400000);
    var predDen   = date.split('-').reverse().join('.');

    rows.push([
      nowCZ,
      predDen,
      dnyPredem,
      d.tempMax,
      d.popis,
      d.srazky,
      _pctRaw(d.srazkyPravdep),
      _spolehlivost(dnyPredem),
      d.rozptyl != null ? d.rozptyl : '',
      d.typUdaje,          // zdroj se liší den od dne, proto je uložený u dne
      d.zdroj,
      _r1n(d.rfMax),
      _r1n(d.rfMin),
      e ? _r1n(e.med)        : '',
      e ? _r1n(e.p10)        : '',
      e ? _r1n(e.p90)        : '',
      e ? _r1n(e.rfMaxMed)   : '',
      e ? _r1n(e.rfMinMed)   : '',
      e ? _r1n(e.precipMed)  : '',
      e ? _pctFromShare(e.pRainAny)   : '',
      e ? _pctFromShare(e.pRainHeavy) : '',
      e ? _pctFromShare(e.pHot)       : '',
      e ? _pctFromShare(e.pCold)      : '',
      storm[date] != null ? _pctFromShare(storm[date]) : '',
      e ? e.n : '',
      e ? _jistota(e) : ''
    ]);
  });

  if (rows.length === 0) {
    Logger.log('Žádné řádky k zápisu.');
    return;
  }

  // Jeden zápis místo appendRow v cyklu — rychlejší a atomičtější
  sheet.getRange(sheet.getLastRow() + 1, 1, rows.length, HEADERS.length)
       .setValues(rows);

  Logger.log('Zapsáno ' + rows.length + ' řádků.');
}

// ── Nastavení triggerů (spusť jednou ručně) ────────────────────
// Založí dva denní běhy — ráno a večer. Idempotentní: nejdřív smaže
// všechny existující triggery pro fetchWeatherAndLog, takže opakované
// spuštění nevede k duplicitním zápisům. Hodiny se mění tady na jednom
// místě, ať nemusíš klikat v UI.

var TRIGGER_HOURS = [6, 18];   // Europe/Prague

function setupTriggers() {
  var existing = ScriptApp.getProjectTriggers();
  var smazano = 0;
  existing.forEach(function(t) {
    if (t.getHandlerFunction() === 'fetchWeatherAndLog') {
      ScriptApp.deleteTrigger(t);
      smazano++;
    }
  });

  TRIGGER_HOURS.forEach(function(h) {
    ScriptApp.newTrigger('fetchWeatherAndLog')
      .timeBased()
      .atHour(h)
      .nearMinute(5)          // cílí k :05, Apps Script si nechává ±15 min
      .everyDays(1)
      .inTimezone('Europe/Prague')
      .create();
  });

  Logger.log('Smazáno ' + smazano + ' starých triggerů, založeno ' +
             TRIGGER_HOURS.length + ' nových: ~' +
             TRIGGER_HOURS.join(':00, ~') + ':00 Europe/Prague.');
}

// ── Forecast API (deterministický model, ≤ 16 dní) ────────────

function _fetchForecast() {
  // Pevné start_date/end_date API odmítne s HTTP 400, dokud není celé
  // okno v dosahu (povoluje jen dnes+15) — a padal tím celý požadavek
  // včetně dnů, které v dosahu byly. Bereme proto 16 dní a filtrujeme.
  var url = 'https://api.open-meteo.com/v1/forecast'
    + '?latitude='  + CONFIG.lat
    + '&longitude=' + CONFIG.lon
    + '&daily=temperature_2m_max,apparent_temperature_max,apparent_temperature_min,'
    + 'precipitation_sum,precipitation_probability_max,weather_code'
    + '&timezone=Europe/Prague'
    + '&forecast_days=16';

  try {
    var json  = _getJSON(url);
    var daily = json.daily;
    if (!daily || !daily.time) return null;

    var byDate = {};
    for (var i = 0; i < daily.time.length; i++) {
      var date = daily.time[i];
      if (CONFIG.targets.indexOf(date) === -1) continue;
      if (daily.temperature_2m_max[i] == null) continue;

      byDate[date] = {
        tempMax:       _r1(daily.temperature_2m_max[i]),
        rfMax:         _val(daily.apparent_temperature_max, i),
        rfMin:         _val(daily.apparent_temperature_min, i),
        popis:         _wmoToCS(daily.weather_code[i]),
        srazky:        _r1(daily.precipitation_sum[i]),
        srazkyPravdep: _val(daily.precipitation_probability_max, i),
        rozptyl:       null,
        typUdaje:      'predikce',
        zdroj:         'open-meteo.com'
      };
    }

    return Object.keys(byDate).length > 0 ? { byDate: byDate } : null;

  } catch (e) {
    Logger.log('Forecast API selhalo: ' + e.message);
    return null;
  }
}

// ── Ensemble API (pravděpodobnosti, ≤ 35 dní) ─────────────────

function _fetchEnsemble() {
  var url = 'https://ensemble-api.open-meteo.com/v1/ensemble'
    + '?latitude='  + CONFIG.lat
    + '&longitude=' + CONFIG.lon
    + '&daily=temperature_2m_max,apparent_temperature_max,apparent_temperature_min,'
    + 'precipitation_sum'
    + '&models=gfs05,ecmwf_ifs025'
    + '&timezone=Europe/Prague'
    + '&start_date=' + CONFIG.targets[0]
    + '&end_date='   + CONFIG.targets[CONFIG.targets.length - 1];

  try {
    var json  = _getJSON(url);
    var daily = json.daily;
    if (!daily || !daily.time) return null;

    var kT  = _memberKeys(daily, 'temperature_2m_max');
    var kRF = _memberKeys(daily, 'apparent_temperature_max');
    var kRN = _memberKeys(daily, 'apparent_temperature_min');
    var kP  = _memberKeys(daily, 'precipitation_sum');

    var out = {};

    for (var i = 0; i < daily.time.length; i++) {
      var date = daily.time[i];
      if (CONFIG.targets.indexOf(date) === -1) continue;

      var temps = _collect(daily, kT, i);
      if (temps.length === 0) continue;

      var rfMax = _collect(daily, kRF, i);
      var rfMin = _collect(daily, kRN, i);
      var prcp  = _collect(daily, kP,  i);

      temps.sort(_num);

      out[date] = {
        n:          temps.length,
        med:        _pctl(temps, 0.50),
        p10:        _pctl(temps, 0.10),
        p90:        _pctl(temps, 0.90),
        rfMaxMed:   rfMax.length ? _median(rfMax) : null,
        rfMinMed:   rfMin.length ? _median(rfMin) : null,
        precipMed:  prcp.length  ? _median(prcp)  : null,
        pRainAny:   _share(prcp,  function(v) { return v > THRESHOLDS.rainAny; }),
        pRainHeavy: _share(prcp,  function(v) { return v > THRESHOLDS.rainHeavy; }),
        pHot:       _share(temps, function(v) { return v > THRESHOLDS.hot; }),
        pCold:      _share(temps, function(v) { return v < THRESHOLDS.cold; })
      };
    }

    return Object.keys(out).length > 0 ? out : null;

  } catch (e) {
    Logger.log('Ensemble API selhalo: ' + e.message);
    return null;
  }
}

// ── Bouřky z hodinového ensemblu (CAPE + srážky) ──────────────
// Na 0,5° mřížce a dva týdny dopředu je CAPE silně vyhlazená, takže
// tohle je orientační signál, ne poctivá pravděpodobnost bouřky.

function _fetchStormEnsemble() {
  var url = 'https://ensemble-api.open-meteo.com/v1/ensemble'
    + '?latitude='  + CONFIG.lat
    + '&longitude=' + CONFIG.lon
    + '&hourly=cape,precipitation'
    + '&models=gfs05'
    + '&timezone=Europe/Prague'
    + '&start_date=' + CONFIG.targets[0]
    + '&end_date='   + CONFIG.targets[CONFIG.targets.length - 1];

  try {
    var json = _getJSON(url);
    var h    = json.hourly;
    if (!h || !h.time) return null;

    var kC = _memberKeys(h, 'cape');
    var kP = _memberKeys(h, 'precipitation');
    var out = {};

    CONFIG.targets.forEach(function(date) {
      var idx = [];
      for (var i = 0; i < h.time.length; i++) {
        if (h.time[i].indexOf(date) === 0) idx.push(i);
      }
      if (idx.length === 0) return;

      var hit = 0, tot = 0;
      for (var m = 0; m < kC.length; m++) {
        var ck = kC[m], pk = kP[m];
        if (!ck || !pk) continue;

        var maxCape = 0, sumP = 0, any = false;
        for (var j = 0; j < idx.length; j++) {
          var c = h[ck][idx[j]], p = h[pk][idx[j]];
          if (c != null) { if (c > maxCape) maxCape = c; any = true; }
          if (p != null) sumP += p;
        }
        if (!any) continue;

        tot++;
        if (maxCape > THRESHOLDS.capeStorm && sumP > THRESHOLDS.capeRain) hit++;
      }
      if (tot > 0) out[date] = hit / tot;
    });

    return Object.keys(out).length > 0 ? out : null;

  } catch (e) {
    Logger.log('Storm ensemble selhalo: ' + e.message);
    return null;
  }
}

// ── Seasonal API (> 35 dní, až 7 měsíců) ──────────────────────

function _fetchSeasonal() {
  var url = 'https://seasonal-api.open-meteo.com/v1/seasonal'
    + '?latitude='  + CONFIG.lat
    + '&longitude=' + CONFIG.lon
    + '&daily=temperature_2m_max,precipitation_sum'
    + '&timezone=Europe/Prague'
    + '&start_date=' + CONFIG.targets[0]
    + '&end_date='   + CONFIG.targets[CONFIG.targets.length - 1];

  try {
    var json  = _getJSON(url);
    var daily = json.daily;
    if (!daily || !daily.time) return null;

    var byDate = {};
    for (var i = 0; i < daily.time.length; i++) {
      var date = daily.time[i];
      if (CONFIG.targets.indexOf(date) === -1) continue;

      var tempMean = _extractValue(daily, 'temperature_2m_max', i);
      var tempSD   = _extractSD(daily, 'temperature_2m_max', i);
      var prcp     = _extractValue(daily, 'precipitation_sum', i);
      if (tempMean == null) continue;

      byDate[date] = {
        tempMax:       _r1(tempMean),
        rfMax:         null,
        rfMin:         null,
        popis:         _precipToCS(prcp),
        srazky:        _r1(prcp),
        srazkyPravdep: null,
        rozptyl:       tempSD != null ? _r1(tempSD) : null,
        typUdaje:      'klimatologie',
        zdroj:         'open-meteo.com (seasonal)'
      };
    }

    return Object.keys(byDate).length > 0 ? { byDate: byDate } : null;

  } catch (e) {
    Logger.log('Seasonal API selhalo: ' + e.message);
    return null;
  }
}

// ── Climate API (fallback – CMIP6 klimatický model) ────────────

function _fetchClimate() {
  var url = 'https://climate-api.open-meteo.com/v1/climate'
    + '?latitude='  + CONFIG.lat
    + '&longitude=' + CONFIG.lon
    + '&daily=temperature_2m_max,precipitation_sum'
    + '&start_date=' + CONFIG.targets[0]
    + '&end_date='   + CONFIG.targets[CONFIG.targets.length - 1]
    + '&models=EC_Earth3P_HR';

  try {
    var json  = _getJSON(url);
    var daily = json.daily;
    if (!daily || !daily.time) return null;

    var byDate = {};
    for (var i = 0; i < daily.time.length; i++) {
      var date = daily.time[i];
      if (CONFIG.targets.indexOf(date) === -1) continue;
      if (daily.temperature_2m_max[i] == null) continue;

      byDate[date] = {
        tempMax:       _r1(daily.temperature_2m_max[i]),
        rfMax:         null,
        rfMin:         null,
        popis:         _precipToCS(daily.precipitation_sum[i]),
        srazky:        _r1(daily.precipitation_sum[i]),
        srazkyPravdep: null,
        rozptyl:       null,
        typUdaje:      'klimatologie',
        zdroj:         'open-meteo.com (climate)'
      };
    }

    return Object.keys(byDate).length > 0 ? { byDate: byDate } : null;

  } catch (e) {
    Logger.log('Climate API selhalo: ' + e.message);
    return null;
  }
}

// ── Jistota z rozptylu ensemblu ────────────────────────────────
// Stejný vzorec jako ensConfidence() v pocasi.html: p10–p90 širší
// než 12 °C = model si nevěří, do 2 °C = velmi vysoká shoda.

function _jistota(e) {
  if (!e || e.p10 == null || e.p90 == null) return '';
  var w = e.p90 - e.p10;
  var c = 100 - (w - 2) * 7.5;
  return Math.max(10, Math.min(95, Math.round(c)));
}

// ── Spolehlivost předpovědi (lookup dle počtu dní předem) ──────

function _spolehlivost(dny) {
  if (dny <= 1)  return 95;
  if (dny <= 3)  return 90;
  if (dny <= 5)  return 80;
  if (dny <= 7)  return 70;
  if (dny <= 10) return 55;
  if (dny <= 14) return 40;
  if (dny <= 21) return 30;
  if (dny <= 30) return 20;
  if (dny <= 46) return 10;
  return 5;
}

// ── WMO kód → český popis ──────────────────────────────────────

function _wmoToCS(code) {
  var m = {
    0:  'Jasno',
    1:  'Skoro jasno',
    2:  'Polojasno',
    3:  'Oblačno',
    45: 'Zataženo',
    48: 'Zataženo',
    51: 'Déšť',
    53: 'Déšť',
    55: 'Déšť',
    56: 'Déšť',
    57: 'Déšť',
    61: 'Déšť',
    63: 'Déšť',
    65: 'Déšť',
    66: 'Déšť',
    67: 'Déšť',
    80: 'Přeháňky',
    81: 'Přeháňky',
    82: 'Přeháňky',
    95: 'Bouřky',
    96: 'Přeháňky a bouřky',
    99: 'Přeháňky a bouřky'
  };
  return m[code] || 'Oblačno';
}

function _precipToCS(mm) {
  if (mm == null || mm <= 0.1) return 'Polojasno';
  if (mm <= 2)  return 'Přeháňky';
  return 'Déšť';
}

// ── Statistika ─────────────────────────────────────────────────

function _num(a, b) { return a - b; }

// Klíče členů ensemblu. Open-Meteo vrací "temperature_2m_max_ncep_gefs05"
// (řídící běh) i "temperature_2m_max_member07_ncep_gefs05" — oba jsou
// platné realizace. Kontrola na '_' brání tomu, aby se do
// temperature_2m_max připletl jiný klíč se stejným začátkem.
function _memberKeys(daily, prefix) {
  var out = [];
  for (var k in daily) {
    if (k === 'time' || k.indexOf(prefix) !== 0) continue;
    var rest = k.slice(prefix.length);
    if (rest === '' || rest.charAt(0) === '_') out.push(k);
  }
  return out;
}

function _collect(daily, keys, i) {
  var vals = [];
  for (var m = 0; m < keys.length; m++) {
    var v = daily[keys[m]][i];
    if (v != null) vals.push(v);
  }
  return vals;
}

function _pctl(sorted, q) {
  if (!sorted.length) return null;
  if (sorted.length === 1) return sorted[0];
  var pos = (sorted.length - 1) * q;
  var lo  = Math.floor(pos), hi = Math.ceil(pos);
  if (lo === hi) return sorted[lo];
  return sorted[lo] + (sorted[hi] - sorted[lo]) * (pos - lo);
}

function _median(arr) {
  return _pctl(arr.slice().sort(_num), 0.50);
}

function _sd(vals) {
  if (vals.length < 2) return null;
  var mean = 0, i;
  for (i = 0; i < vals.length; i++) mean += vals[i];
  mean /= vals.length;
  var sq = 0;
  for (i = 0; i < vals.length; i++) sq += (vals[i] - mean) * (vals[i] - mean);
  return Math.sqrt(sq / vals.length);
}

// Podíl členů, u kterých platí podmínka → skutečná pravděpodobnost
function _share(vals, pred) {
  if (!vals.length) return null;
  var hit = 0;
  for (var i = 0; i < vals.length; i++) if (pred(vals[i])) hit++;
  return hit / vals.length;
}

// ── Pomocné funkce ─────────────────────────────────────────────

function _getJSON(url) {
  var resp = UrlFetchApp.fetch(url, { muteHttpExceptions: true });
  var code = resp.getResponseCode();
  if (code !== 200) {
    throw new Error('HTTP ' + code + ': ' + resp.getContentText().substring(0, 200));
  }
  return JSON.parse(resp.getContentText());
}

// Seasonal API vrací ensemble členy (temperature_2m_max_member01, …).
// Vrátí průměr přes všechny členy.
function _extractValue(daily, key, index) {
  if (daily[key] && typeof daily[key][index] === 'number') {
    return daily[key][index];
  }

  var sum = 0, count = 0;
  for (var k in daily) {
    if (k.indexOf(key + '_member') === 0) {
      var val = daily[k][index];
      if (val != null) { sum += val; count++; }
    }
  }
  return count > 0 ? sum / count : null;
}

// Směrodatná odchylka teploty přes ensemble členy.
// Vysoká = velká nejistota, nízká = modely se shodují.
function _extractSD(daily, key, index) {
  var vals = [];
  for (var k in daily) {
    if (k.indexOf(key + '_member') === 0) {
      var val = daily[k][index];
      if (val != null) vals.push(val);
    }
  }
  return _sd(vals);
}

function _val(arr, i) {
  return (arr && arr[i] != null) ? arr[i] : null;
}

// _r1 vrací 0 pro null (zpětná kompatibilita sloupce srazky_mm),
// _r1n nechá prázdno — chybějící Real Feel nesmí vypadat jako 0 °C.
function _r1(v) {
  if (v == null) return 0;
  return Math.round(v * 10) / 10;
}

function _r1n(v) {
  if (v == null) return '';
  return Math.round(v * 10) / 10;
}

// Podíl 0–1 → procenta. Oddělené od _pctRaw schválně: hodnota 1 znamená
// u podílu 100 %, ale u Forecast API 1 %, a uhádnout to z rozsahu nejde.
function _pctFromShare(v) {
  if (v == null) return '';
  return Math.round(v * 100);
}

// Hodnota, která už v procentech přišla (precipitation_probability_max)
function _pctRaw(v) {
  if (v == null) return '';
  return Math.round(v);
}

function _fmtISO(d) {
  return Utilities.formatDate(d, 'Europe/Prague', 'yyyy-MM-dd');
}

// Přepíše hlavičku na řádku 1, pokud neodpovídá novému formátu.
function _ensureHeaders(sheet) {
  var width = Math.max(sheet.getLastColumn(), HEADERS.length);
  var first = sheet.getRange(1, 1, 1, width).getValues()[0];

  var ok = true;
  for (var i = 0; i < HEADERS.length; i++) {
    if (first[i] !== HEADERS[i]) { ok = false; break; }
  }
  if (ok) return;

  sheet.getRange(1, 1, 1, HEADERS.length).setValues([HEADERS]);
}
