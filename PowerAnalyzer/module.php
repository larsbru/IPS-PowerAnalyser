<?php
declare(strict_types=1);

/**
 * IPS-PowerAnalyzer - module.php
 * Vollständige, stabile Fassung (Konfiguration im WebFront, Timer, Heatmap, Kurzfazit).
 *
 * Voraussetzungen:
 *  - IP-Symcon 6.x+
 *  - Archivsteuerung (Archive Control), Standard-ID 12496 (konfigurierbar)
 *
 * Struktur:
 *  - Instanz-Root:
 *      * Kurzfazit            (String ~HTMLBox)
 *      * Heatmap_Monatsanteile_je_Dezil (String ~HTMLBox)
 *  - Unterordner "Konfiguration":
 *      * PowerVariableID          (Integer, beschreibbar)
 *      * ArchiveControlID         (Integer, beschreibbar; Default 12496)
 *      * YearsToAnalyze           (Integer, beschreibbar; Default 5)
 *      * UpdateIntervalSeconds    (Integer, beschreibbar; Default 300)
 *      * EnableAutoUpdate         (Boolean, Button)
 *      * EnableDebug              (Boolean, Button)
 */

class PowerAnalyzer extends IPSModule
{
    // Idents (Variablen/Kategorien/Timer)
    private const CAT_CFG               = 'CFG_CATEGORY';
    private const VAR_POWER_ID          = 'CFG_POWER_ID';
    private const VAR_AC_ID             = 'CFG_ARCHIVE_ID';
    private const VAR_YEARS             = 'CFG_YEARS';
    private const VAR_INTERVAL          = 'CFG_INTERVAL';
    private const VAR_AUTO              = 'CFG_AUTO';
    private const VAR_DEBUG             = 'CFG_DEBUG';

    private const VAR_HTML_SUMMARY      = 'HTML_SUMMARY';
    private const VAR_HTML_HEATMAP      = 'HTML_HEATMAP';

    private const TIMER_NAME            = 'PA_UpdateTimer';

    // Property-Namen
    private const PROP_POWER_ID         = 'PowerVariableID';
    private const PROP_ARCHIVE_ID       = 'ArchiveControlID';
    private const PROP_YEARS            = 'YearsToAnalyze';
    private const PROP_INTERVAL         = 'UpdateIntervalSeconds';
    private const PROP_AUTO             = 'EnableAutoUpdate';
    private const PROP_DEBUG            = 'EnableDebug';

    public function Create()
    {
        parent::Create();

        // Properties (Defaults)
        $this->RegisterPropertyInteger(self::PROP_POWER_ID, 0);
        $this->RegisterPropertyInteger(self::PROP_ARCHIVE_ID, 12496);
        $this->RegisterPropertyInteger(self::PROP_YEARS, 5);
        $this->RegisterPropertyInteger(self::PROP_INTERVAL, 300);
        $this->RegisterPropertyBoolean(self::PROP_AUTO, true);
        $this->RegisterPropertyBoolean(self::PROP_DEBUG, false);

        // Timer
        $this->RegisterTimer(self::TIMER_NAME, 0, 'PA_Update($_IPS[\'TARGET\']);');

        // Kategorien & Variablen anlegen
        $cfgCatID = $this->EnsureCategory(self::CAT_CFG, 'Konfiguration', $this->InstanceID);

        // Konfiguration-Variablen (beschreibbar via RequestAction)
        $this->EnsureIntegerVar(self::VAR_POWER_ID, 'PowerVariableID', $cfgCatID, '~Number', true);
        $this->EnsureIntegerVar(self::VAR_AC_ID, 'ArchiveControlID', $cfgCatID, '~Number', true);
        $this->EnsureIntegerVar(self::VAR_YEARS, 'YearsToAnalyze', $cfgCatID, '~Number', true);
        $this->EnsureIntegerVar(self::VAR_INTERVAL, 'UpdateIntervalSeconds', $cfgCatID, '~Number', true);
        $this->EnsureBooleanVar(self::VAR_AUTO, 'EnableAutoUpdate', $cfgCatID, true);
        $this->EnsureBooleanVar(self::VAR_DEBUG, 'EnableDebug', $cfgCatID, true);

        // Reports auf Instanz-Ebene
        $this->EnsureHTMLVar(self::VAR_HTML_SUMMARY, 'Kurzfazit', $this->InstanceID);
        $this->EnsureHTMLVar(self::VAR_HTML_HEATMAP, 'Heatmap: Monatsanteile je Dezil', $this->InstanceID);
    }

    public function ApplyChanges()
    {
        parent::ApplyChanges();

        // Properties -> Konfig-Variablen synchronisieren (damit WebFront Werte sieht)
        $this->SetValueSafe(self::VAR_POWER_ID, (int)$this->ReadPropertyInteger(self::PROP_POWER_ID));
        $this->SetValueSafe(self::VAR_AC_ID,    (int)$this->ReadPropertyInteger(self::PROP_ARCHIVE_ID));
        $this->SetValueSafe(self::VAR_YEARS,    (int)$this->ReadPropertyInteger(self::PROP_YEARS));
        $this->SetValueSafe(self::VAR_INTERVAL, (int)$this->ReadPropertyInteger(self::PROP_INTERVAL));
        $this->SetValueSafe(self::VAR_AUTO,     (bool)$this->ReadPropertyBoolean(self::PROP_AUTO));
        $this->SetValueSafe(self::VAR_DEBUG,    (bool)$this->ReadPropertyBoolean(self::PROP_DEBUG));

        // Timer setzen / stoppen
        $interval = max(0, (int)$this->ReadPropertyInteger(self::PROP_INTERVAL));
        $auto     = (bool)$this->ReadPropertyBoolean(self::PROP_AUTO);

        $this->SetTimerInterval(self::TIMER_NAME, ($auto && $interval > 0) ? $interval * 1000 : 0);

        // Optional: beim ApplyChanges einmal aktualisieren (nicht bei jedem Klick)
        if ($auto && $interval > 0) {
            // kleine Verzögerung ist durch Timer gegeben; hier direkte Initial-Aktualisierung:
            $this->Update();
        }
    }

    /**
     * WebFront schreibt Variablen an die Instanz (Buttons / editierbare Felder)
     */
    public function RequestAction($Ident, $Value)
    {
        switch ($Ident) {
            case self::VAR_POWER_ID:
                $this->UpdatePropertyAndVar(self::PROP_POWER_ID, self::VAR_POWER_ID, (int)$Value);
                break;

            case self::VAR_AC_ID:
                $this->UpdatePropertyAndVar(self::PROP_ARCHIVE_ID, self::VAR_AC_ID, (int)$Value);
                break;

            case self::VAR_YEARS:
                $val = max(1, (int)$Value);
                $this->UpdatePropertyAndVar(self::PROP_YEARS, self::VAR_YEARS, $val);
                break;

            case self::VAR_INTERVAL:
                $val = max(10, (int)$Value);
                $this->UpdatePropertyAndVar(self::PROP_INTERVAL, self::VAR_INTERVAL, $val);
                // Timer direkt neu setzen
                $auto = (bool)$this->ReadPropertyBoolean(self::PROP_AUTO);
                $this->SetTimerInterval(self::TIMER_NAME, $auto ? $val * 1000 : 0);
                break;

            case self::VAR_AUTO:
                $this->UpdatePropertyAndVar(self::PROP_AUTO, self::VAR_AUTO, (bool)$Value);
                // Timer an/aus
                $interval = (int)$this->ReadPropertyInteger(self::PROP_INTERVAL);
                $auto     = (bool)$Value;
                $this->SetTimerInterval(self::TIMER_NAME, ($auto && $interval > 0) ? $interval * 1000 : 0);
                break;

            case self::VAR_DEBUG:
                $this->UpdatePropertyAndVar(self::PROP_DEBUG, self::VAR_DEBUG, (bool)$Value);
                break;

            default:
                throw new Exception("Unbekannter Ident: " . $Ident);
        }

        // Nach Einstellungsänderungen neu berechnen
        if (in_array($Ident, [self::VAR_POWER_ID, self::VAR_AC_ID, self::VAR_YEARS], true)) {
            $this->Update();
        }
    }

    /**
     * Manuelle/Timer-Aktualisierung
     */
    public function Update()
    {
        $debug   = (bool)$this->ReadPropertyBoolean(self::PROP_DEBUG);
        $powerID = (int)$this->ReadPropertyInteger(self::PROP_POWER_ID);
        $acID    = (int)$this->ReadPropertyInteger(self::PROP_ARCHIVE_ID);
        $years   = max(1, (int)$this->ReadPropertyInteger(self::PROP_YEARS));

        if ($powerID <= 0 || !IPS_VariableExists($powerID)) {
            $this->WriteHTML(self::VAR_HTML_SUMMARY, $this->WarnBox('PowerVariableID ist ungültig oder existiert nicht.'));
            $this->WriteHTML(self::VAR_HTML_HEATMAP, $this->WarnBox('Keine Berechnung möglich.'));
            return;
        }

        if ($acID <= 0 || !IPS_InstanceExists($acID)) {
            $this->WriteHTML(self::VAR_HTML_SUMMARY, $this->WarnBox('ArchiveControlID ungültig oder Instanz existiert nicht.'));
            $this->WriteHTML(self::VAR_HTML_HEATMAP, $this->WarnBox('Keine Berechnung möglich.'));
            return;
        }

        // Logging prüfen (nur Hinweis, kein Auto-Enable hier im Modul)
        $isLogged = @AC_GetLoggingStatus($acID, $powerID);
        if ($isLogged === false) {
            $this->Log('Hinweis: Die Power-Variable ist im Archiv nicht auf "Standard" gesetzt. Ergebnisse können leer/unvollständig sein.');
        }

        [$yearMax, $decileBounds, $perMonthStats] = $this->BuildDecileStats($acID, $powerID, $years, $debug);

        // HTMLs rendern
        $summary = $this->RenderSummary($yearMax, $perMonthStats);
        $this->WriteHTML(self::VAR_HTML_SUMMARY, $summary);

        $heatmap = $this->RenderHeatmap($decileBounds, $perMonthStats);
        $this->WriteHTML(self::VAR_HTML_HEATMAP, $heatmap);
    }

    /* =========================
     *   Berechnung / Datenzugriff
     * ========================= */

    /**
     * Liefert:
     *  - $yearMax (float)
     *  - $decileBounds (array[11] Grenzwerte 0..max in 10 Dezilbänder)
     *  - $perMonthStats: [
     *        'YYYY-MM' => [
     *            'minutesTotal' => int,
     *            'max'          => float,
     *            'avg'          => float,
     *            'buckets'      => [10] Sekunden in jedem Dezil,
     *        ],
     *        ...
     *    ]
     */
    private function BuildDecileStats(int $acID, int $varID, int $years, bool $debug): array
    {
        $now = time();

        // Zeitraum: letzte N Jahre bis jetzt (Monatsweise)
        $startMonthTS = $this->MonthStart(strtotime('-' . ($years * 12 - 1) . ' months', $this->MonthStart($now)));

        // 1) Jahres-Max über den Zeitraum (für Bänder)
        $yearMax = 0.0;
        $monthPtrs = [];
        $ptr = $startMonthTS;
        while ($ptr <= $now) {
            $monthPtrs[] = $ptr;
            $ptr = $this->MonthStart(strtotime('+1 month', $ptr));
        }

        $perMonthStats = [];

        // Zunächst grob Max je Monat ermitteln und YearMax ableiten
        foreach ($monthPtrs as $mStart) {
            $mEnd = $this->MonthEnd($mStart);
            $maxVal = $this->GetMonthMax($acID, $varID, $mStart, $mEnd);
            if ($maxVal !== null && $maxVal > $yearMax) {
                $yearMax = $maxVal;
            }
        }

        if ($yearMax <= 0.0) {
            $yearMax = 1.0; // Fallback, damit Bänder nicht kollabieren
        }

        // Dezil-Grenzen 0..Max in 10 gleichen Schritten
        $decileBounds = [];
        for ($i = 0; $i <= 10; $i++) {
            $decileBounds[$i] = $yearMax * ($i / 10.0);
        }

        // 2) Monatliche Buckets: Sekunden pro Dezil
        foreach ($monthPtrs as $mStart) {
            $mEnd = $this->MonthEnd($mStart);
            $key  = date('Y-m', $mStart);

            // Sekunden pro Band
            $secsPerBucket = array_fill(0, 10, 0);
            $sumSecs       = 0;
            $sumValSecs    = 0.0;
            $maxVal        = 0.0;

            // Rohdaten: alle Änderungen im Monat
            $raw = @AC_GetLoggedValues($acID, $varID, $mStart, $mEnd, 0);
            if (!is_array($raw)) {
                $raw = [];
            }

            // Falls der Monat mit einem "laufenden" Wert startet, der vor Monatsbeginn gesetzt wurde:
            $firstTS = $mStart;
            $firstVal = null;

            // Letzten Wert vor Monatsbeginn holen (Limit 1)
            $prev = @AC_GetLoggedValues($acID, $varID, 0, $mStart - 1, 1);
            if (is_array($prev) && count($prev) > 0) {
                $firstVal = (float)$prev[0]['Value'];
            } elseif (count($raw) > 0) {
                // sonst ersten Monatswert als Startwert annehmen
                $firstVal = (float)$raw[0]['Value'];
                if ($raw[0]['TimeStamp'] > $mStart) {
                    // Annahme: Wert gilt ab Monatsbeginn bis zum ersten Logeintrag
                }
            } else {
                // Kein Wert -> leerer Monat
                $perMonthStats[$key] = [
                    'minutesTotal' => (int)round(($mEnd - $mStart + 1) / 60),
                    'max'          => 0.0,
                    'avg'          => 0.0,
                    'buckets'      => $secsPerBucket
                ];
                continue;
            }

            $lastTS = $mStart;
            $lastVal = $firstVal;

            // Hilfsfunktion: Wert in Bucket legen
            $bucketIndexOf = function(float $v) use ($decileBounds): int {
                if ($v <= $decileBounds[0]) {
                    return 0;
                }
                for ($i = 1; $i < 10; $i++) {
                    if ($v <= $decileBounds[$i]) {
                        return max(0, $i - 1);
                    }
                }
                return 9; // oberstes Dezil
            };

            $accumulate = function(int $fromTS, int $toTS, float $value) use (&$secsPerBucket, &$sumSecs, &$sumValSecs, &$maxVal, $bucketIndexOf) {
                if ($toTS <= $fromTS) {
                    return;
                }
                $dur = $toTS - $fromTS; // Sekunden
                if ($dur < 1) {
                    return; // Spikes < 1 Sekunde ignorieren
                }
                $b = $bucketIndexOf($value);
                $secsPerBucket[$b] += $dur;
                $sumSecs           += $dur;
                $sumValSecs        += $value * $dur;
                if ($value > $maxVal) {
                    $maxVal = $value;
                }
            };

            // Durch alle Monats-Logeinträge iterieren
            foreach ($raw as $row) {
                $ts = (int)$row['TimeStamp'];
                $val = (float)$row['Value'];

                // Zeitraum bis zum aktuellen Eintrag
                $from = max($lastTS, $mStart);
                $to   = min($ts, $mEnd + 1);
                if ($to > $from) {
                    $accumulate($from, $to, $lastVal);
                }

                // neuer Stand
                $lastTS  = $ts;
                $lastVal = $val;
            }

            // Rest bis Monatsende
            $from = max($lastTS, $mStart);
            $to   = $mEnd + 1;
            if ($to > $from) {
                $accumulate($from, $to, $lastVal);
            }

            // Minuten/AVG berechnen
            $minutesTotal = (int)round($sumSecs / 60);
            $avg = ($sumSecs > 0) ? ($sumValSecs / $sumSecs) : 0.0;

            $perMonthStats[$key] = [
                'minutesTotal' => $minutesTotal,
                'max'          => $maxVal,
                'avg'          => $avg,
                'buckets'      => $secsPerBucket
            ];
        }

        return [$yearMax, $decileBounds, $perMonthStats];
    }

    private function GetMonthMax(int $acID, int $varID, int $startTS, int $endTS): ?float
    {
        $raw = @AC_GetLoggedValues($acID, $varID, $startTS, $endTS, 0);
        if (!is_array($raw) || count($raw) === 0) {
            // Versuche letzten Wert vor Monatsbeginn mitzunehmen
            $prev = @AC_GetLoggedValues($acID, $varID, 0, $startTS - 1, 1);
            if (is_array($prev) && count($prev) > 0) {
                return (float)$prev[0]['Value'];
            }
            return null;
        }
        $max = null;
        foreach ($raw as $r) {
            $v = (float)$r['Value'];
            if ($max === null || $v > $max) {
                $max = $v;
            }
        }
        return $max ?? null;
    }

    /* =========================
     *   Rendering
     * ========================= */

    private function RenderSummary(float $yearMax, array $perMonth): string
    {
        // Kennzahlen: aktueller Monat
        $nowKey = date('Y-m');
        $mKeys  = array_keys($perMonth);
        sort($mKeys);
        $lastKey = end($mKeys) ?: $nowKey;

        $cur = $perMonth[$lastKey] ?? [
            'minutesTotal' => 0,
            'max' => 0.0,
            'avg' => 0.0
        ];

        $maxStr = number_format($cur['max'], 1, ',', '.');
        $avgStr = number_format($cur['avg'], 1, ',', '.');
        $yearMaxStr = number_format($yearMax, 1, ',', '.');

        $css = $this->BaseCSS();

        $html  = '<div class="pa-card">';
        $html .= '<div class="pa-title">Kurzfazit</div>';
        $html .= '<div class="pa-text">';
        $html .= '• <b>Aktueller Monat:</b> Max <b>'.$maxStr.' W</b>, Durchschnitt <b>'.$avgStr.' W</b><br>';
        $html .= '• <b>Jahres-Maximum (Skala):</b> <b>'.$yearMaxStr.' W</b>';
        $html .= '</div></div>';

        return $css . $html;
    }

    private function RenderHeatmap(array $bounds, array $perMonth): string
    {
        // Sortiere Monate (letzte 12 zuerst)
        $keys = array_keys($perMonth);
        sort($keys);
        $last12 = array_slice($keys, -12);

        // Kopfzeile
        $thMonths = '';
        foreach ($last12 as $k) {
            $thMonths .= '<th class="pa-th">'.htmlspecialchars($this->FormatMonthKey($k)).'</th>';
        }
        $thMonths .= '<th class="pa-th">Summe</th>';

        $yearMax = end($bounds); // bounds[10]

        // Tabelle aufbauen
        $rows = '';
        for ($b = 9; $b >= 0; $b--) {
            $low  = $bounds[$b];
            $high = $bounds[$b+1];
            $bandLabel = number_format($low, 1, ',', '.').'–'.number_format($high, 1, ',', '.').' W';

            $tds = '';
            $sumPercent = 0.0;

            foreach ($last12 as $monKey) {
                $st = $perMonth[$monKey] ?? null;
                if (!$st || $st['minutesTotal'] <= 0) {
                    $tds .= $this->Cell(0.0, 0, false);
                    continue;
                }
                $secs = (int)($st['buckets'][$b] ?? 0);
                $percent = ($st['minutesTotal'] > 0) ? (100.0 * ($secs / 60.0) / $st['minutesTotal']) : 0.0;
                $sumPercent += $percent;

                $minutes = $secs / 60.0;
                $tds .= $this->Cell($percent, $minutes, $secs > 0);
            }

            // Summenspalte (über Monate – rein informativ)
            $rows .= '<tr>';
            $rows .= '<th class="pa-rowhead">'.$this->WrapText('Dezilband: '.$bandLabel).'</th>';
            $rows .= $tds;
            $rows .= '<td class="pa-sumcell">'.number_format($sumPercent, 2, ',', '.').'%</td>';
            $rows .= '</tr>';
        }

        // Fußzeile „Summe 100 %“ je Monat + Max/Avg je Monat
        $sumRow = '<tr><th class="pa-foothead">Summe</th>';
        foreach ($last12 as $monKey) {
            $st = $perMonth[$monKey] ?? null;
            if (!$st || $st['minutesTotal'] <= 0) {
                $sumRow .= '<td class="pa-footcell">—</td>';
                continue;
            }
            // Prozent über alle Bänder:
            $totalMinutes = $st['minutesTotal'];
            $sumPerc = 0.0;
            foreach ($perMonth[$monKey]['buckets'] as $secs) {
                $sumPerc += (100.0 * ($secs / 60.0) / $totalMinutes);
            }
            // Auf 100.00% runden (nur Anzeige)
            $sumRow .= '<td class="pa-footcell"><b>'.number_format($sumPerc, 2, ',', '.').'%</b><br>'
                     . '<span class="pa-sub">Ø '.number_format($st['avg'], 1, ',', '.').' W · Max '.number_format($st['max'], 1, ',', '.').' W</span>'
                     . '</td>';
        }
        $sumRow .= '<td class="pa-footcell">—</td></tr>';

        $css = $this->BaseCSS();

        $html  = '<div class="pa-card">';
        $html .= '<div class="pa-title">Monatsanteile je Dezil (letzte 12 Monate)</div>';
        $html .= '<table class="pa-table">';
        $html .= '<thead><tr><th class="pa-th pa-sticky">Dezilband</th>'.$thMonths.'</tr></thead>';
        $html .= '<tbody>'.$rows.'</tbody>';
        $html .= '<tfoot>'.$sumRow.'</tfoot>';
        $html .= '</table></div>';

        return $css . $html;
    }

    private function Cell(float $percent, float $minutes, bool $hasData): string
    {
        $p = max(0.0, $percent);
        $bg = $this->ColorForPercent($p); // <1% hellgrau
        $txt = number_format(ceil($p * 100) / 100, 2, ',', '.').'%' ; // 2 Stellen, „aufrunden“ auf 2. Stelle
        $minStr = $hasData ? '<br><span class="pa-sub">'.number_format($minutes, 2, ',', '.').' min</span>' : '<br><span class="pa-sub">0 min</span>';

        return '<td class="pa-td" style="background: '.$bg.'">'.$txt.$minStr.'</td>';
    }

    private function ColorForPercent(float $p): string
    {
        // < 1% -> sehr helles Grau
        if ($p < 1.0) {
            return '#F0F0F0';
        }
        // Warme Abstufung (ohne dominantes Blau)
        // 1%..100% mappen auf 0..1
        $t = min(1.0, max(0.0, ($p - 1.0) / 99.0));
        // einfache Interpolation zwischen hell (fast weiß) -> orange -> rot
        // Start:   #FDFBE5 (253,251,229)
        // Mitte:   #FDB366 (253,179,102)
        // Ende:    #D84A4A (216,74,74)
        $c1 = [253,251,229];
        $c2 = [253,179,102];
        $c3 = [216, 74, 74];
        if ($t < 0.5) {
            $k = $t / 0.5;
            $r = intval($c1[0] + ($c2[0]-$c1[0])*$k);
            $g = intval($c1[1] + ($c2[1]-$c1[1])*$k);
            $b = intval($c1[2] + ($c2[2]-$c1[2])*$k);
        } else {
            $k = ($t-0.5)/0.5;
            $r = intval($c2[0] + ($c3[0]-$c2[0])*$k);
            $g = intval($c2[1] + ($c3[1]-$c2[1])*$k);
            $b = intval($c2[2] + ($c3[2]-$c2[2])*$k);
        }
        return sprintf('#%02X%02X%02X', $r, $g, $b);
    }

    private function BaseCSS(): string
    {
        return <<<CSS
<style>
.pa-card{background:#121212;color:#FFFFFF;border:1px solid #333;border-radius:8px;padding:12px;margin:8px 0;font-family:-apple-system,Segoe UI,Roboto,Arial,sans-serif;}
.pa-title{font-weight:700;margin-bottom:6px;}
.pa-text{line-height:1.4}
.pa-table{width:100%;border-collapse:separate;border-spacing:0}
.pa-th{position:sticky;top:0;background:#1a1a1a;color:#000;border-bottom:1px solid #333;padding:6px;text-align:center;font-weight:700}
.pa-sticky{left:0;z-index:2}
.pa-rowhead{background:#1a1a1a;color:#000;border-right:1px solid #333;padding:6px;max-width:180px;word-wrap:break-word;white-space:normal}
.pa-td{border-bottom:1px solid #222;border-right:1px solid #222;text-align:center;padding:6px;min-width:80px}
.pa-sumcell{background:#0f0f0f;color:#fff;border-left:1px solid #333;padding:6px;text-align:center;font-weight:700}
.pa-foothead{background:#0f0f0f;color:#fff;border-top:1px solid #333;border-right:1px solid #333;padding:6px}
.pa-footcell{background:#0f0f0f;color:#fff;border-top:1px solid #333;border-right:1px solid #333;padding:6px;text-align:center}
.pa-sub{display:block;font-size:11px;opacity:.9}
.pa-warn{background:#2a1a1a;border:1px solid #662;color:#ffdede;border-radius:6px;padding:8px}
</style>
CSS;
    }

    private function WrapText(string $s): string
    {
        return htmlspecialchars($s);
    }

    private function WarnBox(string $msg): string
    {
        return $this->BaseCSS() . '<div class="pa-warn">'.$this->WrapText($msg).'</div>';
    }

    /* =========================
     *   Hilfsfunktionen
     * ========================= */

    private function EnsureCategory(string $ident, string $name, int $parentID): int
    {
        $id = @IPS_GetObjectIDByIdent($ident, $this->InstanceID);
        if ($id && IPS_ObjectExists($id)) {
            if (IPS_GetObject($id)['ObjectType'] === OBJECTTYPE_CATEGORY) {
                IPS_SetName($id, $name);
                IPS_SetParent($id, $parentID);
                IPS_SetIdent($id, $ident);
                return $id;
            }
        }
        // neu
        $catID = IPS_CreateCategory();
        IPS_SetName($catID, $name);
        IPS_SetParent($catID, $parentID);
        IPS_SetIdent($catID, $ident);
        return $catID;
    }

    private function EnsureHTMLVar(string $ident, string $name, int $parentID): int
    {
        $id = @IPS_GetObjectIDByIdent($ident, $this->InstanceID);
        if (!$id) {
            $id = IPS_CreateVariable(VARIABLETYPE_STRING);
            IPS_SetIdent($id, $ident);
        }
        IPS_SetName($id, $name);
        IPS_SetParent($id, $parentID);
        IPS_SetInfo($id, 'Auto-generiert');
        IPS_SetVariableCustomProfile($id, '~HTMLBox');
        return $id;
    }

    private function EnsureIntegerVar(string $ident, string $name, int $parentID, string $profile, bool $writable): int
    {
        $id = @IPS_GetObjectIDByIdent($ident, $this->InstanceID);
        if (!$id) {
            $id = IPS_CreateVariable(VARIABLETYPE_INTEGER);
            IPS_SetIdent($id, $ident);
        }
        IPS_SetName($id, $name);
        IPS_SetParent($id, $parentID);
        IPS_SetVariableCustomProfile($id, $profile);
        if ($writable) {
            IPS_SetVariableCustomAction($id, $this->InstanceID);
        } else {
            IPS_SetVariableCustomAction($id, 0);
        }
        return $id;
    }

    private function EnsureBooleanVar(string $ident, string $name, int $parentID, bool $writable): int
    {
        $id = @IPS_GetObjectIDByIdent($ident, $this->InstanceID);
        if (!$id) {
            $id = IPS_CreateVariable(VARIABLETYPE_BOOLEAN);
            IPS_SetIdent($id, $ident);
        }
        IPS_SetName($id, $name);
        IPS_SetParent($id, $parentID);
        $this->EnsureBooleanProfile('PA.Toggle', 'Aus', 'An');
        IPS_SetVariableCustomProfile($id, 'PA.Toggle');
        IPS_SetVariableCustomAction($id, $writable ? $this->InstanceID : 0);
        return $id;
    }

    private function EnsureBooleanProfile(string $name, string $off, string $on): void
    {
        if (!IPS_VariableProfileExists($name)) {
            IPS_CreateVariableProfile($name, VARIABLETYPE_BOOLEAN);
        }
        IPS_SetVariableProfileAssociation($name, 0, $off, '', 0x666666);
        IPS_SetVariableProfileAssociation($name, 1, $on,  '', 0x00AA00);
    }

    private function SetValueSafe(string $ident, $value): void
    {
        $id = @IPS_GetObjectIDByIdent($ident, $this->InstanceID);
        if ($id && IPS_VariableExists($id)) {
            @SetValue($id, $value);
        }
    }

    private function UpdatePropertyAndVar(string $prop, string $ident, $value): void
    {
        $ok = $this->UpdatePropertyValue($prop, $value);
        $this->SetValueSafe($ident, $value);
        if ($ok) {
            $this->ApplyChanges(); // neu anwenden
        }
    }

    private function UpdatePropertyValue(string $prop, $value): bool
    {
        // Properties sind schreibgeschützt, aber über die JSON-Config änderbar:
        // Workaround: alte Properties lesen, array anpassen, speichern
        $cfg = json_decode(IPS_GetConfiguration($this->InstanceID), true) ?: [];
        $cfg[$prop] = $value;
        IPS_SetConfiguration($this->InstanceID, json_encode($cfg));
        return true;
    }

    private function MonthStart(int $ts): int
    {
        return (int)strtotime(date('Y-m-01 00:00:00', $ts));
    }
    private function MonthEnd(int $monthStartTS): int
    {
        return (int)strtotime(date('Y-m-t 23:59:59', $monthStartTS));
    }

    private function FormatMonthKey(string $ym): string
    {
        // "YYYY-MM" -> "MM.YYYY"
        $parts = explode('-', $ym);
        if (count($parts) == 2) {
            return $parts[1] . '.' . $parts[0];
        }
        return $ym;
    }

    private function WriteHTML(string $ident, string $html): void
    {
        $id = @IPS_GetObjectIDByIdent($ident, $this->InstanceID);
        if ($id && IPS_VariableExists($id)) {
            @SetValueString($id, $html);
        }
    }

    private function Log(string $msg): void
    {
        if ((bool)$this->ReadPropertyBoolean(self::PROP_DEBUG)) {
            $this->SendDebug('PowerAnalyzer', $msg, 0);
        }
    }
}

