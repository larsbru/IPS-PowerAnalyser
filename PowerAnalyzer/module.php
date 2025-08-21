<?php
declare(strict_types=1);

class PowerAnalyzer extends IPSModule
{
    public function Create()
    {
        // Immer aufrufen!
        parent::Create();

        // Properties (mit deinen Defaults)
        $this->RegisterPropertyInteger('ArchiveControlID', 12496);
        $this->RegisterPropertyInteger('PowerVarID', 0);
        $this->RegisterPropertyInteger('Months', 12);
        $this->RegisterPropertyBoolean('TimerEnabled', true);
        $this->RegisterPropertyInteger('TimerSeconds', 300);
        $this->RegisterPropertyBoolean('Debug', false);

        // Hilfs-Attribut: ID des Konfig-Ordners unter der Instanz
        $this->RegisterAttributeInteger('CfgCategoryID', 0);

        // Timer
        $this->RegisterTimer('Update', 0, 'PA_Update($_IPS[\'TARGET\']);');

        // HTML-Ausgaben (auf Instanzebene)
        $this->RegisterVariableString('ReportDeciles', 'Monatsanteile je Dezil (Heatmap)', '~HTMLBox', 10);
        $this->RegisterVariableString('ReportSummary', 'Kurzfazit', '~HTMLBox', 20);
    }

    public function ApplyChanges()
    {
        parent::ApplyChanges();

        // Konfig-Ordner sicherstellen
        $cfgCatID = $this->EnsureCategory($this->InstanceID, 'Konfiguration', 'CFG');
        $this->WriteAttributeInteger('CfgCategoryID', $cfgCatID);

        // Profile
        $this->EnsureVarProfile('PA.Integer', VARIABLETYPE_INTEGER);

        // Konfig-Variablen im WebFront (Buttons/Eingabefelder)
        $vArchive = $this->EnsureVar($cfgCatID, 'ArchiveControl-ID', VARIABLETYPE_INTEGER, 'ArchiveIDVar', 'PA.Integer');
        $vPower   = $this->EnsureVar($cfgCatID, 'Power-Variable (W)', VARIABLETYPE_INTEGER, 'PowerVarIDVar', 'PA.Integer');
        $vMonths  = $this->EnsureVar($cfgCatID, 'Monate analysieren', VARIABLETYPE_INTEGER, 'MonthsVar', 'PA.Integer');
        $vDebug   = $this->EnsureVar($cfgCatID, 'Debug-Modus', VARIABLETYPE_BOOLEAN, 'DebugVar', '');
        $vTimOn   = $this->EnsureVar($cfgCatID, 'Timer aktiv', VARIABLETYPE_BOOLEAN, 'TimerOnVar', '~Switch');
        $vTimSec  = $this->EnsureVar($cfgCatID, 'Timer (Sekunden)', VARIABLETYPE_INTEGER, 'TimerSecVar', 'PA.Integer');

        // Werte aus Properties -> Variablen spiegeln
        $this->SetValueIfChanged($vArchive, $this->ReadPropertyInteger('ArchiveControlID'));
        $this->SetValueIfChanged($vPower,   $this->ReadPropertyInteger('PowerVarID'));
        $this->SetValueIfChanged($vMonths,  max(1, min(36, $this->ReadPropertyInteger('Months'))));
        $this->SetValueIfChanged($vDebug,   $this->ReadPropertyBoolean('Debug'));
        $this->SetValueIfChanged($vTimOn,   $this->ReadPropertyBoolean('TimerEnabled'));
        $this->SetValueIfChanged($vTimSec,  max(30, min(3600, $this->ReadPropertyInteger('TimerSeconds'))));

        // Variablen beschreibbar machen -> Aktionen kommen an RequestAction()
        foreach ([$vArchive,$vPower,$vMonths,$vDebug,$vTimOn,$vTimSec] as $vid) {
            IPS_SetVariableCustomAction($vid, $this->InstanceID);
        }

        // Timer anwenden
        $this->SetTimerInterval('Update',
            $this->ReadPropertyBoolean('TimerEnabled') ? $this->ReadPropertyInteger('TimerSeconds') * 1000 : 0
        );

        // Optional: beim Übernehmen einmal rechnen (sanft)
        // $this->Update();
    }

    public function RequestAction($Ident, $Value)
    {
        // Änderung aus dem WebFront: Variablen -> Properties -> ApplyChanges -> Update()
        switch ($Ident) {
            case 'ArchiveIDVar':
                $Value = (int)$Value;
                $this->SetConfigAndValue('ArchiveControlID', $Ident, $Value);
                break;

            case 'PowerVarIDVar':
                $Value = (int)$Value;
                $this->SetConfigAndValue('PowerVarID', $Ident, $Value);
                break;

            case 'MonthsVar':
                $Value = max(1, min(36, (int)$Value));
                $this->SetConfigAndValue('Months', $Ident, $Value);
                break;

            case 'TimerOnVar':
                $Value = (bool)$Value;
                $this->SetConfigAndValue('TimerEnabled', $Ident, $Value);
                break;

            case 'TimerSecVar':
                $Value = max(30, min(3600, (int)$Value));
                $this->SetConfigAndValue('TimerSeconds', $Ident, $Value);
                break;

            case 'DebugVar':
                $Value = (bool)$Value;
                $this->SetConfigAndValue('Debug', $Ident, $Value);
                break;

            default:
                throw new Exception('Unknown Ident: ' . $Ident);
        }

        // Änderungen sofort wirksam machen
        $this->ApplyChanges();
        $this->Update();
    }

    /** Button aus der Form */
    public function Update()
    {
        // Analyse durchführen und rendern
        $acID = $this->ReadPropertyInteger('ArchiveControlID');
        $pvID = $this->ReadPropertyInteger('PowerVarID');
        $months = max(1, (int)$this->ReadPropertyInteger('Months'));

        if ($acID <= 0 || !IPS_InstanceExists($acID)) {
            $this->SetSummaryBox('ArchiveControl-ID ungültig.');
            return;
        }
        if ($pvID <= 0 || !IPS_VariableExists($pvID)) {
            $this->SetSummaryBox('Power-Variable ungültig.');
            return;
        }

        // Monate erzeugen
        $now = time();
        $monthsInfo = [];
        for ($i = $months - 1; $i >= 0; $i--) {
            $t = strtotime(date('Y-m-01 00:00:00', strtotime("-$i month", $now)));
            $monthsInfo[] = [
                'start' => $t,
                'end'   => strtotime(date('Y-m-t 23:59:59', $t)),
                'label' => strftime('%b', $t),
                'y'     => date('Y', $t)
            ];
        }

        // Globales Max (echter Max) aus Duration-Quelle
        $globalMax = 0.0;
        foreach ($monthsInfo as $m) {
            $maxM = $this->readMonthMax_duration($acID, $pvID, $m['start'], $m['end']);
            if ($maxM > $globalMax) $globalMax = $maxM;
        }
        if ($globalMax <= 0) $globalMax = 1.0;

        // Bänder
        $bands = [];
        for ($d = 1; $d <= 10; $d++) $bands[] = ($globalMax * $d) / 10.0;

        // Ergebnisse je Monat
        $results = [];
        $monthStats = [];
        foreach ($monthsInfo as $idx => $m) {
            $results[$idx]    = $this->computeMonthDeciles_duration($acID, $pvID, $m['start'], $m['end'], $bands);
            $monthStats[$idx] = $this->computeMonthAvgMax($acID, $pvID, $m['start'], $m['end']);
        }

        // globale Max-Zelle finden
        $globalIdx = 0; $globalVal = -INF;
        foreach ($results as $i => $r) {
            if (($r['maxVal'] ?? 0) > $globalVal) { $globalVal = $r['maxVal']; $globalIdx = $i; }
        }
        $globalMaxMonthIndex = $globalIdx;
        $globalMaxBandIndex  = $results[$globalIdx]['maxBand'] ?? 9;

        // Rendern
        $html = $this->renderHeatmap($monthsInfo, $bands, $results, $monthStats, $globalMaxMonthIndex, $globalMaxBandIndex);
        SetValueString($this->GetIDForIdent('ReportDeciles'), $html);

        $summary = $this->buildSummary($globalMax, $monthsInfo, $results, $acID, $pvID);
        SetValueString($this->GetIDForIdent('ReportSummary'), $summary);
    }

    /* ===================== Helpers: Struktur/Variablen ===================== */

    private function EnsureCategory(int $parentID, string $name, string $ident): int
    {
        $id = @IPS_GetObjectIDByIdent($ident, $parentID);
        if ($id && IPS_GetObject($id)['ObjectType'] === 0) {
            if (IPS_GetObject($id)['ObjectName'] !== $name) IPS_SetName($id, $name);
            return $id;
        }
        // neu
        $id = IPS_CreateCategory();
        IPS_SetParent($id, $parentID);
        IPS_SetName($id, $name);
        IPS_SetIdent($id, $ident);
        return $id;
    }

    private function EnsureVarProfile(string $name, int $type): void
    {
        if (!IPS_VariableProfileExists($name)) {
            IPS_CreateVariableProfile($name, $type);
        }
    }

    private function EnsureVar(int $parentID, string $name, int $type, string $ident, string $profile = '')
    {
        $vid = @IPS_GetObjectIDByIdent($ident, $parentID);
        if ($vid === false) {
            $vid = IPS_CreateVariable($type);
            IPS_SetParent($vid, $parentID);
            IPS_SetIdent($vid, $ident);
            IPS_SetName($vid, $name);
            if ($profile !== '') IPS_SetVariableCustomProfile($vid, $profile);
        } else {
            if ($profile !== '') IPS_SetVariableCustomProfile($vid, $profile);
            IPS_SetName($vid, $name);
        }
        return $vid;
    }

    private function SetValueIfChanged(int $varID, $value): void
    {
        $cur = @GetValue($varID);
        if ($cur !== $value) {
            SetValue($varID, $value);
        }
    }

    private function SetConfigAndValue(string $propName, string $ident, $value): void
    {
        // Variable setzen
        $varID = @IPS_GetObjectIDByIdent($ident, $this->ReadAttributeInteger('CfgCategoryID'));
        if ($varID) SetValue($varID, $value);

        // Property setzen + anwenden
        IPS_SetProperty($this->InstanceID, $propName, $value);
    }

    private function SetSummaryBox(string $text): void
    {
        $box = '<div style="font-family:Arial; color:#fff; background:#333; padding:8px;">'
             . htmlspecialchars($text) . '</div>';
        SetValueString($this->GetIDForIdent('ReportSummary'), $box);
    }

    /* ===================== Analyse-Funktionen (Duration-basiert) ===================== */

    private function readMonthMax_duration(int $acID, int $varID, int $start, int $end): float
    {
        $max = 0.0;
        $rows = @AC_GetLoggedValues($acID, $varID, $start, $end, 0) ?: [];
        foreach ($rows as $r) {
            $v = (float)$r['Value'];
            if ($v > $max) $max = $v;
        }
        return $max;
    }

    private function computeMonthDeciles_duration(int $acID, int $varID, int $start, int $end, array $bands): array
    {
        $sec = array_fill(0, 10, 0);
        $rows = @AC_GetLoggedValues($acID, $varID, $start, $end, 0) ?: [];

        $monthMaxVal  = 0.0;
        $monthMaxBand = 9;

        foreach ($rows as $r) {
            $ts  = (int)$r['TimeStamp'];
            $dur = (int)$r['Duration'];   // Sekunden
            $val = (float)$r['Value'];

            // Clamp auf Monatsfenster
            $segStart = max($ts, $start);
            $segEnd   = min($ts + $dur, $end);
            $d = $segEnd - $segStart;
            if ($d <= 0) continue;

            // Max verfolgen
            if ($val > $monthMaxVal) {
                $monthMaxVal  = $val;
                $monthMaxBand = $this->valueToBandByEdges($val, $bands);
            }

            // Dauer einem Band zuordnen
            $idx = $this->valueToBandByEdges($val, $bands);
            $sec[$idx] += $d;
        }

        $sum = array_sum($sec);
        $pct = array_fill(0, 10, 0.0);

        if ($sum > 0) {
            $raw = [];
            for ($i = 0; $i < 10; $i++) {
                $raw[$i] = ($sec[$i] / $sum) * 100.0;
                if ($raw[$i] <= 0.0) {
                    $pct[$i] = 0.00;
                } else {
                    $up = ceil($raw[$i] * 100.0) / 100.0; // immer aufrunden auf 2 NK
                    if ($up < 0.01) $up = 0.01;
                    $pct[$i] = round($up, 2);
                }
            }
            // Summenkorrektur -> exakt 100,00 %
            $sumRounded = array_sum($pct);
            $diff = round(100.00 - $sumRounded, 2);
            if (abs($diff) >= 0.01) {
                $k = array_keys($pct, max($pct))[0];
                $pct[$k] = round($pct[$k] + $diff, 2);
                if ($pct[$k] < 0.00)   $pct[$k] = 0.00;
                if ($pct[$k] > 100.00) $pct[$k] = 100.00;
            }
        }

        return ['pct' => $pct, 'sec' => $sec, 'maxVal' => $monthMaxVal, 'maxBand' => $monthMaxBand];
    }

    private function valueToBandByEdges(float $v, array $bands): int
    {
        if ($v <= 0) return 0;
        $k = count($bands);
        for ($i = 0; $i < $k; $i++) {
            if ($v <= $bands[$i]) return $i; // inklusiv
        }
        return $k - 1;
    }

    private function computeMonthAvgMax(int $acID, int $varID, int $start, int $end): array
    {
        $vals = @AC_GetAggregatedValues($acID, $varID, 1, $start, $end, 0);
        if (!is_array($vals) || count($vals) == 0) return ['avg' => 0.0, 'max' => 0.0];
        $sum = 0.0; $c = 0; $mx = 0.0;
        foreach ($vals as $v) {
            $sum += (float)$v['Avg']; $c++;
            if ((float)$v['Max'] > $mx) $mx = (float)$v['Max'];
        }
        $avg = $c > 0 ? $sum / $c : 0.0;
        return ['avg' => round($avg, 1), 'max' => round($mx, 1)];
    }

    /* ===================== Rendering ===================== */

    private function renderHeatmap(array $months, array $bands, array $results, array $monthStats, int $gMon, int $gBand): string
    {
        $decLabels = []; $lo = 0.0;
        for ($i = 0; $i < 10; $i++) {
            $hi = $bands[$i];
            $decLabels[$i] = number_format($lo, 1, ',', '.') . '–' . number_format($hi, 1, ',', '.') . ' W';
            $lo = $hi;
        }

        $css = <<<CSS
<style>
.pa_mod { font-family: Arial,sans-serif; background:#fff; color:#000; }
.pa_mod table { border-collapse: collapse; width:100%; table-layout:fixed; }
.pa_mod th,.pa_mod td { border:1px solid #ccc; padding:6px 8px; text-align:center; vertical-align:middle; }
.pa_mod th { background:#f0f0f0; color:#000; font-weight:600; }
.pa_mod .rowhdr { text-align:left; background:#f9f9f9; color:#000; font-weight:600; white-space: normal; word-break: break-word; }
.pa_mod .sum100 { font-weight:bold; color:#000; }
.pa_mod .monstat { display:block; font-size:11px; line-height:1.25; color:#000; margin-top:2px; }
.pa_mod .maxcell { border:2px solid #000 !important; }
</style>
CSS;

        $thead = '<tr><th class="rowhdr">Dezil-Band (W)</th>';
        foreach ($months as $m) {
            $thead .= '<th>' . htmlspecialchars($m['label'] . ' ' . $m['y']) . '</th>';
        }
        $thead .= '</tr>';

        $tbody = '';
        for ($r = 0; $r < 10; $r++) {
            $tbody .= '<tr><td class="rowhdr">' . htmlspecialchars($decLabels[$r]) . '</td>';
            foreach ($months as $idx => $m) {
                $p    = $results[$idx]['pct'][$r] ?? 0.0;
                $secs = $results[$idx]['sec'][$r] ?? 0;
                $pPrint = number_format($p, 2, ',', '.');
                $isZero = (round($p, 2) === 0.00);
                $color  = $isZero ? '#ffffff' : $this->pctToBlueYellow($p);
                $timeTxt = ($secs < 60) ? ($secs . ' s') : (ceil($secs / 60) . ' min');

                $cls = ($idx === $gMon && $r === $gBand) ? ' class="maxcell"' : '';
                $tbody .= '<td' . $cls . ' style="background:' . $color . ';color:#000;">'
                        . $pPrint . ' %'
                        . '<span class="monstat">' . $timeTxt . '</span>'
                        . '</td>';
            }
            $tbody .= '</tr>';
        }

        // Summenzeile + Monatsstats
        $tbody .= '<tr><td class="rowhdr sum100">Summe</td>';
        foreach ($months as $idx => $m) {
            $sum = array_sum($results[$idx]['pct'] ?? []);
            $sumPrint = number_format($sum, 2, ',', '.');
            $avg = $monthStats[$idx]['avg'] ?? 0.0;
            $mx  = $monthStats[$idx]['max'] ?? 0.0;
            $tbody .= '<td class="sum100">' . $sumPrint . ' %'
                    . '<span class="monstat">Ø ' . number_format($avg, 1, ',', '.') . ' W<br>Max ' . number_format($mx, 1, ',', '.') . ' W</span>'
                    . '</td>';
        }
        $tbody .= '</tr>';

        return $css . '<div class="pa_mod"><table><thead>' . $thead . '</thead><tbody>' . $tbody . '</tbody></table></div>';
    }

    private function pctToBlueYellow(float $p): string
    {
        if (round($p, 2) === 0.00) return '#ffffff';
        $p = max(0.0, min(100.0, $p)) / 100.0;
        $b = [30, 144, 255]; $y = [255, 215, 0];
        $r = (int)round($b[0] + ($y[0] - $b[0]) * $p);
        $g = (int)round($b[1] + ($y[1] - $b[1]) * $p);
        $bl= (int)round($b[2] + ($y[2] - $b[2]) * $p);
        return "rgb($r,$g,$bl)";
    }

    private function buildSummary(float $max, array $months, array $res, int $acID, int $varID): string
    {
        $now = time();
        $monthStart = strtotime(date('Y-m-01 00:00:00', $now));
        $monthEnd   = strtotime(date('Y-m-t 23:59:59', $now));
        $dayStart   = strtotime('today 00:00:00');
        $dayEnd     = strtotime('today 23:59:59');

        $avgMonth = $this->computeAvg($acID, $varID, $monthStart, $monthEnd);
        $maxMonth = $this->computeMax($acID, $varID, $monthStart, $monthEnd);
        $avgDay   = $this->computeAvg($acID, $varID, $dayStart, $dayEnd);
        $maxDay   = $this->computeMax($acID, $varID, $dayStart, $dayEnd);

        $firstStart = $months[0]['start'] ?? $monthStart;
        $lastEnd    = $months[count($months)-1]['end'] ?? $monthEnd;
        $avgGlobal  = $this->computeAvg($acID, $varID, $firstStart, $lastEnd);

        $txt  = '<div style="font-family:Arial; color:#fff; background:#333; padding:8px">';
        $txt .= '<div>Jahresmax (Zeitraum): <b>' . number_format($max, 1, ',', '.') . ' W</b></div>';
        $txt .= '<div>Aktueller Monat: ⌀ <b>' . number_format($avgMonth, 1, ',', '.') . ' W</b>, '
              . 'Max <b>' . number_format($maxMonth, 1, ',', '.') . ' W</b></div>';
        $txt .= '<div>Heute: ⌀ <b>' . number_format($avgDay, 1, ',', '.') . ' W</b>, '
              . 'Max <b>' . number_format($maxDay, 1, ',', '.') . ' W</b></div>';
        $txt .= '<div>Gesamt‑Ø (alle Monate): <b>' . number_format($avgGlobal, 1, ',', '.') . ' W</b></div>';
        $txt .= '</div>';
        return $txt;
    }

    private function computeAvg(int $acID, int $varID, int $start, int $end): float
    {
        $vals = @AC_GetAggregatedValues($acID, $varID, 1, $start, $end, 0);
        if (!is_array($vals) || count($vals) == 0) return 0.0;
        $sum = 0.0; $c = 0;
        foreach ($vals as $v) { $sum += (float)$v['Avg']; $c++; }
        return $c > 0 ? $sum / $c : 0.0;
    }

    private function computeMax(int $acID, int $varID, int $start, int $end): float
    {
        $vals = @AC_GetAggregatedValues($acID, $varID, 1, $start, $end, 0);
        if (!is_array($vals) || count($vals) == 0) return 0.0;
        $m = 0.0;
        foreach ($vals as $v) { $vMax = (float)$v['Max']; if ($vMax > $m) $m = $vMax; }
        return $m;
    }
}

// Globale Wrapper-Funktion für Actions im Form.json
function PA_Update(int $InstanceID)
{
    $obj = IPS_GetInstance($InstanceID);
    if ($obj['ModuleInfo']['ModuleID'] !== '{2C9A0D2E-7A36-4F88-9F2A-7F7E6B9E6F01}') {
        throw new Exception('Wrong instance type');
    }
    /** @var PowerAnalyzer $inst */
    $inst = IPS_GetInstanceObject($InstanceID);
    // Fallback, falls Umgebung keine Instanz-Rückgabe hat:
    if (!method_exists($inst, 'Update')) {
        // Normale Instanz-Methode aufrufen
        IPS_RequestAction($InstanceID, 'TimerOnVar', GetValueBoolean(@IPS_GetObjectIDByIdent('TimerOnVar', $InstanceID)));
    }
    $module = new PowerAnalyzer($InstanceID);
    $module->Update();
}
