<?php

/** @noinspection MultiAssignmentUsageInspection */

namespace Shuchkin;

use SimpleXMLElement;

class SimpleXLSXEx
{
    public static $IC = [
        0 => '000000',
        1 => 'FFFFFF',
        2 => 'FF0000',
        3 => '00FF00',
        4 => '0000FF',
        5 => 'FFFF00',
        6 => 'FF00FF',
        7 => '00FFFF',
        8 => '000000',
        9 => 'FFFFFF',
        10 => 'FF0000',
        11 => '00FF00',
        12 => '0000FF',
        13 => 'FFFF00',
        14 => 'FF00FF',
        15 => '00FFFF',
        16 => '800000',
        17 => '008000',
        18 => '000080',
        19 => '808000',
        20 => '800080',
        21 => '008080',
        22 => 'C0C0C0',
        23 => '808080',
        24 => '9999FF',
        25 => '993366',
        26 => 'FFFFCC',
        27 => 'CCFFFF',
        28 => '660066',
        29 => 'FF8080',
        30 => '0066CC',
        31 => 'CCCCFF',
        32 => '000080',
        33 => 'FF00FF',
        34 => 'FFFF00',
        35 => '00FFFF',
        36 => '800080',
        37 => '800000',
        38 => '008080',
        39 => '0000FF',
        40 => '00CCFF',
        41 => 'CCFFFF',
        42 => 'CCFFCC',
        43 => 'FFFF99',
        44 => '99CCFF',
        45 => 'FF99CC',
        46 => 'CC99FF',
        47 => 'FFCC99',
        48 => '3366FF',
        49 => '33CCCC',
        50 => '99CC00',
        51 => 'FFCC00',
        52 => 'FF9900',
        53 => 'FF6600',
        54 => '666699',
        55 => '969696',
        56 => '003366',
        57 => '339966',
        58 => '003300',
        59 => '333300',
        60 => '993300',
        61 => '993366',
        62 => '333399',
        63 => '333333',
        64 => '000000', // System Foreground
        65 => 'FFFFFF', // System Background'
    ];
    public static $CH = [
        0 => 'ANSI_CHARSET',
        1 => 'DEFAULT_CHARSET',
        2 => 'SYMBOL_CHARSET',
        77 => 'MAC_CHARSET',
        128 => 'SHIFTJIS_CHARSET',
        //129 => 'HANGEUL_CHARSET',
        129 => 'HANGUL_CHARSET',
        130 => 'JOHAB_CHARSET',
        134 => 'GB2312_CHARSET',
        136 => 'CHINESEBIG5_CHARSET',
        161 => 'GREEK_CHARSET',
        162 => 'TURKISH_CHARSET',
        163 => 'VIETNAMESE_CHARSET',
        177 => 'HEBREW_CHARSET',
        178 => 'ARABIC_CHARSET',
        186 => 'BALTIC_CHARSET',
        204 => 'RUSSIAN_CHARSET',
        222 => 'THAI_CHARSET',
        238 => 'EASTEUROPE_CHARSET',
        255 => 'OEM_CHARSET'
    ];
    public $xlsx;
    public $themeColors;
    public $fonts;
    public $fills;
    public $borders;
    public $cellStyles;
    public $css;
    public $comments;
    public $hyperlinks;
    public $worksheetIndex;

    public function __construct(SimpleXLSX $xlsx)
    {
        $this->xlsx = $xlsx;
        $this->readThemeColors();
        $this->readFonts();
        $this->readFills();
        $this->readBorders();
        $this->readXfs();
        $this->readHyperlinks();
        $this->readComments();
    }
    public function readThemeColors()
    {
        $this->themeColors = [];
        if (isset($this->xlsx->theme->themeElements->clrScheme)) {
            $colors12 = ['lt1', 'dk1', 'lt2', 'dk2','accent1','accent2','accent3','accent4','accent5',
                'accent6','hlink','folHlink'];
            foreach ($colors12 as $c) {
                $v = $this->xlsx->theme->themeElements->clrScheme->{$c};
                if (isset($v->sysClr)) {
                    $this->themeColors[] = (string) $v->sysClr['lastClr'];
                } elseif (isset($v->srgbClr)) {
                    $this->themeColors[] = (string) $v->srgbClr['val'];
                } else {
                    $this->themeColors[] = null;
                }
            }
        }
    }


    public function readFonts()
    {
        // fonts
        $this->fonts = [];
        if (isset($this->xlsx->styles->fonts->font)) {
            foreach ($this->xlsx->styles->fonts->font as $v) {
                $u = '';
                if (isset($v->u)) {
                    $u = isset($v->u['val']) ? (string) $v->u['val'] : 'single';
                }
                $f = [
                    'b' => isset($v->b) && ($v->b['val'] === null || $v->b['val']),
                    'i' => isset($v->i) && ($v->i['val'] === null || $v->i['val']),
                    'u' => $u,
                    'strike' => isset($v->strike) && ($v->strike['val'] === null || $v->strike['val']),
                    'sz' => isset($v->sz['val']) ? (int) $v->sz['val'] : 11,
                    'color' => $this->getColorValue($v->color),
                    'name' => isset($v->name['val']) ? (string) $v->name['val'] : 'Calibri',
                    'family' => isset($v->family['val']) ? (int) $v->family['val'] : 2,
                    'charset' => isset($v->charset['val']) ? (int) $v->charset['val'] : 1,
                    'scheme' => isset($v->scheme['val']) ? (string) $v->scheme['val'] : 'minor'
                ];
                $this->fonts[] = $f;
            }
        }
    }
    public function readFills()
    {
        // fills
        $this->fills = [];
        if (isset($this->xlsx->styles->fills->fill)) {
            foreach ($this->xlsx->styles->fills->fill as $v) {
                if (isset($v->patternFill)) {
                    $this->fills[] = [
                        'pattern' => isset($v->patternFill['patternType']) ? (string) $v->patternFill['patternType'] : 'none',
                        'fgcolor' => $this->getColorValue($v->patternFill->fgColor),
                        'bgcolor' => $this->getColorValue($v->patternFill->bgColor)
                    ];
                }
            }
        }
    }
    public function readBorders()
    {
        $this->borders = [];
        if (isset($this->xlsx->styles->borders->border)) {
            foreach ($this->xlsx->styles->borders->border as $v) {
                $this->borders[] = [
                    'left' => [
                        'style' => (string) $v->left['style'],
                        'color' => $this->getColorValue($v->left->color)
                    ],
                    'right' => [
                        'style' => (string) $v->right['style'],
                        'color' => $this->getColorValue($v->right->color)
                    ],
                    'top' => [
                        'style' => (string) $v->top['style'],
                        'color' => $this->getColorValue($v->top->color)
                    ],
                    'bottom' => [
                        'style' => (string) $v->bottom['style'],
                        'color' => $this->getColorValue($v->bottom->color)
                    ],
                    'diagonal' => [
                        'style' => (string) $v->diagonal['style'],
                        'color' => $this->getColorValue($v->diagonal->color)
                    ],
                    'horizontal' => [
                        'style' => (string) $v->horizontal['style'],
                        'color' => $this->getColorValue($v->horizontal->color)
                    ],
                    'vertical' => [
                        'style' => (string) $v->vertical['style'],
                        'color' => $this->getColorValue($v->vertical->color)
                    ],
                    'diagonalUp' => (bool) $v['diagonalUp'],
                    'diagonalDown' => (bool) $v['diagonalDown'],
                    'outline' => !(isset($v['outline'])) || $v['outline']
                ];
            }
        }
    }

    public function readXfs()
    {
        // cellStyles
        $this->cellStyles = [];
        if (isset($this->xlsx->styles->cellStyleXfs->xf)) {
            foreach ($this->xlsx->styles->cellStyleXfs->xf as $v) {
                $x = [];
                foreach ($v->attributes() as $k1 => $v1) {
                    $x[ $k1 ] = (int) $v1;
                }
                if (isset($v->alignment)) {
                    foreach ($v->alignment->attributes() as $k1 => $v1) {
                        $x['alignment'][$k1] = (string) $v1;
                    }
                }
                $this->cellStyles[] = $x;
            }
        }
        // css
        $this->css = [];
        // xf
        if (isset($this->xlsx->styles->cellXfs->xf)) {
            $k = 0;
            foreach ($this->xlsx->styles->cellXfs->xf as $v) {
                $cf = &$this->xlsx->cellFormats[$k];

                // alignment
                $alignment = [];
                if (isset($v->alignment)) {
                    foreach ($v->alignment->attributes() as $k1 => $v1) {
                        $alignment[$k1] = (string)$v1;
                    }
                }

                if (isset($cf['xfId'], $this->cellStyles[ $cf['xfId'] ])) {
                    $s = $this->cellStyles[$cf['xfId']];
                    if (!empty($s['applyNumberFormat'])) {
                        $cf['numFmtId'] = $s['numFmtId'];
                    }
                    if (!empty($s['applyFont'])) {
                        $cf['fontId'] = $s['fontId'];
                    }
                    if (!empty($s['applyBorder'])) {
                        $cf['borderId'] = $s['borderId'];
                    }
                    if (!empty($s['applyAlignment']) && !empty($s['alignment'])) {
                        $alignment = $s['alignment'];
                    }
                }
                $cf['alignment'] = $alignment;

                $align = null;
                if (isset($alignment['horizontal'])) {
                    $align = $alignment['horizontal'];
                    if ($align === 'centerContinuous') {
                        $align = 'center';
                    }
                    if ($align === 'distributed') {
                        $align = 'justify';
                    }
                    if ($align === 'general') {
                        $align = null;
                    }
                }
                $cf['align'] = $align;

                $valign = null;
                if (isset($alignment['vertical'])) {
                    $valign = $alignment['vertical'];
                    if ($valign === 'center' || $valign === 'distributed' || $valign === 'justify') {
                        $valign = 'middle';
                    }
                }
                $cf['valign'] = $valign;

                // font
                if (isset($cf['fontId'])) {
                    $cf['font'] = $this->fonts[$cf['fontId']]['name'];
                    $cf['color'] = $this->fonts[$cf['fontId']]['color'];
                    $cf['f-size'] = $this->fonts[$cf['fontId']]['sz'];
                    $cf['f-b'] = $this->fonts[$cf['fontId']]['b'];
                    $cf['f-i'] = $this->fonts[$cf['fontId']]['i'];
                    $cf['f-u'] = $this->fonts[$cf['fontId']]['u'];
                    $cf['f-strike'] = $this->fonts[$cf['fontId']]['strike'];
                } else {
                    $cf['font'] = null;
                    $cf['color'] = null;
                    $cf['f-size'] = null;
                    $cf['f-b'] = null;
                    $cf['f-i'] = null;
                    $cf['f-u'] = null;
                    $cf['f-strike'] = null;
                }

                // fill
                $cf['bgcolor'] = isset($cf['fillId']) ? $this->fills[ $cf['fillId'] ]['fgcolor'] : null;

                // borders
                if (isset($cf['borderId'], $this->borders[ $cf['borderId'] ])) {
                    $border = $this->borders[ $cf['borderId'] ];

                    $borders = ['left', 'right', 'top', 'bottom'];
                    foreach ($borders as $b) {
                        $cf['b-' . $b.'-color'] = $border[$b]['color'];
                        if ($border[$b]['style'] === '' || $border[$b]['style'] === 'none') {
                            $cf['b-' . $b.'-style'] = '';
                            $cf['b-' . $b.'-color'] = '';
                        } elseif ($border[$b]['style'] === 'dashDot'
                            || $border[$b]['style'] === 'dashDotDot'
                            || $border[$b]['style'] === 'dashed'
                        ) {
                            $cf['b-' . $b.'-style'] = 'dashed';
                        } else {
                            $cf['b-' . $b.'-style'] = 'solid';
                        }
                    }
                } else {
                    $cf['b-top-style'] = null;
                    $cf['b-right-style'] = null;
                    $cf['b-bottom-style'] = null;
                    $cf['b-left-style'] = null;
                }

                $css = '';

                if ($cf['color']) {
                    $css .= 'color: #'.$cf['color'].';';
                }
                if ($cf['font']) {
                    $css .= 'font-family: '.$cf['font'].';';
                }
                if ($cf['f-size']) {
//                    $css .= 'font-size: '.($cf['f-size'] * 0.352806).'mm;';
                    $css .= 'font-size: '.(round($cf['f-size'] * 1.3333) + 2).'px;';
                }
                if ($cf['f-b']) {
                    $css .= 'font-weight: bold;';
                }
                if ($cf['f-i']) {
                    $css .= 'font-style: italic;';
                }
                if ($cf['f-u']) {
                    $css .= 'text-decoration: underline;';
                }
                if ($cf['f-strike']) {
                    $css .= 'text-decoration: line-through;';
                }
                if ($cf['bgcolor']) {
                    $css .= 'background-color: #' . $cf['bgcolor'] . ';';
                }
                if ($cf['align']) {
                    $css .= 'text-align: '.$cf['align'].';';
                }
                if ($cf['valign']) {
                    $css .= 'vertical-align: '.$cf['valign'].';';
                }
                if ($cf['b-top-style']) {
                    $css .= 'border-top-style: '.$cf['b-top-style'].';';
                    $css .= 'border-top-color: #'.$cf['b-top-color'].';';
                    $css .= 'border-top-width: thin;';
                }
                if ($cf['b-right-style']) {
                    $css .= 'border-right-style: '.$cf['b-right-style'].';';
                    $css .= 'border-right-color: #'.$cf['b-right-color'].';';
                    $css .= 'border-right-width: thin;';
                }
                if ($cf['b-bottom-style']) {
                    $css .= 'border-bottom-style: '.$cf['b-bottom-style'].';';
                    $css .= 'border-bottom-color: #'.$cf['b-bottom-color'].';';
                    $css .= 'border-bottom-width: thin;';
                }
                if ($cf['b-left-style']) {
                    $css .= 'border-left-style: '.$cf['b-left-style'].';';
                    $css .= 'border-left-color: #'.$cf['b-left-color'].';';
                    $css .= 'border-left-width: thin;';
                }
                $this->css[$k] = $css;
                
                $k++;
            }
        }
    }
    public function readComments()
    {
        $this->comments = [];
        foreach ($this->xlsx->sheetRels as $index => $xml) {
            foreach ($xml->Relationship as $rel) {
                $rel_type = basename(trim((string) $rel['Type']));
                $rel_target = (string) $rel['Target'];
                if ($rel_type === 'comments') {
                    $d = dirname($this->xlsx->sheetFiles[$index]);
                    $com_file = SimpleXLSX::getTarget($d, $rel_target);
                    if ($com_xml = $this->xlsx->getEntryXML($com_file)) {
                        foreach ($com_xml->commentList->comment as $com) {
                            $this->comments[$index][(string)$com['ref']] = SimpleXLSX::parseRichText($com->text);
                        }
                    }
                }
            }
        }
    }

    public function readRowsEx($worksheetIndex = 0, $limit = 0)
    {
        if (($ws = $this->xlsx->worksheet($worksheetIndex)) === false) {
            return;
        }
        $this->worksheetIndex = $worksheetIndex;

        $dim = $this->xlsx->dimension($worksheetIndex);
        $numCols = $dim[0];
        $numRows = $dim[1];

        /*$emptyRow = array();
        for ($i = 0; $i < $numCols; $i++) {
            $emptyRow[] = null;
        }
        */
        $cols = [];
        for ($i = 0; $i < $numCols; $i++) {
            $cols[] = ['s' => 0, 'hidden' => false, 'width' => 0];
        }
//        $hiddenCols = [];
        /* @var SimpleXMLElement $ws */
        if (isset($ws->cols)) {
            foreach ($ws->cols->col as $col) {
                $min = (int)$col['min'];
                $max = (int)$col['max'];
                if (($max-$min) > 100) {
                    $max = $min;
                }
                for ($i = $min; $i <= $max; $i++) {
                    $cols[$i-1] = [
                        's' => (int)$col['style'],
                        'hidden' => (bool)$col['hidden'],
                        'width' => $col['customWidth'] ? (float) $col['width'] : 0
                    ];
                }
            }
        }

        $curR = 0;
        $_limit = $limit;

        foreach ($ws->sheetData->row as $row) {
            $curC = 0;

            $r_idx = (int)$row['r'];
            $r_style = ['s' => 0, 'hidden' => (bool)$row['hidden'], 'height' => 0];
            if ($row['customFormat']) {
                $r_style['s'] = (int)$row['s'];
            }
            if ($row['customHeight']) {
                $r_style['height'] = (int)$row['ht'];
            }

            $cells = [];
            for ($i = 0; $i < $numCols; $i++) {
                $cells[] = null;
            }

            foreach ($row->c as $c) {
                $r = (string)$c['r'];
                $t = (string)$c['t'];
                $s = (int)$c['s'];

                $idx = $this->xlsx->getIndex($r);
                $x = $idx[0];
                $y = $idx[1];

                if ($x > -1) {
                    $curC = $x;
                    if ($curC >= $numCols) {
                        $numCols = $curC + 1;
                    }
                    while ($curR < $y) {
                        $emptyRow = [];
                        for ($i = 0; $i < $numCols; $i++) {
                            $emptyRow[] = $this->valueEx($cols[$i], $i, $curR);
                        }
                        yield $emptyRow;
                        $curR++;

                        $_limit--;
                        if ($_limit === 0) {
                            return;
                        }
                    }
                }

                $data = [
                    'type' => $t,
                    'name' => $r,
                    'value' => $this->xlsx->value($c),
                    'f' => (string)$c->f,
                    'r' => $r_idx,
                    's' => ($s > 0) ? $s : $cols[$curC]['s'],
                    'hidden' => $r_style['hidden'] || $cols[$curC]['hidden'],
                    'width' => $cols[$curC]['width'],
                    'height' => $r_style['height']
                ];
                $cells[$curC] = $this->valueEx($data, $curC, $curR);

                $curC++;
            }
            // check empty cells
            for ($i = 0; $i < $numCols; $i++) {
                if ($cells[$i] === null) {
                    if ($r_style['s'] > 0) {
                        $data = $r_style;
                    } else {
                        $data = $cols[$i];
                    }
                    $data['width'] = $cols[$i]['width'];
                    $data['height'] = $r_style['height'];
                    $cells[$i] = $this->valueEx($data, $i, $curR);
                }
            }

            yield $cells;

            $curR++;
            $_limit--;
            if ($_limit === 0) {
                break;
            }
        }

        while ($curR < $numRows) {
            $emptyRow = [];
            for ($i = 0; $i < $numCols; $i++) {
                $data = $cols[$i];
                $emptyRow[] = $this->valueEx($data, $i, $curR);
            }
            yield $emptyRow;
            $curR++;
            $_limit--;
            if ($_limit === 0) {
                return;
            }
        }
    }

    protected function valueEx($data, $x = null, $y = null)
    {

        $r = [
            'type' => '',
            'name' => '',
            'value' => '',
            'href' => '',
            'f' => '',
            'format' => '',
            's' => 0,
            'css' => '',
            'r' => '',
            'hidden' => false,
            'width' => 0,
            'height' => 0,
        ];
        foreach ($data as $k => $v) {
            if (isset($r[$k])) {
                $r[$k] = $v;
            }
        }
        $st = &$this->xlsx->cellFormats[$r['s']];
        $r['format'] = $st['format'];
        $r['css'] = &$this->css[ $r['s'] ];
        if ($r['value'] !== '' && !$st['align'] && !in_array($r['type'], ['s','str','inlineStr','e'], true)) {
            $r['css'] .= 'text-align: right;';
        }

        if (!$r['name']) {
            $c = '';
            for ($k = $x; $k >= 0; $k = (int)($k / 26) - 1) {
                $c = chr($k % 26 + 65) . $c;
            }
            $r['name'] = $c . ($y + 1);
            $r['r'] = $y+1;
        }
        $r['href'] = isset($this->hyperlinks[$this->worksheetIndex][$r['name']]) ? $this->hyperlinks[$this->worksheetIndex][$r['name']] : '';
        $r['comment'] = isset($this->comments[$this->worksheetIndex][$r['name']]) ? $this->comments[$this->worksheetIndex][$r['name']] : '';

        return $r;
    }
    public function getColorValue(SimpleXMLElement $a = null, $default = '')
    {
        if ($a === null) {
            return $default;
        }
        $c = $default; // auto
        if ($a['rgb'] !== null) {
            $c = substr((string) $a['rgb'], 2); // FFCCBBAA -> CCBBAA
        } elseif ($a['indexed'] !== null && isset(static::$IC[ (int) $a['indexed'] ])) {
            $c = static::$IC[ (int) $a['indexed'] ];
        } elseif ($a['theme'] !== null && isset($this->themeColors[ (int) $a['theme'] ])) {
            $c = $this->themeColors[ (int) $a['theme'] ];
        }
        if ($a['tint'] !== null) {
            list($r,$g,$b) = array_map('hexdec', str_split($c, 2));
            $tint = (float) $a['tint'];
            if ($tint > 0) {
                $r += (255 - $r) * $tint;
                $g += (255 - $g) * $tint;
                $b += (255 - $b) * $tint;
            } else {
                $r += $r * $tint;
                $g += $g * $tint;
                $b += $b * $tint;
            }
            $c = strtoupper(
                str_pad(dechex((int) $r), 2, '0', 0) .
                str_pad(dechex((int) $g), 2, '0', 0) .
                str_pad(dechex((int) $b), 2, '0', 0)
            );
        }
        return $c;
    }

    public function readHyperlinks()
    {
        $this->hyperlinks = [];
        foreach ($this->xlsx->sheetRels as $index => $xml) {
            $sheet = $this->xlsx->sheets[$index];
            $link_ids = [];
            // hyperlink
            foreach ($xml->Relationship as $rel) {
                $rel_type = basename(trim((string)$rel['Type']));
                $rel_target = (string)$rel['Target'];
                if ($rel_type === 'hyperlink') {
                    $rel_id = (string)$rel['Id'];
                    $link_ids[$rel_id] = $rel_target;
                }
            }

            if (isset($sheet->hyperlinks)) {
                foreach ($sheet->hyperlinks->hyperlink as $hyperlink) {
                    $ref = (string)$hyperlink['ref'];
                    if (SimpleXLSX::strpos($ref, ':') > 0) { // A1:A8 -> A1
                        $ref = explode(':', $ref);
                        $ref = $ref[0];
                    }
                    $loc = (string)$hyperlink['location'];
                    $id = (string)$hyperlink['id'];
                    if ($id) {
                        $href = $link_ids[$id] . ($loc ? '#' . $loc : '');
                    } else {
                        $href = $loc;
                    }
                    $this->hyperlinks[$index][$ref] = $href;
                }
            }
        }
    }
}
