/**
 * VBA Format() 関数の実装
 * MS-VBAL §6.1.2.11.1.8 / VBA Language Reference: Format Function
 */

// ------------------------------------------------------------------ //
// セクション分割（; 区切り、\x / "text" エスケープを考慮）
// ------------------------------------------------------------------ //
export function splitFormatSections(pattern: string): string[] {
    const sections: string[] = [];
    let cur = '';
    let i = 0;
    while (i < pattern.length) {
        if (pattern[i] === '\\') {
            cur += pattern[i++];
            if (i < pattern.length) cur += pattern[i++];
        } else if (pattern[i] === '"') {
            const start = i++;
            while (i < pattern.length && pattern[i] !== '"') i++;
            cur += pattern.slice(start, i + 1);
            if (i < pattern.length) i++;
        } else if (pattern[i] === ';') {
            sections.push(cur); cur = ''; i++;
        } else {
            cur += pattern[i++];
        }
    }
    sections.push(cur);
    return sections;
}

// ------------------------------------------------------------------ //
// 文字列フォーマット（@, &, <, >, !）
// ------------------------------------------------------------------ //
export function formatString(s: string, fmt: string): string {
    if (!fmt) return s;
    const sections = splitFormatSections(fmt);
    // 2 セクション: 第1は通常文字列、第2は null/空文字列
    const section = sections.length >= 2 && s === '' ? sections[1] : sections[0];
    if (!section) return s;

    // 大文字 / 小文字変換（単独でもパターン中に含まれていても有効）
    if (section.includes('>')) return s.toUpperCase();
    if (section.includes('<')) return s.toLowerCase();

    // @, & プレースホルダー
    const hasExcl = section.includes('!');
    const phs: Array<'@' | '&'> = [];
    for (const ch of section) {
        if (ch === '@') phs.push('@');
        else if (ch === '&') phs.push('&');
    }
    if (phs.length === 0) return s;

    const chars = [...s]; // Unicode 対応
    let result = '';
    if (hasExcl) {
        // ! 付き: 左から右へ埋める
        let ci = 0;
        for (const ph of phs) {
            if (ci < chars.length) result += chars[ci++];
            else if (ph === '@') result += ' '; // @ は空スペース、& は何も出力しない
        }
    } else {
        // デフォルト: 右から左へ埋める
        let ci = chars.length - 1;
        for (let j = phs.length - 1; j >= 0; j--) {
            if (ci >= 0) result = chars[ci--] + result;
            else if (phs[j] === '@') result = ' ' + result;
        }
    }
    return result;
}

// ------------------------------------------------------------------ //
// 日付フォーマット
// ------------------------------------------------------------------ //
export function formatDate(d: Date, pattern: string): string {
    const pLower = pattern.toLowerCase();

    const pad2 = (n: number) => String(n).padStart(2, '0');
    const yyyy = String(d.getFullYear());
    const MM   = pad2(d.getMonth() + 1);
    const dd   = pad2(d.getDate());
    const HH   = pad2(d.getHours());
    const mm   = pad2(d.getMinutes());
    const ss   = pad2(d.getSeconds());

    const months      = ['January','February','March','April','May','June','July','August','September','October','November','December'];
    const monthsShort = ['Jan','Feb','Mar','Apr','May','Jun','Jul','Aug','Sep','Oct','Nov','Dec'];
    const days        = ['Sunday','Monday','Tuesday','Wednesday','Thursday','Friday','Saturday'];
    const daysShort   = ['Sun','Mon','Tue','Wed','Thu','Fri','Sat'];

    const h12  = d.getHours() % 12 || 12;
    const ampm = d.getHours() >= 12 ? 'PM' : 'AM';

    // 名前付きフォーマット（VBA 名前付き日付書式。ロケール非依存の固定英語フォーマット）
    if (pLower === 'general date') {
        const hasTime = d.getHours() !== 0 || d.getMinutes() !== 0 || d.getSeconds() !== 0;
        return hasTime ? `${yyyy}/${MM}/${dd} ${HH}:${mm}:${ss}` : `${yyyy}/${MM}/${dd}`;
    }
    if (pLower === 'long date')   return `${days[d.getDay()]}, ${months[d.getMonth()]} ${d.getDate()}, ${yyyy}`;
    if (pLower === 'medium date') return `${dd}-${monthsShort[d.getMonth()]}-${yyyy.slice(-2)}`;
    if (pLower === 'short date')  return `${yyyy}/${MM}/${dd}`;
    if (pLower === 'long time')   return `${HH}:${mm}:${ss}`;
    if (pLower === 'medium time') return `${h12}:${mm} ${ampm}`;
    if (pLower === 'short time')  return `${HH}:${mm}`;

    // 年の通算日（1〜366）
    const dayOfYear = (): number => {
        const jan1 = new Date(d.getFullYear(), 0, 1);
        return Math.floor((d.getTime() - jan1.getTime()) / 86400000) + 1;
    };

    // 年の通算週（1〜54、日曜始まり・1/1 を含む週 = 第1週）
    const weekOfYear = (): number => {
        const jan1DayOfWeek = new Date(d.getFullYear(), 0, 1).getDay();
        return Math.ceil((dayOfYear() + jan1DayOfWeek) / 7);
    };

    // ddddd（短い日付）/ dddddd（長い日付）/ ttttt（長い時刻）/ c（汎用日付時刻）用ヘルパー
    const shortDateStr = () => `${d.getMonth() + 1}/${d.getDate()}/${yyyy.slice(-2)}`;
    const longDateStr  = () => `${months[d.getMonth()]} ${d.getDate()}, ${yyyy}`;
    const longTimeStr  = () => `${d.getHours()}:${mm}:${ss}`;
    const generalDateStr = () => {
        const hasTime = d.getHours() !== 0 || d.getMinutes() !== 0 || d.getSeconds() !== 0;
        const isBaseDate = d.getFullYear() === 1899 && d.getMonth() === 11 && d.getDate() === 30;
        if (!isBaseDate && hasTime) return `${shortDateStr()} ${longTimeStr()}`;
        if (!isBaseDate) return shortDateStr();
        return longTimeStr();
    };

    // トークン化（\x / "text" エスケープ対応、長いトークンを先に）
    const tokenRe = /\\.|"[^"]*"|ampm|am\/pm|a\/p|ttttt|dddddd|ddddd|dddd|ddd|dd|d|yyyy|yy|mmmm|mmm|mm|m|ww|w|q|y|hh|h|nn|n|ss|s|c|[^a-zA-Z\\"]+|[a-zA-Z]/gi;
    const toks = pattern.match(tokenRe) || [];

    // AM/PM 指定がある場合は 12 時間制を使う
    const use12Hour = toks.some(t => /^(am\/pm|ampm|a\/p)$/i.test(t));
    const hourPad  = use12Hour ? pad2(h12) : HH;
    const hourNoP  = use12Hour ? String(h12) : String(d.getHours());

    let prevHour = false;

    return toks.map(tok => {
        const tl = tok.toLowerCase();

        // エスケープ文字 / 引用符リテラルは prevHour をリセットしない（hh:mm の : と同様）
        if (tok.length === 2 && tok[0] === '\\') return tok[1];
        if (tok.startsWith('"') && tok.endsWith('"')) return tok.slice(1, -1);

        switch (tl) {
            case 'yyyy':   prevHour = false; return yyyy;
            case 'yy':     prevHour = false; return yyyy.slice(-2);
            case 'y':      prevHour = false; return String(dayOfYear());
            case 'mmmm':   prevHour = false; return months[d.getMonth()];
            case 'mmm':    prevHour = false; return monthsShort[d.getMonth()];
            case 'mm': {
                const isMins = prevHour; prevHour = false;
                return isMins ? mm : pad2(d.getMonth() + 1);
            }
            case 'm': {
                const isMins = prevHour; prevHour = false;
                return isMins ? String(d.getMinutes()) : String(d.getMonth() + 1);
            }
            case 'dddddd': prevHour = false; return longDateStr();
            case 'ddddd':  prevHour = false; return shortDateStr();
            case 'dddd':   prevHour = false; return days[d.getDay()];
            case 'ddd':    prevHour = false; return daysShort[d.getDay()];
            case 'dd':     prevHour = false; return pad2(d.getDate());
            case 'd':      prevHour = false; return String(d.getDate());
            case 'ww':     prevHour = false; return String(weekOfYear());
            case 'w':      prevHour = false; return String(d.getDay() + 1); // 日曜=1
            case 'q':      prevHour = false; return String(Math.ceil((d.getMonth() + 1) / 3));
            case 'hh':     prevHour = true;  return hourPad;
            case 'h':      prevHour = true;  return hourNoP;
            case 'nn':     prevHour = false; return mm;
            case 'n':      prevHour = false; return String(d.getMinutes());
            case 'ss':     prevHour = false; return ss;
            case 's':      prevHour = false; return String(d.getSeconds());
            case 'ttttt':  prevHour = false; return longTimeStr();
            case 'am/pm':  prevHour = false; return tok[0] === 'A' ? ampm : ampm.toLowerCase();
            case 'a/p':    prevHour = false; return tok[0] === 'A' ? ampm[0] : ampm[0].toLowerCase();
            case 'ampm':   prevHour = false; return ampm;
            case 'c':      prevHour = false; return generalDateStr();
            // リテラル文字（:, /, スペース, 不明な英字）は prevHour をリセットしない
            default:       return tok;
        }
    }).join('');
}

// ------------------------------------------------------------------ //
// 数値フォーマット（セクション選択 → セクション書式化）
// ------------------------------------------------------------------ //
export function formatNumber(n: number, pattern: string): string {
    const pLower = pattern.toLowerCase();

    // 千の位区切り付き文字列ヘルパー
    const withThousands = (num: number, decimals: number): string => {
        const fixed = Math.abs(num).toFixed(decimals);
        const [intPart, decPart] = fixed.split('.');
        const grouped = intPart.replace(/\B(?=(\d{3})+(?!\d))/g, ',');
        const sign = num < 0 ? '-' : '';
        return decPart !== undefined ? `${sign}${grouped}.${decPart}` : `${sign}${grouped}`;
    };

    // 名前付きフォーマット
    switch (pLower) {
        case 'general number': return String(n);
        case 'currency':   return (n < 0 ? '-$' : '$') + withThousands(n, 2).replace(/^-/, '');
        case 'fixed':      return n.toFixed(2);
        case 'standard':   return withThousands(n, 2);
        case 'percent':    return (n * 100).toFixed(2) + '%';
        case 'scientific': return n.toExponential(2).replace(/e([+-])(\d+)$/, (_, sign, d) => 'E' + sign + d.padStart(2, '0'));
        case 'true/false': return n !== 0 ? 'True' : 'False';
        case 'yes/no':     return n !== 0 ? 'Yes' : 'No';
        case 'on/off':     return n !== 0 ? 'On' : 'Off';
    }

    // セクション分割（正/負/ゼロ/Null の最大4セクション）
    const sections = splitFormatSections(pattern);

    let section: string;
    let autoNegSign: boolean;

    if (sections.length === 1) {
        section = sections[0];
        autoNegSign = n < 0;
    } else if (sections.length === 2) {
        if (n >= 0) {
            section = sections[0]; autoNegSign = false;
        } else {
            // 負セクションが空 → 正セクションを流用して自動符号付き
            section = sections[1] !== '' ? sections[1] : sections[0];
            autoNegSign = sections[1] === '';
        }
    } else {
        if (n > 0)      { section = sections[0]; autoNegSign = false; }
        else if (n < 0) {
            section = sections[1] !== '' ? sections[1] : sections[0];
            autoNegSign = sections[1] === '';
        } else {
            // n === 0（第3セクション。空なら正セクション流用）
            section = sections[2] !== '' ? sections[2] : sections[0];
            autoNegSign = false;
        }
        // 第4セクション（Null）は formatFunc で null/Empty を弾いているので到達しない
    }

    return formatNumberSection(Math.abs(n), section, autoNegSign);
}

// 1セクション分の数値書式化（\x / "text" / リテラル文字 / E+E- 対応）
function formatNumberSection(absN: number, section: string, addNegSign: boolean): string {
    type Tok = { k: 'lit'; v: string } | { k: 'fmt'; v: string };
    const toks: Tok[] = [];
    let i = 0;
    while (i < section.length) {
        const c = section[i];
        if (c === '\\') {
            i++;
            toks.push({ k: 'lit', v: i < section.length ? section[i++] : '' });
        } else if (c === '"') {
            i++;
            let lit = '';
            while (i < section.length && section[i] !== '"') lit += section[i++];
            if (i < section.length) i++;
            toks.push({ k: 'lit', v: lit });
        } else if ('0#.,'.includes(c)) {
            toks.push({ k: 'fmt', v: c }); i++;
        } else if (c === '%') {
            toks.push({ k: 'fmt', v: '%' }); i++;
        } else if ((c === 'E' || c === 'e') && i + 1 < section.length && '+-'.includes(section[i + 1])) {
            // 科学記数法: E+ / E- / e+ / e- の後に続く 0/# を込みで1トークンにする
            let sci = c + section[i + 1]; i += 2;
            let expDigits = '';
            while (i < section.length && '0#'.includes(section[i])) expDigits += section[i++];
            toks.push({ k: 'fmt', v: sci + expDigits });
        } else {
            // $, -, +, (, ), スペース, その他の文字はすべてリテラル
            toks.push({ k: 'lit', v: c }); i++;
        }
    }

    const firstFmt = toks.findIndex(t => t.k === 'fmt');
    let lastFmt = -1;
    for (let j = toks.length - 1; j >= 0; j--) {
        if (toks[j].k === 'fmt') { lastFmt = j; break; }
    }

    // 書式指定子なし（\Z\e\r\o のような純リテラルセクション）
    if (firstFmt === -1) {
        return (addNegSign ? '-' : '') + toks.map(t => t.v).join('');
    }

    const prefix   = toks.slice(0, firstFmt).map(t => t.v).join('');
    const coreToks = toks.slice(firstFmt, lastFmt + 1);
    const suffix   = toks.slice(lastFmt + 1).map(t => t.v).join('');
    const coreFmt  = coreToks.map(t => t.v).join('');

    // ---- 科学記数法 ----
    const sciMatch = coreFmt.match(/^([\d#,.%]*)([Ee][+\-])(0+|#+)$/);
    if (sciMatch) {
        const mFmt    = sciMatch[1] || '0';
        const sciTok  = sciMatch[2]; // e.g. "E+" or "e-"
        const eFmt    = sciMatch[3]; // e.g. "00"
        let   n       = absN;
        const pct     = mFmt.includes('%');
        if (pct) n *= 100;

        const exp      = n === 0 ? 0 : Math.floor(Math.log10(n));
        const mantissa = n === 0 ? 0 : n / Math.pow(10, exp);

        const mf       = mFmt.replace(/%/g, '').replace(/,/g, '');
        const mParts   = mf.split('.');
        const mIntFmt  = mParts[0] || '';
        const mDecFmt  = mParts[1] || '';
        const mMaxDec  = mDecFmt.length;
        const mMinDec  = (mDecFmt.match(/0/g) || []).length;
        const mMinInt  = (mIntFmt.match(/0/g) || []).length;
        const mFixed   = mantissa.toFixed(mMaxDec);
        const [mInt, mDec] = mFixed.split('.');
        const mPadded  = mInt.padStart(mMinInt, '0');
        const mDisplay = mMinInt === 0 && mPadded === '0' ? '' : mPadded;
        const mDecDisp = mDec
            ? mDec.replace(/0+$/, '').padEnd(mMinDec, '0') || (mMinDec > 0 ? '0'.repeat(mMinDec) : '')
            : '';
        const mStr = mDisplay + (mDecDisp ? '.' + mDecDisp : '');

        const expWidth = eFmt.length;
        const expAbs   = String(Math.abs(exp)).padStart(expWidth, '0');
        const expSign  = exp < 0 ? '-' : (sciTok[1] === '+' ? '+' : '');
        const eLetter  = sciTok[0]; // E or e: 書式の大小文字を保持
        const pctStr   = pct ? '%' : '';

        return (addNegSign ? '-' : '') + prefix + mStr + pctStr + eLetter + expSign + expAbs + suffix;
    }

    // ---- 通常の数値書式 ----
    let n = absN;
    const isPercent = coreFmt.includes('%');
    if (isPercent) n *= 100;
    const workFmt = coreFmt.replace(/%/g, '');

    // , が数字プレースホルダーの後に続く場合のみ千の位区切りとして扱う
    const hasThousands = /,[0#]/.test(workFmt);

    const stripped     = workFmt.replace(/,/g, '');
    const parts        = stripped.split('.');
    const intFormatPart = parts[0];
    const decimalPart   = parts[1] || '';

    const minIntegers  = (intFormatPart.match(/0/g) || []).length;
    const hashOnlyInt  = minIntegers === 0; // # のみ → ゼロの整数部を省略
    const minDecimals  = (decimalPart.match(/0/g) || []).length;
    const maxDecimals  = decimalPart.length;

    const absFixed = n.toFixed(maxDecimals);
    const [intPart, decPart] = absFixed.split('.');
    const intPadded   = intPart.padStart(minIntegers, '0');
    const intDisplay  = hashOnlyInt && intPadded === '0' ? '' : intPadded;
    const intFormatted = hasThousands
        ? intDisplay.replace(/\B(?=(\d{3})+(?!\d))/g, ',')
        : intDisplay;
    const decFormatted = decPart !== undefined
        ? decPart.replace(/0+$/, '').padEnd(minDecimals, '0') || (minDecimals > 0 ? '0'.repeat(minDecimals) : '')
        : '';

    const sign   = addNegSign ? '-' : '';
    const numStr = intFormatted + (decFormatted ? '.' + decFormatted : '');
    const pctStr = isPercent ? '%' : '';

    return sign + prefix + numStr + pctStr + suffix;
}
