/**
 * Excel 依存オブジェクトの一覧から __mocks__/ExcelObjects.bas のひな形を生成する。
 *
 * 使い方:
 *   generateExcelMockBas(['ActiveSheet', 'Range', 'Cells', 'Rows'])
 *   → VBA クラス・関数定義を含む .bas 文字列を返す
 *
 * 生成されたファイルは __mocks__/ に配置すると injectExcelStub のサイレント注入を
 * 部分的に上書きできる。ユーザーはセル値や戻り値を書き換えてカスタマイズする。
 */

interface GenerateOptions {
    /** 対象プロシージャ名（コメントに記載） */
    procName?: string;
    /** 生成日時（コメントに記載、省略時は today） */
    date?: string;
}

/**
 * 検出された Excel オブジェクト名の集合から .bas ひな形文字列を生成する。
 * オブジェクト名は大文字小文字を問わない（'activesheet' でも 'ActiveSheet' でも可）。
 */
export function generateExcelMockBas(objects: string[], opts: GenerateOptions = {}): string {
    const needed = new Set(objects.map(o => o.toLowerCase()));
    const date = opts.date ?? new Date().toISOString().slice(0, 10);
    const procLabel = opts.procName ? ` (${opts.procName})` : '';

    const lines: string[] = [];

    const L = (...args: string[]) => lines.push(...args);

    L(
        `' __mocks__/ExcelObjects.bas`,
        `' 自動生成: vba-runner.generateMocks ${date}${procLabel}`,
        `' 必要に応じてセル値・戻り値を書き換えてください`,
        `' このファイルは injectExcelStub のサイレント注入を上書きします。`,
        ``,
    );

    // ---- MockRange -------------------------------------------------------
    if (needed.has('cells') || needed.has('range')) {
        L(
            `Class MockRange`,
            `    Public Value As Variant   ' ← 読み取り値をここで設定`,
            `    Public Row As Long`,
            `    Public Column As Long`,
            `    ' ※ Range.End(xlUp) は VBA キーワード "End" とぶつかるため`,
            `    '   VBA スタブでは定義できません。JS/TS モックが必要な場合は`,
            `    '   __mocks__.js の __addCreateObject__ を使ってください。`,
            `End Class`,
            ``,
        );
    }

    // ---- MockRows --------------------------------------------------------
    if (needed.has('rows')) {
        L(
            `Class MockRows`,
            `    Public Count As Long`,
            `    Sub Class_Initialize()`,
            `        Count = 1048576`,
            `    End Sub`,
            `End Class`,
            ``,
        );
    }

    // ---- MockColumns -----------------------------------------------------
    if (needed.has('columns')) {
        L(
            `Class MockColumns`,
            `    Public Count As Long`,
            `    Sub Class_Initialize()`,
            `        Count = 16384`,
            `    End Sub`,
            `End Class`,
            ``,
        );
    }

    // ---- MockWorksheet ---------------------------------------------------
    const needsSheet = needed.has('activesheet') || needed.has('sheets')
        || needed.has('worksheets') || needed.has('cells') || needed.has('range');

    if (needsSheet) {
        L(`Class MockWorksheet`);
        L(`    Public Property Get Name() As String`);
        L(`        Name = "Sheet1"   ' ← シート名を変えたい場合はここを編集`);
        L(`    End Property`);
        L(``);

        if (needed.has('cells') || needed.has('range')) {
            L(
                `    Public Function Cells(row, col)`,
                `        Dim r As New MockRange`,
                `        r.Row = row`,
                `        r.Column = col`,
                `        ' r.Value = 0  ← テスト用の値はここに設定`,
                `        Set Cells = r`,
                `    End Function`,
                ``,
                `    Public Function Range(addr)`,
                `        Dim r As New MockRange`,
                `        ' r.Value = 0  ← テスト用の値はここに設定`,
                `        Set Range = r`,
                `    End Function`,
                ``,
            );
        }

        if (needed.has('rows')) {
            L(
                `    Public Property Get Rows()`,
                `        Set Rows = New MockRows`,
                `    End Property`,
                ``,
            );
        }

        if (needed.has('columns')) {
            L(
                `    Public Property Get Columns()`,
                `        Set Columns = New MockColumns`,
                `    End Property`,
                ``,
            );
        }

        L(`End Class`, ``);
    }

    // ---- MockApplication ------------------------------------------------
    if (needed.has('application')) {
        L(
            `Class MockApplication`,
            `    ' Application.ScreenUpdating / Calculation / EnableEvents のノーオプ`,
            `    Public ScreenUpdating As Boolean`,
            `    Public DisplayAlerts As Boolean`,
            `    Public EnableEvents As Boolean`,
            `    Public Calculation As Long`,
            ``,
            `    Sub Class_Initialize()`,
            `        ScreenUpdating = True`,
            `        DisplayAlerts = True`,
            `        EnableEvents = True`,
            `        Calculation = -4105  ' xlCalculationAutomatic`,
            `    End Sub`,
            `End Class`,
            ``,
            `Function Application()`,
            `    Set Application = New MockApplication`,
            `End Function`,
            ``,
        );
    }

    // ---- MockWorkbook ---------------------------------------------------
    if (needed.has('activeworkbook') || needed.has('thisworkbook') || needed.has('workbooks')) {
        L(
            `Class MockWorkbook`,
            `    Public Property Get Name() As String`,
            `        Name = "MockWorkbook.xlsx"`,
            `    End Property`,
            `    Public Property Get Path() As String`,
            `        Path = ""`,
            `    End Property`,
            `End Class`,
            ``,
        );
    }

    // ---- Global functions -----------------------------------------------
    if (needed.has('activesheet')) {
        L(
            `Function ActiveSheet()`,
            `    Set ActiveSheet = New MockWorksheet`,
            `End Function`,
            ``,
        );
    }

    if (needed.has('sheets')) {
        L(
            `Function Sheets(nameOrIndex)`,
            `    Dim ws As New MockWorksheet`,
            `    Set Sheets = ws`,
            `End Function`,
            ``,
        );
    }

    if (needed.has('worksheets')) {
        L(
            `Function Worksheets(nameOrIndex)`,
            `    Set Worksheets = New MockWorksheet`,
            `End Function`,
            ``,
        );
    }

    if (needed.has('activeworkbook')) {
        L(
            `Function ActiveWorkbook()`,
            `    Set ActiveWorkbook = New MockWorkbook`,
            `End Function`,
            ``,
        );
    }

    if (needed.has('thisworkbook')) {
        L(
            `Function ThisWorkbook()`,
            `    Set ThisWorkbook = New MockWorkbook`,
            `End Function`,
            ``,
        );
    }

    if (needed.has('workbooks')) {
        L(
            `Function Workbooks(nameOrIndex)`,
            `    Set Workbooks = New MockWorkbook`,
            `End Function`,
            ``,
        );
    }

    return lines.join('\n');
}

/**
 * vba-analyzer --json の出力から Excel 依存オブジェクト名の集合を抽出する。
 * files[].procedures[].excelObjectsUsed の全プロシージャの和集合を返す。
 */
export function extractObjectsFromAnalyzerJson(analyzerJson: any): string[] {
    const objects = new Set<string>();
    for (const file of analyzerJson?.files ?? []) {
        for (const proc of file?.procedures ?? []) {
            for (const obj of proc?.excelObjectsUsed ?? []) {
                objects.add(String(obj));
            }
        }
    }
    return [...objects].sort();
}
