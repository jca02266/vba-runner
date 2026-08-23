const path = require('node:path');

// Extend the injected built-in Excel application so Worksheets/Range state is
// shared with the other host globals instead of being replaced by a plain
// object. The fallback keeps this fixture usable without excelStub as well.
module.exports = (context = {}) => {
    const app = context.excel?.Application ?? {
        PathSeparator: '/',
        DisplayAlerts: true,
        Worksheets: () => ({ Name: 'Sheet1' }),
    };
    app.PathSeparator = '/';
    return {
        Application: app,
        Worksheets: typeof app.Worksheets === 'function'
            ? app.Worksheets.bind(app)
            : () => ({ Name: 'Sheet1' }),
        ThisWorkbook: {
            Path: context.sourceDirectory ?? path.dirname(__dirname),
        },
    };
};

// Keep the host identifier visible to the static VS Code mock scanner. The
// runtime factory above supplies the actual bound method.
module.exports.Worksheets = function Worksheets() {};
