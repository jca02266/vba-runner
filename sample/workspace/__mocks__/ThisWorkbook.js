'use strict';

module.exports = ({ excel }) => {
    const app = excel.Application;
    const getSheet = app.Sheets.bind(app);
    const names = new Set(['Sheet1']);
    let nextSheet = 2;
    const wrappedSheets = new WeakMap();

    const wrapSheet = (sheet) => {
        if (wrappedSheets.has(sheet)) return wrappedSheets.get(sheet);

        let currentName = sheet.name;
        const wrapped = new Proxy(sheet, {
            get(target, property) {
                return property === 'name' ? currentName : target[property];
            },
            set(target, property, value) {
                if (property === 'name') {
                    names.delete(currentName);
                    currentName = value;
                    names.add(currentName);
                }
                target[property] = value;
                sheets.count = names.size;
                return true;
            },
        });
        wrappedSheets.set(sheet, wrapped);
        return wrapped;
    };

    const sheets = function (nameOrIndex) {
        const name = typeof nameOrIndex === 'string'
            ? nameOrIndex
            : `Sheet${nameOrIndex}`;
        if (!names.has(name)) {
            throw new Error(`Worksheet not found: ${name}`);
        }
        return wrapSheet(getSheet(name));
    };

    sheets.count = names.size;
    sheets.add = function () {
        const sheet = wrapSheet(getSheet(`Sheet${nextSheet++}`));
        names.add(sheet.name);
        sheets.count = names.size;
        return sheet;
    };
    // VBA's Sheets.Add accepts named Before/After arguments.  JavaScript
    // parameter names are not available to the runner, so publish the VBA
    // contract explicitly for named-argument validation.
    sheets.add.__vbaParamSpec__ = [
        { name: 'Before', optional: true },
        { name: 'After', optional: true },
    ];

    return {
        // Excel.XlDirection.xlUp. Host constants are normalized to lowercase
        // by vba-runner, just like object and member names.
        xlup: -4162,
        ThisWorkbook: {
            Sheets: sheets,
        },
    };
};
