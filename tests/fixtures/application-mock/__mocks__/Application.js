module.exports = ({ excel }) => {
    // Extend the injected built-in instance; do not replace it with a plain
    // object, otherwise Worksheets/Range/ThisWorkbook state is disconnected.
    excel.Application.PathSeparator = '/';
    let displayAlerts = true;
    Object.defineProperty(excel.Application, 'DisplayAlerts', {
        configurable: true,
        get: () => displayAlerts,
        set: (value) => { displayAlerts = value; },
    });
    return {
        Application: excel.Application,
    };
};
