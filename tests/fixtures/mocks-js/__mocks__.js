// __mocks__.js: MsgBox と VBScript.RegExp のスタブ

class VBScriptRegExp {
    constructor() {
        this.Pattern = '';
        this.Global = false;
        this.IgnoreCase = false;
    }

    _flags() {
        return (this.Global ? 'g' : '') + (this.IgnoreCase ? 'i' : '');
    }

    Test(str) {
        return new RegExp(this.Pattern, this._flags()).test(String(str));
    }

    Replace(str, replacement) {
        return String(str).replace(new RegExp(this.Pattern, this._flags()), replacement);
    }
}

module.exports = {
    MsgBox: (prompt, _buttons, _title) => {
        console.log(`[MsgBox mock] ${prompt}`);
        return 99;
    },
    __addCreateObject__: {
        'VBScript.RegExp': () => new VBScriptRegExp(),
    },
};
