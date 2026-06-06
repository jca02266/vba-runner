module.exports = {
    MsgBox: (prompt, _buttons, _title) => {
        console.log(`[MsgBox dir mock] ${prompt}`);
        return 77;
    },
};
