"use strict";
Object.defineProperty(exports, "__esModule", { value: true });
exports.activate = activate;
exports.deactivate = deactivate;
const language_1 = require("./language");
async function activate(context) {
    await (0, language_1.languageActivate)(context);
}
function deactivate() { }
//# sourceMappingURL=extension.js.map