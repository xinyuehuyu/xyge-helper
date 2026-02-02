"use strict";
var __createBinding = (this && this.__createBinding) || (Object.create ? (function(o, m, k, k2) {
    if (k2 === undefined) k2 = k;
    var desc = Object.getOwnPropertyDescriptor(m, k);
    if (!desc || ("get" in desc ? !m.__esModule : desc.writable || desc.configurable)) {
      desc = { enumerable: true, get: function() { return m[k]; } };
    }
    Object.defineProperty(o, k2, desc);
}) : (function(o, m, k, k2) {
    if (k2 === undefined) k2 = k;
    o[k2] = m[k];
}));
var __setModuleDefault = (this && this.__setModuleDefault) || (Object.create ? (function(o, v) {
    Object.defineProperty(o, "default", { enumerable: true, value: v });
}) : function(o, v) {
    o["default"] = v;
});
var __importStar = (this && this.__importStar) || (function () {
    var ownKeys = function(o) {
        ownKeys = Object.getOwnPropertyNames || function (o) {
            var ar = [];
            for (var k in o) if (Object.prototype.hasOwnProperty.call(o, k)) ar[ar.length] = k;
            return ar;
        };
        return ownKeys(o);
    };
    return function (mod) {
        if (mod && mod.__esModule) return mod;
        var result = {};
        if (mod != null) for (var k = ownKeys(mod), i = 0; i < k.length; i++) if (k[i] !== "default") __createBinding(result, mod, k[i]);
        __setModuleDefault(result, mod);
        return result;
    };
})();
var __importDefault = (this && this.__importDefault) || function (mod) {
    return (mod && mod.__esModule) ? mod : { "default": mod };
};
Object.defineProperty(exports, "__esModule", { value: true });
exports.languageActivate = languageActivate;
const vscode = __importStar(require("vscode"));
const ts = __importStar(require("typescript"));
const path_1 = __importDefault(require("path"));
const exceljs_1 = __importDefault(require("exceljs"));
const fs = __importStar(require("fs"));
const iconv = __importStar(require("iconv-lite"));
const stream_1 = require("stream");
const pinyin_1 = __importDefault(require("pinyin"));
async function languageActivate(context) {
    // ========== Hover Provider ==========
    const hoverProvider = vscode.languages.registerHoverProvider(['typescript', 'typescriptreact'], {
        async provideHover(document, position, token) {
            function escapeMarkdown(text) {
                // 这里的正则匹配了所有 Markdown 中有特殊含义的字符
                return text.replace(/[\\`*_{}\[\]()#+\-.!|<>]/g, '\\$&');
            }
            const langId = document.languageId;
            if (langId !== 'typescript' && langId !== 'typescriptreact') {
                return null;
            }
            const stringInfo = getStringLiteralWrappedByL(document, position);
            if (!stringInfo) {
                return null;
            }
            // 创建 Hover 内容
            const contents = new vscode.MarkdownString();
            // contents.appendCodeblock(stringInfo.text, 'text');
            // 获取excelData中key为stringInfo.text的value
            if (stringInfo.isTemplateExpression) {
                // contents.appendCodeblock("L()方法不应框涵带插值的模板字符串", 'text');
                contents.appendMarkdown(`<i>${escapeMarkdown("L()方法不应框涵带插值的模板字符串")}</i>`);
            }
            else {
                const translation = excelData.get(stringInfo.text);
                if (translation) {
                    contents.appendCodeblock(translation, 'text');
                }
                else {
                    // contents.appendCodeblock("多语言中未找到此key，是否需要将其添加到多语言中", 'text');
                    contents.appendMarkdown(`<i>${escapeMarkdown("多语言中未找到此key，是否需要将其添加到多语言中")}</i>`);
                    contents.appendMarkdown(`\n\n`);
                    const params = {
                        uri: document.uri.toString(),
                        startLine: stringInfo.range.start.line,
                        startChar: stringInfo.range.start.character,
                        endLine: stringInfo.range.end.line,
                        endChar: stringInfo.range.end.character,
                        sStr: stringInfo.text,
                    };
                    const encodedParams = encodeURIComponent(JSON.stringify(params));
                    if (!translation) {
                        contents.appendMarkdown(`[添加多语言](command:extension.addKV?${encodedParams})`);
                        contents.appendMarkdown(` | `);
                    }
                }
                contents.appendMarkdown(`[打开多语言表](command:extension.openLanguageFile)`);
            }
            contents.supportHtml = true;
            contents.isTrusted = true; // 允许执行命令链接（VS Code 1.64+）
            return new vscode.Hover(contents, stringInfo.range);
        }
    });
    // ========== 注册替换命令 ==========
    const replaceCommand = vscode.commands.registerCommand('extension.openLanguageFile', async () => {
        const document = await vscode.workspace.openTextDocument(targetUri, { encoding: 'gbk' });
        await vscode.window.showTextDocument(document);
    });
    const replaceCommand2 = vscode.commands.registerCommand('extension.addKV', async (arg) => {
        // VS Code 会自动尝试 JSON.parse 参数，所以可能收到字符串或已解析的对象
        // let params: { uri: string; startLine: number; startChar: number; endLine: number; endChar: number };
        try {
            // 获取文档
            const uri = vscode.Uri.parse(arg.uri);
            const document = await vscode.workspace.openTextDocument(uri);
            const editor = await vscode.window.showTextDocument(document);
            arg.tStr = getPinyinDefaultKey(arg.sStr, document.fileName);
            // 1. 弹出输入框，允许用户修改默认值
            // arg.tStr 是我们传入的默认拼音 Key
            const userConfirmedKey = await vscode.window.showInputBox({
                title: `添加多语言 (${arg.sStr})`,
                value: arg.tStr, // 默认值
                prompt: `推荐Key: ${arg.tStr}    `,
                placeHolder: "请输入key...",
                validateInput: (input) => {
                    if (!input || input.length === 0) {
                        return "Key不能为空";
                    }
                    if (input.startsWith(" ") || input.endsWith(" ")) {
                        return "Key开头或末尾不能有空格";
                    }
                    if (/^\d/.test(input)) {
                        return "Key不能以数字开头";
                    }
                    // key不能在excelData中
                    if (excelData.has(input)) {
                        return "Key已存在！ key:" + input + " value:" + excelData.get(input);
                    }
                    return null;
                }
            });
            // 2. 如果用户取消了输入（按了ESC），直接返回，不做任何操作
            if (userConfirmedKey === undefined) {
                return;
            }
            // 重建 Range
            const range = new vscode.Range(arg.startLine, arg.startChar, arg.endLine, arg.endChar);
            WriteExcel(targetFile, arg.tStr, arg.sStr);
            // 执行替换
            await editor.edit(editBuilder => {
                editBuilder.replace(range, "\"" + arg.tStr + "\"");
            });
            // vscode.window.showInformationMessage('✅ 已替换为 "11111111"');
        }
        catch (error) {
            const choice = await vscode.window.showErrorMessage(`替换失败: ${error.message || error}`);
        }
    });
    // 配置表监听
    const workspaceFolders = vscode.workspace.workspaceFolders;
    if (!workspaceFolders?.length) {
        return;
    }
    const projectName = getProjectName();
    if (projectName === "") {
        return;
    }
    const targetFile = path_1.default.join(workspaceFolders[0].uri.fsPath, `res/${projectName}/config/format/excel/language/dic_language_ts_cn.csv`);
    const targetUri = vscode.Uri.file(targetFile);
    const watcher = vscode.workspace.createFileSystemWatcher(targetFile, // glob 模式或绝对路径
    false, // ignoreCreateEvents: 不监听创建
    false, // ignoreChangeEvents: 不监听修改 ← 设为 false 才监听修改！
    true // ignoreDeleteEvents: 不监听删除
    );
    watcher.onDidChange(uri => {
        if (uri.fsPath === targetUri.fsPath) {
            vscode.window.showInformationMessage('检测到多语言文件改变，重新读取');
            readExcel(targetFile);
        }
    });
    watcher.onDidCreate(uri => {
        if (uri.fsPath === targetUri.fsPath) {
            vscode.window.showInformationMessage('检测到多语言文件生成，重新读取');
            readExcel(targetFile);
        }
    });
    context.subscriptions.push(hoverProvider, replaceCommand, replaceCommand2);
    readExcel(targetFile);
}
/**
 * 获取被 L() 包裹的字符串字面量信息
 */
function getStringLiteralWrappedByL(document, position) {
    try {
        const sourceFile = ts.createSourceFile(document.fileName || 'temp.ts', document.getText(), ts.ScriptTarget.Latest, true, ts.ScriptKind.TS);
        let result = null;
        function findStringLiteral(node) {
            if (result) {
                return;
            }
            const isTemplateExpression = ts.isTemplateExpression(node);
            if (ts.isStringLiteral(node) ||
                ts.isNoSubstitutionTemplateLiteral(node) ||
                isTemplateExpression) {
                const start = document.positionAt(node.pos);
                const end = document.positionAt(node.end);
                const nodeRange = new vscode.Range(start, end);
                if (nodeRange.contains(position)) {
                    const parent = node.parent;
                    if (parent && ts.isCallExpression(parent)) {
                        if (ts.isIdentifier(parent.expression) && parent.expression.text === 'L') {
                            const text = node.getText(sourceFile).slice(1, -1);
                            result = {
                                text,
                                range: nodeRange,
                                isTemplateExpression
                            };
                        }
                    }
                }
            }
            ts.forEachChild(node, findStringLiteral);
        }
        ts.forEachChild(sourceFile, findStringLiteral);
        return result;
    }
    catch (error) {
        vscode.window.showErrorMessage(`ts语言解析失败\n${error}`);
        return null;
    }
}
let excelData = new Map();
async function readExcel(targetFile) {
    excelData.clear();
    const oldWorkbook = new exceljs_1.default.Workbook();
    try {
        if (fs.existsSync(targetFile)) {
            const fileBuffer = fs.readFileSync(targetFile);
            const utf8Content = iconv.decode(fileBuffer, 'gbk');
            const stream = stream_1.Readable.from([utf8Content]);
            await oldWorkbook.csv.read(stream);
            const oldWorksheet = oldWorkbook.getWorksheet(1);
            if (oldWorksheet) {
                // 从第6行开始读取数据（前5行是固定行）
                for (let i = 6; i <= oldWorksheet.rowCount; i++) {
                    const row = oldWorksheet.getRow(i);
                    if (row.getCell('B').value && row.getCell('C').value) {
                        const key = String(row.getCell('B').value);
                        const value = String(row.getCell('C').value);
                        excelData.set(key, value);
                    }
                }
            }
        }
        else {
            vscode.window.showErrorMessage(`多语言文件不存在：${targetFile}`);
        }
    }
    catch (error) {
        vscode.window.showErrorMessage(`多语言文件读取失败，请检查多语言表是否存在或是否被占用\n多语言表路径：${targetFile}\n${error}`);
        return;
    }
}
async function WriteExcel(targetFile, key, value) {
    // 检查是否存在旧的Excel文件
    let oldData = new Map();
    const oldWorkbook = new exceljs_1.default.Workbook();
    try {
        if (fs.existsSync(targetFile)) {
            const fileBuffer = fs.readFileSync(targetFile);
            const utf8Content = iconv.decode(fileBuffer, 'gbk');
            const stream = stream_1.Readable.from([utf8Content]);
            await oldWorkbook.csv.read(stream);
            const oldWorksheet = oldWorkbook.getWorksheet(1);
            if (oldWorksheet) {
                // 从第6行开始读取数据（前5行是固定行）
                for (let i = 6; i <= oldWorksheet.rowCount; i++) {
                    const row = oldWorksheet.getRow(i);
                    if (row.getCell('B').value && row.getCell('C').value) {
                        const key = String(row.getCell('B').value);
                        const value = String(row.getCell('C').value);
                        oldData.set(key, value);
                    }
                }
            }
        }
    }
    catch (error) {
        vscode.window.showErrorMessage(`多语言文件读取失败，请检查多语言表是否存在或是否被占用\n多语言表路径：${targetFile}\n${error}`);
        return;
    }
    const workbook = new exceljs_1.default.Workbook();
    const worksheet = workbook.addWorksheet(`${path_1.default.basename(targetFile, '.csv')}`);
    // 设置表头
    worksheet.columns = [
        { header: '#name', key: 'hand', width: 20 },
        { header: 'key', key: 'key', width: 100 },
        { header: 'value', key: 'value', width: 100 },
    ];
    // 添加固定行
    worksheet.addRow({
        hand: '#visbility',
        key: 'cs',
        value: 'cs',
    });
    worksheet.addRow({
        hand: '#comments',
        key: '',
        value: '',
    });
    worksheet.addRow({
        hand: '#type',
        key: 'string',
        value: 'string',
    });
    worksheet.addRow({
        hand: '#index',
        key: '',
        value: '',
    });
    worksheet.addRow({
        hand: '#foreign',
        key: '',
        value: '',
    });
    // 添加已有kv
    for (const [key, value] of oldData.entries()) {
        worksheet.addRow({
            hand: '',
            key: key,
            value: value
        });
    }
    // 添加新增kv
    worksheet.addRow({
        hand: '',
        key: key,
        value: value
    });
    // 保存文件
    if (fs.existsSync(targetFile)) {
        fs.unlinkSync(targetFile);
    }
    const buffer = await workbook.csv.writeBuffer();
    const utf8Content = buffer.toString();
    const gbkBuffer = iconv.encode(utf8Content, 'gbk');
    fs.writeFileSync(targetFile, gbkBuffer);
}
function getProjectName() {
    const workspaceFolders = vscode.workspace.workspaceFolders;
    if (!workspaceFolders?.length) {
        return null;
    }
    const nameFile = path_1.default.join(workspaceFolders[0].uri.fsPath, `client/Assets/Res/Resources/productName.txt`);
    if (fs.existsSync(nameFile)) {
        return fs.readFileSync(nameFile, 'utf-8').trim();
    }
    else {
        vscode.window.showErrorMessage(`项目名称文件未找到，请修复并重启编辑器：${nameFile}`);
        return "";
    }
}
function getPinyinDefaultKey(text, fileName) {
    let result = '';
    for (let i = 0; i < text.length; i++) {
        // 限制最大长度为 8
        if (result.length >= 8) {
            break;
        }
        const char = text[i];
        // 1. 判断是否为英文字母 或 数字
        if (/[a-zA-Z0-9]/.test(char)) {
            result += char;
        }
        // 2. 判断是否为中文字符
        else if (/[\u4e00-\u9fa5]/.test(char)) {
            const pyArray = (0, pinyin_1.default)(char, {
                style: 3,
                heteronym: false
            });
            if (pyArray.length > 0 && pyArray[0].length > 0) {
                result += pyArray[0][0];
            }
        }
        // 3. 其他情况略过
    }
    // 新增：处理无有效字符的情况
    if (!result) {
        let hash = 0;
        // 计算字符串的哈希值 (djb2 算法变种)
        for (let i = 0; i < text.length; i++) {
            const char = text.charCodeAt(i);
            hash = ((hash << 5) - hash) + char;
            hash |= 0; // Convert to 32bit integer
        }
        // 取绝对值并转为 16 进制字符串，取前 8 位
        const hexHash = Math.abs(hash).toString(16);
        result = hexHash.substring(0, 8);
    }
    // 最终兜底（理论上 hash 不可能为空，除非 text 为空）
    result = path_1.default.basename(fileName, '.ts') + '_' + (result || 'key');
    return result;
}
//# sourceMappingURL=language.js.map