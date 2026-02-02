import * as vscode from 'vscode';
import { languageActivate } from './language';

export async function activate(context: vscode.ExtensionContext) {
    await languageActivate(context);
}
export function deactivate() {}