// The module 'vscode' contains the VS Code extensibility API
// Import the module and reference it with the alias vscode in your code below
import * as vscode from "vscode";

// This method is called when your extension is activated
// Your extension is activated the very first time the command is executed
export function activate(context: vscode.ExtensionContext) {
  // Use the console to output diagnostic information (console.log) and errors (console.error)
  // This line of code will only be executed once when your extension is activated
  console.log(
    'Congratulations, your extension "vscode-purview-dlp-oversharing-dialogs" is now active!',
  );

  // The command has been defined in the package.json file
  // Now provide the implementation of the command with registerCommand
  // The commandId parameter must match the command field in package.json
  const disposable = vscode.commands.registerCommand(
    "vscode-purview-dlp-oversharing-dialogs.helloWorld",
    () => {
      // The code you place here will be executed every time your command is executed
      // Display a message box to the user
      vscode.window.showInformationMessage(
        "Hello World from vscode-purview-dlp-oversharing-dialogs!",
      );
    },
  );

  const createTemplateDisposable = vscode.commands.registerCommand(
    "vscode-purview-dlp-oversharing-dialogs.createTemplate",
    async () => {
      // Create the JSON template structure
      const template = {
        LocalizationData: [
          {
            Language: "en-US",
            Title: "",
            Body: "",
            Options: ["Option 1", "Option 2"],
          },
        ],
        HasFreeTextOption: true,
        DefaultLanguage: "en-US",
      };

      // Create a new untitled document with JSON language
      const document = await vscode.workspace.openTextDocument({
        content: JSON.stringify(template, null, 2),
        language: "json",
      });

      // Show the document in the editor
      await vscode.window.showTextDocument(document);

      vscode.window.showInformationMessage(
        "Purview DLP oversharing dialog template created!",
      );
    },
  );

  context.subscriptions.push(disposable);
  context.subscriptions.push(createTemplateDisposable);
}

// This method is called when your extension is deactivated
export function deactivate() {}
