// The module 'vscode' contains the VS Code extensibility API
// Import the module and reference it with the alias vscode in your code below
import * as vscode from "vscode";

// Microsoft language codes with display names
const MICROSOFT_LANGUAGES = [
  { code: "en-US", label: "English (United States)" },
  { code: "es-ES", label: "Spanish (Spain)" },
  { code: "fr-FR", label: "French (France)" },
  { code: "de-DE", label: "German (Germany)" },
  { code: "ja-JP", label: "Japanese (Japan)" },
  { code: "pt-BR", label: "Portuguese (Brazil)" },
  { code: "ru-RU", label: "Russian (Russia)" },
  { code: "zh-CN", label: "Chinese (Simplified, China)" },
  { code: "zh-TW", label: "Chinese (Traditional, Taiwan)" },
  { code: "it-IT", label: "Italian (Italy)" },
  { code: "ko-KR", label: "Korean (Korea)" },
  { code: "nl-NL", label: "Dutch (Netherlands)" },
  { code: "pl-PL", label: "Polish (Poland)" },
  { code: "tr-TR", label: "Turkish (Turkey)" },
  { code: "sv-SE", label: "Swedish (Sweden)" },
  { code: "cs-CZ", label: "Czech (Czech Republic)" },
  { code: "da-DK", label: "Danish (Denmark)" },
  { code: "fi-FI", label: "Finnish (Finland)" },
  { code: "el-GR", label: "Greek (Greece)" },
  { code: "hu-HU", label: "Hungarian (Hungary)" },
  { code: "no-NO", label: "Norwegian (Norway)" },
  { code: "pt-PT", label: "Portuguese (Portugal)" },
  { code: "ro-RO", label: "Romanian (Romania)" },
  { code: "sk-SK", label: "Slovak (Slovakia)" },
  { code: "th-TH", label: "Thai (Thailand)" },
  { code: "uk-UA", label: "Ukrainian (Ukraine)" },
  { code: "ar-SA", label: "Arabic (Saudi Arabia)" },
  { code: "he-IL", label: "Hebrew (Israel)" },
  { code: "id-ID", label: "Indonesian (Indonesia)" },
  { code: "ms-MY", label: "Malay (Malaysia)" },
  { code: "vi-VN", label: "Vietnamese (Vietnam)" },
  { code: "bg-BG", label: "Bulgarian (Bulgaria)" },
  { code: "hr-HR", label: "Croatian (Croatia)" },
  { code: "et-EE", label: "Estonian (Estonia)" },
  { code: "lv-LV", label: "Latvian (Latvia)" },
  { code: "lt-LT", label: "Lithuanian (Lithuania)" },
  { code: "sl-SI", label: "Slovenian (Slovenia)" },
  { code: "sr-Latn-RS", label: "Serbian (Latin, Serbia)" },
  { code: "ca-ES", label: "Catalan (Spain)" },
  { code: "eu-ES", label: "Basque (Spain)" },
  { code: "gl-ES", label: "Galician (Spain)" },
];

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
      // Step 1: Select languages (maximum 10 as per Purview limit)
      const selectedLanguages = await vscode.window.showQuickPick(
        MICROSOFT_LANGUAGES.map((lang) => ({
          label: lang.label,
          description: lang.code,
          picked: lang.code === "en-US", // Pre-select English
        })),
        {
          canPickMany: true,
          placeHolder:
            "Select languages to include in the template (max 10, default: en-US)",
          title: "Purview DLP Template Languages",
        },
      );

      if (!selectedLanguages || selectedLanguages.length === 0) {
        vscode.window.showWarningMessage(
          "Template creation cancelled: No languages selected",
        );
        return;
      }

      // Enforce Purview's 10-language limit
      if (selectedLanguages.length > 10) {
        vscode.window.showErrorMessage(
          `Too many languages selected (${selectedLanguages.length}). Purview DLP supports a maximum of 10 languages per template.`,
        );
        return;
      }

      // Step 2: Select default language (or auto-select if only one language chosen)
      let defaultLanguageCode: string;

      if (selectedLanguages.length === 1) {
        // If only one language selected, use it as the default
        defaultLanguageCode = selectedLanguages[0].description!;
      } else {
        // Multiple languages selected, ask user to choose default
        const defaultLanguage = await vscode.window.showQuickPick(
          selectedLanguages.map((lang) => ({
            label: lang.label,
            description: lang.description,
          })),
          {
            placeHolder: "Select the default/fallback language",
            title: "Default Language",
          },
        );

        if (!defaultLanguage) {
          vscode.window.showWarningMessage(
            "Template creation cancelled: No default language selected",
          );
          return;
        }

        defaultLanguageCode = defaultLanguage.description!;
      }

      // Create the JSON template structure
      const template = {
        LocalizationData: selectedLanguages.map((lang) => ({
          Language: lang.description,
          Title: "",
          Body: "",
          Options: ["Option 1", "Option 2"],
        })),
        HasFreeTextOption: true,
        DefaultLanguage: defaultLanguageCode,
      };

      // Create a new untitled document with purview-dlp-json language
      const document = await vscode.workspace.openTextDocument({
        content: JSON.stringify(template, null, 2),
        language: "purview-dlp-json",
      });

      // Show the document in the editor
      await vscode.window.showTextDocument(document);

      vscode.window.showInformationMessage(
        `Purview DLP template created with ${selectedLanguages.length} language(s)!`,
      );
    },
  );

  // Register completion provider for token IntelliSense
  const tokenCompletionProvider =
    vscode.languages.registerCompletionItemProvider(
      "purview-dlp-json",
      {
        provideCompletionItems(
          document: vscode.TextDocument,
          position: vscode.Position,
        ) {
          // Check if we're typing after %%
          const linePrefix = document
            .lineAt(position)
            .text.slice(0, position.character);
          if (!linePrefix.endsWith("%%")) {
            return undefined;
          }

          // Create completion items for each token
          const tokens = [
            {
              label: "%%MatchedRecipientsList%%",
              detail: "Matched Recipients List",
              documentation: `Display the matched recipients for a given DLP rule for these conditions:

- Recipient is
- Recipient domain is
- Recipient is a member of
- Content is shared from Microsoft 365`,
            },
            {
              label: "%%MatchedLabelName%%",
              detail: "Matched Label Name",
              documentation: "Inserts the name of the sensitivity label that was matched",
            },
            {
              label: "%%MatchedConditions%%",
              detail: "Matched Conditions",
              documentation: "Inserts the conditions that were matched by the DLP policy",
            },
          ];

          return tokens.map((token) => {
            const item = new vscode.CompletionItem(
              token.label,
              vscode.CompletionItemKind.Constant,
            );
            item.detail = token.detail;
            item.documentation = new vscode.MarkdownString(token.documentation);
            item.insertText = token.label;
            return item;
          });
        },
      },
      "%", // Trigger on % character
    );

  context.subscriptions.push(disposable);
  context.subscriptions.push(createTemplateDisposable);
  context.subscriptions.push(tokenCompletionProvider);
}

// This method is called when your extension is deactivated
export function deactivate() {}
