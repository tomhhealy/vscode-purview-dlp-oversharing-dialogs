// The module 'vscode' contains the VS Code extensibility API
// Import the module and reference it with the alias vscode in your code below
import * as vscode from "vscode";

// Interface for LocalizationData item
interface LocalizationDataItem {
  Language: string;
  Title: string;
  Body: string;
  Options: string[];
}

// Interface for Purview DLP template
interface PurviewDlpTemplate {
  LocalizationData: LocalizationDataItem[];
  HasFreeTextOption: boolean;
  DefaultLanguage: string;
}

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

// Generate WinForms-style HTML for the dialog preview
function generateDialogHtml(
  title: string,
  body: string,
  options: string[],
  hasFreeText: boolean,
  languageDisplay: string,
): string {
  // Escape HTML to prevent XSS
  const escapeHtml = (text: string): string => {
    return text
      .replace(/&/g, "&amp;")
      .replace(/</g, "&lt;")
      .replace(/>/g, "&gt;")
      .replace(/"/g, "&quot;")
      .replace(/'/g, "&#039;");
  };

  // Convert newlines to <br> for display
  const formatText = (text: string): string => {
    return escapeHtml(text).replace(/\n/g, "<br>");
  };

  // Generate radio button options
  const optionsHtml = options
    .map(
      (option, index) => `
      <div class="option-item">
        <input type="radio" id="option-${index}" name="dialog-option" ${index === 0 ? "checked" : ""}>
        <label for="option-${index}">${escapeHtml(option)}</label>
      </div>
    `,
    )
    .join("");

  // Generate free text area if enabled
  const freeTextHtml = hasFreeText
    ? `
    <div class="free-text-section">
      <label for="free-text">Provide a business justification:</label>
      <textarea id="free-text" placeholder="Enter your justification here..."></textarea>
    </div>
  `
    : "";

  return `<!DOCTYPE html>
<html lang="en">
<head>
  <meta charset="UTF-8">
  <meta name="viewport" content="width=device-width, initial-scale=1.0">
  <meta http-equiv="Content-Security-Policy" content="default-src 'none'; style-src 'unsafe-inline';">
  <title>Purview DLP Dialog Preview</title>
  <style>
    * {
      box-sizing: border-box;
      margin: 0;
      padding: 0;
    }

    body {
      font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif;
      font-size: 14px;
      background-color: #e0e0e0;
      padding: 20px;
      display: flex;
      flex-direction: column;
      align-items: center;
      min-height: 100vh;
    }

    .language-badge {
      background-color: #0078d4;
      color: white;
      padding: 4px 12px;
      border-radius: 4px;
      font-size: 12px;
      margin-bottom: 16px;
    }

    .dialog-frame {
      background-color: #ffffff;
      border: 1px solid #d1d1d1;
      border-radius: 8px;
      box-shadow: 0 8px 32px rgba(0, 0, 0, 0.15), 0 2px 8px rgba(0, 0, 0, 0.1);
      max-width: 500px;
      width: 100%;
      overflow: hidden;
    }

    .title-bar {
      background: linear-gradient(180deg, #f0f0f0 0%, #e8e8e8 100%);
      border-bottom: 1px solid #d1d1d1;
      padding: 8px 12px;
      display: flex;
      align-items: center;
      justify-content: space-between;
    }

    .title-bar-text {
      font-size: 12px;
      color: #333;
      font-weight: 400;
    }

    .title-bar-buttons {
      display: flex;
      gap: 4px;
    }

    .title-bar-button {
      width: 28px;
      height: 20px;
      border: none;
      background: #f0f0f0;
      border-radius: 3px;
      font-size: 14px;
      cursor: default;
      display: flex;
      align-items: center;
      justify-content: center;
      color: #666;
    }

    .title-bar-button.close {
      background: #c42b1c;
      color: white;
    }

    .dialog-content {
      padding: 20px 24px;
    }

    .warning-header {
      display: flex;
      align-items: flex-start;
      gap: 16px;
      margin-bottom: 16px;
    }

    .warning-icon {
      width: 32px;
      height: 32px;
      flex-shrink: 0;
    }

    .warning-icon svg {
      width: 100%;
      height: 100%;
    }

    .dialog-title {
      font-size: 16px;
      font-weight: 600;
      color: #1a1a1a;
      line-height: 1.4;
    }

    .dialog-body {
      color: #444;
      line-height: 1.5;
      margin-bottom: 20px;
      padding-left: 48px;
    }

    .options-section {
      margin-bottom: 20px;
      padding-left: 48px;
    }

    .options-label {
      font-weight: 600;
      color: #333;
      margin-bottom: 12px;
    }

    .option-item {
      display: flex;
      align-items: center;
      gap: 8px;
      margin-bottom: 8px;
    }

    .option-item input[type="radio"] {
      width: 16px;
      height: 16px;
      accent-color: #0078d4;
    }

    .option-item label {
      color: #333;
      cursor: default;
    }

    .free-text-section {
      padding-left: 48px;
      margin-bottom: 20px;
    }

    .free-text-section label {
      display: block;
      font-weight: 600;
      color: #333;
      margin-bottom: 8px;
    }

    .free-text-section textarea {
      width: 100%;
      min-height: 80px;
      padding: 8px 12px;
      border: 1px solid #ccc;
      border-radius: 4px;
      font-family: inherit;
      font-size: 13px;
      resize: vertical;
    }

    .button-bar {
      display: flex;
      justify-content: flex-end;
      gap: 8px;
      padding: 16px 24px;
      background-color: #f5f5f5;
      border-top: 1px solid #e0e0e0;
    }

    .button {
      padding: 6px 20px;
      font-family: inherit;
      font-size: 13px;
      border-radius: 4px;
      cursor: default;
      min-width: 80px;
    }

    .button-primary {
      background: linear-gradient(180deg, #0078d4 0%, #006cbe 100%);
      border: 1px solid #005a9e;
      color: white;
      font-weight: 500;
    }

    .button-secondary {
      background: linear-gradient(180deg, #ffffff 0%, #f5f5f5 100%);
      border: 1px solid #c0c0c0;
      color: #333;
    }

    .empty-state {
      color: #888;
      font-style: italic;
    }
  </style>
</head>
<body>
  <div class="language-badge">Preview: ${escapeHtml(languageDisplay)}</div>

  <div class="dialog-frame">
    <div class="title-bar">
      <span class="title-bar-text">Microsoft Outlook</span>
      <div class="title-bar-buttons">
        <div class="title-bar-button">─</div>
        <div class="title-bar-button">□</div>
        <div class="title-bar-button close">×</div>
      </div>
    </div>

    <div class="dialog-content">
      <div class="warning-header">
        <div class="warning-icon">
          <svg viewBox="0 0 32 32" fill="none" xmlns="http://www.w3.org/2000/svg">
            <path d="M16 3L2 28H30L16 3Z" fill="#FFC107" stroke="#F57C00" stroke-width="1.5"/>
            <text x="16" y="24" text-anchor="middle" font-size="16" font-weight="bold" fill="#333">!</text>
          </svg>
        </div>
        <div class="dialog-title">${title ? formatText(title) : '<span class="empty-state">(No title)</span>'}</div>
      </div>

      <div class="dialog-body">
        ${body ? formatText(body) : '<span class="empty-state">(No body text)</span>'}
      </div>

      ${
        options.length > 0
          ? `
      <div class="options-section">
        <div class="options-label">Select an option:</div>
        ${optionsHtml}
      </div>
      `
          : ""
      }

      ${freeTextHtml}
    </div>

    <div class="button-bar">
      <button class="button button-primary">Override</button>
      <button class="button button-secondary">Cancel</button>
    </div>
  </div>
</body>
</html>`;
}

// Module-level state for preview panel
let currentPanel: vscode.WebviewPanel | undefined;
let currentLanguage: string | undefined;
let documentChangeListener: vscode.Disposable | undefined;
let debounceTimer: ReturnType<typeof setTimeout> | undefined;

// Function to update webview content - defined at module level for access by multiple commands
function updateWebview(): void {
  if (!currentPanel) {
    return;
  }

  const activeEditor = vscode.window.activeTextEditor;
  if (!activeEditor) {
    return;
  }

  let currentTemplate: PurviewDlpTemplate;
  try {
    currentTemplate = JSON.parse(activeEditor.document.getText());
  } catch {
    // Don't update on invalid JSON - keep showing last valid state
    return;
  }

  if (
    !currentTemplate.LocalizationData ||
    !Array.isArray(currentTemplate.LocalizationData)
  ) {
    return;
  }

  // Find localization data for selected language
  let locData = currentTemplate.LocalizationData.find(
    (l) => l.Language === currentLanguage,
  );
  if (!locData) {
    // Fall back to default language or first available
    locData =
      currentTemplate.LocalizationData.find(
        (l) => l.Language === currentTemplate.DefaultLanguage,
      ) || currentTemplate.LocalizationData[0];
    if (locData) {
      currentLanguage = locData.Language;
    }
  }

  if (!locData) {
    return;
  }

  // Get configuration values for token replacement
  const config = vscode.workspace.getConfiguration("purviewDlp.testDialog");
  const matchedRecipientsList = config.get<string>(
    "matchedRecipientsList",
    "example@example.com",
  );
  const matchedLabelName = config.get<string>(
    "matchedLabelName",
    "Confidential",
  );
  const matchedConditions = config.get<string>(
    "matchedConditions",
    "Credit Card Number detected",
  );

  // Replace tokens in title and body
  const replaceTokens = (text: string): string => {
    return text
      .replace(/%%MatchedRecipientsList%%/g, matchedRecipientsList)
      .replace(/%%MatchedLabelName%%/g, matchedLabelName)
      .replace(/%%MatchedConditions%%/g, matchedConditions);
  };

  const title = replaceTokens(locData.Title || "");
  const body = replaceTokens(locData.Body || "");
  const options = locData.Options || [];
  const hasFreeText = currentTemplate.HasFreeTextOption;

  // Get language display name
  const langCode = currentLanguage || locData.Language;
  const langInfo = MICROSOFT_LANGUAGES.find((ml) => ml.code === langCode);
  const languageDisplay = langInfo?.label || langCode;

  currentPanel.webview.html = generateDialogHtml(
    title,
    body,
    options,
    hasFreeText,
    languageDisplay,
  );
}

// This method is called when your extension is activated
// Your extension is activated the very first time the command is executed
export function activate(context: vscode.ExtensionContext) {
  // Use the console to output diagnostic information (console.log) and errors (console.error)
  // This line of code will only be executed once when your extension is activated
  console.log(
    'Congratulations, your extension "purview-dlp-oversharing-dialogs" is now active!',
  );

  const createTemplateDisposable = vscode.commands.registerCommand(
    "purview-dlp-oversharing-dialogs.createTemplate",
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

  // Test Dialog command - opens a webview preview with live auto-update
  const testDialogDisposable = vscode.commands.registerCommand(
    "purview-dlp-oversharing-dialogs.testDialog",
    async () => {
      const editor = vscode.window.activeTextEditor;
      if (!editor) {
        vscode.window.showErrorMessage(
          "No active editor found. Please open a Purview DLP template file.",
        );
        return;
      }

      // Parse JSON content
      let template: PurviewDlpTemplate;
      try {
        template = JSON.parse(editor.document.getText());
      } catch {
        vscode.window.showErrorMessage(
          "Invalid JSON in the current file. Please fix syntax errors.",
        );
        return;
      }

      // Validate template structure
      if (
        !template.LocalizationData ||
        !Array.isArray(template.LocalizationData) ||
        template.LocalizationData.length === 0
      ) {
        vscode.window.showErrorMessage(
          "Invalid template: LocalizationData array is required.",
        );
        return;
      }

      // Select language if multiple languages available
      let selectedLanguage: string;
      if (template.LocalizationData.length === 1) {
        selectedLanguage = template.LocalizationData[0].Language;
      } else if (
        currentLanguage &&
        template.LocalizationData.some((l) => l.Language === currentLanguage)
      ) {
        // Use previously selected language if still available
        selectedLanguage = currentLanguage;
      } else {
        // Show QuickPick for language selection
        const languageItems = template.LocalizationData.map((l) => {
          const langInfo = MICROSOFT_LANGUAGES.find(
            (ml) => ml.code === l.Language,
          );
          return {
            label: langInfo?.label || l.Language,
            description: l.Language,
            isDefault: l.Language === template.DefaultLanguage,
          };
        });

        // Sort to put default language first
        languageItems.sort((a, b) => {
          if (a.isDefault) {
            return -1;
          }
          if (b.isDefault) {
            return 1;
          }
          return 0;
        });

        const selected = await vscode.window.showQuickPick(
          languageItems.map((item) => ({
            label: item.label + (item.isDefault ? " (Default)" : ""),
            description: item.description,
          })),
          {
            placeHolder: "Select a language to preview",
            title: "Dialog Preview Language",
          },
        );

        if (!selected) {
          return; // User cancelled
        }

        selectedLanguage = selected.description!;
      }

      currentLanguage = selectedLanguage;

      // Create or reveal the webview panel
      if (currentPanel) {
        currentPanel.reveal(vscode.ViewColumn.Beside);
      } else {
        currentPanel = vscode.window.createWebviewPanel(
          "purviewDlpDialogPreview",
          "Purview DLP Dialog Preview",
          vscode.ViewColumn.Beside,
          {
            enableScripts: false,
            retainContextWhenHidden: true,
          },
        );

        // Clean up when panel is closed
        currentPanel.onDidDispose(() => {
          currentPanel = undefined;
          if (documentChangeListener) {
            documentChangeListener.dispose();
            documentChangeListener = undefined;
          }
          if (debounceTimer) {
            clearTimeout(debounceTimer);
            debounceTimer = undefined;
          }
        });
      }

      // Initial update
      updateWebview();

      // Set up document change listener for live auto-update
      if (documentChangeListener) {
        documentChangeListener.dispose();
      }

      documentChangeListener = vscode.workspace.onDidChangeTextDocument(
        (event) => {
          // Only update if the changed document is the active editor
          const activeEditor = vscode.window.activeTextEditor;
          if (!activeEditor || event.document !== activeEditor.document) {
            return;
          }

          // Debounce updates to avoid excessive re-renders
          if (debounceTimer) {
            clearTimeout(debounceTimer);
          }
          debounceTimer = setTimeout(() => {
            updateWebview();
          }, 300);
        },
      );

      context.subscriptions.push(documentChangeListener);
    },
  );

  // Switch Preview Language command - allows changing the preview language while panel is open
  const switchPreviewLanguageDisposable = vscode.commands.registerCommand(
    "purview-dlp-oversharing-dialogs.switchPreviewLanguage",
    async () => {
      // Check if preview panel is open
      if (!currentPanel) {
        vscode.window.showWarningMessage(
          "No preview panel is open. Use 'Purview DLP: Preview Dialog' first to open a preview.",
        );
        return;
      }

      const editor = vscode.window.activeTextEditor;
      if (!editor) {
        vscode.window.showErrorMessage(
          "No active editor found. Please open a Purview DLP template file.",
        );
        return;
      }

      // Parse JSON content
      let template: PurviewDlpTemplate;
      try {
        template = JSON.parse(editor.document.getText());
      } catch {
        vscode.window.showErrorMessage(
          "Invalid JSON in the current file. Please fix syntax errors.",
        );
        return;
      }

      // Validate template structure
      if (
        !template.LocalizationData ||
        !Array.isArray(template.LocalizationData) ||
        template.LocalizationData.length === 0
      ) {
        vscode.window.showErrorMessage(
          "Invalid template: LocalizationData array is required.",
        );
        return;
      }

      // Check if there's only one language
      if (template.LocalizationData.length === 1) {
        vscode.window.showInformationMessage(
          "Only one language is available in this template.",
        );
        return;
      }

      // Build language selection items
      const languageItems = template.LocalizationData.map((l) => {
        const langInfo = MICROSOFT_LANGUAGES.find(
          (ml) => ml.code === l.Language,
        );
        const isCurrent = l.Language === currentLanguage;
        const isDefault = l.Language === template.DefaultLanguage;
        let suffix = "";
        if (isCurrent && isDefault) {
          suffix = " (Current, Default)";
        } else if (isCurrent) {
          suffix = " (Current)";
        } else if (isDefault) {
          suffix = " (Default)";
        }
        return {
          label: (langInfo?.label || l.Language) + suffix,
          description: l.Language,
          languageCode: l.Language,
        };
      });

      // Sort to put current language first, then default
      languageItems.sort((a, b) => {
        if (a.languageCode === currentLanguage) {
          return -1;
        }
        if (b.languageCode === currentLanguage) {
          return 1;
        }
        if (a.languageCode === template.DefaultLanguage) {
          return -1;
        }
        if (b.languageCode === template.DefaultLanguage) {
          return 1;
        }
        return 0;
      });

      const selected = await vscode.window.showQuickPick(languageItems, {
        placeHolder: "Select a language to preview",
        title: "Switch Preview Language",
      });

      if (!selected) {
        return; // User cancelled
      }

      // Update the current language and refresh the preview
      currentLanguage = selected.languageCode;
      updateWebview();
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
          // Check if token completion is enabled
          const enabled = vscode.workspace
            .getConfiguration("purviewDlp")
            .get<boolean>("enableTokenCompletion", true);
          if (!enabled) {
            return undefined;
          }

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
              documentation:
                "Inserts the name of the sensitivity label that was matched",
            },
            {
              label: "%%MatchedConditions%%",
              detail: "Matched Conditions",
              documentation:
                "Inserts the conditions that were matched by the DLP policy",
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

  context.subscriptions.push(createTemplateDisposable);
  context.subscriptions.push(testDialogDisposable);
  context.subscriptions.push(switchPreviewLanguageDisposable);
  context.subscriptions.push(tokenCompletionProvider);
}

// This method is called when your extension is deactivated
export function deactivate() {}
