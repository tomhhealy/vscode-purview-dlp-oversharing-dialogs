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
  titleLength: number,
  bodyLength: number,
  optionsClassLength: number,
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

  // Format text for Title (no formatting tags, just escape)
  const formatTitle = (text: string): string => {
    return escapeHtml(text);
  };

  // Format text for Body (supports formatting tags)
  const formatBody = (text: string): string => {
    // First escape all HTML
    let formatted = escapeHtml(text);

    // Convert formatting tags (which are now escaped) to actual HTML
    // Bold: <Bold>text</Bold>
    formatted = formatted.replace(
      /&lt;Bold&gt;(.*?)&lt;\/Bold&gt;/gi,
      "<strong>$1</strong>",
    );

    // Underline: <Underline>text</Underline>
    formatted = formatted.replace(
      /&lt;Underline&gt;(.*?)&lt;\/Underline&gt;/gi,
      "<u>$1</u>",
    );

    // Italic: <Italic>text</Italic>
    formatted = formatted.replace(
      /&lt;Italic&gt;(.*?)&lt;\/Italic&gt;/gi,
      "<em>$1</em>",
    );

    // Line breaks: <LineBreak /> or <br> or <br />
    formatted = formatted.replace(/&lt;LineBreak\s*\/&gt;/gi, "<br>");
    formatted = formatted.replace(/&lt;br\s*\/?&gt;/gi, "<br>");

    return formatted;
  };

  // Generate radio button options (only show first 3)
  const optionsToDisplay = options.slice(0, 3);
  const optionsHtml = optionsToDisplay
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

  // Calculate stat classes based on limits
  const titleClass =
    titleLength > 75 ? "error" : titleLength > 60 ? "warning" : "";
  const bodyClass =
    bodyLength > 800 ? "error" : bodyLength > 640 ? "warning" : "";
  const optionsClass = options.length > 3 ? "error" : "";
  const optionsItemClass = optionsClassLength > 100 ? "error" : optionsClassLength > 80 ? "warning" : "";

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
      margin-bottom: 8px;
    }

    .stats-badge {
      background-color: #f5f5f5;
      border: 1px solid #d1d1d1;
      border-radius: 4px;
      padding: 8px 12px;
      font-size: 11px;
      margin-bottom: 16px;
      display: flex;
      gap: 16px;
      justify-content: center;
      flex-wrap: wrap;
    }

    .stat-item {
      display: flex;
      align-items: center;
      gap: 6px;
    }

    .stat-label {
      color: #666;
      font-weight: 500;
    }

    .stat-value {
      color: #333;
      font-weight: 600;
    }

    .stat-value.warning {
      color: #f59e0b;
    }

    .stat-value.error {
      color: #ef4444;
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

    .dialog-title strong,
    .dialog-title b {
      font-weight: 700;
    }

    .dialog-title em,
    .dialog-title i {
      font-style: italic;
    }

    .dialog-title u {
      text-decoration: underline;
    }

    .dialog-body {
      color: #444;
      line-height: 1.5;
      margin-bottom: 20px;
      padding-left: 48px;
    }

    .dialog-body strong,
    .dialog-body b {
      font-weight: 600;
      color: #333;
    }

    .dialog-body em,
    .dialog-body i {
      font-style: italic;
    }

    .dialog-body u {
      text-decoration: underline;
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

  <div class="stats-badge">
    <div class="stat-item">
      <span class="stat-label">Title:</span>
      <span class="stat-value ${titleClass}">${titleLength}/75</span>
    </div>
    <div class="stat-item">
      <span class="stat-label">Body:</span>
      <span class="stat-value ${bodyClass}">${bodyLength}/800</span>
    </div>
    <div class="stat-item">
      <span class="stat-label">Options:</span>
      <span class="stat-value ${optionsClass}">${options.length}/3</span>
    </div>
    <div class="stat-item">
      <span class="stat-label">Longest Option:</span>
      <span class="stat-value ${optionsItemClass}">${optionsClassLength}/100</span>
    </div>
  </div>

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
        <div class="dialog-title">${title ? formatTitle(title) : '<span class="empty-state">(No title)</span>'}</div>
      </div>

      <div class="dialog-body">
        ${body ? formatBody(body) : '<span class="empty-state">(No body text)</span>'}
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

// Diagnostics collection for validation
let diagnosticCollection: vscode.DiagnosticCollection;

// Validate document and create diagnostics for field length violations
function validateDocument(document: vscode.TextDocument): void {
  if (!diagnosticCollection) {
    return;
  }

  const diagnostics: vscode.Diagnostic[] = [];
  const text = document.getText();

  // Only validate purview-dlp-json files
  if (document.languageId !== "purview-dlp-json") {
    diagnosticCollection.set(document.uri, []);
    return;
  }

  try {
    const template: PurviewDlpTemplate = JSON.parse(text);

    if (
      !template.LocalizationData ||
      !Array.isArray(template.LocalizationData)
    ) {
      diagnosticCollection.set(document.uri, []);
      return;
    }

    // Parse the JSON to find line positions
    const lines = text.split("\n");

    // Track which LocalizationData item we're in
    let inLocalizationData = false;
    let locDataIndex = -1;
    let currentProperty = "";

    for (let lineNum = 0; lineNum < lines.length; lineNum++) {
      const line = lines[lineNum];

      // Check if we're entering LocalizationData array
      if (line.includes('"LocalizationData"')) {
        inLocalizationData = true;
        locDataIndex = -1;
        continue;
      }

      // Check if we found an object start in LocalizationData
      if (inLocalizationData && line.trim() === "{") {
        locDataIndex++;
        continue;
      }

      // Check if we're leaving LocalizationData
      if (inLocalizationData && line.includes("]") && !line.includes("[")) {
        inLocalizationData = false;
        continue;
      }

      if (inLocalizationData && locDataIndex >= 0) {
        const locData = template.LocalizationData[locDataIndex];
        if (!locData) {
          continue;
        }

        // Check for Title property
        if (line.includes('"Title"')) {
          currentProperty = "Title";
          const titleLength = locData.Title.length;
          if (titleLength > 75) {
            const range = new vscode.Range(lineNum, 0, lineNum, line.length);
            const diagnostic = new vscode.Diagnostic(
              range,
              `Title exceeds 75 character limit (current: ${titleLength} characters)`,
              vscode.DiagnosticSeverity.Error,
            );
            diagnostic.code = "title-too-long";
            diagnostics.push(diagnostic);
          }
        }

        // Check for Body property
        if (line.includes('"Body"')) {
          currentProperty = "Body";
          const bodyLength = locData.Body.length;
          if (bodyLength > 800) {
            const range = new vscode.Range(lineNum, 0, lineNum, line.length);
            const diagnostic = new vscode.Diagnostic(
              range,
              `Body exceeds 800 character limit (current: ${bodyLength} characters)`,
              vscode.DiagnosticSeverity.Error,
            );
            diagnostic.code = "body-too-long";
            diagnostics.push(diagnostic);
          }
        }

        // Check for Options property
        if (line.includes('"Options"')) {
          currentProperty = "Options";
          const optionsLength = locData.Options?.length || 0;
          if (optionsLength > 3) {
            const range = new vscode.Range(lineNum, 0, lineNum, line.length);
            const diagnostic = new vscode.Diagnostic(
              range,
              `Options array exceeds maximum of 3 items (current: ${optionsLength} options)`,
              vscode.DiagnosticSeverity.Error,
            );
            diagnostic.code = "too-many-options";
            diagnostics.push(diagnostic);
          }

          // Check individual option lengths
          if (locData.Options && Array.isArray(locData.Options)) {
            locData.Options.forEach((option, optIndex) => {
              if (option.length > 100) {
                // Find the line with this specific option
                for (let i = lineNum + 1; i < lines.length; i++) {
                  const optLine = lines[i];
                  if (optLine.includes(option.substring(0, 20))) {
                    const range = new vscode.Range(i, 0, i, optLine.length);
                    const diagnostic = new vscode.Diagnostic(
                      range,
                      `Option ${optIndex + 1} exceeds 100 character limit (current: ${option.length} characters)`,
                      vscode.DiagnosticSeverity.Error,
                    );
                    diagnostic.code = "option-too-long";
                    diagnostics.push(diagnostic);
                    break;
                  }
                }
              }
            });
          }
        }
      }
    }
  } catch (error) {
    // Invalid JSON - don't add diagnostics, JSON validation will handle it
  }

  diagnosticCollection.set(document.uri, diagnostics);
}

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
  const matchedAttachmentName = config.get<string>(
    "matchedAttachmentName",
    "Financial_Report.xlsx",
  );

  // Replace tokens in title and body
  const replaceTokens = (text: string): string => {
    return text
      .replace(/%%MatchedRecipientsList%%/g, matchedRecipientsList)
      .replace(/%%MatchedLabelName%%/g, matchedLabelName)
      .replace(/%%MatchedAttachmentName%%/g, matchedAttachmentName);
  };

  const title = replaceTokens(locData.Title || "");
  const body = replaceTokens(locData.Body || "");
  const options = locData.Options || [];
  const hasFreeText = currentTemplate.HasFreeTextOption;

  // Get original lengths (before token replacement) for character count display
  const titleLength = locData.Title?.length || 0;
  const bodyLength = locData.Body?.length || 0;
  const optionsClassLength = options.length > 0
    ? Math.max(...options.map(opt => opt.length))
    : 0;

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
    titleLength,
    bodyLength,
    optionsClassLength,
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

  // Initialize diagnostics collection
  diagnosticCollection =
    vscode.languages.createDiagnosticCollection("purview-dlp");
  context.subscriptions.push(diagnosticCollection);

  // Validate all open documents on activation
  vscode.workspace.textDocuments.forEach((document) => {
    if (document.languageId === "purview-dlp-json") {
      validateDocument(document);
    }
  });

  // Listen for document changes
  context.subscriptions.push(
    vscode.workspace.onDidChangeTextDocument((event) => {
      if (event.document.languageId === "purview-dlp-json") {
        validateDocument(event.document);
      }
    }),
  );

  // Listen for document opens
  context.subscriptions.push(
    vscode.workspace.onDidOpenTextDocument((document) => {
      if (document.languageId === "purview-dlp-json") {
        validateDocument(document);
      }
    }),
  );

  // Listen for document closes to clear diagnostics
  context.subscriptions.push(
    vscode.workspace.onDidCloseTextDocument((document) => {
      diagnosticCollection.delete(document.uri);
    }),
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
          Options: ["Option 1", "Option 2", "Option 3"],
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
          "No preview panel is open. Use 'Purview: Preview DLP Oversharing Dialog' first to open a preview.",
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
                "Inserts the name of the sensitivity label that was matched. Supported condition: Content contains sensitivity label",
            },
            {
              label: "%%MatchedAttachmentName%%",
              detail: "Matched Attachment Name",
              documentation:
                "Displays matched attachments. Supported conditions: Content contains sensitive information, Content contains sensitivity label, Attachment is not labeled, File extension is",
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

  // Register completion provider for formatting tags
  const formattingCompletionProvider =
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

          // Check if we're typing after <
          const linePrefix = document
            .lineAt(position)
            .text.slice(0, position.character);
          if (!linePrefix.endsWith("<")) {
            return undefined;
          }

          // Check if we're inside a Body string value (formatting only supported in Body)
          const line = document.lineAt(position).text;
          if (!line.includes('"Body"')) {
            // Look at previous lines to see if we're in a Body multi-line string
            let inBody = false;
            for (let i = position.line - 1; i >= 0; i--) {
              const prevLine = document.lineAt(i).text;
              if (prevLine.includes('"Body"')) {
                inBody = true;
                break;
              }
              if (
                prevLine.includes('"Title"') ||
                prevLine.includes('"Options"') ||
                prevLine.includes("}") ||
                prevLine.includes("]")
              ) {
                break;
              }
            }
            if (!inBody) {
              return undefined;
            }
          }

          // Create completion items for formatting tags
          const tags = [
            {
              label: "Bold",
              insertText: "Bold>${1:text}</Bold>",
              detail: "Bold text (Body only)",
              documentation: "Makes the enclosed text bold. Note: Formatting tags are only supported in the Body field.",
            },
            {
              label: "Italic",
              insertText: "Italic>${1:text}</Italic>",
              detail: "Italic text (Body only)",
              documentation: "Makes the enclosed text italic. Note: Formatting tags are only supported in the Body field.",
            },
            {
              label: "Underline",
              insertText: "Underline>${1:text}</Underline>",
              detail: "Underlined text (Body only)",
              documentation: "Underlines the enclosed text. Note: Formatting tags are only supported in the Body field.",
            },
            {
              label: "LineBreak",
              insertText: "LineBreak />",
              detail: "Line break (Body only)",
              documentation: "Inserts a line break. Note: Formatting tags are only supported in the Body field.",
            },
            {
              label: "br",
              insertText: "br />",
              detail: "Line break alternative (Body only)",
              documentation: "Inserts a line break (alternative syntax). Note: Formatting tags are only supported in the Body field.",
            },
          ];

          return tags.map((tag) => {
            const item = new vscode.CompletionItem(
              tag.label,
              vscode.CompletionItemKind.Snippet,
            );
            item.detail = tag.detail;
            item.documentation = new vscode.MarkdownString(tag.documentation);
            item.insertText = new vscode.SnippetString(tag.insertText);
            return item;
          });
        },
      },
      "<", // Trigger on < character
    );

  // Toggle Token Auto-Completion command
  const toggleTokenCompletionDisposable = vscode.commands.registerCommand(
    "purview-dlp-oversharing-dialogs.toggleTokenCompletion",
    async () => {
      const config = vscode.workspace.getConfiguration("purviewDlp");
      const current = config.get<boolean>("enableTokenCompletion", true);
      await config.update(
        "enableTokenCompletion",
        !current,
        vscode.ConfigurationTarget.Global,
      );
      vscode.window.showInformationMessage(
        `Purview DLP token auto-completion ${!current ? "enabled" : "disabled"}.`,
      );
    },
  );

  context.subscriptions.push(createTemplateDisposable);
  context.subscriptions.push(testDialogDisposable);
  context.subscriptions.push(switchPreviewLanguageDisposable);
  context.subscriptions.push(toggleTokenCompletionDisposable);
  context.subscriptions.push(tokenCompletionProvider);
  context.subscriptions.push(formattingCompletionProvider);
}

// This method is called when your extension is deactivated
export function deactivate() {}
