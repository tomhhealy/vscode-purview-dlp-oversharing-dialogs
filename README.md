# Purview DLP Oversharing Dialogs

A Visual Studio Code extension for creating and previewing Microsoft Purview DLP oversharing dialog templates.

## Features

- **Create Template**: Generate a new Purview DLP oversharing dialog template with support for multiple languages (up to 10)
- **Preview Dialog**: Live preview of your dialog template as it would appear in Outlook
- **Switch Preview Language**: Change the language being previewed without reopening the panel
- **Token IntelliSense**: Autocomplete for DLP tokens (`%%MatchedRecipientsList%%`, `%%MatchedLabelName%%`, `%%MatchedConditions%%`)
- **JSON Schema Validation**: Automatic validation of template structure

## Commands

- `Purview: Create oversharing dialog template` - Create a new template with selected languages
- `Purview: Preview Dialog` - Open a live preview of the current template
- `Purview: Switch Preview Language` - Change the preview language

## Extension Settings

This extension contributes the following settings for token placeholder values in preview:

- `purviewDlp.testDialog.matchedRecipientsList`: Placeholder for `%%MatchedRecipientsList%%` token
- `purviewDlp.testDialog.matchedLabelName`: Placeholder for `%%MatchedLabelName%%` token
- `purviewDlp.testDialog.matchedConditions`: Placeholder for `%%MatchedConditions%%` token

## File Association

Files with the `.purview-dlp.json` extension are automatically recognized and get syntax highlighting, validation, and IntelliSense.
