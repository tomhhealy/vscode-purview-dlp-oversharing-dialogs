import * as assert from "assert";
import * as vscode from "vscode";

suite("Extension Test Suite", () => {
  suiteSetup(async () => {
    // Ensure the extension is activated before running tests
    const ext = vscode.extensions.getExtension(
      "tomhhealy.purview-dlp-oversharing-dialogs",
    );
    if (ext && !ext.isActive) {
      await ext.activate();
    }
  });

  test("Extension should be present", () => {
    const extension = vscode.extensions.getExtension(
      "tomhhealy.purview-dlp-oversharing-dialogs",
    );
    assert.ok(extension, "Extension should be installed");
  });

  test("Extension should activate", async () => {
    const extension = vscode.extensions.getExtension(
      "tomhhealy.purview-dlp-oversharing-dialogs",
    );
    assert.ok(extension, "Extension should be installed");

    if (!extension.isActive) {
      await extension.activate();
    }
    assert.strictEqual(extension.isActive, true, "Extension should be active");
  });

  suite("Command Registration", () => {
    test("createTemplate command should be registered", async () => {
      const commands = await vscode.commands.getCommands(true);
      assert.ok(
        commands.includes("purview-dlp-oversharing-dialogs.createTemplate"),
        "createTemplate command should be registered",
      );
    });

    test("testDialog command should be registered", async () => {
      const commands = await vscode.commands.getCommands(true);
      assert.ok(
        commands.includes("purview-dlp-oversharing-dialogs.testDialog"),
        "testDialog command should be registered",
      );
    });

    test("switchPreviewLanguage command should be registered", async () => {
      const commands = await vscode.commands.getCommands(true);
      assert.ok(
        commands.includes(
          "purview-dlp-oversharing-dialogs.switchPreviewLanguage",
        ),
        "switchPreviewLanguage command should be registered",
      );
    });

    test("toggleTokenCompletion command should be registered", async () => {
      const commands = await vscode.commands.getCommands(true);
      assert.ok(
        commands.includes(
          "purview-dlp-oversharing-dialogs.toggleTokenCompletion",
        ),
        "toggleTokenCompletion command should be registered",
      );
    });
  });

  suite("Language Registration", () => {
    test("purview-dlp-json language should be registered", () => {
      const languages = vscode.languages.getLanguages();
      // getLanguages returns a Thenable, so we need to handle it
      return languages.then((langs) => {
        assert.ok(
          langs.includes("purview-dlp-json"),
          "purview-dlp-json language should be registered",
        );
      });
    });
  });

  suite("Configuration", () => {
    test("purviewDlp configuration section should exist", () => {
      const config = vscode.workspace.getConfiguration("purviewDlp");
      assert.ok(config, "purviewDlp configuration should exist");
    });

    test("testDialog configuration defaults should be set", () => {
      const config = vscode.workspace.getConfiguration("purviewDlp.testDialog");

      const matchedRecipientsList = config.get<string>("matchedRecipientsList");
      assert.strictEqual(
        matchedRecipientsList,
        "example@example.com",
        "matchedRecipientsList should have default value",
      );

      const matchedLabelName = config.get<string>("matchedLabelName");
      assert.strictEqual(
        matchedLabelName,
        "Confidential",
        "matchedLabelName should have default value",
      );

      const matchedConditions = config.get<string>("matchedConditions");
      assert.strictEqual(
        matchedConditions,
        "Credit Card Number detected",
        "matchedConditions should have default value",
      );
    });

    test("enableTokenCompletion should default to true", () => {
      const config = vscode.workspace.getConfiguration("purviewDlp");
      const enableTokenCompletion = config.get<boolean>("enableTokenCompletion");
      assert.strictEqual(
        enableTokenCompletion,
        true,
        "enableTokenCompletion should default to true",
      );
    });
  });
});
