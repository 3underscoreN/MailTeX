/* global document, Office, console, HTMLElement, HTMLTextAreaElement, HTMLInputElement, HTMLButtonElement, HTMLImageElement */

import texToPngImage from "../utils/tex2img";
import removeInlineMathDelimiters from "../utils/removeMathDelimiters";

const fallbackBackground = "#FFFFFF";
const fallbackForeground = "#000000";

function isHexColor(value: string | undefined): value is string {
  return Boolean(value && /^#[0-9a-f]{6}$/i.test(value));
}

function applyOfficeTheme(theme?: Office.OfficeTheme): string {
  const backgroundColor = isHexColor(theme?.bodyBackgroundColor)
    ? theme.bodyBackgroundColor
    : fallbackBackground;
  const foregroundColor = isHexColor(theme?.bodyForegroundColor)
    ? theme.bodyForegroundColor
    : fallbackForeground;
  const root = document.documentElement;
  root.style.setProperty("--office-background", backgroundColor);
  root.style.setProperty("--office-foreground", foregroundColor);
  root.style.setProperty(
    "--office-muted",
    foregroundColor === fallbackForeground ? "#616161" : "#bdbdbd"
  );
  root.style.setProperty(
    "--office-border",
    foregroundColor === fallbackForeground ? "#d1d1d1" : "#5a5a5a"
  );
  document.body.style.backgroundColor = backgroundColor;
  document.body.style.color = foregroundColor;
  return backgroundColor;
}

Office.onReady(async (info) => {
  if (info.host !== Office.HostType.Outlook) {
    return;
  }

  const editorElement = document.getElementById("tex-input");
  const blockModeElement = document.getElementById("block-mode");
  const previewElement = document.getElementById("preview");
  const statusElement = document.getElementById("render-status");
  const insertButtonElement = document.getElementById("insert-button");
  if (
    !(editorElement instanceof HTMLTextAreaElement) ||
    !(blockModeElement instanceof HTMLInputElement) ||
    !(previewElement instanceof HTMLElement) ||
    !(statusElement instanceof HTMLElement) ||
    !(insertButtonElement instanceof HTMLButtonElement)
  ) {
    throw new Error("The MailTex task pane is missing required interface elements.");
  }
  const editor = editorElement;
  const blockMode = blockModeElement;
  const preview = previewElement;
  const status = statusElement;
  const insertButton = insertButtonElement;

  let backgroundColor = fallbackBackground;
  let previewImage: HTMLImageElement | undefined;
  let renderSequence = 0;

  if (Office.context.requirements.isSetSupported("Mailbox", "1.14")) {
    backgroundColor = applyOfficeTheme(Office.context.officeTheme);
    Office.context.mailbox.addHandlerAsync(
      Office.EventType.OfficeThemeChanged,
      (event: Office.OfficeThemeChangedEventArgs) => {
        backgroundColor = applyOfficeTheme(event.officeTheme);
        void updatePreview();
      },
      (result) => {
        if (result.status === Office.AsyncResultStatus.Failed) {
          console.error("Failed to listen for Office theme changes:", result.error.message);
        }
      }
    );
  } else {
    backgroundColor = applyOfficeTheme();
  }

  async function updatePreview(): Promise<void> {
    const currentSequence = ++renderSequence;
    const source = removeInlineMathDelimiters(editor.value);
    const mode = blockMode.checked ? "DISPLAY" : "INLINE";
    previewImage = undefined;
    insertButton.disabled = true;

    if (!source.tex) {
      preview.replaceChildren();
      const placeholder = document.createElement("span");
      placeholder.className = "preview-placeholder";
      placeholder.textContent = "Your rendered formula will appear here.";
      preview.appendChild(placeholder);
      status.textContent = "Enter TeX to see a preview.";
      return;
    }

    status.textContent = "Rendering preview…";
    try {
      const image = await texToPngImage(source.tex, mode, backgroundColor);
      if (currentSequence !== renderSequence) {
        return;
      }
      if (!image) {
        preview.replaceChildren();
        status.textContent = "This TeX expression is invalid.";
        return;
      }

      previewImage = image;
      preview.replaceChildren(image);
      status.textContent = "Preview ready.";
      insertButton.disabled = false;
    } catch (error) {
      if (currentSequence !== renderSequence) {
        return;
      }
      preview.replaceChildren();
      status.textContent = "Preview could not be rendered.";
      console.error("Failed to render TeX preview:", error);
    }
  }

  editor.addEventListener("input", () => {
    void updatePreview();
  });

  blockMode.addEventListener("change", () => {
    void updatePreview();
  });

  insertButton.addEventListener("click", () => {
    const item = Office.context.mailbox.item;
    if (!item || !previewImage) {
      status.textContent = "Enter valid TeX before inserting it.";
      return;
    }

    insertButton.disabled = true;
    status.textContent = "Inserting formula…";
    item.body.setSelectedDataAsync(
      previewImage.outerHTML,
      { coercionType: Office.CoercionType.Html },
      (result) => {
        if (result.status === Office.AsyncResultStatus.Failed) {
          status.textContent = "Could not insert the formula into the message.";
          console.error("Failed to insert TeX image:", result.error.message);
          insertButton.disabled = !previewImage;
          return;
        }
        status.textContent = "Formula inserted.";
        insertButton.disabled = !previewImage;
      }
    );
  });

  await updatePreview();
});
