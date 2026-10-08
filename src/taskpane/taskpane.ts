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

function addInlineImage(item: Office.MessageCompose, image: HTMLImageElement): Promise<string> {
  const dataUrlPrefix = "data:image/png;base64,";
  if (!image.src.startsWith(dataUrlPrefix)) {
    return Promise.reject(new Error("The rendered formula is not a PNG data URL."));
  }

  const base64Image = image.src.slice(dataUrlPrefix.length);
  const attachmentName = `mailtex-formula-${Date.now()}.png`;
  return new Promise((resolve, reject) => {
    item.addFileAttachmentFromBase64Async(
      base64Image,
      attachmentName,
      { isInline: true },
      (result) => {
        if (result.status === Office.AsyncResultStatus.Failed) {
          reject(new Error(result.error.message));
        } else {
          resolve(result.value);
        }
      }
    );
  });
}

function getInlineImageContentId(
  item: Office.MessageCompose,
  attachmentId: string
): Promise<string> {
  return new Promise((resolve, reject) => {
    item.getAttachmentsAsync((result) => {
      if (result.status === Office.AsyncResultStatus.Failed) {
        reject(new Error(result.error.message));
        return;
      }

      const attachment = result.value.find(
        (candidate) => candidate.id === attachmentId && candidate.isInline
      );
      if (!attachment?.contentId) {
        reject(new Error("Outlook did not return a content ID for the inline image."));
        return;
      }
      resolve(attachment.contentId);
    });
  });
}

function removeInlineImage(item: Office.MessageCompose, attachmentId: string): Promise<void> {
  return new Promise((resolve, reject) => {
    item.removeAttachmentAsync(attachmentId, (result) => {
      if (result.status === Office.AsyncResultStatus.Failed) {
        reject(new Error(result.error.message));
      } else {
        resolve();
      }
    });
  });
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

  insertButton.addEventListener("click", async () => {
    const item = Office.context.mailbox.item;
    if (!item || !previewImage) {
      status.textContent = "Enter valid TeX before inserting it.";
      return;
    }

    if (!Office.context.requirements.isSetSupported("Mailbox", "1.16")) {
      status.textContent = "CID image insertion requires Outlook Mailbox requirement set 1.16.";
      return;
    }

    insertButton.disabled = true;
    status.textContent = "Inserting formula…";
    let attachmentId: string | undefined;
    try {
      if (item.itemType !== Office.MailboxEnums.ItemType.Message) {
        throw new Error("Inline formula insertion is only supported for messages.");
      }

      const message = item as Office.MessageCompose;
      attachmentId = await addInlineImage(message, previewImage);
      const contentId = await getInlineImageContentId(message, attachmentId);
      const image = previewImage.cloneNode(false) as HTMLImageElement;
      image.setAttribute("src", `cid:${contentId}`);
      await new Promise<void>((resolve, reject) => {
        message.body.setSelectedDataAsync(
          image.outerHTML,
          { coercionType: Office.CoercionType.Html },
          (result) => {
            if (result.status === Office.AsyncResultStatus.Failed) {
              reject(new Error(result.error.message));
            } else {
              resolve();
            }
          }
        );
      });
      status.textContent = "Formula inserted.";
    } catch (error) {
      status.textContent =
        error instanceof Error
          ? `Could not insert the formula: ${error.message}`
          : "Could not insert the formula into the message.";
      console.error("Failed to insert TeX image:", error);
      if (attachmentId) {
        try {
          await removeInlineImage(item as Office.MessageCompose, attachmentId);
        } catch (removeError) {
          console.error("Failed to remove the unused inline TeX attachment:", removeError);
        }
      }
    } finally {
      insertButton.disabled = !previewImage;
    }
  });

  await updatePreview();
});
