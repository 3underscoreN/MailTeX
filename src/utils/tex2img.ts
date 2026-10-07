/* global document, HTMLImageElement, HTMLElement */

import katex from "katex";
import type { RenderMode } from "../types/RenderMode";

const imageScale = 2;

function getContrastingTextColor(backgroundColor: string): string {
  const match = /^#?([0-9a-f]{6})$/i.exec(backgroundColor);
  if (!match) {
    throw new Error(`Unsupported background color: ${backgroundColor}`);
  }

  const hex = match[1];
  const channels = [0, 2, 4].map((index) => parseInt(hex.slice(index, index + 2), 16) / 255);
  const luminance = channels
    .map((channel) => (channel <= 0.04045 ? channel / 12.92 : ((channel + 0.055) / 1.055) ** 2.4))
    .reduce((sum, channel, index) => sum + channel * [0.2126, 0.7152, 0.0722][index], 0);

  return luminance > 0.179 ? "#000000" : "#FFFFFF";
}

/**
 * Renders TeX as a PNG image. Invalid TeX returns undefined; rendering failures reject.
 */
export default async function texToPngImage(
  tex: string,
  mode: RenderMode,
  backgroundColor: string
): Promise<HTMLImageElement | undefined> {
  const renderTarget = document.createElement("div");
  renderTarget.style.position = "fixed";
  renderTarget.style.left = "-10000px";
  renderTarget.style.top = "0";
  renderTarget.style.display = "inline-block";
  renderTarget.style.backgroundColor = backgroundColor;
  document.body.appendChild(renderTarget);

  try {
    const textColor = getContrastingTextColor(backgroundColor);
    katex.render(tex, renderTarget, {
      throwOnError: true,
      output: "html",
      displayMode: mode === "DISPLAY",
    });
    const katexElement = renderTarget.querySelector<HTMLElement>(".katex");
    if (!katexElement) {
      throw new Error("KaTeX did not produce renderable output.");
    }
    katexElement.style.color = textColor;

    if (document.fonts) {
      await document.fonts.ready;
    }
    const { default: html2canvas } = await import("html2canvas-pro");
    const canvas = await html2canvas(renderTarget, {
      backgroundColor: null,
      scale: imageScale,
    });

    const image = document.createElement("img");
    image.src = canvas.toDataURL("image/png");
    image.alt = tex;
    image.dataset.mailtexMode = mode;
    image.width = Math.ceil(canvas.width / imageScale);
    image.height = Math.ceil(canvas.height / imageScale);
    if (mode === "DISPLAY") {
      image.style.display = "block";
      image.style.margin = "0 auto";
    }
    return image;
  } catch (error) {
    if (error instanceof katex.ParseError) {
      return undefined;
    }
    throw error;
  } finally {
    renderTarget.remove();
  }
}
