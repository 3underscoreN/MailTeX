import type { RenderMode } from "../types/RenderMode";

function removeInlineMathDelimiters(selection: string): { tex: string; mode: RenderMode } {
  const tex = selection
    .replace(/[\u200B-\u200D\uFEFF]/g, "")
    .replace(/\u00A0/g, " ")
    .replace(/\r\n?|\u2028|\u2029/g, "\n")
    .trim();

  if (tex.startsWith("\\[") && tex.endsWith("\\]")) {
    return { tex: tex.slice(2, -2).trim(), mode: "DISPLAY" };
  }

  if (tex.startsWith("$$") && tex.endsWith("$$")) {
    return { tex: tex.slice(2, -2).trim(), mode: "DISPLAY" };
  }

  if (tex.startsWith("$") && tex.endsWith("$") && !tex.startsWith("$$") && !tex.endsWith("$$")) {
    return { tex: tex.slice(1, -1).trim(), mode: "INLINE" };
  }

  if (tex.startsWith("\\(") && tex.endsWith("\\)")) {
    return { tex: tex.slice(2, -2).trim(), mode: "INLINE" };
  }

  return { tex, mode: "INLINE" };
}

export default removeInlineMathDelimiters;
