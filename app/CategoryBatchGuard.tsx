"use client";

import { useEffect } from "react";

const BUTTON_ID = "postsheet-category-batch-save";

export default function CategoryBatchGuard() {
  useEffect(() => {
    let applying = false;

    const isCategoryEditor = () => {
      const active = document.querySelector<HTMLButtonElement>(".stepbar button.active");
      return active?.textContent?.includes("카테고리") ?? false;
    };

    const installBatchButton = () => {
      const editor = document.querySelector<HTMLElement>(".group-workspace .editor-card .editor-fields");
      if (!editor || !isCategoryEditor()) {
        document.getElementById(BUTTON_ID)?.remove();
        return;
      }

      const applyFields = Array.from(editor.querySelectorAll<HTMLElement>(".apply-field"));
      if (applyFields.length !== 3) return;

      const oldButtons = applyFields
        .map((field) => field.querySelector<HTMLButtonElement>("button"))
        .filter((button): button is HTMLButtonElement => Boolean(button));

      oldButtons.forEach((button) => {
        button.style.display = "none";
      });

      if (document.getElementById(BUTTON_ID)) return;

      const button = document.createElement("button");
      button.id = BUTTON_ID;
      button.type = "button";
      button.className = "primary";
      button.textContent = "카테고리 3개 코드 한 번에 저장";
      button.style.width = "100%";
      button.style.marginTop = "16px";

      button.addEventListener("click", () => {
        if (applying) return;

        const currentEditor = document.querySelector<HTMLElement>(".group-workspace .editor-card .editor-fields");
        const currentFields = currentEditor
          ? Array.from(currentEditor.querySelectorAll<HTMLElement>(".apply-field"))
          : [];
        const inputs = currentFields
          .map((field) => field.querySelector<HTMLInputElement>("input"))
          .filter((input): input is HTMLInputElement => Boolean(input));

        if (inputs.length !== 3 || inputs.some((input) => !input.value.trim())) {
          window.alert("ESM 카테고리, G마켓 노출코드, 옥션 노출코드를 모두 입력한 뒤 저장해 주세요.");
          return;
        }

        applying = true;
        button.disabled = true;
        button.textContent = "저장 중…";

        const clickNext = (index: number) => {
          const latestEditor = document.querySelector<HTMLElement>(".group-workspace .editor-card .editor-fields");
          const latestButtons = latestEditor
            ? Array.from(latestEditor.querySelectorAll<HTMLButtonElement>(".apply-field button"))
            : [];
          const target = latestButtons[index];

          if (!target) {
            applying = false;
            return;
          }

          target.click();

          if (index < 2) {
            window.setTimeout(() => clickNext(index + 1), 80);
          } else {
            window.setTimeout(() => {
              applying = false;
              installBatchButton();
            }, 150);
          }
        };

        clickNext(0);
      });

      editor.appendChild(button);
    };

    const observer = new MutationObserver(() => installBatchButton());
    observer.observe(document.body, { childList: true, subtree: true });
    installBatchButton();

    return () => observer.disconnect();
  }, []);

  return null;
}
