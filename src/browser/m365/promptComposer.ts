import { buildClickDispatcher } from "../actions/domEvents.js";
import { logDomFailure } from "../domDebug.js";
import type { BrowserLogger, ChromeClient } from "../types.js";
import { delay } from "../utils.js";
import { M365_INPUT_SELECTORS, M365_SEND_BUTTON_SELECTORS } from "./constants.js";

const ENTER_KEY_EVENT = {
	key: "Enter",
	code: "Enter",
	windowsVirtualKeyCode: 13,
	nativeVirtualKeyCode: 13,
} as const;
const ENTER_KEY_TEXT = "\r";

export async function submitPrompt(
	Runtime: ChromeClient["Runtime"],
	Input: ChromeClient["Input"],
	promptText: string,
	logger: BrowserLogger,
): Promise<void> {
	const focusResult = await Runtime.evaluate({
		expression: `(() => {
			${buildClickDispatcher()}
			const selectors = ${JSON.stringify(M365_INPUT_SELECTORS)};
			const focusNode = (node) => {
				if (!node) return false;
				dispatchClickSequence(node);
				if (typeof node.focus === "function") node.focus();
				const doc = node.ownerDocument;
				const selection = doc?.getSelection?.();
				if (selection) {
					const range = doc.createRange();
					range.selectNodeContents(node);
					range.collapse(false);
					selection.removeAllRanges();
					selection.addRange(range);
				}
				return true;
			};
			for (const selector of selectors) {
				const node = document.querySelector(selector);
				if (node && focusNode(node)) {
					return { focused: true };
				}
			}
			return { focused: false };
		})()`,
		returnByValue: true,
		awaitPromise: true,
	});

	if (!focusResult.result?.value?.focused) {
		await logDomFailure(Runtime, logger, "m365-focus-textarea");
		throw new Error("Failed to focus M365 Copilot prompt textarea");
	}

	await Input.insertText({ text: promptText });

	// Wait for button to appear and become enabled
	// The button appears after typing but is disabled briefly
	await delay(500);

	// Retry logic: wait for button to be enabled before clicking
	const maxRetries = 10;
	let clicked = false;

	for (let attempt = 0; attempt < maxRetries; attempt++) {
		if (attempt > 0) {
			await delay(300);
		}

		const clickResult = await Runtime.evaluate({
			expression: `(() => {
				${buildClickDispatcher()}
				const selectors = ${JSON.stringify(M365_SEND_BUTTON_SELECTORS)};
				for (const selector of selectors) {
					const button = document.querySelector(selector);
					if (!button) continue;
					const ariaDisabled = button.getAttribute("aria-disabled");
					const dataDisabled = button.getAttribute("data-disabled");
					const style = window.getComputedStyle(button);
					const disabled =
						button.hasAttribute("disabled") ||
						ariaDisabled === "true" ||
						dataDisabled === "true" ||
						style.pointerEvents === "none" ||
						style.display === "none" ||
						style.visibility === "hidden";
					if (disabled) {
						return { found: true, enabled: false };
					}
					dispatchClickSequence(button);
					return { found: true, enabled: true, clicked: true };
				}
				return { found: false };
			})()`,
			returnByValue: true,
		});

		const result = clickResult.result?.value;
		if (result?.clicked) {
			clicked = true;
			logger("Clicked send button (M365)");
			break;
		} else if (result?.found && !result?.enabled) {
			// Button found but still disabled, keep waiting
			logger(`Send button found but disabled, waiting... (attempt ${attempt + 1}/${maxRetries})`);
		}
	}

	if (!clicked) {
		// Fallback: Try submitting via form or Enter key on focused element
		const submitted = await Runtime.evaluate({
			expression: `(() => {
				// Try to submit the form if button is in a form
				const button = document.querySelector('button[type="submit"][aria-label*="send" i]');
				if (button) {
					const form = button.closest('form');
					if (form) {
						form.requestSubmit(button);
						return { method: 'form' };
					}
				}

				// Try clicking the button even if it seems disabled
				${buildClickDispatcher()}
				const selectors = ${JSON.stringify(M365_SEND_BUTTON_SELECTORS)};
				for (const selector of selectors) {
					const btn = document.querySelector(selector);
					if (btn) {
						dispatchClickSequence(btn);
						return { method: 'force-click' };
					}
				}

				return { method: 'none' };
			})()`,
			returnByValue: true,
			awaitPromise: true,
		});

		const method = submitted.result?.value?.method;
		if (method === "form") {
			logger("Submitted prompt via form submit (M365)");
		} else if (method === "force-click") {
			logger("Force-clicked send button (M365)");
		} else {
			// Last resort: dispatch Enter key on focused element
			await Input.dispatchKeyEvent({
				type: "keyDown",
				...ENTER_KEY_EVENT,
				text: ENTER_KEY_TEXT,
				unmodifiedText: ENTER_KEY_TEXT,
			});
			await Input.dispatchKeyEvent({
				type: "keyUp",
				...ENTER_KEY_EVENT,
			});
			logger("Submitted prompt via Enter key (M365)");
		}
	}
}
