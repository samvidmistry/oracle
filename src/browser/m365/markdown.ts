import { buildClickDispatcher } from "../actions/domEvents.js";
import { logDomFailure } from "../domDebug.js";
import type { BrowserLogger, ChromeClient } from "../types.js";
import { M365_COPY_BUTTON_SELECTORS } from "./constants.js";

export async function captureAssistantMarkdown(
	Runtime: ChromeClient["Runtime"],
	meta: { messageId?: string | null; turnId?: string | null },
	logger: BrowserLogger,
): Promise<string | null> {
	const expression = `(() => {
		${buildClickDispatcher()}
		const selectors = ${JSON.stringify(M365_COPY_BUTTON_SELECTORS)};
		const hint = ${JSON.stringify(meta ?? {})};
		const candidates = [];
		const findScoped = (root) => {
			if (!root) return null;
			for (const selector of selectors) {
				const button = root.querySelector(selector);
				if (button) return button;
			}
			return null;
		};
		if (hint?.messageId) {
			const node = document.querySelector('[data-message-id="' + hint.messageId + '"]');
			const scoped = findScoped(node);
			if (scoped) candidates.push(scoped);
		}
		if (hint?.turnId) {
			const node = document.querySelector('[data-testid="' + hint.turnId + '"]');
			const scoped = findScoped(node);
			if (scoped) candidates.push(scoped);
		}
		for (const selector of selectors) {
			document.querySelectorAll(selector).forEach((node) => candidates.push(node));
		}
		const button = candidates.find(Boolean);
		if (!button) return { status: 'missing' };
		const clipboard = navigator.clipboard;
		if (!clipboard) return { status: 'no-clipboard' };
		const state = { text: '', updatedAt: 0 };
		const originalWriteText = clipboard.writeText?.bind(clipboard);
		const originalWrite = clipboard.write?.bind(clipboard);
		clipboard.writeText = (value) => {
			state.text = typeof value === 'string' ? value : '';
			state.updatedAt = Date.now();
			return Promise.resolve();
		};
		clipboard.write = async (items) => {
			try {
				const list = Array.isArray(items) ? items : items ? [items] : [];
				for (const item of list) {
					const types = Array.isArray(item.types) ? item.types : [];
					if (types.includes('text/plain') && typeof item.getType === 'function') {
						const blob = await item.getType('text/plain');
						const text = await blob.text();
						state.text = text ?? '';
						state.updatedAt = Date.now();
						break;
					}
				}
			} catch {
				state.text = '';
				state.updatedAt = Date.now();
			}
			return Promise.resolve();
		};
		button.scrollIntoView({ block: 'center', behavior: 'instant' });
		dispatchClickSequence(button);
		return new Promise((resolve) => {
			const deadline = Date.now() + 4000;
			const tick = () => {
				if (state.text && Date.now() - state.updatedAt > 200) {
					resolve({ status: 'success', markdown: state.text });
					return;
				}
				if (Date.now() > deadline) {
					resolve({ status: 'timeout' });
					return;
				}
				setTimeout(tick, 120);
			};
			tick();
		}).finally(() => {
			if (originalWriteText) clipboard.writeText = originalWriteText;
			if (originalWrite) clipboard.write = originalWrite;
		});
	})()`;
	const { result } = await Runtime.evaluate({
		expression,
		returnByValue: true,
		awaitPromise: true,
	});
	const value = result?.value as { status?: string; markdown?: string } | undefined;
	if (value?.status === "success" && typeof value.markdown === "string" && value.markdown.trim()) {
		return value.markdown;
	}
	if (value?.status && value.status !== "missing") {
		logger(`M365 markdown copy status: ${value.status}`);
		await logDomFailure(Runtime, logger, "m365-copy-markdown");
	}
	return null;
}
