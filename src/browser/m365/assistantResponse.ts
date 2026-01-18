import { logDomFailure } from "../domDebug.js";
import type { BrowserLogger, ChromeClient } from "../types.js";
import { delay } from "../utils.js";
import {
	M365_CHAT_CONTAINER_SELECTORS,
	M365_LOADING_SELECTORS,
	M365_STOP_BUTTON_SELECTORS,
	M365_SUGGESTION_SELECTORS,
	M365_TURN_SELECTORS,
} from "./constants.js";

export interface M365AssistantSnapshot {
	text: string;
	html?: string;
	turnId?: string | null;
	messageId?: string | null;
	turnIndex?: number | null;
}

function buildSnapshotExpression(minTurnIndex?: number): string {
	const minTurnLiteral =
		typeof minTurnIndex === "number" && Number.isFinite(minTurnIndex) && minTurnIndex >= 0
			? Math.floor(minTurnIndex)
			: -1;
	return `(() => {
		const MIN_TURN_INDEX = ${minTurnLiteral};
		const containerSelectors = ${JSON.stringify(M365_CHAT_CONTAINER_SELECTORS)};
		const turnSelectors = ${JSON.stringify(M365_TURN_SELECTORS)};
		const resolveContainer = () => {
			for (const selector of containerSelectors) {
				const node = document.querySelector(selector);
				if (node) return node;
			}
			return document.body;
		};
		const container = resolveContainer();
		const nodes = Array.from(container.querySelectorAll(turnSelectors.join(',')));
		const candidates = nodes.filter((node) => {
			if (!(node instanceof HTMLElement)) return false;
			const role = (node.getAttribute('data-message-author-role') || node.getAttribute('data-turn') || '').toLowerCase();
			if (role === 'assistant') return true;
			const label = (node.getAttribute('aria-label') || '').toLowerCase();
			if (label.includes('copilot')) return true;
			const testId = (node.getAttribute('data-testid') || '').toLowerCase();
			if (testId.includes('assistant')) return true;
			// Check for M365 Copilot specific structure
			if (testId === 'lastchatmessage') return true;
			if (node.querySelector('[data-testid="markdown-reply"]')) return true;
			return node.querySelector('[data-message-author-role="assistant"], [data-turn="assistant"], [data-testid*=assistant]') != null;
		});
		for (let index = candidates.length - 1; index >= 0; index -= 1) {
			const node = candidates[index];
			const turnIndex = nodes.indexOf(node);
			if (MIN_TURN_INDEX >= 0 && turnIndex >= 0 && turnIndex < MIN_TURN_INDEX) {
				continue;
			}
			const text = node.innerText?.trim() || node.textContent?.trim() || '';
			if (!text) continue;
			return {
				text,
				html: node.innerHTML ?? '',
				turnId: node.getAttribute('data-testid'),
				messageId: node.getAttribute('data-message-id'),
				turnIndex,
			};
		}
		return null;
	})()`;
}

async function isStillStreaming(
	Runtime: ChromeClient["Runtime"],
	logger?: BrowserLogger,
): Promise<boolean> {
	const expression = `(() => {
		// Check for suggestions - if present, response is complete
		const suggestionSelectors = ${JSON.stringify(M365_SUGGESTION_SELECTORS)};
		for (const sel of suggestionSelectors) {
			const elem = document.querySelector(sel);
			if (elem) return { streaming: false, reason: 'suggestions-present', selector: sel };
		}

		// Check for stop button - indicates still streaming
		const stopSelectors = ${JSON.stringify(M365_STOP_BUTTON_SELECTORS)};
		for (const sel of stopSelectors) {
			const elem = document.querySelector(sel);
			if (elem) return { streaming: true, reason: 'stop-button', selector: sel };
		}

		// Check for loading indicators
		const loadingSelectors = ${JSON.stringify(M365_LOADING_SELECTORS)};
		const containerSelectors = ${JSON.stringify(M365_CHAT_CONTAINER_SELECTORS)};
		let root = document.body;
		for (const selector of containerSelectors) {
			const node = document.querySelector(selector);
			if (node) { root = node; break; }
		}
		for (const sel of loadingSelectors) {
			const elem = root.querySelector(sel);
			if (elem) return { streaming: true, reason: 'loading-indicator', selector: sel };
		}

		return { streaming: false };
	})()`;
	const { result } = await Runtime.evaluate({ expression, returnByValue: true });
	const value = result?.value as { streaming?: boolean; reason?: string; selector?: string } | boolean | undefined;
	if (typeof value === 'object' && value !== null) {
		if (value.streaming && logger) {
			logger(`Still streaming: ${value.reason} (${value.selector})`);
		} else if (!value.streaming && value.reason === 'suggestions-present' && logger) {
			logger('Response complete: suggestions appeared');
		}
		return Boolean(value.streaming);
	}
	return Boolean(value);
}

export async function readAssistantSnapshot(
	Runtime: ChromeClient["Runtime"],
	minTurnIndex?: number,
): Promise<M365AssistantSnapshot | null> {
	const { result } = await Runtime.evaluate({
		expression: buildSnapshotExpression(minTurnIndex),
		returnByValue: true,
	});
	const value = result?.value as M365AssistantSnapshot | null | undefined;
	if (!value || typeof value !== "object") {
		return null;
	}
	return value.text ? value : null;
}

export async function waitForAssistantResponse(
	Runtime: ChromeClient["Runtime"],
	timeoutMs: number,
	logger: BrowserLogger,
	minTurnIndex?: number,
): Promise<{
	text: string;
	html?: string;
	meta: { turnId?: string | null; messageId?: string | null };
}> {
	const deadline = Date.now() + timeoutMs;
	let lastText = "";
	let stableCount = 0;
	const requiredStableChecks = 5;
	const pollInterval = 500;
	while (Date.now() < deadline) {
		const snapshot = await readAssistantSnapshot(Runtime, minTurnIndex);
		const currentText = snapshot?.text ?? "";
		const streaming = await isStillStreaming(Runtime, logger);
		if (!streaming && currentText && currentText === lastText) {
			stableCount += 1;
			if (stableCount >= requiredStableChecks) {
				logger("M365 Copilot response complete (stabilized)");
				return {
					text: currentText,
					html: snapshot?.html ?? undefined,
					meta: {
						turnId: snapshot?.turnId ?? undefined,
						messageId: snapshot?.messageId ?? undefined,
					},
				};
			}
		} else {
			stableCount = 0;
			if (currentText) {
				lastText = currentText;
			}
		}
		await delay(pollInterval);
	}
	await logDomFailure(Runtime, logger, "m365-assistant-response");
	throw new Error("M365 Copilot response timeout");
}
