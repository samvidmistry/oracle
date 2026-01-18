import { logDomFailure } from "../domDebug.js";
import type { BrowserLogger, ChromeClient } from "../types.js";
import { delay } from "../utils.js";
import { M365_COPILOT_URL, M365_INPUT_SELECTORS } from "./constants.js";

const AUTH_HOSTS = [
	"login.microsoftonline.com",
	"login.live.com",
	"login.windows.net",
	"account.live.com",
];

const AUTH_PATH_MARKERS = ["/signin", "/oauth2", "/authorize", "/consent", "/mfa"];

function isAuthUrl(rawUrl: string): boolean {
	try {
		const parsed = new URL(rawUrl);
		if (AUTH_HOSTS.some((host) => parsed.hostname.includes(host))) {
			return true;
		}
		const path = parsed.pathname.toLowerCase();
		return AUTH_PATH_MARKERS.some((marker) => path.includes(marker));
	} catch {
		return false;
	}
}

async function currentUrl(Runtime: ChromeClient["Runtime"]): Promise<string | null> {
	const { result } = await Runtime.evaluate({
		expression: "typeof location === 'object' && location.href ? location.href : null",
		returnByValue: true,
	});
	return typeof result?.value === "string" ? result.value : null;
}

async function hasLoginCta(Runtime: ChromeClient["Runtime"]): Promise<boolean> {
	const { result } = await Runtime.evaluate({
		expression: `(() => {
			const labels = ["sign in", "signin", "log in", "login", "pick an account", "continue"];
			const buttons = document.querySelectorAll("button, a, [role=\\"button\\"]");
			for (const btn of buttons) {
				const text = (btn.textContent || "").toLowerCase().trim();
				if (labels.some((label) => text.includes(label))) return true;
			}
			return false;
		})()`,
		returnByValue: true,
	});

	return Boolean(result?.value);
}

export async function ensureLoggedIn(
	Runtime: ChromeClient["Runtime"],
	logger: BrowserLogger,
): Promise<void> {
	const url = await currentUrl(Runtime);
	if (url && isAuthUrl(url)) {
		throw new Error(
			"M365 auth page detected. Run with --browser-provider m365-copilot and sign in in the opened Chrome window.",
		);
	}
	const cta = await hasLoginCta(Runtime);
	if (cta) {
		throw new Error(
			"M365 login required. Run with --browser-provider m365-copilot and sign in in the opened Chrome window.",
		);
	}
	logger("M365 Copilot login check passed");
}

export async function ensurePromptReady(
	Runtime: ChromeClient["Runtime"],
	timeoutMs: number,
	logger: BrowserLogger,
): Promise<void> {
	const deadline = Date.now() + timeoutMs;
	while (Date.now() < deadline) {
		const { result } = await Runtime.evaluate({
			expression: `(() => {
				const selectors = ${JSON.stringify(M365_INPUT_SELECTORS)};
				for (const selector of selectors) {
					const node = document.querySelector(selector);
					if (node && !node.hasAttribute("disabled")) {
						return true;
					}
				}
				return false;
			})()`,
			returnByValue: true,
		});
		if (result?.value) {
			return;
		}
		await delay(200);
	}
	await logDomFailure(Runtime, logger, "m365-prompt-textarea");
	throw new Error("M365 Copilot prompt textarea did not appear before timeout");
}

export async function navigateToCopilot(
	Page: ChromeClient["Page"],
	Runtime: ChromeClient["Runtime"],
	logger: BrowserLogger,
	url: string = M365_COPILOT_URL,
): Promise<void> {
	logger(`Navigating to ${url}`);
	await Page.navigate({ url });
	const deadline = Date.now() + 45_000;
	while (Date.now() < deadline) {
		const { result } = await Runtime.evaluate({
			expression: "document.readyState",
			returnByValue: true,
		});
		if (result?.value === "complete" || result?.value === "interactive") {
			return;
		}
		await delay(100);
	}
	throw new Error("M365 Copilot page did not reach ready state in time");
}
