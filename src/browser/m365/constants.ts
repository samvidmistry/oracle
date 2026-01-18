export const M365_COPILOT_URL = "https://copilot.cloud.microsoft/";

export const M365_INPUT_SELECTORS = [
	'[data-testid="composer-input"]',
	"textarea#userInput",
	'textarea[role="textbox"]',
	'textarea[placeholder*="copilot" i]',
	'[role="textbox"]:not([aria-disabled="true"])',
	"textarea:not([disabled])",
];

export const M365_SEND_BUTTON_SELECTORS = [
	'button.fai-SendButton[type="submit"]',
	'button[aria-label="Send"][type="submit"]',
	'button[data-testid*="send"]',
	'button[aria-label*="send" i]:not([disabled])',
	'button[aria-label*="submit" i]:not([disabled])',
	'button[type="submit"]:not([disabled])',
];

export const M365_CHAT_CONTAINER_SELECTORS = [
	"main",
	'[role="main"]',
	'[data-testid*="chat"]',
	'[data-testid*="conversation"]',
];

export const M365_TURN_SELECTORS = [
	'[data-testid="lastChatMessage"]',
	'[data-testid*="message"]',
	'[data-testid*="turn"]',
	'[role="article"]',
	'[role="listitem"]',
	"[data-message-author-role]",
	"[data-turn]",
];

export const M365_STOP_BUTTON_SELECTORS = [
	'button[aria-label*="stop" i]',
	'button[aria-label*="cancel" i]',
	'button[data-testid*="stop"]',
];

export const M365_LOADING_SELECTORS = [
	'[role="progressbar"]',
	'[aria-busy="true"]',
	'[aria-live="polite"]:not(:empty)',
	'[aria-live="assertive"]:not(:empty)',
	'[data-loading="true"]',
];

export const M365_COPY_BUTTON_SELECTORS = [
	'button[aria-label*="copy" i]',
	'button[title*="copy" i]',
	'button[data-testid*="copy"]',
];

export const M365_SUGGESTION_SELECTORS = [
	'[role="toolbar"][aria-label*="suggestion" i]',
	'button[data-testid="chat-suggestion"]',
	'.fai-SuggestionList',
];
