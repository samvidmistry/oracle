# M365 Copilot Browser Automation Plan

This document outlines the architecture and implementation plan for adding Microsoft 365 Copilot Chat support to Oracle's browser automation, alongside the existing ChatGPT provider.

## Goal

Enable Oracle to query M365 Copilot Chat using the same browser automation approach currently used for ChatGPT.

---

## Copilot Variants

There are **two distinct Copilot surfaces** with different DOM structures:

| Variant | URL | Target Users | Auth |
|---------|-----|--------------|------|
| **Consumer Copilot** | `copilot.microsoft.com` | Personal Microsoft accounts | MSA (Microsoft Account) |
| **M365 Copilot Chat** | `copilot.cloud.microsoft` | Enterprise/work accounts | Entra ID (Azure AD) |

The consumer version loads directly with a full React app. The enterprise version redirects through `login.microsoftonline.com` for MSAL-based authentication before loading the chat UI.

---

## Comparison: ChatGPT vs M365 Copilot

| Component | ChatGPT | M365 Copilot (Consumer) |
|-----------|---------|-------------------------|
| **URL** | `chatgpt.com` | `copilot.microsoft.com` |
| **Auth Check** | `/backend-api/me` probe | DOM-based (login CTAs, URL detection) |
| **Cookie Domains** | `chatgpt.com`, `chat.openai.com` | `copilot.microsoft.com`, `login.live.com` |
| **Input Selector** | `#prompt-textarea` | `[data-testid="composer-input"]` |
| **Input ID** | `prompt-textarea` | `userInput` |
| **Input Role** | varies | `role="textbox"` |
| **Send Button** | `data-testid="send-button"` | TBD (may use Enter key) |
| **File Input** | `input[type="file"]` | `[data-testid="composer-file-input"]` |
| **Supported Files** | Images, PDFs, code | `.png,.jpg,.jpeg,.pdf,.docx,.xlsx,.pptx,.json,.csv,.md,.txt` |
| **New Chat** | varies | `[data-testid="sidebar-new-conversation-button"]` |
| **Completion** | Stop button + copy button | `aria-live` regions + text stabilization |
| **Markdown Copy** | Copy button → clipboard | Likely unavailable |
| **Model Selection** | Required (GPT-5.2 Pro) | Not needed |
| **Accessibility** | `data-message-author-role` | `aria-live="polite"`, `aria-live="assertive"` |

---

## Current Architecture

The ChatGPT browser automation lives in `src/browser/` and follows this flow:

```
Launch Chrome → Sync Cookies → Navigate → Ensure Logged In → Select Model → Submit Prompt → Wait for Response → Capture Markdown
```

### Key Files

| File | Purpose |
|------|---------|
| `src/browser/index.ts` | Main orchestrator (`runBrowserMode`) |
| `src/browser/constants.ts` | ChatGPT-specific DOM selectors |
| `src/browser/actions/navigation.ts` | Navigation, login detection, prompt-ready checks |
| `src/browser/actions/assistantResponse.ts` | Response polling, completion detection, markdown capture |
| `src/browser/actions/promptComposer.ts` | Prompt submission |
| `src/browser/cookies.ts` | Cookie synchronization |
| `src/browser/chromeLifecycle.ts` | Chrome launch/connect |

### ChatGPT-Specific Elements

- **URL**: `https://chatgpt.com/`
- **Auth check**: `/backend-api/me` endpoint probe
- **Input selector**: `#prompt-textarea`, `data-id="prompt-textarea"`
- **Send button**: `data-testid="send-button"`
- **Assistant turns**: `data-message-author-role="assistant"`
- **Completion detection**: Stop button disappears + copy button appears
- **Markdown capture**: Click copy button → intercept clipboard

---

## Proposed Multi-Provider Architecture

### Provider Interface

Create a `BrowserProvider` interface that abstracts all provider-specific behavior:

```ts
// src/browser/types.ts (extend existing file)

export interface BrowserProvider {
  id: 'chatgpt' | 'm365-copilot';
  
  // URLs and cookie domains
  defaultUrl: string;
  fallbackUrl?: string;
  cookieUrls: string[];
  
  // Navigation + auth
  ensureNotBlocked(Runtime: ChromeClient['Runtime'], headless: boolean, logger: BrowserLogger): Promise<void>;
  ensureLoggedIn(Runtime: ChromeClient['Runtime'], logger: BrowserLogger, opts?: { appliedCookies?: number }): Promise<void>;
  ensurePromptReady(Runtime: ChromeClient['Runtime'], timeoutMs: number, logger: BrowserLogger): Promise<void>;
  
  // Prompt submission
  submitPrompt(Runtime: ChromeClient['Runtime'], Input: ChromeClient['Input'], promptText: string, logger: BrowserLogger): Promise<void>;
  clearPromptComposer?(Runtime: ChromeClient['Runtime'], logger: BrowserLogger): Promise<void>;
  
  // Attachments (optional capability)
  supportsAttachments?: boolean;
  uploadAttachmentFile?(Runtime: ChromeClient['Runtime'], filePath: string, logger: BrowserLogger): Promise<void>;
  waitForAttachmentCompletion?(Runtime: ChromeClient['Runtime'], logger: BrowserLogger): Promise<void>;
  
  // Response capture
  readAssistantSnapshot(Runtime: ChromeClient['Runtime'], minTurnIndex?: number): Promise<AssistantSnapshot | null>;
  waitForAssistantResponse(Runtime: ChromeClient['Runtime'], timeoutMs: number, logger: BrowserLogger, minTurnIndex?: number): Promise<AssistantResponse>;
  captureAssistantMarkdown?(Runtime: ChromeClient['Runtime'], meta: ResponseMeta, logger: BrowserLogger): Promise<string | null>;
  
  // Model selection (optional capability)
  supportsModelSelection?: boolean;
  ensureModelSelection?(Runtime: ChromeClient['Runtime'], logger: BrowserLogger, targetModel: string, strategy: string): Promise<void>;
}
```

### Directory Structure

```
src/browser/
├── providers/
│   ├── index.ts                 # Provider registry + resolver
│   ├── interface.ts             # BrowserProvider interface
│   ├── chatgpt/
│   │   ├── index.ts             # ChatGPT provider implementation
│   │   ├── selectors.ts         # ChatGPT DOM selectors (from constants.ts)
│   │   ├── navigation.ts        # ChatGPT auth/nav logic
│   │   └── response.ts          # ChatGPT response capture
│   └── m365Copilot/
│       ├── index.ts             # M365 Copilot provider implementation
│       ├── selectors.ts         # M365 Copilot DOM selectors
│       ├── navigation.ts        # M365 Copilot auth/nav logic
│       └── response.ts          # M365 Copilot response capture
├── index.ts                     # Orchestrator (modified to use provider)
├── constants.ts                 # Shared constants only
└── ... (other existing files)
```

### Orchestrator Changes

Modify `src/browser/index.ts` to:

1. Accept `--browser-provider chatgpt|m365-copilot` option
2. Resolve provider from config/env
3. Call provider methods instead of direct imports

```ts
// In runBrowserMode():
const provider = resolveProvider(config.provider ?? 'chatgpt');

// Replace direct calls:
// Before: await ensureLoggedIn(Runtime, logger, opts);
// After:
await provider.ensureLoggedIn(Runtime, logger, opts);

// Before: await ensureModelSelection(...);
// After:
if (provider.supportsModelSelection) {
  await provider.ensureModelSelection(Runtime, logger, targetModel, strategy);
}
```

---

## M365 Copilot Provider Specification

### URLs and Domains

```ts
// src/browser/providers/m365Copilot/selectors.ts

export const COPILOT_URL = 'https://copilot.cloud.microsoft/';
export const COOKIE_URLS = [
  'https://copilot.cloud.microsoft',
  'https://copilot.microsoft.com',
  'https://login.microsoftonline.com',
  'https://login.live.com',
];
```

### Authentication Strategy

**Primary approach: Manual login with persistent profile**

M365 Copilot uses Entra ID (Azure AD) with MSAL, which often involves:
- Multi-factor authentication
- Conditional Access policies
- Device compliance checks
- Short-lived tokens with refresh

Cookie sync from a consumer Chrome profile is unreliable for enterprise tenants.

**Recommended flow:**
1. Default to `--manual-login` mode for M365 Copilot
2. Reuse persistent profile at `~/.oracle/m365-browser-profile`
3. Cookie sync remains available but not guaranteed

**Login detection (no API probe available):**

```ts
// src/browser/providers/m365Copilot/navigation.ts

function isAuthPage(url: string): boolean {
  const authHostnames = [
    'login.microsoftonline.com',
    'login.live.com',
    'login.windows.net',
    'account.live.com',
  ];
  const authPaths = ['/signin', '/oauth2', '/authorize', '/consent', '/mfa'];
  
  try {
    const parsed = new URL(url);
    if (authHostnames.some(h => parsed.hostname.includes(h))) return true;
    if (authPaths.some(p => parsed.pathname.toLowerCase().includes(p))) return true;
    return false;
  } catch {
    return false;
  }
}

async function hasLoginCta(Runtime): Promise<boolean> {
  // Detect "Sign in", "Pick an account", "Continue" buttons
  const expression = `(() => {
    const labels = ['sign in', 'signin', 'log in', 'login', 'pick an account', 'continue'];
    const buttons = document.querySelectorAll('button, a, [role="button"]');
    for (const btn of buttons) {
      const text = (btn.textContent || '').toLowerCase().trim();
      if (labels.some(l => text.includes(l))) return true;
    }
    return false;
  })()`;
  const { result } = await Runtime.evaluate({ expression, returnByValue: true });
  return Boolean(result?.value);
}

export async function ensureLoggedIn(Runtime, logger, opts): Promise<void> {
  const url = await getCurrentUrl(Runtime);
  if (isAuthPage(url)) {
    throw new Error('M365 auth page detected. Run with --manual-login and sign in.');
  }
  
  const hasCta = await hasLoginCta(Runtime);
  if (hasCta) {
    throw new Error('M365 login required. Run with --manual-login and sign in.');
  }
  
  logger('M365 Copilot login check passed');
}
```

**Handle common interstitials:**

```ts
// Dismiss "Stay signed in?" and similar prompts
async function dismissBlockingUi(Runtime, logger): Promise<boolean> {
  const expression = `(() => {
    const labels = ['yes', 'no', 'stay signed in', 'continue', 'next', 'accept', 'ok'];
    const buttons = document.querySelectorAll('button, [role="button"]');
    for (const btn of buttons) {
      const text = (btn.textContent || '').toLowerCase().trim();
      if (labels.includes(text)) {
        btn.click();
        return { dismissed: true, label: text };
      }
    }
    return { dismissed: false };
  })()`;
  // ...
}
```

### DOM Selectors (Based on Real DOM Analysis)

Analysis of `copilot.microsoft.com` (consumer Copilot) reveals these actual selectors:

#### Discovered `data-testid` Attributes
```
data-testid="composer"              - Main composer container
data-testid="composer-content"      - Composer content area
data-testid="composer-input"        - The actual input textarea
data-testid="composer-file-input"   - File upload input
data-testid="composer-create-button" - Attach/create button
data-testid="sidebar-container"     - Sidebar
data-testid="sidebar-toggle-button" - Toggle sidebar
data-testid="sidebar-new-conversation-button" - New chat button
data-testid="sidebar-discover-button" - Discover link
data-testid="sidebar-action-menu-button" - Actions menu
data-testid="sidebar-copilot-brand-button" - Home/brand button
```

#### Discovered `aria-label` Attributes
```
aria-label="Message Copilot"        - Input placeholder/label
aria-label="Open sidebar"           - Sidebar toggle
aria-label="Start new chat"         - New conversation
aria-label="Open actions menu"      - Actions dropdown
aria-label="Attach files, connect apps, or make something with Copilot." - Create button
aria-label="Discover"               - Discover nav
aria-label="Library"                - Library nav
aria-label="Labs"                   - Labs nav
aria-label="Imagine (New)"          - Image generation
```

#### Discovered `role` Attributes
```
role="textbox"     - The input area
role="navigation"  - Sidebar container
role="link"        - Brand/home button
```

#### Actual HTML Structure (Key Elements)

**Input textarea:**
```html
<textarea 
  class="..." 
  placeholder="Message Copilot" 
  id="userInput" 
  role="textbox" 
  aria-autocomplete="both" 
  spellCheck="false" 
  enterKeyHint="enter" 
  autofocus="" 
  data-testid="composer-input">
</textarea>
```

**File input (hidden):**
```html
<input 
  type="file" 
  accept=".png,.jpg,.jpeg,.pdf,.docx,.xlsx,.pptx,.json,.csv,.md,.txt" 
  class="hidden" 
  aria-hidden="true" 
  multiple="" 
  data-testid="composer-file-input">
```

**Accessibility regions:**
```html
<div class="sr-only">
  <div aria-live="polite" aria-atomic="true"></div>
  <div aria-live="assertive" aria-atomic="true"></div>
</div>
```

#### Recommended Selectors for M365 Copilot Provider

```ts
// src/browser/providers/m365Copilot/selectors.ts

export const COPILOT_URL = 'https://copilot.microsoft.com/';
export const COPILOT_CLOUD_URL = 'https://copilot.cloud.microsoft/';

// Cookie domains (both consumer and enterprise)
export const COOKIE_URLS = [
  'https://copilot.microsoft.com',
  'https://copilot.cloud.microsoft',
  'https://login.microsoftonline.com',
  'https://login.live.com',
];

// Prompt input - using real discovered selectors
export const INPUT_SELECTORS = [
  '[data-testid="composer-input"]',           // Primary: exact testid
  'textarea#userInput',                        // Fallback: id
  'textarea[role="textbox"]',                  // Fallback: role
  'textarea[placeholder*="Message Copilot" i]', // Fallback: placeholder
  '[role="textbox"]:not([aria-disabled="true"])',
  'textarea:not([disabled])',
];

// Send button - need to discover via live DOM inspection
// The composer may use Enter key or a send button
export const SEND_BUTTON_SELECTORS = [
  'button[data-testid*="send"]',
  'button[aria-label*="send" i]:not([disabled])',
  'button[aria-label*="submit" i]:not([disabled])',
  'button[type="submit"]:not([disabled])',
];

// File input for attachments
export const FILE_INPUT_SELECTOR = '[data-testid="composer-file-input"]';

// Attach/create button
export const CREATE_BUTTON_SELECTOR = '[data-testid="composer-create-button"]';

// New conversation
export const NEW_CHAT_SELECTORS = [
  '[data-testid="sidebar-new-conversation-button"]',
  'button[aria-label="Start new chat"]',
];

// Conversation container - for scoping turn detection
export const CHAT_CONTAINER_SELECTORS = [
  'main',
  '[role="main"]',
  '[data-testid*="chat"]',
  '[data-testid*="conversation"]',
];

// Conversation turns - will need live DOM inspection to refine
// Copilot likely uses different structure than ChatGPT
export const TURN_SELECTORS = [
  '[data-testid*="message"]',
  '[data-testid*="turn"]',
  '[role="article"]',
  '[role="listitem"]',
];

// Stop/cancel button (for streaming detection)
export const STOP_BUTTON_SELECTORS = [
  'button[aria-label*="stop" i]',
  'button[aria-label*="cancel" i]',
  'button[data-testid*="stop"]',
];

// Loading/streaming indicators
export const LOADING_SELECTORS = [
  '[role="progressbar"]',
  '[aria-busy="true"]',
  '[aria-live="polite"]:not(:empty)',  // Accessibility live region with content
  '[aria-live="assertive"]:not(:empty)',
  '[data-loading="true"]',
];

// Sidebar toggle (useful for ensuring chat area is visible)
export const SIDEBAR_TOGGLE_SELECTOR = '[data-testid="sidebar-toggle-button"]';
```

#### Notes on M365 Enterprise (copilot.cloud.microsoft)

The enterprise version redirects to `login.microsoftonline.com` for authentication. Key observations:

1. **Auth flow uses MSAL**: The redirect includes OAuth2 parameters targeting `login.microsoftonline.com`
2. **Different from consumer**: The authenticated app loads a different UI (Office Hub based)
3. **Cookie domains**: `*.microsoft.com`, `login.microsoftonline.com`, `login.live.com`

For enterprise, after login, the chat UI may differ from consumer. Live DOM inspection post-authentication will be needed to finalize selectors.

### Response Capture

**Completion detection (no copy button reliance):**

```ts
// src/browser/providers/m365Copilot/response.ts

export async function waitForAssistantResponse(
  Runtime,
  timeoutMs: number,
  logger,
  minTurnIndex?: number
): Promise<AssistantResponse> {
  const deadline = Date.now() + timeoutMs;
  let lastText = '';
  let stableCount = 0;
  const requiredStableChecks = 5;
  const pollInterval = 500;
  
  while (Date.now() < deadline) {
    // Check if still streaming
    const isStreaming = await isStillStreaming(Runtime);
    
    // Get current response text
    const snapshot = await readAssistantSnapshot(Runtime, minTurnIndex);
    const currentText = snapshot?.text ?? '';
    
    if (!isStreaming && currentText && currentText === lastText) {
      stableCount++;
      if (stableCount >= requiredStableChecks) {
        logger('M365 Copilot response complete (stabilized)');
        return { text: currentText, html: snapshot?.html, meta: {} };
      }
    } else {
      stableCount = 0;
      lastText = currentText;
    }
    
    await delay(pollInterval);
  }
  
  throw new Error('M365 Copilot response timeout');
}

async function isStillStreaming(Runtime): Promise<boolean> {
  const expression = `(() => {
    // Check for stop button
    const stopSelectors = ${JSON.stringify(STOP_BUTTON_SELECTORS)};
    for (const sel of stopSelectors) {
      if (document.querySelector(sel)) return true;
    }
    
    // Check for loading indicators in chat area
    const loadingSelectors = ${JSON.stringify(LOADING_SELECTORS)};
    const chatContainer = document.querySelector('[role="main"]') || document.body;
    for (const sel of loadingSelectors) {
      if (chatContainer.querySelector(sel)) return true;
    }
    
    return false;
  })()`;
  const { result } = await Runtime.evaluate({ expression, returnByValue: true });
  return Boolean(result?.value);
}
```

**Markdown capture (fallback to plain text):**

```ts
export async function captureAssistantMarkdown(Runtime, meta, logger): Promise<string | null> {
  // M365 Copilot may not have a copy-as-markdown button
  // Return null to fall back to innerText
  logger('M365 Copilot: markdown copy not available, using plain text');
  return null;
}
```

### Prompt Submission

```ts
// src/browser/providers/m365Copilot/promptComposer.ts

export async function submitPrompt(Runtime, Input, promptText, logger): Promise<void> {
  // Find and focus input
  const inputExpression = buildFindInputExpression();
  const { result } = await Runtime.evaluate({ expression: inputExpression, returnByValue: true });
  
  if (!result?.value?.found) {
    throw new Error('M365 Copilot prompt input not found');
  }
  
  // Type text (use existing typing utilities)
  await typeText(Input, promptText);
  
  // Find and click send button
  const sendExpression = buildFindSendButtonExpression();
  await Runtime.evaluate({ expression: sendExpression });
  
  logger('M365 Copilot prompt submitted');
}
```

---

## CLI and Configuration

### New Options

```ts
// src/cli/browserConfig.ts

interface BrowserConfig {
  // Existing options...
  
  // New: provider selection
  provider?: 'chatgpt' | 'm365-copilot';
}
```

### CLI Flags

```
--browser-provider <provider>   Browser automation provider (chatgpt, m365-copilot)
```

### Environment Variables

```
ORACLE_BROWSER_PROVIDER=m365-copilot
```

---

## Implementation Phases

### Phase 1: Provider Abstraction (1-2 hours)

1. Create `BrowserProvider` interface in `src/browser/providers/interface.ts`
2. Create provider registry in `src/browser/providers/index.ts`
3. Move ChatGPT-specific code into `src/browser/providers/chatgpt/`
4. Update `index.ts` to use provider interface
5. Verify ChatGPT still works

### Phase 2: M365 Copilot Provider Skeleton (2-3 hours)

1. Create `src/browser/providers/m365Copilot/` directory
2. Implement basic selectors (will need DOM inspection to refine)
3. Implement auth detection (URL-based + DOM CTAs)
4. Implement prompt submission
5. Implement basic response capture

### Phase 3: Response Capture Refinement (2-4 hours)

1. Test against live M365 Copilot
2. Refine selectors based on actual DOM
3. Implement robust completion detection
4. Handle edge cases (streaming, errors, timeouts)

### Phase 4: Testing and Polish (2-3 hours)

1. Add CLI flag and documentation
2. Test with various enterprise configurations
3. Add error messages with actionable guidance
4. Update `docs/browser-mode.md`

**Total estimated effort: 1-2 days**

---

## Risks and Mitigations

| Risk | Impact | Mitigation |
|------|--------|------------|
| Conditional Access / MFA blocks automation | High | Default to `--manual-login`; clear error messages |
| DOM changes frequently | Medium | Use ARIA/role selectors; add `logDomFailure` hooks |
| Iframe embedding | Medium | Detect and error; frame targeting is a bigger change |
| No markdown copy available | Low | Fall back to `innerText`; document limitation |
| Multiple Copilot surfaces (Teams, Outlook) | Medium | Start with `copilot.cloud.microsoft` only |
| Enterprise compliance concerns | Medium | Document that this is UI automation; recommend sanctioned APIs where available |

---

## Testing Strategy

### Manual Testing Checklist

1. [ ] Navigate to M365 Copilot (fresh profile)
2. [ ] Navigate to M365 Copilot (manual login profile)
3. [ ] Detect "not logged in" state correctly
4. [ ] Submit a simple prompt
5. [ ] Capture response text
6. [ ] Handle streaming (long response)
7. [ ] Handle "Stay signed in?" interstitial
8. [ ] Error gracefully on auth failure

### Automated Tests

- Unit tests for selector building
- Unit tests for auth URL detection
- Integration tests with mock Runtime (if feasible)

---

## Future Enhancements

- Support for file attachments (if M365 Copilot supports them)
- Support for other Copilot surfaces (Teams, Outlook)
- Cookie sync improvements for specific tenant configurations
- Frame targeting for embedded Copilot UIs
