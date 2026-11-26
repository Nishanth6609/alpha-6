# Connect A.L.P.H.A Cloud

ALPHA Cloud uses OpenRouter's `openrouter/free` model router. There is no model ID or provider URL to configure.

1. Open [ALPHA Cloud](https://nishanth6609.github.io/alpha-6/cloud/).
2. Open **Settings**.
3. Create or copy your personal key from [OpenRouter API Keys](https://openrouter.ai/settings/keys).
4. Enter it in **OpenRouter API Key**, then choose **Connect Key**.
5. Return to chat and send a message. Use **Forget Key** to remove it from the current browser tab.

ALPHA keeps the key in the browser tab's `sessionStorage` and sends it directly to OpenRouter. It is not saved in ALPHA's GitHub files or sent to an ALPHA server. Do not paste your key into chat or use a shared browser profile.

OpenRouter's free router selects an available free model automatically. Free models may be busy or rate-limited; OpenRouter's current free-tier limits and availability can change. A `503` means the selected model or service is temporarily unavailable; wait a few minutes and try again. A `429` indicates a usage or rate limit.

Prompts and recent conversation context are sent to OpenRouter and the model provider selected for a reply. Do not send passwords, private codes, payment details, or other sensitive information. Chat history stays in the browser. This web version cannot control the computer or read local files.

The original Ollama browser page remains at the site root; its setup guide is [here](../SETUP.md).
