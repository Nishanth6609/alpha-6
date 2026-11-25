# Use A.L.P.H.A - 6 Cloud

The cloud browser version uses OpenRouter's free model router. It does not need Ollama. Each person uses their own OpenRouter API key; do not put a shared key in GitHub source code.

## Connect

1. Create an OpenRouter account and an API key at [openrouter.ai/settings/keys](https://openrouter.ai/settings/keys).
2. Open [A.L.P.H.A - 6 Cloud](https://nishanth6609.github.io/alpha-6/cloud/).
3. Open **Settings**, paste your key, and select **Connect key**.
4. Start a chat. Use **Forget key** to remove the key from this browser tab.

ALPHA keeps the key in this tab's session storage. It is sent to OpenRouter for key validation and chat requests; it is not saved in the repository or sent to an ALPHA server. Closing the tab clears the saved key.

## Free use and privacy

- ALPHA requests `openrouter/free`, which routes to an available free model. Models and availability can change.
- Free use is rate-limited, not unlimited. OpenRouter currently lists a limit of 50 requests per day for free accounts; the model provider may apply additional limits.
- Chat messages and recent context are sent to OpenRouter and the model provider selected for each reply. Do not send passwords, private codes, payment details, or other sensitive information.
- Recent chat history is saved in this browser's local storage. Use **New chat** to clear it.
- The cloud page cannot control the computer, read local files, or use desktop-only device features.

The original Ollama browser page remains at the site root and its setup guide is [here](../SETUP.md).
