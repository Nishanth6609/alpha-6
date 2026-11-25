# Use A.L.P.H.A - 6 Cloud with your AI provider

ALPHA Cloud can save up to 20 separate provider profiles in the current browser tab. Each profile can use a different API key, endpoint, and model. It calls the provider directly from your browser; there is no ALPHA server.

## Add a provider

1. Open [A.L.P.H.A - 6 Cloud](https://nishanth6609.github.io/alpha-6/cloud/) and choose **Settings**.
2. Choose a provider template, or select **Custom OpenAI-compatible API**.
3. Enter a profile name, the provider's API base URL, model ID, and your own API key.
4. Choose **Check Provider** to try the provider's `/models` endpoint and load model IDs when supported.
5. Choose **Save & Select**, then start a chat.

To add another key, choose **New Profile** and save another profile. Select a saved profile to switch providers. **Remove Profile** deletes that profile and its key from this browser tab.

## Supported API format

ALPHA sends streaming requests using the OpenAI Chat Completions format at `<base URL>/chat/completions` and checks `<base URL>/models`. This works with providers that offer this API format, including OpenAI, Groq, OpenRouter, and Google's Gemini OpenAI-compatible endpoint. For Gemini, use `https://generativelanguage.googleapis.com/v1beta/openai` as the base URL. For other services, use the provider's documented compatible URL and model ID.

An API key alone is not enough to connect an unrelated API: the endpoint must support this format and allow requests from a browser (CORS). A service with a different API format needs a compatible endpoint or a dedicated adapter. You can use an OpenRouter key with its `openrouter/free` model router if you want a free-model option; its availability and limits are set by OpenRouter and the selected model provider.

## Key storage and privacy

- Provider profiles, including keys, are kept in this tab's `sessionStorage`. They are not written to ALPHA's files or sent to an ALPHA server. They are sent directly to the selected provider for model checks and chat replies.
- Browser session storage is not a secure vault. Use only your own key, keep it private, and remove the profile on a shared device.
- Chat history is stored in this browser's `localStorage`. Use **New chat** to clear it.
- Your provider may charge for API use or set its own limits. Free access is not guaranteed or unlimited.
- Prompts and recent conversation context go to the provider you selected. Do not include passwords, private codes, payment details, or other sensitive information.
- The browser edition cannot control the computer, read local files, or use desktop-only device features.

The original Ollama-based assistant remains at the site root. Its setup guide is [here](../SETUP.md).
