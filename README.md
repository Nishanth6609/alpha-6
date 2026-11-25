# A.L.P.H.A - 6

A.L.P.H.A - 6 is a Windows laptop assistant with a dark dashboard, chat, optional voice input, local memory, and a small set of confirmed desktop actions.

## Browser versions

- [Local Ollama version](https://nishanth6609.github.io/alpha-6/) uses Ollama installed on the visitor's laptop. See [the web setup guide](docs/SETUP.md).
- [Cloud version](https://nishanth6609.github.io/alpha-6/cloud/) supports multiple API key profiles for OpenAI-compatible providers, including custom endpoints and Gemini's compatible endpoint. Keys stay in the browser tab's session storage and are sent directly to the selected provider. See [the cloud setup guide](docs/cloud/SETUP.md).

Both browser versions keep chat history in the visitor's browser and cannot control the computer, read local files, or use the desktop app's optional voice and device features. Cloud providers set their own API usage fees and limits.

## Get started

1. Download the latest `ALPHA-6-Windows.zip` from the [Releases](../../releases) page and extract it. If no release is available yet, use **Code → Download ZIP** to get the source.
2. Install [Ollama for Windows](https://ollama.com/download) and run `ollama pull llama3.2` in PowerShell.
3. For the source version, install Python 3.11 or newer, open PowerShell in the extracted folder, then run:

   ```powershell
   py -3.11 -m venv .venv
   .\.venv\Scripts\Activate.ps1
   python -m pip install -r requirements.txt
   python gui_app.py
   ```

The Windows release includes the assistant app but still needs Ollama and the `llama3.2` model installed on the same laptop. The dashboard source is `gui_app.py`.

## Optional features

- Voice input needs a microphone and internet access; its current speech recognition service sends audio to Google for transcription.
- Offline wake-word listening additionally needs the Vosk English model folder `vosk-model-small-en-us-0.15` beside `gui_app.py`.
- Gesture controls need a webcam and the optional packages in `requirements-gesture.txt`.
- Voice cloning is optional and asks the user to review the Coqui Public Model License before downloading its model.

## Privacy

Chat history and settings are stored locally on each laptop. Do not put passwords or private codes in chat. Do not upload local databases, voice recordings, or downloaded model files to GitHub. PC actions that need extra permission show a confirmation before they run.

## Build a Windows release

Pushing a version tag such as `v1.0.0` runs the GitHub Actions workflow and attaches `ALPHA-6-Windows.zip` to a GitHub Release. A manual workflow run creates a downloadable build artifact.

## Project status

The Windows desktop dashboard starts with `gui_app.py`. The static browser interfaces are in `docs/` and are deployed with GitHub Actions when Pages is enabled for the repository.
