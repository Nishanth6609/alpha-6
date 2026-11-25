# Use A.L.P.H.A - 6 in your browser

This page is a static website. It does not run a shared AI server. Chat requests go from your browser to Ollama on your own computer, and conversation history stays in this browser's local storage.

## Windows setup

1. Install Ollama from [ollama.com/download](https://ollama.com/download).
2. Open PowerShell and run: ollama pull llama3.2
3. Allow the A.L.P.H.A Pages origin to connect to Ollama. In PowerShell, run:

       setx OLLAMA_ORIGINS "https://nishanth6609.github.io"

4. Fully exit Ollama from the system tray, then start it again. The setting applies to newly started Ollama processes.
5. Reopen the A.L.P.H.A page and choose **Check connection**.

If the site uses a different domain, replace the origin above with the exact website origin. Do not use a wildcard; keep this limited to the trusted A.L.P.H.A site.

## Privacy and limits

- The web page sends chat messages to Ollama running on your own laptop. There is no A.L.P.H.A chat server.
- Recent messages are saved in this browser's local storage. Use **New chat** to clear them.
- Do not enter passwords, sign-in codes, payment details, or private keys.
- This browser version answers questions. It cannot control your laptop, read local files, or use the desktop app's optional voice and device features.
- A downloaded model uses your computer's memory and storage. Performance depends on your computer.

See the [main project page](https://github.com/Nishanth6609/alpha-6) for the Windows desktop version.
