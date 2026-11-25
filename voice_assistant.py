from voice import VoiceEngine
from brain import AlphaBrain

import threading


class AlphaAssistant:

    def __init__(self):

        print("Starting A.L.P.H.A - 6 AI...")

        self.voice = VoiceEngine()
        self.brain = AlphaBrain()

        self.running = True

        print("A.L.P.H.A - 6 is ready!")

        self.voice.speak(
            "Hello! A.L.P.H.A - 6 is online. How can I help you?"
        )

    def process_voice(self):

        while self.running:

            question = self.voice.listen()

            if not question:
                continue

            if question.lower().strip() in [
                "exit", "shutdown alpha", "stop alpha"
            ]:
                self.running = False
                self.voice.speak("Goodbye!")
                break

            print("\nA.L.P.H.A - 6 is thinking...")

            answer = self.brain.ask(question)

            print("\nA.L.P.H.A - 6:", answer)

            self.voice.speak(answer)

    def process_text(self):

        while self.running:

            question = input("\nType your question: ")

            if question.lower().strip() == "exit":
                self.running = False
                break

            answer = self.brain.ask(question)

            print("\nA.L.P.H.A - 6:", answer)

            self.voice.speak(answer)

    def start(self):

        try:
            self.process_voice()

        except KeyboardInterrupt:
            self.running = False
            print("\nA.L.P.H.A - 6 stopped.")

        finally:
            self.running = False


if __name__ == "__main__":

    alpha = AlphaAssistant()

    alpha.start()