import sounddevice as sd
import speech_recognition as sr
import pyttsx3
import time

# Initialize A.L.P.H.A - 6 voice
engine = pyttsx3.init()
engine.setProperty("rate", 180)
engine.setProperty("volume", 1.0)

# Find an available female-sounding voice
voices = engine.getProperty("voices")

for voice in voices:
    name = (voice.name + " " + voice.id).lower()
    if any(word in name for word in ["zira", "aria", "jenny", "female"]):
        engine.setProperty("voice", voice.id)
        print("Selected voice:", voice.name)
        break


def speak(message):
    print("A.L.P.H.A - 6:", message)
    engine.say(message)
    engine.runAndWait()


def listen():
    recognizer = sr.Recognizer()
    sample_rate = 16000

    print("Listening for 5 seconds...")

    try:
        audio = sd.rec(
            int(5 * sample_rate),
            samplerate=sample_rate,
            channels=1,
            dtype="int16"
        )
        sd.wait()

        recorded = sr.AudioData(
            audio.tobytes(),
            sample_rate,
            2
        )

        print("Recognizing speech...")

        command = recognizer.recognize_google(recorded)

        print("YOU:", command)
        return command

    except sr.UnknownValueError:
        print("Sorry, I couldn't understand.")
    except sr.RequestError as error:
        print("Internet recognition error:", error)
    except Exception as error:
        print("Microphone error:", error)

    return ""


if __name__ == "__main__":
    speak("Hello! I am Alpha. My voice system is ready.")

    while True:
        command = listen()

        if command:
            if command.lower() in ["stop", "exit", "goodbye"]:
                speak("Goodbye!")
                break

            speak("I heard you say " + command)

        time.sleep(0.3)