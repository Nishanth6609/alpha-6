from voice import VoiceEngine

alpha = VoiceEngine()

alpha.speak("Hello! I am Alpha. My voice system is connected.")

command = alpha.listen()

if command:
    alpha.speak("I heard you say " + command)