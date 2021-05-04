from win32com.client import Dispatch

def speak(str):
    speak=Dispatch(("SAPI.spvoice"))
    speak.Speak(str)
    print("Happy birthday to you")

if __name__=="__main__":
    with open("rough.txt") as f:
        for items in f.readlines():
            speak(items)