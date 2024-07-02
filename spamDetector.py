import pickle
import streamlit as st
import pythoncom
from win32com.client import Dispatch

def speak(text):
    speak = Dispatch("SAPI.SpVoice")
    speak.Speak(text)

# Load model and vectorizer
model = pickle.load(open("C:\\Users\\Acer\\Email Detection\\spam.pkl", "rb"))
cv = pickle.load(open("C:\\Users\\Acer\\Email Detection\\vectorizer.pkl", "rb"))

def main():
    pythoncom.CoInitialize()
    result = None  # Initialize result to None
    try:
        st.title("Email Spam Detection Model")
        st.subheader("Check if the E-mail is Spam or Legit...")
        msg = st.text_input("Enter any E-mail")
        if st.button("Predict"):
            data = [msg]
            vector = cv.transform(data).toarray()
            prediction = model.predict(vector)
            result = prediction[0]
            if result == 1:
                st.error("This is a Spam mail")
                speak("This is a Spam mail")
            else:
                st.success("This is a Legit mail")
                speak("This is a Legit mail")

        st.markdown("---")
        st.markdown("**Contact Information**")
        st.markdown("**Email:** aitech2508@gmail.com")
        st.markdown("**Phone:** 8872083040, 6283074753")
        st.markdown("**Organization:** AI Tech")
    
    finally:
        pythoncom.CoUninitialize()

if __name__ == "__main__":
    main()