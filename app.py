import streamlit as st
from google import genai
from google.genai import types

st.set_page_config(page_title="AssistU", page_icon="🎓")

client = genai.Client(api_key=st.secrets["GEMINI_API_KEY"])
st.title("AssistU")
st.caption("Your AI study assistant")

if "history" not in st.session_state:
    st.session_state.history = []

for role, text in st.session_state.history:
    st.chat_message(role).write(text)

if prompt := st.chat_input("Ask me anything"):
    st.chat_message("user").write(prompt)
    st.session_state.history.append(("user", prompt))

    convo = "\n".join(f"{r}: {t}" for r, t in st.session_state.history)
    reply = client.models.generate_content(
        model="gemini-3.8-flash",
        contents=convo,
        config=types.GenerateContentConfig(
            system_instruction="You are AssistU, a friendly assistant for students. Explain simply with examples."
        ),
    ).text

    st.chat_message("assistant").write(reply)
    st.session_state.history.append(("assistant", reply))