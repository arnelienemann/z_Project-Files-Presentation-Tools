import streamlit as st
from langdetect import detect
from googletrans import Translator
from wordcloud import WordCloud

st.set_page_config(page_title="Text to Word Cloud", page_icon=":guardsman:", layout="wide")

def create_wordcloud(text):
    translator = Translator()
    translated_lines = []
    for line in text.splitlines():
        lang = detect(line)

        if lang != 'en':
            translated_lines.append(translator.translate(line, dest='en').text)
        else:
            translated_lines.append(line)

    text = " ".join(translated_lines)
    wordcloud = WordCloud().generate(text)
    st.image(wordcloud, caption='Wordcloud', use_column_width=True)

text = st.text_area("Enter your text:")
if text:
    create_wordcloud(text)