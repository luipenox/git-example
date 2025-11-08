import streamlit as st
import pandas as pd

st.title('Moje webová prezentace')

st.info('Velmi důležitá informace')

# Generování ukázkových dat pro bar chart
data_all = pd.DataFrame({
    'Měsíc': ['Leden', 'Únor', 'Březen', 'Duben', 'Květen', 'Červen'],
    'Prodeje': [120, 150, 180, 140, 200, 170],
    'Náklady': [80, 90, 100, 85, 110, 95]
})

st.bar_chart(data_all.set_index('Měsíc'))