import streamlit as st
import pandas as pd
from datetime import datetime

st.set_page_config(page_title="Quiz dla Par", page_icon="💖", layout="centered")

st.title("💖 Przedspotkaniowy Check-in dla Par")
st.write("Odpowiedzcie na poniższe pytania, aby dostosować dzisiejsze spotkanie do Waszych aktualnych potrzeb i nastrojów.")

# Wybór osoby
osoba = st.radio("Kto teraz wypełnia quiz?", ["Osoba A", "Osoba B"], horizontal=True)

st.divider()

# 1. Nastrój i Energia
st.header("1. Nastrój i Energia")
energia = st.slider(f"{osoba}: Jaki masz poziom energii na dzisiejsze spotkanie?", 1, 10, 5)
nastroj_slowo = st.selectbox(
    f"{osoba}: Określ swój obecny nastrój jednym słowem:",
    ["Zrelaksowany/a", "Podekscytowany/a", "Zmęczony/a", "Czule nastrojony/a", "Stresujący dzień", "Gotowy/a na wszystko"]
)

# 2. Emocje i Oczekiwania
st.header("2. Emocje i Granice")
komfort = st.text_area(f"{osoba}: Czy jest coś, o czym Twój partner/partnerka powinien wiedzieć o Twoim dzisiejszym stanie emocjonalnym?")
potrzeba = st.selectbox(
    f"{osoba}: Czego najbardziej dzisiaj potrzebujesz?",
    ["Wspólnego wyciszenia i rozmowy", "Czułości i przytulasów", "Dobrej zabawy i śmiechu", "Bliskości fizycznej / Intymności"]
)

# 3. Seksualność i Intymność
st.header("3. Seksualność i Intymność")
seks_ochota = st.select_slider(
    f"{osoba}: Jak dużą masz dziś ochotę na zbliżenia intymne?",
    options=["Brak ochoty / Tylko czułość", "Otwarty/a, jeśli powoli", "Umiarkowana", "Duża", "100% ogień! 🔥"]
)

# 4. Lista zachcianek (Yes/No/Maybe)
st.header("4. Menu Zachcianek na dziś")
st.write("Zaznacz to, na co masz dzisiaj ochotę:")

ochoty_opcje = [
    "Długie przytulanie na kanapie",
    "Masaż pleców / karku",
    "Zmysłowy masaż całego ciała",
    "Gorący prysznic / kąpiel we dwoje",
    "Gra wstępna z naciskiem na dotyk",
    "Pełna intymność / seks",
    "Eksperymenty / nowe rzeczy"
]

zaznaczone_ochoty = []
for opcja in ochoty_opcje:
    if st.checkbox(opcja, key=f"{osoba}_{opcja}"):
        zaznaczone_ochoty.append(opcja)

# 5. Pytania niestandardowe (własne)
st.header("5. Twoje własne pytania")
wlasne_pytanie = st.text_input(f"{osoba}: Masz jakieś specjalne pytanie, które chcesz zadać drugiej osobie?")

st.divider()

# Podsumowanie i Generowanie Excela
if st.button("Zapisz i przygotuj plik Excel"):
    st.success(f"Dziękujemy, {osoba}! Twoje odpowiedzi zostały przygotowane do pobrania.")
    
    # Przygotowanie danych do tabeli (Dataframe)
    dane = {
        "Data": [datetime.now().strftime("%Y-%m-%d %H:%M")],
        "Osoba": [osoba],
        "Poziom Energii": [energia],
        "Nastrój": [nastroj_slowo],
        "Stan emocjonalny / Komfort": [komfort if komfort else "Brak uwag"],
        "Główna potrzeba": [potrzeba],
        "Ochota na intymność": [seks_ochota],
        "Wybrane zachcianki": [", ".join(zaznaczone_ochoty)],
        "Własne pytanie": [wlasne_pytanie if wlasne_pytanie else "Brak"]
    }
    
    df = pd.DataFrame(dane)
    
    # Zapis do pliku Excel w pamięci (funkcja do pobrania)
    from io import BytesIO
    output = BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        df.to_excel(writer, index=False, sheet_name='Wyniki Quizu')
    excel_data = output.getvalue()
    
    # Przycisk do pobrania pliku Excel
    st.download_button(
        label="📥 Pobierz wyniki jako plik Excel (.xlsx)",
        data=excel_data,
        file_name=f"quiz_wyniki_{osoba.lower().replace(' ', '_')}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )

    with st.expander("Zobacz podsumowanie na ekranie"):
        st.write(f"**Poziom energii:** {energia}")
        st.write(f"**Nastrój:** {nastroj_slowo}")
        st.write(f"**Komfort/Emocje:** {komfort if komfort else 'Brak uwag'}")
        st.write(f"**Główna potrzeba:** {potrzeba}")
        st.write(f"**Ochota na intymność:** {seks_ochota}")
        st.write(f"**Wybrane aktywności:** {zaznaczone_ochoty}")
        if wlasne_pytanie:
            st.write(f"**Twoje specjalne pytanie:** {wlasne_pytanie}")
