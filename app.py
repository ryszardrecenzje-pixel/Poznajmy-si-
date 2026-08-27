import streamlit as st

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

# Podsumowanie
if st.button("Zapisz / Podsumuj moje odpowiedzi"):
    st.success(f"Dziękujemy, {osoba}! Twoje odpowiedzi zostały zarejestrowane.")
    
    with st.expander("Zobacz podsumowanie swoich wyborów"):
        st.write(**Poziom energii:**, energia)
        st.write(**Nastrój:**, nastroj_slowo)
        st.write(**Komfort/Emocje:**, komfort if komfort else "Brak uwag")
        st.write(**Główna potrzeba:**, potrzeba)
        st.write(**Ochota na intymność:**, seks_ochota)
        st.write(**Wybrane aktywności:**, zaznaczone_ochoty)
        if wlasne_pytanie:
            st.write(**Twoje specjalne pytanie:**, wlasne_pytanie)
