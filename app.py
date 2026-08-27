import streamlit as st
import pandas as pd
from datetime import datetime
from io import BytesIO
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side

st.set_page_config(page_title="Intymny Check-in dla Par", page_icon="💖", layout="centered")

st.title("💖 Przedspotkaniowy Check-in dla Par")
st.write("Narzędzie do odkrywania nastrojów, potrzeb i granic – blisko siebie lub na odległość.")

# Wybór trybu relacji
tryb_relacji = st.radio(
    "Wybierz tryb:",
    ["💞 Jesteśmy razem (na jednym urządzeniu / porównanie na żywo)", "✈️ Relacja na odległość (każdy wypełnia u siebie i przesyła plik)"],
    horizontal=False
)

st.divider()

# Inicjalizacja stanu
if "imie_1" not in st.session_state:
    st.session_state.imie_1 = "Osoba A"
if "imie_2" not in st.session_state:
    st.session_state.imie_2 = "Osoba B"
if "odpowiedzi" not in st.session_state:
    st.session_state.odpowiedzi = {}

# Konfiguracja imion
col1, col2 = st.columns(2)
with col1:
    st.session_state.imie_1 = st.text_input("Imię Pierwszej Osoby:", value=st.session_state.imie_1)
with col2:
    st.session_state.imie_2 = st.text_input("Imię Drugiej Osoby:", value=st.session_state.imie_2)

st.divider()

# Wybór kto teraz wypełnia
aktualna_osoba = st.radio("Kto teraz wypełnia quiz?", [st.session_state.imie_1, st.session_state.imie_2], horizontal=True)

st.divider()

# --- FORMULARZ PYTAŃ ---
st.header("1. Nastrój, Energia i Bezpieczeństwo")
energia = st.slider(f"{aktualna_osoba}: Jaki masz poziom energii?", 1, 10, 5)
nastroj_slowo = st.selectbox(
    f"{aktualna_osoba}: Określ swój obecny nastrój jednym słowem:",
    ["Zrelaksowany/a", "Podekscytowany/a", "Zmęczony/a", "Czule nastrojony/a", "Stresujący dzień", "Gotowy/a na wszystko"]
)
bezpiecz_pytanie = st.text_input(f"{aktualna_osoba}: Co mogę dzisiaj zrobić, żebyś poczuł/a się przy mnie w 100% bezpiecznie?")

st.header("2. Emocje i Granice")
komfort = st.text_area(f"{aktualna_osoba}: Czy jest coś o Twoim stanie emocjonalnym, o czym partner powinien wiedzieć?")
granice = st.text_input(f"{aktualna_osoba}: Czego dzisiaj ZUPEŁNIE NIE CHCESZ / jakie są Twoje granice?")
potrzeba = st.selectbox(
    f"{aktualna_osoba}: Czego najbardziej dzisiaj potrzebujesz?",
    ["Wspólnego wyciszenia i rozmowy", "Czułości i przytulasów", "Dobrej zabawy i śmiechu", "Bliskości fizycznej / Intymności"]
)

st.header("3. Seksualność, Fantazje i Oczekiwania")
seks_ochota = st.select_slider(
    f"{aktualna_osoba}: Jak dużą masz dziś ochotę na zbliżenia?",
    options=["Brak ochoty / Tylko czułość", "Otwarty/a, jeśli powoli", "Umiarkowana", "Duża", "100% ogień! 🔥"]
)
pikantna_niespodzianka = st.text_input(f"{aktualna_osoba}: Na jaką małą, pikantną niespodziankę masz ochotę?")
motyw_przewodni = st.text_input(f"{aktualna_osoba}: Gdyby ten wieczór miał motyw przewodni, to jaki?")

st.header("4. Jestem gotowy/wa na...")
st.write("Zaznacz aktywności, na które czujesz się dziś w pełni gotowy/a:")
gotowosc_opcje = [
    "Długie przytulanie i bliskość emocjonalną",
    "Masaż pleców / karku",
    "Zmysłowy masaż całego ciała",
    "Gorący prysznic / kąpiel we dwoje",
    "Grę wstępną z naciskiem na dotyk",
    "Pełną intymność i zbliżenie",
    "Eksperymenty i nowe doznania"
]

zaznaczone_gotowosci = []
for opcja in gotowosc_opcje:
    if st.checkbox(opcja, key=f"{aktualna_osoba}_{opcja}"):
        zaznaczone_gotowosci.append(opcja)

st.header("5. Pytania i Docenienie")
komplement = st.text_input(f"{aktualna_osoba}: Jaki komplement chcesz usłyszeć dzisiaj wieczorem?")
wlasne_pytanie = st.text_input(f"{aktualna_osoba}: Masz specjalne pytanie lub prośbę do partnera?")

st.divider()

# Przycisk zapisu w sesji (głównie dla trybu razem)
if st.button(f"Zapisz odpowiedzi dla: {aktualna_osoba}"):
    st.session_state.odpowiedzi[aktualna_osoba] = {
        "Data": datetime.now().strftime("%Y-%m-%d %H:%M"),
        "Energia": energia,
        "Nastrój": nastroj_slowo,
        "Bezpieczeństwo": bezpiecz_pytanie if bezpiecz_pytanie else "Brak uwag",
        "Komfort/Emocje": komfort if komfort else "Brak uwag",
        "Granice": granice if granice else "Brak ograniczeń",
        "Potrzeba": potrzeba,
        "Ochota na intymność": seks_ochota,
        "Pikantna niespodzianka": pikantna_niespodzianka if pikantna_niespodzianka else "Brak",
        "Motyw przewodni": motyw_przewodni if motyw_przewodni else "Romantyczny wieczór",
        "Gotowość": zaznaczone_gotowosci,
        "Komplement": komplement if komplement else "Miłe słowo",
        "Własne pytanie": wlasne_pytanie if wlasne_pytanie else "Brak"
    }
    st.success(f"Zapisano pomyślnie odpowiedzi dla {aktualna_osoba}! 🎉")

st.divider()

# --- FUNKCJA GENEROWANIA EXCELA (DARK MODE) ---
def generuj_excel_karty(dane_karty, imie_wypelniajacego):
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Karta Spotkania"
    ws.views.sheetView[0].showGridLines = False
    
    DARK_BG = "000000"
    CARD_BG = "1A1A1A"
    HEADER_BG = "2C2C2C"
    TEXT_WHITE = "FFFFFF"
    TEXT_ACCENT = "FFD700"
    BORDER_DARK = "333333"
    
    title_font = Font(name="Segoe UI", size=14, bold=True, color="FF69B4")
    header_font = Font(name="Segoe UI", size=10, bold=True, color=TEXT_WHITE)
    header_fill = PatternFill(start_color=HEADER_BG, end_color=HEADER_BG, fill_type="solid")
    
    data_font = Font(name="Segoe UI", size=10, color=TEXT_WHITE)
    accent_data_font = Font(name="Segoe UI", size=10, bold=True, color=TEXT_ACCENT)
    data_fill = PatternFill(start_color=CARD_BG, end_color=CARD_BG, fill_type="solid")
    
    thin_border = Border(
        left=Side(style='thin', color=BORDER_DARK),
        right=Side(style='thin', color=BORDER_DARK),
        top=Side(style='thin', color=BORDER_DARK),
        bottom=Side(style='thin', color=BORDER_DARK)
    )
    
    ws.merge_cells('A1:B1')
    ws['A1'] = f"💖 KARTA SPOTKANIA: {imie_wypelniajacego.upper()}"
    ws['A1'].font = title_font
    ws['A1'].alignment = Alignment(horizontal="left", vertical="center")
    ws.row_dimensions[1].height = 40
    
    fields = [
        ("📅 Data i Czas", dane_karty["Data"]),
        ("👤 Kto odpowiada", imie_wypelniajacego),
        ("⚡ Poziom Energii", f"{dane_karty['Energia']} / 10"),
        ("✨ Nastrój", dane_karty["Nastrój"]),
        ("🛡️ Bezpieczeństwo", dane_karty["Bezpieczeństwo"]),
        ("💭 Stan emocjonalny", dane_karty["Komfort/Emocje"]),
        ("🚫 Granice", dane_karty["Granice"]),
        ("💖 Główna Potrzeba", dane_karty["Potrzeba"]),
        ("🔥 Ochota na intymność", dane_karty["Ochota na intymność"]),
        ("🎁 Pikantna niespodzianka", dane_karty["Pikantna niespodzianka"]),
        ("🌙 Motyw przewodni", dane_karty["Motyw przewodni"]),
        ("🌹 Jestem gotowy/wa na...", ", ".join(dane_karty["Gotowość"]) if dane_karty["Gotowość"] else "Czułość i bliskość"),
        ("✨ Komplement", dane_karty["Komplement"]),
        ("💬 Pytanie do partnera", dane_karty["Własne pytanie"])
    ]
    
    for idx, (label, val) in enumerate(fields, start=3):
        ws.row_dimensions[idx].height = 30
        
        cell_lbl = ws.cell(row=idx, column=1, value=label)
        cell_lbl.font = header_font
        cell_lbl.fill = header_fill
        cell_lbl.alignment = Alignment(horizontal="left", vertical="center", indent=1)
        cell_lbl.border = thin_border
        
        cell_val = ws.cell(row=idx, column=2, value=val)
        if "intymność" in label.lower() or "gotowy" in label.lower() or "granice" in label.lower():
            cell_val.font = accent_data_font
        else:
            cell_val.font = data_font
            
        cell_val.fill = data_fill
        cell_val.alignment = Alignment(horizontal="left", vertical="center", wrap_text=True, indent=1)
        cell_val.border = thin_border
        
    ws.column_dimensions['A'].width = 35
    ws.column_dimensions['B'].width = 60
        
    output = BytesIO()
    wb.save(output)
    return output.getvalue()


# --- OBSŁUGA POBIERANIA / PORÓWNANIA W ZALEŻNOŚCI OD TRYBU ---
if "Jesteśmy razem" in tryb_relacji:
    st.subheader("📥 Pobierz kartę wybranej osoby")
    osoba_do_pobrania = st.selectbox("Wybierz czyją kartę pobrać:", [st.session_state.imie_1, st.session_state.imie_2], key="pobierz_razem")
    
    if st.button("Pobierz zmysłowy plik Excel"):
        if osoba_do_pobrania in st.session_state.odpowiedzi:
            dane = st.session_state.odpowiedzi[osoba_do_pobrania]
            excel_ bytes = generuj_excel_karty(dane, osoba_do_pobrania)
            st.download_button(
                label=f"Pobierz kartę dla: {osoba_do_pobrania}",
                data=excel_bytes,
                file_name=f"karta_spotkania_{osoba_do_pobrania.lower()}_dark.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
        else:
            # Tworzymy tymczasowe dane z bieżących wyborów suwaków, jeśli nie kliknięto jeszcze "Zapisz"
            dane_tymczasowe = {
                "Data": datetime.now().strftime("%Y-%m-%d %H:%M"),
                "Energia": energia, "Nastrój": nastroj_slowo,
                "Bezpieczeństwo": bezpiecz_pytanie or "Brak uwag",
                "Komfort/Emocje": komfort or "Brak uwag",
                "Granice": granice or "Brak ograniczeń",
                "Potrzeba": potrzeba, "Ochota na intymność": seks_ochota,
                "Pikantna niespodzianka": pikantna_niespodzianka or "Brak",
                "Motyw przewodni": motyw_przewodni or "Romantyczny wieczór",
                "Gotowość": zaznaczone_gotowosci, "Komplement": komplement or "Miłe słowo",
                "Własne pytanie": wlasne_pytanie or "Brak"
            }
            excel_bytes = generuj_excel_karty(dane_tymczasowe, osoba_do_pobrania)
            st.download_button(
                label=f"Pobierz kartę dla: {osoba_do_pobrania}",
                data=excel_bytes,
                file_name=f"karta_spotkania_{osoba_do_pobrania.lower()}_dark.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )

    # Sekcja porównania dostępna TYLKO w tym trybie, gdy są dane
    if len(st.session_state.odpowiedzi) >= 2:
        st.divider()
        st.header("✨ Porównanie Waszych odpowiedzi na żywo")
        oA = st.session_state.odpowiedzi.get(st.session_state.imie_1)
        oB = st.session_state.odpowiedzi.get(st.session_state.imie_2)
        
        c1, c2 = st.columns(2)
        with c1:
            st.markdown(f"**{st.session_state.imie_1}**")
            st.write(f"Energia: {oA['Energia']}/10 | Ochota: {oA['Ochota na intymność']}")
            st.write(f"Granice: {oA['Granice']}")
        with c2:
            st.markdown(f"**{st.session_state.imie_2}**")
            st.write(f"Energia: {oB['Energia']}/10 | Ochota: {oB['Ochota na intymność']}")
            st.write(f"Granice: {oB['Granice']}")

else:
    # Tryb relacji na odległość
    st.subheader("✈️ Tryb relacji na odległość")
    st.write("Wypełnij formularz jako Ty, pobierz swoją zmysłową kartę w pliku Excel i wyślij ją partnerowi (np. mailem lub komunikatorem).")
    
    if st.button("Pobierz moją kartę spotkania (Excel)"):
        dane_odleglosc = {
            "Data": datetime.now().strftime("%Y-%m-%d %H:%M"),
            "Energia": energia, "Nastrój": nastroj_slowo,
            "Bezpieczeństwo": bezpiecz_pytanie or "Brak uwag",
            "Komfort/Emocje": komfort or "Brak uwag",
            "Granice": granice or "Brak ograniczeń",
            "Potrzeba": potrzeba, "Ochota na intymność": seks_ochota,
            "Pikantna niespodzianka": pikantna_niespodzianka or "Brak",
            "Motyw przewodni": motyw_przewodni or "Romantyczny wieczór",
            "Gotowość": zaznaczone_gotowosci, "Komplement": komplement or "Miłe słowo",
            "Własne pytanie": wlasne_pytanie or "Brak"
        }
        excel_bytes = generuj_excel_karty(dane_odleglosc, aktualna_osoba)
        st.download_button(
            label=f"💾 Pobierz plik dla: {aktualna_osoba}",
            data=excel_bytes,
            file_name=f"karta_spotkania_{aktualna_osoba.lower()}_odleglosc.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
