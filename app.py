import streamlit as st
import pandas as pd
from datetime import datetime
from io import BytesIO
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side

st.set_page_config(page_title="Intymny Quiz dla Par", page_icon="💖", layout="centered")

st.title("💖 Przedspotkaniowy Check-in dla Par")
st.write("Odpowiedzcie na pytania, aby dostosować dzisiejsze spotkanie do Waszych aktualnych potrzeb, granic i pragnień.")

# Inicjalizacja pamięci sesji
if "imie_1" not in st.session_state:
    st.session_state.imie_1 = "Osoba A"
if "imie_2" not in st.session_state:
    st.session_state.imie_2 = "Osoba B"
if "odpowiedzi" not in st.session_state:
    st.session_state.odpowiedzi = {}

# Sekcja ustawienia imion
col1, col2 = st.columns(2)
with col1:
    st.session_state.imie_1 = st.text_input("Imię Pierwszej Osoby:", value=st.session_state.imie_1)
with col2:
    st.session_state.imie_2 = st.text_input("Imię Drugiej Osoby:", value=st.session_state.imie_2)

st.divider()

# Wybór kto teraz wypełnia
aktualna_osoba = st.radio("Kto teraz wypełnia quiz?", [st.session_state.imie_1, st.session_state.imie_2], horizontal=True)

st.divider()

# 1. Nastrój, Energia i Bezpieczeństwo
st.header("1. Nastrój, Energia i Bezpieczeństwo")
energia = st.slider(f"{aktualna_osoba}: Jaki masz poziom energii na dzisiejsze spotkanie?", 1, 10, 5)
nastroj_slowo = st.selectbox(
    f"{aktualna_osoba}: Określ swój obecny nastrój jednym słowem:",
    ["Zrelaksowany/a", "Podekscytowany/a", "Zmęczony/a", "Czule nastrojony/a", "Stresujący dzień", "Gotowy/a na wszystko"]
)
bezpiecz_pytanie = st.text_input(f"{aktualna_osoba}: Co mogę dzisiaj zrobić, żebyś poczuł/a się przy mnie w 100% bezpiecznie i swobodnie?")

# 2. Emocje i Granice (Bardzo ważne!)
st.header("2. Emocje i Granice")
komfort = st.text_area(f"{aktualna_osoba}: Czy jest coś, o czym partner powinien wiedzieć o Twoim dzisiejszym stanie emocjonalnym?")
granice = st.text_input(f"{aktualna_osoba}: Czego dzisiaj ZUPEŁNIE NIE CHCESZ / jakie są Twoje granice na ten wieczór?")
potrzeba = st.selectbox(
    f"{aktualna_osoba}: Czego najbardziej dzisiaj potrzebujesz?",
    ["Wspólnego wyciszenia i rozmowy", "Czułości i przytulasów", "Dobrej zabawy i śmiechu", "Bliskości fizycznej / Intymności"]
)

# 3. Seksualność, Fantazje i Oczekiwania
st.header("3. Seksualność i Intymność")
seks_ochota = st.select_slider(
    f"{aktualna_osoba}: Jak dużą masz dziś ochotę na zbliżenia intymne?",
    options=["Brak ochoty / Tylko czułość", "Otwarty/a, jeśli powoli", "Umiarkowana", "Duża", "100% ogień! 🔥"]
)
pikantna_niespodzianka = st.text_input(f"{aktualna_osoba}: Na jaką małą, pikantną niespodziankę masz dzisiaj ochotę?")
motyw_przewodni = st.text_input(f"{aktualna_osoba}: Gdyby ten wieczór miał swój zmysłowy motyw przewodni, to jak by brzmiał?")

# 4. Tabela: Jestem gotowy/wa na...
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

# 5. Pytania niestandardowe i docenienie
st.header("5. Pytania i Docenienie")
komplement = st.text_input(f"{aktualna_osoba}: Jaki komplement chciałbyś/chciałabyś usłyszeć ode mnie dzisiaj wieczorem?")
wlasne_pytanie = st.text_input(f"{aktualna_osoba}: Masz jakieś specjalne pytanie lub prośbę, którą chcesz przekazać drugiej osobie?")

st.divider()

# Zapisanie odpowiedzi w pamięci sesji dla danej osoby
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

# Sekcja porównania i generowania pliku Excel, gdy obie osoby odpowiedziały
if len(st.session_state.odpowiedzi) >= 2:
    st.divider()
    st.header("✨ Porównanie i Dopasowanie Waszych Nastrojów")
    
    osoba_A_dane = st.session_state.odpowiedzi.get(st.session_state.imie_1)
    osoba_B_dane = st.session_state.odpowiedzi.get(st.session_state.imie_2)
    
    if osoba_A_dane and osoba_B_dane:
        col_A, col_B = st.columns(2)
        with col_A:
            st.subheader(f"👤 {st.session_state.imie_1}")
            st.write(f"**Energia:** {osoba_A_dane['Energia']}/10")
            st.write(f"**Nastrój:** {osoba_A_dane['Nastrój']}")
            st.write(f"**Granice:** {osoba_A_dane['Granice']}")
            st.write(f"**Ochota:** {osoba_A_dane['Ochota na intymność']}")
            st.write(f"**Pytanie:** {osoba_A_dane['Własne pytanie']}")
            
        with col_B:
            st.subheader(f"👤 {st.session_state.imie_2}")
            st.write(f"**Energia:** {osoba_B_dane['Energia']}/10")
            st.write(f"**Nastrój:** {osoba_B_dane['Nastrój']}")
            st.write(f"**Granice:** {osoba_B_dane['Granice']}")
            st.write(f"**Ochota:** {osoba_B_dane['Ochota na intymność']}")
            st.write(f"**Pytanie:** {osoba_B_dane['Własne pytanie']}")
            
        st.info("💡 **Wskazówka na randkę:** Przed spotkaniem omówcie szczególnie Wasze granice oraz to, na co oboje wyraziliście gotowość!")

    st.divider()
    
    # Generowanie zmysłowej karty w Dark Mode dla wybranej osoby
    wybrana_do_pobrania = st.selectbox("Wybierz czyją kartę chcesz pobrać jako plik Excel:", [st.session_state.imie_1, st.session_state.imie_2])
    
    if st.button("📥 Pobierz zmysłową kartę spotkania (Dark Mode)"):
        dane_karty = st.session_state.odpowiedzi[wybrana_do_pobrania]
        
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
        ws['A1'] = f"💖 KARTA SPOTKANIA: {wybrana_do_pobrania.upper()}"
        ws['A1'].font = title_font
        ws['A1'].alignment = Alignment(horizontal="left", vertical="center")
        ws.row_dimensions[1].height = 40
        
        fields = [
            ("📅 Data i Czas", dane_karty["Data"]),
            ("👤 Kto odpowiada", wybrana_do_pobrania),
            ("⚡ Poziom Energii", f"{dane_karty['Energia']} / 10"),
            ("✨ Nastrój", dane_karty["Nastrój"]),
            ("🛡️ Jak zadbać o bezpieczeństwo", dane_karty["Bezpieczeństwo"]),
            ("💭 Stan emocjonalny / Komfort", dane_karty["Komfort/Emocje"]),
            ("🚫 Granice / Czego nie chcę", dane_karty["Granice"]),
            ("💖 Główna Potrzeba", dane_karty["Potrzeba"]),
            ("🔥 Ochota na intymność", dane_karty["Ochota na intymność"]),
            ("🎁 Pikantna niespodzianka", dane_karty["Pikantna niespodzianka"]),
            ("🌙 Motyw przewodni wieczoru", dane_karty["Motyw przewodni"]),
            ("🌹 Jestem gotowy/wa na...", ", ".join(dane_karty["Gotowość"]) if dane_karty["Gotowość"] else "Czułość i bliskość"),
            ("✨ Pożądany komplement", dane_karty["Komplement"]),
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
        excel_data = output.getvalue()
        
        st.download_button(
            label=f"📥 Pobierz plik sformatowany dla: {wybrana_do_pobrania}",
            data=excel_data,
            file_name=f"karta_spotkania_{wybrana_do_pobrania.lower()}_dark.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
else:
    st.info("ℹ️ Wskazówka: Gdy obie osoby wypełnią i zapiszą swoje odpowiedzi, u dołu ekranu pojawi się sekcja porównania Waszych nastrojów oraz możliwość pobrania kart w stylu Dark Mode!")
