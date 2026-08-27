import streamlit as st
import pandas as pd
from datetime import datetime
from io import BytesIO
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side

st.set_page_config(page_title="Intymny Quiz dla Par", page_icon="💖", layout="centered")

st.title("💖 Przedspotkaniowy Check-in dla Par")
st.write("Wypełnijcie quiz, aby otworzyć się na siebie i sprawdzić, na co macie dziś ochotę.")

# Sekcja ustawienia imion
if "imie_1" not in st.session_state:
    st.session_state.imie_1 = "Osoba A"
if "imie_2" not in st.session_state:
    st.session_state.imie_2 = "Osoba B"

col1, col2 = st.columns(2)
with col1:
    st.session_state.imie_1 = st.text_input("Imię Pierwszej Osoby:", value=st.session_state.imie_1)
with col2:
    st.session_state.imie_2 = st.text_input("Imię Drugiej Osoby:", value=st.session_state.imie_2)

st.divider()

# Wybór kto teraz wypełnia
aktualna_osoba = st.radio("Kto teraz wypełnia quiz?", [st.session_state.imie_1, st.session_state.imie_2], horizontal=True)

st.divider()

# 1. Nastrój i Energia
st.header("1. Nastrój i Energia")
energia = st.slider(f"{aktualna_osoba}: Jaki masz poziom energii na dzisiejsze spotkanie?", 1, 10, 5)
nastroj_slowo = st.selectbox(
    f"{aktualna_osoba}: Określ swój obecny nastrój jednym słowem:",
    ["Zrelaksowany/a", "Podekscytowany/a", "Zmęczony/a", "Czule nastrojony/a", "Stresujący dzień", "Gotowy/a na wszystko"]
)

# 2. Emocje i Oczekiwania
st.header("2. Emocje i Granice")
komfort = st.text_area(f"{aktualna_osoba}: Czy jest coś, o czym partner powinien wiedzieć o Twoim dzisiejszym stanie emocjonalnym?")
potrzeba = st.selectbox(
    f"{aktualna_osoba}: Czego najbardziej dzisiaj potrzebujesz?",
    ["Wspólnego wyciszenia i rozmowy", "Czułości i przytulasów", "Dobrej zabawy i śmiechu", "Bliskości fizycznej / Intymności"]
)

# 3. Seksualność i Intymność
st.header("3. Seksualność i Intymność")
seks_ochota = st.select_slider(
    f"{aktualna_osoba}: Jak dużą masz dziś ochotę na zbliżenia intymne?",
    options=["Brak ochoty / Tylko czułość", "Otwarty/a, jeśli powoli", "Umiarkowana", "Duża", "100% ogień! 🔥"]
)

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

# 5. Pytania niestandardowe (własne)
st.header("5. Własne pytania")
wlasne_pytanie = st.text_input(f"{aktualna_osoba}: Masz jakieś specjalne pytanie lub prośbę, którą chcesz przekazać drugiej osobie?")

st.divider()

# Podsumowanie i Generowanie Ulotki w Stylu "Dark Mode"
if st.button("Zapisz i pobierz kartę spotkania (Dark Mode)"):
    st.success(f"Dziękujemy, {aktualna_osoba}! Twoja karta spotkania została przygotowana.")
    
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Karta Spotkania"
    
    # Całkowite wyłączenie widoczności siatki (efekt czystej ulotki)
    ws.views.sheetView[0].showGridLines = False
    
    # Paleta Dark Mode (Głęboka czerń + grafitowe karty + różowe akcenty)
    DARK_BG = "000000"          # Czern tła
    CARD_BG = "1A1A1A"          # Grafit na dane
    HEADER_BG = "2C2C2C"        # Odcień na etykiety
    TEXT_WHITE = "FFFFFF"       # Biały tekst
    TEXT_ACCENT = "FFD700"      # Złoto-różowy akcent na najważniejsze pola
    BORDER_DARK = "333333"      # Ciemne, subtelne obramowanie
    
    title_font = Font(name="Segoe UI", size=14, bold=True, color="FF69B4")
    header_font = Font(name="Segoe UI", size=11, bold=True, color=TEXT_WHITE)
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
    
    # Elegancki nagłówek ulotki
    ws.merge_cells('A1:B1')
    ws['A1'] = f"💖 KARTA SPOTKANIA: {aktualna_osoba.upper()}"
    ws['A1'].font = title_font
    ws['A1'].alignment = Alignment(horizontal="left", vertical="center")
    ws.row_dimensions[1].height = 40
    
    # Dane do karty mobilnej (pionowy układ w dwóch kolumnach)
    fields = [
        ("📅 Data i Czas", datetime.now().strftime("%Y-%m-%d %H:%M")),
        ("👤 Kto odpowiada", aktualna_osoba),
        ("⚡ Poziom Energii", f"{energia} / 10"),
        ("✨ Nastrój", nastroj_slowo),
        ("💭 Stan emocjonalny / Komfort", komfort if komfort else "Wszystko w porządku"),
        ("💖 Główna Potrzeba", potrzeba),
        ("🔥 Ochota na intymność", seks_ochota),
        ("🌹 Jestem gotowy/wa na...", ", ".join(zaznaczone_gotowosci) if zaznaczone_gotowosci else "Czułość i bliskość"),
        ("💬 Pytanie do partnera", wlasne_pytanie if wlasne_pytanie else "Brak pytań na ten moment")
    ]
    
    for idx, (label, val) in enumerate(fields, start=3):
        ws.row_dimensions[idx].height = 32
        
        # Etykieta (Lewa kolumna)
        cell_lbl = ws.cell(row=idx, column=1, value=label)
        cell_lbl.font = header_font
        cell_lbl.fill = header_fill
        cell_lbl.alignment = Alignment(horizontal="left", vertical="center", indent=1)
        cell_lbl.border = thin_border
        
        # Wartość (Prawa kolumna)
        cell_val = ws.cell(row=idx, column=2, value=val)
        # Podświetlenie najważniejszych sekcji na wyróżniający kolor
        if "intymność" in label.lower() or "gotowy" in label.lower():
            cell_val.font = accent_data_font
        else:
            cell_val.font = data_font
            
        cell_val.fill = data_fill
        cell_val.alignment = Alignment(horizontal="left", vertical="center", wrap_text=True, indent=1)
        cell_val.border = thin_border
        
    # Idealne dopasowanie szerokości pod ekran telefonu
    ws.column_dimensions['A'].width = 30
    ws.column_dimensions['B'].width = 55
        
    output = BytesIO()
    wb.save(output)
    excel_data = output.getvalue()
    
    st.download_button(
        label=f"🖤 Pobierz kartę spotkania (Dark Mode) - {aktualna_osoba}",
        data=excel_data,
        file_name=f"karta_spotkania_{aktualna_osoba.lower()}_dark.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )

    with st.expander("Zobacz podsumowanie na ekranie"):
        st.write(f"**Poziom energii:** {energia}/10")
        st.write(f"**Nastrój:** {nastroj_slowo}")
        st.write(f"**Komfort/Emocje:** {komfort if komfort else 'Brak uwag'}")
        st.write(f"**Główna potrzeba:** {potrzeba}")
        st.write(f"**Ochota na intymność:** {seks_ochota}")
        st.write(f"**Jestem gotowy/wa na:** {zaznaczone_gotowosci}")
        if wlasne_pytanie:
            st.write(f"**Twoje specjalne pytanie:** {wlasne_pytanie}")
