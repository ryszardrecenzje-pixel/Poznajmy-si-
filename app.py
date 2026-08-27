import streamlit as st
import pandas as pd
from datetime import datetime
from io import BytesIO
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

st.set_page_config(page_title="Quiz dla Par", page_icon="💖", layout="centered")

st.title("💖 Przedspotkaniowy Check-in dla Par")
st.write("Skonfigurujcie aplikację swoimi imionami, a następnie wypełnijcie quiz, aby dopasować spotkanie do Waszych potrzeb.")

# Sekcja ustawienia imion (zapisujemy w sesji Streamlita, żeby pamiętało oba imiona)
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
st.header("5. Twoje własne pytania")
wlasne_pytanie = st.text_input(f"{aktualna_osoba}: Masz jakieś specjalne pytanie lub prośbę, którą chcesz przekazać drugiej osobie?")

st.divider()

# Podsumowanie i Generowanie Excela
if st.button("Zapisz i pobierz ładny plik Excel"):
    st.success(f"Dziękujemy, {aktualna_osoba}! Twoje odpowiedzi zostały przygotowane.")
    
    # Tworzenie arkusza przez openpyxl
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Wyniki Quizu"
    ws.views.sheetView[0].showGridLines = True
    
    HEADER_BG = "34495E"
    HEADER_FG = "FFFFFF"
    BORDER_COLOR = "D0D3D4"
    
    header_font = Font(name="Calibri", size=11, bold=True, color=HEADER_FG)
    header_fill = PatternFill(start_color=HEADER_BG, end_color=HEADER_BG, fill_type="solid")
    header_alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    thin_border = Border(
        left=Side(style='thin', color=BORDER_COLOR),
        right=Side(style='thin', color=BORDER_COLOR),
        top=Side(style='thin', color=BORDER_COLOR),
        bottom=Side(style='thin', color=BORDER_COLOR)
    )
    
    headers = [
        "Data", "Osoba", "Poziom Energii", "Nastrój", 
        "Stan emocjonalny / Komfort", "Główna potrzeba", 
        "Ochota na intymność", "Jestem gotowy/wa na...", "Własne pytanie"
    ]
    
    ws.append(headers)
    ws.row_dimensions[1].height = 28
    
    for col_num, header in enumerate(headers, 1):
        cell = ws.cell(row=1, column=col_num)
        cell.font = header_font
        cell.fill = header_fill
        cell.alignment = header_alignment
        cell.border = thin_border
        
    row_data = [
        datetime.now().strftime("%Y-%m-%d %H:%M"),
        aktualna_osoba,
        energia,
        nastroj_slowo,
        komfort if komfort else "Brak uwag",
        potrzeba,
        seks_ochota,
        ", ".join(zaznaczone_gotowosci) if zaznaczone_gotowosci else "Brak",
        wlasne_pytanie if wlasne_pytanie else "Brak"
    ]
    
    ws.append(row_data)
    ws.row_dimensions[2].height = 35
    
    for col_num in range(1, len(headers) + 1):
        cell = ws.cell(row=2, column=col_num)
        cell.alignment = Alignment(vertical="center", wrap_text=True)
        cell.border = thin_border
        
    for col in ws.columns:
        max_len = max(len(str(cell.value or '')) for cell in col)
        col_letter = get_column_letter(col[0].column)
        ws.column_dimensions[col_letter].width = max(max_len + 4, 18)
        
    output = BytesIO()
    wb.save(output)
    excel_data = output.getvalue()
    
    st.download_button(
        label=f"📥 Pobierz plik Excel z wynikami ({aktualna_osoba})",
        data=excel_data,
        file_name=f"quiz_wyniki_{aktualna_osoba.lower()}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )

    with st.expander("Zobacz podsumowanie na ekranie"):
        st.write(f"**Poziom energii:** {energia}")
        st.write(f"**Nastrój:** {nastroj_slowo}")
        st.write(f"**Komfort/Emocje:** {komfort if komfort else 'Brak uwag'}")
        st.write(f"**Główna potrzeba:** {potrzeba}")
        st.write(f"**Ochota na intymność:** {seks_ochota}")
        st.write(f"**Jestem gotowy/wa na:** {zaznaczone_gotowosci}")
        if wlasne_pytanie:
            st.write(f"**Twoje specjalne pytanie:** {wlasne_pytanie}")
