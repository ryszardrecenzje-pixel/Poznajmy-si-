import streamlit as st
import pandas as pd
from datetime import datetime
from io import BytesIO
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

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

# Podsumowanie i Generowanie Ładnego Excela
if st.button("Zapisz i pobierz ładny plik Excel"):
    st.success(f"Dziękujemy, {osoba}! Twoje odpowiedzi zostały przygotowane.")
    
    # 1. Tworzenie arkusza przez openpyxl
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Wyniki Quizu"
    ws.views.sheetView[0].showGridLines = True
    
    # Stylizacja nagłówków (ciemny, elegancki kolor)
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
        "Ochota na intymność", "Wybrane zachcianki", "Własne pytanie"
    ]
    
    ws.append(headers)
    ws.row_dimensions[1].height = 28
    
    for col_num, header in enumerate(headers, 1):
        cell = ws.cell(row=1, column=col_num)
        cell.font = header_font
        cell.fill = header_fill
        cell.alignment = header_alignment
        cell.border = thin_border
        
    # Wstawianie danych użytkownika
    row_data = [
        datetime.now().strftime("%Y-%m-%d %H:%M"),
        osoba,
        energia,
        nastroj_slowo,
        komfort if komfort else "Brak uwag",
        potrzeba,
        seks_ochota,
        ", ".join(zaznaczone_ochoty) if zaznaczone_ochoty else "Brak",
        wlasne_pytanie if wlasne_pytanie else "Brak"
    ]
    
    ws.append(row_data)
    ws.row_dimensions[2].height = 35
    
    for col_num in range(1, len(headers) + 1):
        cell = ws.cell(row=2, column=col_num)
        cell.alignment = Alignment(vertical="center", wrap_text=True)
        cell.border = thin_border
        
    # Automatyczne dopasowanie szerokości kolumn
    for col in ws.columns:
        max_len = max(len(str(cell.value or '')) for cell in col)
        col_letter = get_column_letter(col[0].column)
        ws.column_dimensions[col_letter].width = max(max_len + 4, 18)
        
    # Zapis do pamięci jako plik binarny do pobrania
    output = BytesIO()
    wb.save(output)
    excel_data = output.getvalue()
    
    # Przycisk pobierania
    st.download_button(
        label="📥 Pobierz sformatowany plik Excel (.xlsx)",
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
