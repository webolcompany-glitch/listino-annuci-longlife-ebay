import streamlit as st
import pandas as pd
import unicodedata
import io

st.title("Generatore File eBay con Description a paragrafi (senza immagini)")

uploaded_file = st.file_uploader("Carica file Excel", type=["xlsx"])

# -------------------------
# NORMALIZZAZIONE COLONNE
# -------------------------
def normalize_col(col):
    col_norm = ''.join(
        c for c in unicodedata.normalize('NFKD', col)
        if not unicodedata.combining(c)
    )
    return col_norm.strip().lower()

# -------------------------
# HTML DESCRIPTION
# -------------------------
def format_html_ebay(title, desc):
    if pd.isna(desc):
        desc = ""
    lines = [line.strip() for line in str(desc).split("\n") if line.strip()]
    html_desc = ""
    for line in lines:
        html_desc += f"<p>{line}</p>\n"
    html = f"<h2>{title}</h2>\n{html_desc}"
    return html

# -------------------------
# FORMATTING CAPACITA' (C:Capienza)
# -------------------------
def format_capienza(x):
    try:
        x_float = float(x)
        if x_float == 1:
            return "1 Litro"
        else:
            return f"{int(x_float)} Litri"
    except (ValueError, TypeError):
        return "Sconosciuto"

# -------------------------
# INIZIO ELABORAZIONE
# -------------------------
if uploaded_file is not None:

    df = pd.read_excel(uploaded_file)
    df.columns = [normalize_col(c) for c in df.columns]

    st.write("Colonne rilevate:", df.columns.tolist())

    # -------------------------
    # COLONNE OBBLIGATORIE
    # -------------------------
    required_cols = [
        "sku",
        "formato (l)",
        "nome olio",
        "viscosita",
        "tipologia",
        "acea",
        "marca",
        "prezzo marketplace",
        "codice prodotto",
        "utilizzo"
    ]

    missing = [c for c in required_cols if c not in df.columns]

    if missing:
        st.error(f"Colonne mancanti nel file Excel: {missing}")
        st.stop()

    # -------------------------
    # STRUTTURA OUTPUT EBAY
    # -------------------------
    ebay_columns = [
        "Action(SiteID=Italy|Country=IT|Currency=EUR|Version=1193)",
        "Custom label (SKU)",
        "Title",
        "Start price",
        "Quantity",
        "Item photo URL",
        "Description",
        "Format",
        "Duration",
        "Buy It Now price",
        "VAT%",
        "Location",
        "Returns within option",
        "Return shipping cost paid by",
        "C:Marca",
        "C:MPN",
        "C:Viscosità SAE",
        "C:Marca veicolo",
        "C:Capienza",
        "C:Utilizzo",
        "C:Tipologia",
        "Manufacturer Name",
        "Manufacturer AddressLine1",
        "Manufacturer City",
        "Manufacturer Country",
        "Manufacturer PostalCode",
        "Manufacturer Email",
        "Responsible Person 1",
        "Responsible Person 1 Type",
        "Responsible Person 1 City",
        "Responsible Person 1 Country",
        "Responsible Person 1 PostalCode"
    ]

    output = pd.DataFrame(columns=ebay_columns)

    # -------------------------
    # CAMPI BASE
    # -------------------------
    output["Action(SiteID=Italy|Country=IT|Currency=EUR|Version=1193)"] = "Add"
    output["Custom label (SKU)"] = df["sku"]

    output["Title"] = (
        "Olio Motore Auto "
        + df["formato (l)"].astype(str)
        + " L di "
        + df["nome olio"].astype(str) + " "
        + df["viscosita"].astype(str) + " "
        + df["tipologia"].astype(str) + " "
        + df["acea"].astype(str) + " "
        + df["marca"].astype(str)
    )

    output["Start price"] = df["prezzo marketplace"]
    output["Buy It Now price"] = df["prezzo marketplace"]
    output["Quantity"] = 10

    # -------------------------
    # IMMAGINI
    # -------------------------
    img_cols = [c for c in df.columns if c.startswith("img")]

    def join_images(row):
        imgs = [str(row[col]) for col in img_cols if pd.notna(row[col])]
        return "|".join(imgs)

    output["Item photo URL"] = df.apply(join_images, axis=1)

    # -------------------------
    # DESCRIZIONE HTML
    # -------------------------
    descriptions = []

    for idx in df.index:
        title = output.at[idx, "Title"]
        desc_text = df.at[idx, "descrizione"] if "descrizione" in df.columns else ""
        descriptions.append(format_html_ebay(title, desc_text))

    output["Description"] = descriptions

    # -------------------------
    # VALORI FISSI
    # -------------------------
    output["Format"] = "FixedPrice"
    output["Duration"] = "GTC"
    output["VAT%"] = 22
    output["Location"] = "82030"
    output["Returns within option"] = "Days_14"
    output["Return shipping cost paid by"] = "Seller"

    # -------------------------
    # ATTRIBUTI C
    # -------------------------
    output["C:Marca"] = df["marca"]
    output["C:MPN"] = df["codice prodotto"]
    output["C:Viscosità SAE"] = df["viscosita"]
    output["C:Marca veicolo"] = "Leggere descrizione per specifiche"

    output["C:Capienza"] = df["formato (l)"].apply(format_capienza)
    output["C:Utilizzo"] = df["utilizzo"]
    output["C:Tipologia"] = df["tipologia"]

    # -------------------------
    # DATI PRODUTTORE
    # -------------------------
    output["Manufacturer Name"] = "TAMOIL ITALIA S.p.A."
    output["Manufacturer AddressLine1"] = "Via Andrea Costa 17"
    output["Manufacturer City"] = "Milano"
    output["Manufacturer Country"] = "IT"
    output["Manufacturer PostalCode"] = "20131"
    output["Manufacturer Email"] = "tamoil.italia@pec.tamoil.it"

    output["Responsible Person 1"] = "TAMOIL ITALIA S.p.A."
    output["Responsible Person 1 Type"] = "Via Andrea Costa 17"
    output["Responsible Person 1 City"] = "Milano"
    output["Responsible Person 1 Country"] = "IT"
    output["Responsible Person 1 PostalCode"] = "20131"

    # -------------------------
    # EXPORT EXCEL
    # -------------------------
    excel_buffer = io.BytesIO()

    with pd.ExcelWriter(excel_buffer, engine='openpyxl') as writer:
        output.to_excel(writer, index=False, sheet_name="eBay")

    excel_buffer.seek(0)

    st.success("File eBay generato correttamente!")

    st.download_button(
        "Scarica file Excel",
        data=excel_buffer,
        file_name="ebay_output.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )

    # -------------------------
    # ANTEPRIMA HTML
    # -------------------------
    st.subheader("Anteprima HTML della Descrizione")

    for i, html_desc in enumerate(descriptions):
        st.markdown(f"**Prodotto {i+1}:**", unsafe_allow_html=True)
        st.markdown(html_desc, unsafe_allow_html=True)
        st.markdown("---")
