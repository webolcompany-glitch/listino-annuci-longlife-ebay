import streamlit as st
import pandas as pd
import unicodedata
import io

st.title("Generatore File eBay con Description a paragrafi (senza immagini)")

uploaded_file = st.file_uploader("Carica file Excel", type=["xlsx"])

# Normalizzazione colonne
def normalize_col(col):
    col_norm = ''.join(
        c for c in unicodedata.normalize('NFKD', col)
        if not unicodedata.combining(c)
    )
    return col_norm.strip().lower()

# Funzione per creare HTML compatibile eBay con paragrafi
def format_html_ebay(title, desc):
    """
    Genera HTML compatibile eBay:
    - Titolo in <h2>
    - Ogni riga della descrizione diventa un <p>
    """
    if pd.isna(desc):
        desc = ""
    lines = [line.strip() for line in desc.split("\n") if line.strip()]
    html_desc = ""
    for line in lines:
        html_desc += f"<p>{line}</p>\n"
    html = f"<h2>{title}</h2>\n{html_desc}"
    return html

if uploaded_file is not None:
    df = pd.read_excel(uploaded_file)
    df.columns = [normalize_col(c) for c in df.columns]

    # Definizione colonne eBay
    ebay_columns = [
        "Action(SiteID=Italy|Country=IT|Currency=EUR|Version=1193)",
        "Custom label (SKU)",
        "Category name",
        "Title",
        "Relationship",
        "Relationship details",
        "Schedule Time",
        "P:EPID",
        "Start price",
        "Quantity",
        "Item photo URL",
        "VideoID",
        "Condition ID",
        "Description",
        "Format",
        "Duration",
        "Buy It Now price",
        "Best Offer Enabled",
        "Best Offer Auto Accept Price",
        "Minimum Best Offer Price",
        "VAT%",
        "Immediate pay required",
        "Location",
        "Shipping service 1 option",
        "Shipping service 1 cost",
        "Shipping service 1 priority",
        "Shipping service 2 option",
        "Shipping service 2 cost",
        "Shipping service 2 priority",
        "Max dispatch time",
        "Returns accepted option",
        "Returns within option",
        "Refund option",
        "Return shipping cost paid by",
        "Shipping profile name",
        "Return profile name",
        "Payment profile name",
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

    # Fissi e colonne base
    output["Action(SiteID=Italy|Country=IT|Currency=EUR|Version=1193)"] = "Add"
    output["Custom label (SKU)"] = df["sku"]
    output["Title"] = (
        " Olio Motore Auto " "x " + df["formato (l)"].astype(str) + "L" + "di " +
        df["nome olio"].astype(str) + " " +
        df["viscosita"].astype(str) + " " +
        df["tipologia"].astype(str) + " " +
        df["acea"].astype(str) + " " +
        df["marca"].astype(str)
    )
    output["Start price"] = df["prezzo marketplace"]
    output["Buy It Now price"] = df["prezzo marketplace"]
    output["Quantity"] = 10

    # Item photo URL (tutte le immagini separate da "|")
    img_cols = [c for c in df.columns if c.startswith("img")]
    def join_images(row):
        imgs = [str(row[col]) for col in img_cols if pd.notna(row[col])]
        return "|".join(imgs)
    output["Item photo URL"] = df.apply(join_images, axis=1)

    # HTML Descrizione compatibile eBay senza immagini
    descriptions = []
    for idx in df.index:
        title = output.at[idx, "Title"]
        desc_text = df.at[idx, "descrizione"] if "descrizione" in df.columns else ""
        descriptions.append(format_html_ebay(title, desc_text))
    output["Description"] = descriptions

    # Altri valori fissi
    output["Format"] = "FixedPrice"
    output["Duration"] = "GTC"
    output["VAT%"] = 22
    output["Location"] = "82030"
    output["Returns within option"] = "Days_14"
    output["Return shipping cost paid by"] = "Seller"

    # Attributi C
    output["C:Marca"] = df["marca"]
    output["C:MPN"] = df["codice prodotto"]
    output["C:Viscosità SAE"] = df["viscosita"]
    output["C:Marca veicolo"] = "Leggere descrizione per specifiche"
    output["C:Capienza"] = df["formato (l)"].apply(lambda x: "1 Litro" if float(x)==1 else f"{int(float(x))} Litri")
    output["C:Utilizzo"] = df["utilizzo"]
    output["C:Tipologia"] = df["tipologia"]

    # Manufacturer e Responsible Person
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

    # Esportazione Excel
    excel_buffer = io.BytesIO()
    with pd.ExcelWriter(excel_buffer, engine='openpyxl') as writer:
        output.to_excel(writer, index=False, sheet_name="eBay")
    excel_buffer.seek(0)

    st.success("File eBay generato correttamente senza immagini nella Description!")

    st.download_button(
        "Scarica file Excel",
        data=excel_buffer,
        file_name="ebay_output_senza_img.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )

    # Anteprima HTML live
    st.subheader("Anteprima HTML della Descrizione")
    for i, html_desc in enumerate(descriptions):
        st.markdown(f"**Prodotto {i+1}:**", unsafe_allow_html=True)
        st.markdown(html_desc, unsafe_allow_html=True)
        st.markdown("---")
