"""
app.py
======
Streamlit-webapp: upload een sjabloon (opgemaakt referentiedocument) en een
brondocument (ongeformatteerde tekst), en download het resultaat waarbij de
tekst automatisch in de opmaakstructuur van het sjabloon is gegoten.

Lokaal starten:
    pip install -r requirements.txt
    streamlit run app.py

Online zetten (gratis, geen server nodig):
    1. Zet dit project in een GitHub-repository (app.py, devis_engine.py,
       requirements.txt).
    2. Ga naar https://share.streamlit.io en meld je aan met je GitHub-account.
    3. Kies "New app", selecteer de repository en het bestand app.py.
    4. Klik "Deploy". Je krijgt een permanente link (bv. iets.streamlit.app)
       die je met de dossierbeheerster kan delen.
"""

import streamlit as st
from devis_engine import fill_template_from_bytes

st.set_page_config(page_title="Devis in sjabloon gieten", page_icon="📄", layout="centered")

st.title("📄 Word-document in sjabloon gieten")
st.write(
    "Upload een ongeformatteerd Word-document (de ruwe tekst) en het "
    "opgemaakte referentiedocument (het sjabloon). De app giet de tekst "
    "automatisch in de juiste opmaak: lettertype, vet/cursief/onderstreping, "
    "kaders en uitlijning van bedragen."
)

st.divider()

col1, col2 = st.columns(2)

with col1:
    st.subheader("1. Sjabloon (opmaak)")
    template_file = st.file_uploader(
        "Referentiedocument met de gewenste opmaak",
        type=["docx"],
        key="template",
        help="Dit is het document waarvan de lay-out (lettertype, header, kaders) wordt overgenomen."
    )

with col2:
    st.subheader("2. Brondocument (tekst)")
    input_file = st.file_uploader(
        "Nieuw document met de ruwe, ongeformatteerde tekst",
        type=["docx"],
        key="input",
        help="Dit is de tekst die in de opmaak van het sjabloon gegoten wordt."
    )

st.divider()

if template_file and input_file:
    if st.button("🚀 Verwerk document", type="primary", use_container_width=True):
        try:
            with st.spinner("Bezig met verwerken..."):
                template_bytes = template_file.read()
                input_bytes = input_file.read()
                result_bytes, num_items = fill_template_from_bytes(template_bytes, input_bytes)

            st.success(f"Klaar! {num_items} alinea's correct opgemaakt.")

            output_name = input_file.name.rsplit(".", 1)[0] + "_in_sjabloon.docx"
            st.download_button(
                label="⬇️ Download resultaat",
                data=result_bytes,
                file_name=output_name,
                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                use_container_width=True,
            )

        except ValueError as e:
            st.error(
                f"Kon de structuur niet herkennen: {e}\n\n"
                "Controleer of het sjabloon minstens een voorbeeld bevat van elk "
                "structuuronderdeel (titel, subtitel A./B./..., verdiepingsniveau "
                "zoals REZ-DE-CHAUSSEE, ruimtenaam gevolgd door Contenant/Contenu, "
                "en minstens een bullet-regel met een ':' erin)."
            )
        except Exception as e:
            st.error(f"Er ging iets mis tijdens de verwerking: {e}")
else:
    st.info("Upload beide bestanden om te starten.")

st.divider()
with st.expander("ℹ️ Hoe werkt dit precies?"):
    st.markdown(
        """
        Deze app herkent in het **sjabloon** negen soorten opmaakblokken:
        hoofdtitel, onderlijnde titel, subtitel (A./B./C.…), verdiepingsniveau
        (bv. REZ-DE-CHAUSSEE), ruimtenaam in kader, het label
        "Contenant"/"Contenu", het rechts uitgelijnde bedrag, een inleidende
        lijst, en gewone bullet-tekst.

        Vervolgens leest de app het **brondocument** regel per regel, herkent
        welke rol elke regel heeft (op basis van patronen zoals hoofdletters,
        het woord "Contenant", of een bedrag met €-teken), en plakt de tekst
        in een kopie van het bijhorende opmaakblok uit het sjabloon.

        Werkt dit niet correct voor jouw documenttype? Pas de patronen in
        `devis_engine.py` (FLOOR_PATTERNS, MAIN_TITLE_PATTERNS, classify_line)
        aan aan jouw structuur.
        """
    )
