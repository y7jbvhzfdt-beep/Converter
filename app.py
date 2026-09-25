"""
app.py
======
Streamlit-webapp: upload enkel het brondocument (ongeformatteerde tekst).
Het sjabloon (opmaak, logo, header/footer) is vast ingebouwd in het project
als sjabloon.docx en wordt automatisch gebruikt bij elke verwerking.

BELANGRIJK: dit bestand mag GEEN tweede st.file_uploader voor het sjabloon
bevatten. Als je in de live app nog een upload-veld voor het sjabloon ziet,
staat er een oudere versie van dit bestand in je GitHub-repository — upload
dit bestand opnieuw met dezelfde naam (app.py) om het te overschrijven, en
zorg dat sjabloon.docx in dezelfde map staat.

Lokaal starten:
    pip install -r requirements.txt
    streamlit run app.py

Online zetten (gratis, geen server nodig):
    1. Zet dit project in een GitHub-repository (app.py, devis_engine.py,
       requirements.txt, EN sjabloon.docx) — alle 4 in dezelfde map.
    2. Ga naar https://share.streamlit.io en meld je aan met je GitHub-account.
    3. Kies "New app", selecteer de repository en het bestand app.py.
    4. Klik "Deploy". Je krijgt een permanente link (bv. iets.streamlit.app)
       die je met de dossierbeheerster kan delen.
"""

import streamlit as st
from pathlib import Path
from devis_engine import fill_template_from_bytes

st.set_page_config(page_title="Devis in sjabloon gieten", page_icon="📄", layout="centered")

st.title("📄 Word-document in sjabloon gieten")
st.write(
    "Upload een ongeformatteerd Word-document (de ruwe tekst). De app giet "
    "de tekst automatisch in de vaste bedrijfsopmaak: lettertype, "
    "vet/cursief/onderstreping, kaders, uitlijning van bedragen, logo en "
    "footer."
)

st.divider()

TEMPLATE_PATH = Path(__file__).parent / "sjabloon.docx"

if not TEMPLATE_PATH.exists():
    st.error(
        "Het vaste sjabloon (sjabloon.docx) werd niet gevonden naast app.py. "
        "Zorg dat dit bestand mee is geüpload naar de GitHub-repository, in "
        "dezelfde map als app.py."
    )
    st.stop()

st.subheader("Brondocument (tekst)")
input_file = st.file_uploader(
    "Nieuw document met de ruwe, ongeformatteerde tekst",
    type=["docx"],
    key="input",
    help="Dit is de tekst die in de vaste bedrijfsopmaak gegoten wordt."
)

st.divider()

if input_file:
    if st.button("🚀 Verwerk document", type="primary", use_container_width=True):
        try:
            with st.spinner("Bezig met verwerken..."):
                template_bytes = TEMPLATE_PATH.read_bytes()
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
                f"Kon de tekst niet correct verwerken: {e}\n\n"
                "Controleer of het brondocument de juiste structuur volgt "
                "(titel, sectie A. met verdiepingen en ruimtes, Contenant/Contenu "
                "met bedragen, en beschrijvingen met een ':' erin)."
            )
        except Exception as e:
            st.error(f"Er ging iets mis tijdens de verwerking: {e}")
else:
    st.info("Upload het brondocument om te starten.")

st.divider()
with st.expander("ℹ️ Hoe werkt dit precies?"):
    st.markdown(
        """
        Deze app gebruikt een **vast, ingebouwd sjabloon** (met logo, header
        en footer) dat niet meer hoeft te worden geüpload. Je hoeft enkel het
        **brondocument** met de ruwe tekst te uploaden.

        De app leest het brondocument regel per regel, herkent welke rol
        elke regel heeft (titel, subtitel A./B./…, verdiepingsniveau,
        ruimtenaam, Contenant/Contenu-label met bedrag, of een gewone
        beschrijving), en plakt de tekst in de bijhorende opmaak van het
        vaste sjabloon — inclusief de witruimte tussen structuurelementen.

        Wil je de vaste opmaak wijzigen? Vervang het bestand `sjabloon.docx`
        in de GitHub-repository door een nieuwe versie met dezelfde
        structuur (titel, onderlijnde titel, subtitel A./B./…,
        verdiepingsniveau, ruimtenaam in kader, Contenant/Contenu, en
        minstens één voorbeeldregel met een ':' erin).
        """
    )
