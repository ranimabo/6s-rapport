import calendar
import json
import os
import tempfile
from pathlib import Path

import streamlit as st
from engine import generate_report

ROOT = Path(__file__).resolve().parent
st.set_page_config(page_title='Rapport 6S — Menzel Hayet', page_icon='📊', layout='centered')
st.markdown('''<style>
.block-container{max-width:860px;padding-top:2rem}
.title-small{color:#5b7892;font-size:14px;letter-spacing:1.1px;font-weight:700}
.hero{padding:1.2rem 1.4rem;border-radius:12px;background:linear-gradient(110deg,#082c53,#096ca8);color:white;margin-bottom:1.2rem}
.hero h1{color:white;font-size:2.0rem;margin:0}
.hero p{color:#ddeefd;margin:.3rem 0 0}
</style>''', unsafe_allow_html=True)
st.markdown('<div class="hero"><div class="title-small" style="color:#bbdcf5">OPERATIONS EXCELLENCE · MH</div><h1>Rapport 6S — LEONI</h1><p>Générez votre présentation mensuelle en un clic.</p></div>', unsafe_allow_html=True)

# Scores can be provisioned by an administrator as scores.json next to app.py.
# Otherwise they are uploaded per session, without being added to the GitHub repository.
scored_file = ROOT / 'scores.json'
if scored_file.exists():
    source_bytes=scored_file.read_bytes()
    st.success('Données des scores disponibles.')
else:
    uploaded=st.file_uploader('Fichier des scores Power BI (JSON)',type=['json'],help="Vous pouvez ajouter un fichier scores.json au dépôt privé pour éviter de le téléverser à chaque utilisation.")
    source_bytes=uploaded.getvalue() if uploaded is not None else None

try:
    rows=json.loads(source_bytes.decode('utf-8-sig')) if source_bytes else []
    if source_bytes and not isinstance(rows,list):raise ValueError('Le JSON doit contenir une liste.')
except Exception as ex:
    st.error(f'JSON invalide : {ex}')
    st.stop()

months=sorted({str(row.get('MH_BI[Month]',''))[:7] for row in rows if row.get('MH_BI[Month]')}, reverse=True)
if not months:
    st.info('Ajoutez votre fichier JSON pour afficher les mois disponibles et générer un rapport.')
    st.stop()

month=st.selectbox('Choisir le mois du rapport', months, format_func=lambda x: f'{calendar.month_name[int(x[5:7])]} {x[:4]}')
segments={str(row.get('MH_BI[Segment ]','')).strip() for row in rows if str(row.get('MH_BI[Month]','')).startswith(month)}
st.caption(f'{len(segments)} segments trouvés pour ce mois. Le rapport comprend les résultats de tous les segments.')
photo_zip=st.file_uploader('Photos par segment — ZIP (facultatif)',type=['zip'],help='Structure : 2026-06/Doors/photo1.jpg, 2026-06/FORD/photo2.jpg. Jusqu’à 6 photos par diapositive.')
with st.expander('Options de calcul'):
    missing_safety=st.checkbox('Afficher Sécurité à N/A lorsque sa valeur est 0',value=True,help='Les champs S6 du JSON fourni sont tous à zéro. Cette option évite de les interpréter comme des audits échoués sans confirmation.')
    st.caption('Le score global reprend le champ Power BI C_V 2 (moyenne des critères S1 à S5 dans le fichier fourni).')

if st.button('Générer le rapport PowerPoint',type='primary',use_container_width=True):
    try:
        with st.spinner('Création de la présentation…'):
            with tempfile.TemporaryDirectory() as tmp:
                json_path=Path(tmp)/'scores.json'
                json_path.write_bytes(source_bytes)
                zip_path=None
                if photo_zip is not None:
                    zip_path=Path(tmp)/'photos.zip'
                    zip_path.write_bytes(photo_zip.getvalue())
                report, count, photos, slides=generate_report(str(ROOT/'Template.pptx'),str(json_path),month,str(zip_path) if zip_path else None,missing_safety)
        st.session_state['result']=(month,report,count,photos,slides)
    except Exception as exc:
        st.error(f'La génération a échoué : {exc}')

if 'result' in st.session_state:
    selected,report,count,photos,slides=st.session_state['result']
    st.success(f'Rapport {selected} prêt · {count} segments · {slides} diapositives · {photos} photos')
    st.download_button('⬇️ Télécharger le rapport PowerPoint',data=report,file_name=f'6S_Audit_Menzel_Hayet_{selected}.pptx',mime='application/vnd.openxmlformats-officedocument.presentationml.presentation',type='primary',use_container_width=True)

st.divider()
st.caption('La génération se fait sur le serveur hébergeant cette application. Utilisez uniquement des données dont le traitement externe est autorisé.')
