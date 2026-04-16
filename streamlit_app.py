import streamlit as st
import pandas as pd
import io
from openpyxl import Workbook
from openpyxl.styles import PatternFill, Font, Alignment, Border, Side
from openpyxl.utils import get_column_letter

st.set_page_config(page_title="RoC - Import Ville", page_icon="🏙️", layout="centered")
st.title("🏙️ Rise of Cultures — Import Ville")
st.caption("Convertit le CSV exporté depuis le jeu en fichier Excel pour l'optimiseur.")

# ── Catalogue dimensions bâtiments ──
DIMS = {
    'CultureSite_Large': (3,3), 'CultureSite_Moderate': (2,2),
    'CultureSite_Compact': (2,1), 'CultureSite_Small': (1,2),
    'CultureSite_Little': (1,1), 'CultureSite_Luxurious': (2,2),
    'Home_Small': (2,2), 'Home_Average': (2,2), 'Home_Large': (3,2), 'Home_Medium': (2,2),
    'Farm_Rural': (3,2), 'Farm_Domestic': (2,2), 'Farm_Pastoral': (3,3),
    'Barracks_Infantry': (3,3), 'Barracks_Ranged': (3,3),
    'Barracks_Cavalry': (3,3), 'Barracks_Siege': (5,6),
    'CityHall': (4,4), 'Evolving': (3,3), 'Collectable': (2,2), 'Wonder': (4,4),
    'Irrigation_Noria': (2,2), 'Irrigation_Large': (3,2), 'Irrigation_Medium': (2,2),
    'Irrigation_Small': (1,2), 'Irrigation_Oasis': (3,3),
    'Merchant_Average': (2,2), 'CamelFarm_Average': (3,3),
    'Workshop_Coffee': (2,2), 'Workshop_Incense': (2,2),
    'Workshop_OilLamp': (2,2), 'Workshop_Carpet': (2,2),
    'DEFAULT': (2,2),
}

CATS = {
    'CultureSite': 'Culturel', 'Barracks': 'Caserne',
    'Farm': 'Production', 'Home': 'Habitation',
    'CityHall': 'Neutre', 'Evolving': 'Producteur',
    'Collectable': 'Neutre', 'Wonder': 'Neutre',
    'Irrigation': 'Neutre', 'Merchant': 'Production',
    'CamelFarm': 'Production', 'Workshop': 'Production',
    'DEFAULT': 'Neutre',
}

CULTURE_VAL = {
    'CultureSite_Large': 600, 'CultureSite_Moderate': 350,
    'CultureSite_Compact': 200, 'CultureSite_Small': 150,
    'CultureSite_Little': 100, 'CultureSite_Luxurious': 700,
}
CULTURE_RANGE = {
    'CultureSite_Large': 3, 'CultureSite_Moderate': 2,
    'CultureSite_Compact': 1, 'CultureSite_Small': 2,
    'CultureSite_Little': 1, 'CultureSite_Luxurious': 2,
}

CAT_COLORS = {
    'Culturel': 'FFE699', 'Caserne': 'FF9999',
    'Production': 'C6EFCE', 'Habitation': 'BDD7EE',
    'Producteur': 'FFCCFF', 'Neutre': 'EFEFEF',
}

def get_key(name, mapping):
    for k in mapping:
        if k == 'DEFAULT': continue
        if k in name: return k
    return 'DEFAULT'

def short_name(name):
    parts = name.replace('Building_', '').split('_')
    if parts and parts[-1].isdigit(): parts = parts[:-1]
    return '_'.join(parts[-3:] if len(parts) > 3 else parts)

def build_excel(df, ville):
    rows_df = df[df['Ville'] == ville].copy()
    if rows_df.empty:
        return None

    n_rows = int(rows_df['Ligne'].max()) + 6
    n_cols = int(rows_df['Colonne'].max()) + 6

    wb = Workbook()
    thin = Side(style='thin', color='CCCCCC')
    brd  = Border(left=thin, right=thin, top=thin, bottom=thin)
    hdr_fill = PatternFill('solid', start_color='366092')
    hdr_font = Font(bold=True, color='FFFFFF', name='Arial', size=10)
    ctr = Alignment(horizontal='center', vertical='center')
    lft = Alignment(horizontal='left', vertical='center')

    # ── Onglet Terrain ──
    ws_t = wb.active
    ws_t.title = 'Terrain'

    ws_t.cell(1, 1).value = ''
    for c in range(n_cols):
        cell = ws_t.cell(1, c + 2)
        cell.value = c
        cell.font = Font(bold=True, name='Arial', size=9)
        cell.alignment = ctr

    # Grille des bâtiments placés
    grid = {}
    for _, b in rows_df.iterrows():
        bkey = get_key(b['Nom_complet'], DIMS)
        w, h = DIMS[bkey]
        cat  = CATS[get_key(b['Nom_complet'], CATS)]
        lbl  = short_name(b['Nom_complet'])
        for dr in range(h):
            for dc in range(w):
                grid[(int(b['Ligne']) + dr, int(b['Colonne']) + dc)] = {
                    'cat': cat,
                    'label': lbl if (dr == 0 and dc == 0) else '',
                }

    for r in range(n_rows):
        ws_t.cell(r + 2, 1).value = r
        ws_t.cell(r + 2, 1).font = Font(bold=True, name='Arial', size=9)
        ws_t.cell(r + 2, 1).alignment = ctr
        for c in range(n_cols):
            cell = ws_t.cell(r + 2, c + 2)
            if (r, c) in grid:
                info = grid[(r, c)]
                color = CAT_COLORS.get(info['cat'], 'EFEFEF')
                cell.fill = PatternFill('solid', start_color=color)
                if info['label']:
                    cell.value = info['label']
                    cell.font = Font(name='Arial', size=7, bold=True)
            else:
                cell.fill = PatternFill('solid', start_color='FFFFFF')
            cell.border = brd
            cell.alignment = ctr

    ws_t.column_dimensions['A'].width = 5
    for c in range(n_cols):
        ws_t.column_dimensions[get_column_letter(c + 2)].width = 4
    for r in range(n_rows + 1):
        ws_t.row_dimensions[r + 1].height = 15

    # ── Onglet Batiments ──
    ws_b = wb.create_sheet('Batiments')
    headers = ['Nom', 'Ligne', 'Colonne', 'Hauteur', 'Largeur',
               'Type', 'Culture', 'Rayonnement', 'Ere', 'Nom_complet']

    for i, h in enumerate(headers, 1):
        cell = ws_b.cell(1, i)
        cell.value = h
        cell.fill = hdr_fill
        cell.font = hdr_font
        cell.alignment = ctr
        cell.border = brd

    for row_idx, (_, b) in enumerate(rows_df.iterrows(), 2):
        bkey = get_key(b['Nom_complet'], DIMS)
        w, h = DIMS[bkey]
        cat  = CATS[get_key(b['Nom_complet'], CATS)]
        cult = CULTURE_VAL.get(bkey, 0)
        ray  = CULTURE_RANGE.get(bkey, 0)
        color = CAT_COLORS.get(cat, 'EFEFEF')

        row_data = [
            short_name(b['Nom_complet']),
            int(b['Ligne']),
            int(b['Colonne']),
            h, w, cat, cult, ray,
            str(b.get('Ere', '')),
            b['Nom_complet'],
        ]
        for col_idx, val in enumerate(row_data, 1):
            cell = ws_b.cell(row_idx, col_idx)
            cell.value = val
            cell.font = Font(name='Arial', size=10)
            cell.fill = PatternFill('solid', start_color=color)
            cell.alignment = lft if col_idx in (1, 6, 9, 10) else ctr
            cell.border = brd

    for i, w in enumerate([35, 8, 10, 8, 8, 12, 10, 12, 18, 55], 1):
        ws_b.column_dimensions[get_column_letter(i)].width = w

    buf = io.BytesIO()
    wb.save(buf)
    buf.seek(0)
    return buf


# ── Interface ──
uploaded = st.file_uploader(
    "Uploadez le fichier CSV exporté depuis le jeu",
    type=["csv"],
    help="Fichier généré par le userscript RoC dans Safari sur iPad"
)

if uploaded:
    try:
        df = pd.read_csv(uploaded)
        for col in df.columns:
            if df[col].dtype == object:
                df[col] = df[col].str.strip('"')

        st.success(f"✓ {len(df)} bâtiments chargés")

        villes = sorted(df['Ville'].unique().tolist())
        default_idx = villes.index('City_Capital') if 'City_Capital' in villes else 0
        ville_sel = st.selectbox("Choisissez la ville à exporter", villes, index=default_idx)

        df_ville = df[df['Ville'] == ville_sel]
        col1, col2, col3 = st.columns(3)
        col1.metric("Bâtiments", len(df_ville))
        col2.metric("Colonnes max", int(df_ville['Colonne'].max()))
        col3.metric("Lignes max", int(df_ville['Ligne'].max()))

        with st.expander("Aperçu des bâtiments"):
            st.dataframe(
                df_ville[['Nom_complet', 'Ligne', 'Colonne']].reset_index(drop=True),
                height=300
            )

        buf = build_excel(df, ville_sel)
        if buf:
            st.download_button(
                label="⬇️ Télécharger le fichier Excel",
                data=buf,
                file_name=f"ville_{ville_sel}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                type="primary",
            )
            st.info("Ce fichier Excel est prêt à être uploadé dans l'optimiseur de ville.")
        else:
            st.error("Aucun bâtiment trouvé pour cette ville.")

    except Exception as e:
        st.error(f"Erreur lors de la lecture du CSV : {e}")
        st.exception(e)

else:
    st.info("👆 Uploadez le CSV pour commencer.")
    with st.expander("Comment obtenir le CSV ?"):
        st.markdown("""
1. Installez **RoC_export_ville.user.js** dans Userscripts sur iPad
2. Ouvrez Safari → `https://u0.riseofcultures.com/`
3. Connectez-vous et attendez que votre ville soit chargée
4. Un panneau vert apparaît en bas à droite
5. Tapez **⬇ Télécharger CSV**
6. Uploadez ce fichier ici
        """)
