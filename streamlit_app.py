import streamlit as st
import pandas as pd
import io
from openpyxl import Workbook
from openpyxl.styles import PatternFill, Font, Alignment, Border, Side
from openpyxl.utils import get_column_letter

st.set_page_config(page_title="RoC - Import Ville", page_icon="🏙️", layout="centered")
st.title("🏙️ Rise of Cultures — Import Ville")
st.caption("Convertit le CSV exporté depuis le jeu en fichier Excel pour l'optimiseur.")

# ── Catalogue dimensions (largeur, hauteur) ──
# Ordre important : les clés plus spécifiques d'abord
DIMS = {
    # Culture Sites
    'CultureSite_Large':     (3, 3),
    'CultureSite_Luxurious': (2, 2),
    'CultureSite_Moderate':  (2, 2),
    'CultureSite_Compact':   (2, 1),
    'CultureSite_Small':     (1, 2),
    'CultureSite_Little':    (1, 1),
    # Maisons
    'Home_Large':    (3, 2),
    'Home_Average':  (2, 2),
    'Home_Small':    (2, 2),
    'Home_Medium':   (2, 2),
    # Fermes — corrections
    'Farm_Domestic': (4, 4),
    'Farm_Premium':  (3, 3),
    'Farm_Rural':    (3, 2),
    'Farm_Pastoral': (3, 3),
    # Casernes
    'Barracks_Siege':         (5, 6),
    'Barracks_HeavyInfantry': (3, 3),
    'Barracks_Infantry':      (3, 3),
    'Barracks_Ranged':        (3, 3),
    'Barracks_Cavalry':       (3, 3),
    # Workshops
    'Workshop_Alchemist':    (3, 3),
    'Workshop_Glassblower':  (3, 3),
    'Workshop_Jeweler':      (3, 3),
    'Workshop_Carpenter':    (3, 3),
    'Workshop_Scribe':       (3, 3),
    'Workshop_SpiceMerchant':(3, 3),
    'Workshop_Coffee':       (2, 2),
    'Workshop_Incense':      (2, 2),
    'Workshop_OilLamp':      (2, 2),
    'Workshop_Carpet':       (2, 2),
    # Génériques
    'CityHall':    (4, 4),
    # Bâtiments évolutifs (dimensions spécifiques connues)
    'Evolving_DraculaCastle': (4, 3),
    'Evolving_CryptOfTheCount': (4, 3),
    'Evolving_MadScientistsLab': (4, 3),
    'Evolving_HotAirBalloon': (2, 2),
    'Evolving_SleighStation': (2, 2),
    'Evolving_WinterMarket': (3, 3),
    'Evolving_Bakery': (2, 2),
    'Evolving_RoyalTemple': (3, 3),
    'Evolving_AztecMainTemple': (4, 4),
    'Evolving_Aqueduct': (3, 3),
    'Evolving_MotherTree': (3, 3),
    'Evolving_TravellingYurt': (2, 2),
    'Evolving_YurtOfTheKhan': (3, 3),
    'Evolving_MongolianFeast': (3, 3),
    'Evolving_CelticArch': (3, 3),
    'Evolving_GrandSmithy': (3, 3),
    'Evolving_CelticBroch': (3, 3),
    'Evolving_GriotCourt': (3, 3),
    'Evolving_Madrasa': (4, 3),
    'Evolving_DovecoteTower': (2, 3),
    'Evolving_DrumTower': (3, 3),
    'Evolving_FountainOfYouth': (3, 3),
    'Evolving_PirateFortress': (4, 4),
    'Evolving_TreasureWreck': (3, 3),
    'Evolving_ElysianField': (3, 3),
    'Evolving_GreatGarden': (3, 3),
    'Evolving_Conservatory': (3, 3),
    'Evolving_AncientLibrary': (3, 3),
    'Evolving_Hydra': (3, 3),
    'Evolving_TrojanHorse': (3, 3),
    'Evolving_Exhibition': (4, 3),
    'Evolving_NaturalHistoryMuseum': (4, 3),
    'Evolving_ShrineOfReflection': (3, 3),
    'Evolving_ShirazWineHouse': (3, 3),
    'Evolving_MosaicBath': (3, 3),
    'Evolving': (3, 3),  # fallback générique
    'Collectable': (2, 2),
    'Wonder':      (4, 4),
    'Harbor':      (3, 3),
    # Arabia
    'Irrigation_Noria':  (2, 2),
    'Irrigation_Large':  (3, 2),
    'Irrigation_Medium': (2, 2),
    'Irrigation_Small':  (1, 2),
    'Irrigation_Oasis':  (3, 3),
    'Merchant_Average':  (2, 2),
    'CamelFarm_Average': (3, 3),
    # Fallback
    'DEFAULT': (2, 2),
}

CATS = {
    'CultureSite': 'Culturel',
    'Barracks':    'Caserne',
    'Farm':        'Production',
    'Home':        'Habitation',
    'CityHall':    'Neutre',
    'Evolving':    'Producteur',
    'Collectable': 'Neutre',
    'Wonder':      'Neutre',
    'Irrigation':  'Neutre',
    'Merchant':    'Production',
    'CamelFarm':   'Production',
    'Workshop':    'Production',
    'Harbor':      'Neutre',
    'DEFAULT':     'Neutre',
}

# Valeurs culture et rayonnement lues depuis le CSV (extraites du GameDesignResponse)
# Ces catalogues sont conservés comme fallback si le CSV ne contient pas ces colonnes
CULTURE_VAL = {}   # vide : on lit depuis le CSV
CULTURE_RANGE = {} # vide : on lit depuis le CSV

CAT_COLORS = {
    'Culturel':    'FFE699',
    'Caserne':     'FF9999',
    'Production':  'C6EFCE',
    'Habitation':  'BDD7EE',
    'Producteur':  'FFCCFF',
    'Neutre':      'EFEFEF',
}

def get_key(name, mapping):
    """Cherche la première clé qui apparaît dans le nom."""
    for k in mapping:
        if k == 'DEFAULT': continue
        if k in name: return k
    return 'DEFAULT'

def get_cat(name):
    """Retourne la catégorie d'un bâtiment depuis son nom."""
    return CATS.get(get_key(name, CATS), 'Neutre')

def short_name(name):
    parts = name.replace('Building_', '').split('_')
    if parts and parts[-1].isdigit():
        parts = parts[:-1]
    return '_'.join(parts[-3:] if len(parts) > 3 else parts)

def build_excel(df, ville):
    rows_df = df[df['Ville'] == ville].copy()
    if rows_df.empty:
        return None

    # Calcule les vraies limites du terrain
    # En tenant compte des dimensions des bâtiments
    max_r_used = 0
    max_c_used = 0
    for _, b in rows_df.iterrows():
        bkey = get_key(b['Nom_complet'], DIMS)
        w, h = DIMS[bkey]
        max_r_used = max(max_r_used, int(b['Ligne']) + h - 1)
        max_c_used = max(max_c_used, int(b['Colonne']) + w - 1)

    # Terrain avec marge de 3 cases
    n_rows = max_r_used + 4
    n_cols = max_c_used + 4

    wb = Workbook()

    # ── Styles ──
    thin      = Side(style='thin',   color='AAAAAA')
    thick     = Side(style='medium', color='333333')
    brd_cell  = Border(left=thin, right=thin, top=thin, bottom=thin)
    brd_inner = Border(left=thin, right=thin, top=thin, bottom=thin)

    ctr = Alignment(horizontal='center', vertical='center', wrap_text=True)
    lft = Alignment(horizontal='left',   vertical='center')

    hdr_fill = PatternFill('solid', start_color='366092')
    hdr_font = Font(bold=True, color='FFFFFF', name='Arial', size=10)
    free_fill   = PatternFill('solid', start_color='F5F5F5')
    border_fill = PatternFill('solid', start_color='D0D0D0')

    # ── Onglet Terrain ──
    ws_t = wb.active
    ws_t.title = 'Terrain'

    # En-têtes colonnes (ligne 1)
    ws_t.cell(1, 1).value = ''
    for c in range(n_cols):
        cell = ws_t.cell(1, c + 2)
        cell.value = c
        cell.font = Font(bold=True, name='Arial', size=8)
        cell.alignment = ctr
        cell.fill = PatternFill('solid', start_color='E8E8E8')

    # Numéros de lignes (colonne A)
    for r in range(n_rows):
        cell = ws_t.cell(r + 2, 1)
        cell.value = r
        cell.font = Font(bold=True, name='Arial', size=8)
        cell.alignment = ctr
        cell.fill = PatternFill('solid', start_color='E8E8E8')

    # Remplit toutes les cases en "libre" avec bordure fine
    for r in range(n_rows):
        for c in range(n_cols):
            cell = ws_t.cell(r + 2, c + 2)
            cell.fill = free_fill
            cell.border = brd_cell

    # Déduplique par position top-left pour éviter les chevauchements
    placed = {}
    for _, b in rows_df.iterrows():
        # Dimensions : depuis le CSV si disponibles, sinon fallback catalogue
        if 'Largeur' in b and 'Hauteur' in b and pd.notna(b['Largeur']) and pd.notna(b['Hauteur']) and int(b['Largeur']) > 0:
            w, h = int(b['Largeur']), int(b['Hauteur'])
        else:
            bkey = get_key(b['Nom_complet'], DIMS)
            w, h = DIMS[bkey]
        r0 = int(b['Ligne'])
        c0 = int(b['Colonne'])
        cat   = get_cat(b['Nom_complet'])
        lbl   = short_name(b['Nom_complet'])
        color = CAT_COLORS.get(cat, 'EFEFEF')
        if (r0, c0) not in placed:
            placed[(r0, c0)] = (lbl, color, w, h)

    from openpyxl.cell.cell import MergedCell as _MC
    for (r0, c0), (lbl, color, w, h) in placed.items():
        fill     = PatternFill('solid', start_color=color)
        excel_r1 = r0 + 2
        excel_c1 = c0 + 2
        excel_r2 = r0 + h + 1
        excel_c2 = c0 + w + 1
        # Vérifie que la cellule top-left n'est pas déjà dans une fusion
        if isinstance(ws_t.cell(excel_r1, excel_c1), _MC):
            continue
        if h > 1 or w > 1:
            try:
                ws_t.merge_cells(
                    start_row=excel_r1, start_column=excel_c1,
                    end_row=excel_r2,   end_column=excel_c2
                )
            except Exception:
                continue
        top_left = ws_t.cell(excel_r1, excel_c1)
        if isinstance(top_left, _MC):
            continue
        top_left.value     = lbl
        top_left.font      = Font(name='Arial', size=7, bold=True)
        top_left.alignment = ctr
        top_left.fill      = fill
        top_left.border    = Border(left=thick, right=thick, top=thick, bottom=thick)

    # Dimensions des colonnes et lignes
    ws_t.column_dimensions['A'].width = 4
    for c in range(n_cols):
        ws_t.column_dimensions[get_column_letter(c + 2)].width = 4
    for r in range(n_rows + 1):
        ws_t.row_dimensions[r + 1].height = 20

    # Fige les volets (ligne 1 + colonne A)
    ws_t.freeze_panes = 'B2'

    # ── Onglet Batiments ──
    ws_b = wb.create_sheet('Batiments')
    headers = ['Nom', 'Ligne', 'Colonne', 'Hauteur', 'Largeur',
               'Type', 'Culture', 'Rayonnement', 'Ere', 'Nom_complet']

    for i, h_txt in enumerate(headers, 1):
        cell = ws_b.cell(1, i)
        cell.value = h_txt
        cell.fill = hdr_fill
        cell.font = hdr_font
        cell.alignment = ctr
        cell.border = brd_cell

    for row_idx, (_, b) in enumerate(rows_df.iterrows(), 2):
        if 'Largeur' in b and 'Hauteur' in b and pd.notna(b['Largeur']) and pd.notna(b['Hauteur']) and int(b['Largeur']) > 0:
            w, h = int(b['Largeur']), int(b['Hauteur'])
        else:
            bkey = get_key(b['Nom_complet'], DIMS)
            w, h = DIMS[bkey]
        cat  = get_cat(b['Nom_complet'])
        # Culture et rayonnement : depuis le CSV si disponibles
        cult = int(b['Culture'])    if 'Culture'     in b.index and pd.notna(b.get('Culture'))    else 0
        ray  = int(b['Rayonnement']) if 'Rayonnement' in b.index and pd.notna(b.get('Rayonnement')) else 0
        color = CAT_COLORS.get(cat, 'EFEFEF')
        fill  = PatternFill('solid', start_color=color)

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
            cell.fill = fill
            cell.alignment = lft if col_idx in (1, 6, 9, 10) else ctr
            cell.border = brd_cell

    col_widths = [35, 8, 10, 8, 8, 12, 10, 12, 18, 55]
    for i, cw in enumerate(col_widths, 1):
        ws_b.column_dimensions[get_column_letter(i)].width = cw

    buf = io.BytesIO()
    wb.save(buf)
    buf.seek(0)
    return buf


# ── Interface Streamlit ──
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

        # Filtre les valeurs aberrantes (uint32 overflow de coordonnées négatives)
        df['Ligne']   = pd.to_numeric(df['Ligne'],   errors='coerce')
        df['Colonne'] = pd.to_numeric(df['Colonne'], errors='coerce')
        df = df[
            (df['Ligne']   >= 0) & (df['Ligne']   < 10000) &
            (df['Colonne'] >= 0) & (df['Colonne'] < 10000)
        ].copy()

        st.success(f"✓ {len(df)} bâtiments chargés")

        villes = sorted(df['Ville'].unique().tolist())
        default_idx = villes.index('City_Capital') if 'City_Capital' in villes else 0
        ville_sel = st.selectbox("Choisissez la ville à exporter", villes, index=default_idx)

        df_ville = df[df['Ville'] == ville_sel]
        col1, col2, col3 = st.columns(3)
        col1.metric("Bâtiments", len(df_ville))
        col2.metric("Colonnes max", int(df_ville['Colonne'].max()))
        col3.metric("Lignes max",   int(df_ville['Ligne'].max()))

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
                use_container_width=True,
            )
            st.info("Ce fichier Excel est prêt à être uploadé dans l'optimiseur de ville.")
        else:
            st.error("Aucun bâtiment trouvé pour cette ville.")

    except Exception as e:
        st.error(f"Erreur : {e}")
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
