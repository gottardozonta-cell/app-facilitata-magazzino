import io
import re
from datetime import datetime
from urllib.parse import parse_qs, urlparse

import pandas as pd
import streamlit as st

# ---------------------------------------------------------------------------
# Pagina e stile Kids
# ---------------------------------------------------------------------------
st.set_page_config(
    page_title="Magazzino Cecap",
    page_icon="factory",
    layout="wide",
    initial_sidebar_state="expanded",
)

st.markdown(
    """
<style>
    .stApp { background-color: #FFF8E1; }
    div[data-testid="stMetric"] {
        background: #fff; border-radius: 12px; padding: 8px 12px;
        border: 2px solid #FFE082;
    }
    .card-ok { background:#C8E6C9; border-radius:16px; padding:14px; border:3px solid #66BB6A; }
    .card-low { background:#FFF9C4; border-radius:16px; padding:14px; border:3px solid #FFA726; }
    .card-empty { background:#FFCDD2; border-radius:16px; padding:14px; border:3px solid #EF5350; }
    .hero-ok { background:#66BB6A; color:white; padding:16px; border-radius:12px; text-align:center; font-size:1.4rem; font-weight:700; }
    .hero-low { background:#FFA726; color:white; padding:16px; border-radius:12px; text-align:center; font-size:1.4rem; font-weight:700; }
    .hero-empty { background:#EF5350; color:white; padding:16px; border-radius:12px; text-align:center; font-size:1.4rem; font-weight:700; }
    .qty-big { font-size: 2.8rem; font-weight: 800; line-height: 1.1; }
    div.stButton > button {
        border-radius: 14px; font-weight: 700; min-height: 3rem; border: 0;
    }
</style>
""",
    unsafe_allow_html=True,
)

# ---------------------------------------------------------------------------
# Costanti colonne
# ---------------------------------------------------------------------------
COLONNA_NOME = "Nome Prodotto"
COLONNA_CATEGORIA = "Categoria"
COLONNE_EXTRA = ["Codice", "Codice a Barre", "Prezzo", "Quantita", "Giacenza", "Scorta minima", "Foto"]
NOMI_FOTO = ["foto", "immagine", "image", "img", "picture", "url foto", "link foto"]


# ---------------------------------------------------------------------------
# Utility
# ---------------------------------------------------------------------------
def converti_link_drive(url):
    url = (url or "").strip().strip("<>")
    if not url:
        return None
    if not re.match(r"^https?://", url, re.IGNORECASE):
        url = f"https://{url}"
    parsed = urlparse(url)
    host = (parsed.hostname or "").lower()

    if host in {"docs.google.com", "www.docs.google.com"}:
        m = re.match(r"^/spreadsheets/(?:u/\d+/)?d/([a-zA-Z0-9_-]+)", parsed.path or "")
        if m:
            return f"https://docs.google.com/spreadsheets/d/{m.group(1)}/export?format=xlsx"

    if host in {"drive.google.com", "www.drive.google.com"}:
        m = re.match(r"^/file/d/([a-zA-Z0-9_-]+)", parsed.path or "")
        id_file = m.group(1) if m else parse_qs(parsed.query).get("id", [None])[0]
        if id_file and re.fullmatch(r"[a-zA-Z0-9_-]+", id_file):
            return f"https://drive.google.com/uc?export=download&confirm=t&id={id_file}"
    return None


def pulisci(testo):
    return re.sub(r"\s+", " ", str(testo)).strip().casefold()


def senza_accenti(s):
    for a, b in [("à", "a"), ("è", "e"), ("é", "e"), ("ì", "i"), ("ò", "o"), ("ù", "u")]:
        s = s.replace(a, b)
    return s


def trova_colonna(df, nome):
    obiettivo = senza_accenti(pulisci(nome))
    for c in df.columns:
        if senza_accenti(pulisci(c)) == obiettivo:
            return c
    return None


def normalizza_url_immagine(valore):
    if valore is None or (isinstance(valore, float) and pd.isna(valore)):
        return None
    testo = str(valore).strip()
    if not testo or testo.lower() in {"nan", "none", ""}:
        return None
    m = re.search(r'IMAGE\s*\(\s*["\']?(https?://[^"\')\s]+)', testo, re.I)
    if m:
        testo = m.group(1)
    if not re.match(r"^https?://", testo, re.I):
        return None
    parsed = urlparse(testo)
    host = (parsed.hostname or "").lower()
    if "drive.google.com" in host:
        m = re.match(r"^/file/d/([a-zA-Z0-9_-]+)", parsed.path or "")
        id_file = m.group(1) if m else parse_qs(parsed.query).get("id", [None])[0]
        if id_file:
            return f"https://drive.google.com/uc?export=view&id={id_file}"
    return testo


def normalizza_colonne(df):
    df = df.copy()
    df.columns = [str(c).strip() for c in df.columns]
    keep = [
        c
        for c in df.columns
        if not (c.startswith("Unnamed") or c == "" or c.lower() == "nan")
    ]
    df = df.loc[:, keep].dropna(axis=1, how="all")
    return df


def to_int(val, default=0):
    try:
        if val is None or (isinstance(val, float) and pd.isna(val)):
            return default
        s = str(val).strip().replace(",", ".")
        if not s or s.lower() == "nan":
            return default
        return max(0, int(float(s)))
    except Exception:
        return default


def status_of(qty, min_q):
    if qty <= 0:
        return "empty", "FINITO", "#EF5350"
    if qty <= min_q:
        return "low", "POCO", "#FFA726"
    return "ok", "OK", "#66BB6A"


@st.cache_data(show_spinner="Caricamento file...")
def leggi_bytes(contenuto, header_index):
    return pd.read_excel(io.BytesIO(contenuto), header=header_index, engine="openpyxl")


@st.cache_data(ttl=300, show_spinner="Download da Google...")
def leggi_url(url, header_index):
    return pd.read_excel(url, header=header_index, engine="openpyxl")


def dataframe_da_sorgente(file_excel, link_drive, header_row):
    header_index = int(header_row) - 1
    if file_excel is not None:
        return leggi_bytes(file_excel.getvalue(), header_index)
    link = converti_link_drive(link_drive)
    if not link:
        raise ValueError("Link Google non riconosciuto.")
    return leggi_url(link, header_index)


def prepara_df_lavoro(df_raw):
    """Converte il foglio in un DataFrame standard per l'app."""
    df = normalizza_colonne(df_raw)
    col_nome = trova_colonna(df, COLONNA_NOME) or trova_colonna(df, "Prodotto") or trova_colonna(df, "Nome")
    col_cat = trova_colonna(df, COLONNA_CATEGORIA)
    if col_nome is None:
        raise ValueError(
            "Manca la colonna 'Nome Prodotto'. Trovate: " + ", ".join(map(str, df.columns))
        )

    col_cod = trova_colonna(df, "Codice")
    col_bar = trova_colonna(df, "Codice a Barre") or trova_colonna(df, "Barcode")
    col_pre = trova_colonna(df, "Prezzo")
    col_fmt = trova_colonna(df, "Quantita") or trova_colonna(df, "Formato")
    col_gia = (
        trova_colonna(df, "Giacenza")
        or trova_colonna(df, "Stock")
        or trova_colonna(df, "Qta magazzino")
    )
    col_min = trova_colonna(df, "Scorta minima") or trova_colonna(df, "Scorta Min.")
    col_foto = None
    for n in NOMI_FOTO:
        col_foto = trova_colonna(df, n)
        if col_foto:
            break
    if col_foto is None:
        # cerca colonna con URL
        for c in df.columns:
            campione = df[c].dropna().astype(str).head(12)
            if not len(campione):
                continue
            n_url = sum(1 for v in campione if normalizza_url_immagine(v))
            if n_url >= max(1, len(campione) // 3):
                col_foto = c
                break

    rows = []
    for _, r in df.iterrows():
        nome = str(r.get(col_nome, "")).strip()
        if not nome or nome.lower() in {"nan", "totale", ""}:
            continue
        formato = str(r.get(col_fmt, "")).strip() if col_fmt else ""
        if formato.lower() == "nan":
            formato = ""
        giacenza = to_int(r.get(col_gia), 0) if col_gia else 0
        # se non c'e' Giacenza e Quantita e' numerica, usala come stock
        if not col_gia and col_fmt:
            raw = r.get(col_fmt)
            try:
                if raw is not None and str(raw).strip().replace(".", "").replace(",", "").isdigit():
                    giacenza = to_int(raw, 0)
                    formato = ""
            except Exception:
                pass
        scorta = to_int(r.get(col_min), 5) if col_min else 5
        foto = normalizza_url_immagine(r.get(col_foto)) if col_foto else None
        rows.append(
            {
                "id": len(rows) + 1,
                "Nome Prodotto": nome,
                "Categoria": (str(r.get(col_cat, "")).strip() if col_cat else "") or "Senza categoria",
                "Codice": str(r.get(col_cod, "")).strip() if col_cod else "",
                "Codice a Barre": str(r.get(col_bar, "")).strip() if col_bar else "",
                "Prezzo": str(r.get(col_pre, "")).strip() if col_pre else "",
                "Formato": formato,
                "Giacenza": giacenza,
                "Scorta minima": scorta,
                "Foto": foto,
            }
        )
    out = pd.DataFrame(rows)
    if out.empty:
        out = pd.DataFrame(
            columns=[
                "id", "Nome Prodotto", "Categoria", "Codice", "Codice a Barre",
                "Prezzo", "Formato", "Giacenza", "Scorta minima", "Foto",
            ]
        )
    # pulizia nan testuali
    for c in ["Codice", "Codice a Barre", "Prezzo", "Formato", "Categoria"]:
        out[c] = out[c].replace({"nan": "", "None": ""}).fillna("")
    return out


def init_state():
    defaults = {
        "page": "home",
        "df": None,
        "loaded": False,
        "filter_mode": "all",
        "selected_id": None,
        "msg": "",
    }
    for k, v in defaults.items():
        if k not in st.session_state:
            st.session_state[k] = v


def go(page):
    st.session_state.page = page
    st.session_state.msg = ""


def get_df():
    return st.session_state.df


def set_df(df):
    st.session_state.df = df.reset_index(drop=True)
    st.session_state.loaded = True


def counts(df):
    if df is None or df.empty:
        return 0, 0, 0, 0
    total = len(df)
    empty = int((df["Giacenza"] <= 0).sum())
    low = int(((df["Giacenza"] > 0) & (df["Giacenza"] <= df["Scorta minima"])).sum())
    ok = total - empty - low
    return total, ok, low, empty


def df_to_xlsx_bytes(df):
    export = df.copy()
    cols = [
        "Categoria", "Foto", "Nome Prodotto", "Codice a Barre", "Codice",
        "Formato", "Prezzo", "Giacenza", "Scorta minima",
    ]
    for c in cols:
        if c not in export.columns:
            export[c] = ""
    export = export[cols].rename(columns={"Formato": "Quantita"})
    buf = io.BytesIO()
    with pd.ExcelWriter(buf, engine="openpyxl") as writer:
        export.to_excel(writer, index=False, sheet_name="Magazzino")
    return buf.getvalue()


# ---------------------------------------------------------------------------
# Init
# ---------------------------------------------------------------------------
init_state()

# ---------------------------------------------------------------------------
# Sidebar: carica dati
# ---------------------------------------------------------------------------
with st.sidebar:
    st.header("Configurazione")
    link_drive = st.text_input(
        "Link Google Sheets / Drive",
        value=st.session_state.get("last_link", ""),
        placeholder="https://docs.google.com/spreadsheets/d/...",
    )
    file_excel = st.file_uploader("Oppure file Excel (.xlsx)", type=["xlsx"])
    header_row = st.number_input("Riga intestazioni", min_value=1, max_value=50, value=1, step=1)

    c1, c2 = st.columns(2)
    with c1:
        load_btn = st.button("Carica dati", type="primary", use_container_width=True)
    with c2:
        if st.button("Svuota cache", use_container_width=True):
            leggi_url.clear()
            leggi_bytes.clear()
            st.success("Cache svuotata")

    if load_btn:
        if not file_excel and not link_drive:
            st.error("Inserisci un link o carica un file.")
        else:
            try:
                raw = dataframe_da_sorgente(file_excel, link_drive, header_row)
                lavoro = prepara_df_lavoro(raw)
                set_df(lavoro)
                st.session_state.last_link = link_drive
                st.session_state.page = "home"
                st.success(f"Caricati {len(lavoro)} prodotti")
                st.rerun()
            except Exception as e:
                st.error(f"Errore: {e}")

    st.divider()
    if st.session_state.loaded and st.session_state.df is not None:
        st.caption(f"Prodotti in memoria: {len(st.session_state.df)}")
        xbytes = df_to_xlsx_bytes(st.session_state.df)
        st.download_button(
            "Scarica Excel aggiornato",
            data=xbytes,
            file_name=f"magazzino_{datetime.now():%Y%m%d_%H%M%S}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True,
        )
        st.info(
            "Le modifiche restano finche' tieni aperta la pagina. "
            "Scarica l'Excel per conservarle, poi rimpiazza il file su Drive se serve."
        )

# ---------------------------------------------------------------------------
# HOME
# ---------------------------------------------------------------------------
def page_home():
    st.title("Magazzino Cecap")
    st.caption("Interfaccia web stile Kids — consultabile e modificabile online")

    df = get_df()
    if df is None:
        st.info(
            "Benvenuto! Dal menu a sinistra carica il foglio Google oppure un file Excel, "
            "poi usa i bottoni qui sotto."
        )
        st.markdown(
            """
**Come condividere l'app**
1. Usa l'indirizzo Streamlit Cloud dell'app
2. Crea un collegamento sul desktop a quell'indirizzo
3. Invia lo stesso link ad altri (devono poter aprire il foglio se e' su Drive)
"""
        )
        return

    total, ok, low, empty = counts(df)
    if empty:
        st.markdown(f'<div class="hero-empty">{empty} prodotti FINITI</div>', unsafe_allow_html=True)
    elif low:
        st.markdown(f'<div class="hero-low">{low} prodotti in SCARSEZZA</div>', unsafe_allow_html=True)
    else:
        st.markdown('<div class="hero-ok">Tutto OK</div>', unsafe_allow_html=True)

    m1, m2, m3, m4 = st.columns(4)
    m1.metric("Totale", total)
    m2.metric("OK", ok)
    m3.metric("Poco", low)
    m4.metric("Finiti", empty)

    st.write("")
    b1, b2, b3 = st.columns(3)
    with b1:
        if st.button("AGGIUNGI PRODOTTO", use_container_width=True):
            go("add")
            st.rerun()
        st.markdown(
            "<div style='background:#66BB6A;color:white;text-align:center;padding:8px;border-radius:10px;margin-top:-8px'>Nuovo articolo</div>",
            unsafe_allow_html=True,
        )
    with b2:
        if st.button("GUARDA MAGAZZINO", use_container_width=True):
            go("view")
            st.rerun()
        st.markdown(
            "<div style='background:#42A5F5;color:white;text-align:center;padding:8px;border-radius:10px;margin-top:-8px'>Schede e filtri</div>",
            unsafe_allow_html=True,
        )
    with b3:
        if st.button("MUOVI QUANTITA", use_container_width=True):
            go("move")
            st.rerun()
        st.markdown(
            "<div style='background:#FFA726;color:white;text-align:center;padding:8px;border-radius:10px;margin-top:-8px'>+ / - giacenza</div>",
            unsafe_allow_html=True,
        )

    b4, b5, b6 = st.columns(3)
    with b4:
        if st.button("TABELLA MODIFICABILE", use_container_width=True):
            go("table")
            st.rerun()
    with b5:
        if st.button("ESPORTA / SALVA", use_container_width=True):
            go("export")
            st.rerun()
    with b6:
        if st.button("RICARICA DATI", use_container_width=True):
            leggi_url.clear()
            st.session_state.loaded = False
            st.session_state.df = None
            go("home")
            st.rerun()


# ---------------------------------------------------------------------------
# VIEW schede
# ---------------------------------------------------------------------------
def page_view():
    st.title("Guarda magazzino")
    if st.button("HOME"):
        go("home")
        st.rerun()

    df = get_df()
    if df is None or df.empty:
        st.warning("Nessun prodotto. Carica i dati dalla sidebar.")
        return

    f1, f2, f3, f4 = st.columns(4)
    with f1:
        if st.button("TUTTI", use_container_width=True):
            st.session_state.filter_mode = "all"
            st.rerun()
    with f2:
        if st.button("POCO", use_container_width=True):
            st.session_state.filter_mode = "low"
            st.rerun()
    with f3:
        if st.button("FINITI", use_container_width=True):
            st.session_state.filter_mode = "empty"
            st.rerun()
    with f4:
        cats = ["(tutte)"] + sorted(df["Categoria"].dropna().astype(str).unique().tolist())
        cat = st.selectbox("Categoria", cats, label_visibility="collapsed")

    mode = st.session_state.filter_mode
    view = df.copy()
    if mode == "low":
        view = view[(view["Giacenza"] > 0) & (view["Giacenza"] <= view["Scorta minima"])]
    elif mode == "empty":
        view = view[view["Giacenza"] <= 0]
    if cat != "(tutte)":
        view = view[view["Categoria"].astype(str) == cat]

    st.write(f"Prodotti mostrati: **{len(view)}**")
    if view.empty:
        st.info("Nessun prodotto con questo filtro.")
        return

    cols = st.columns(3)
    for i, row in enumerate(view.itertuples(index=False)):
        d = row._asdict() if hasattr(row, "_asdict") else dict(zip(view.columns, row))
        # compat tuple
        if not isinstance(d, dict):
            d = {c: getattr(row, c) for c in view.columns}
        # itertuples rinomina spazi
        def g(name, default=""):
            key = name.replace(" ", "_")
            if hasattr(row, key):
                return getattr(row, key)
            if name in view.columns:
                return view.loc[view["id"] == getattr(row, "id"), name].values[0]
            return default

        pid = g("id")
        nome = g("Nome Prodotto")
        catv = g("Categoria")
        qty = to_int(g("Giacenza"), 0)
        minq = to_int(g("Scorta minima"), 5)
        foto = g("Foto", None)
        stato, label, color = status_of(qty, minq)
        css = {"ok": "card-ok", "low": "card-low", "empty": "card-empty"}[stato]

        with cols[i % 3]:
            with st.container():
                st.markdown(f'<div class="{css}">', unsafe_allow_html=True)
                if foto and isinstance(foto, str) and foto.startswith("http"):
                    st.image(foto, use_container_width=True)
                st.markdown(f"**{nome}**")
                st.caption(str(catv))
                st.markdown(
                    f'<div class="qty-big" style="color:{color}">{qty}</div>',
                    unsafe_allow_html=True,
                )
                st.markdown(f"**{label}** · min {minq}")
                c_a, c_b = st.columns(2)
                with c_a:
                    if st.button("Dettaglio", key=f"det_{pid}", use_container_width=True):
                        st.session_state.selected_id = int(pid)
                        go("detail")
                        st.rerun()
                with c_b:
                    if st.button("Muovi", key=f"mv_{pid}", use_container_width=True):
                        st.session_state.selected_id = int(pid)
                        go("move_one")
                        st.rerun()
                st.markdown("</div>", unsafe_allow_html=True)
            st.write("")


# ---------------------------------------------------------------------------
# DETTAGLIO
# ---------------------------------------------------------------------------
def page_detail():
    st.title("Dettaglio prodotto")
    if st.button("Indietro"):
        go("view")
        st.rerun()

    df = get_df()
    pid = st.session_state.selected_id
    if df is None or pid is None:
        st.warning("Nessun prodotto selezionato.")
        return
    match = df[df["id"] == pid]
    if match.empty:
        st.error("Prodotto non trovato.")
        return
    r = match.iloc[0]
    qty = to_int(r["Giacenza"], 0)
    minq = to_int(r["Scorta minima"], 5)
    stato, label, color = status_of(qty, minq)
    css = {"ok": "card-ok", "low": "card-low", "empty": "card-empty"}[stato]

    st.markdown(f'<div class="{css}">', unsafe_allow_html=True)
    left, right = st.columns([1, 2])
    with left:
        if r.get("Foto") and str(r.get("Foto")).startswith("http"):
            st.image(str(r["Foto"]), use_container_width=True)
        else:
            st.info("Nessuna foto")
    with right:
        st.subheader(r["Nome Prodotto"])
        st.write(f"Categoria: **{r['Categoria']}**")
        st.markdown(f'<div class="qty-big" style="color:{color}">{qty}</div>', unsafe_allow_html=True)
        st.markdown(f"### {label}")
        st.write(f"Scorta minima: {minq}")
        st.write(f"Codice: {r.get('Codice','')} | Barcode: {r.get('Codice a Barre','')}")
        st.write(f"Prezzo: {r.get('Prezzo','')} | Formato: {r.get('Formato','')}")
    st.markdown("</div>", unsafe_allow_html=True)

    c1, c2, c3 = st.columns(3)
    with c1:
        if st.button("Modifica", use_container_width=True):
            go("edit")
            st.rerun()
    with c2:
        if st.button("Muovi quantita", use_container_width=True):
            go("move_one")
            st.rerun()
    with c3:
        if st.button("Elimina", use_container_width=True):
            nuovo = df[df["id"] != pid].copy()
            set_df(nuovo)
            st.success("Prodotto eliminato")
            go("view")
            st.rerun()


# ---------------------------------------------------------------------------
# MUOVI (lista o singolo)
# ---------------------------------------------------------------------------
def page_move_select():
    st.title("Muovi quantita")
    if st.button("HOME"):
        go("home")
        st.rerun()
    df = get_df()
    if df is None or df.empty:
        st.warning("Nessun prodotto.")
        return
    st.caption("Scegli un prodotto")
    for _, r in df.sort_values("Giacenza").iterrows():
        qty = to_int(r["Giacenza"], 0)
        minq = to_int(r["Scorta minima"], 5)
        _, label, _ = status_of(qty, minq)
        cols = st.columns([4, 1, 1])
        cols[0].write(f"**{r['Nome Prodotto']}** — {r['Categoria']} — giacenza **{qty}** ({label})")
        if cols[2].button("Scegli", key=f"pick_{r['id']}"):
            st.session_state.selected_id = int(r["id"])
            go("move_one")
            st.rerun()


def page_move_one():
    st.title("Aggiorna giacenza")
    if st.button("Indietro"):
        go("move")
        st.rerun()

    df = get_df()
    pid = st.session_state.selected_id
    if df is None or pid is None:
        st.warning("Nessun prodotto.")
        return
    idx = df.index[df["id"] == pid]
    if len(idx) == 0:
        st.error("Prodotto non trovato")
        return
    i = idx[0]
    r = df.loc[i]
    qty = to_int(r["Giacenza"], 0)
    minq = to_int(r["Scorta minima"], 5)
    stato, label, color = status_of(qty, minq)

    if r.get("Foto") and str(r.get("Foto")).startswith("http"):
        st.image(str(r["Foto"]), width=220)
    st.subheader(r["Nome Prodotto"])
    st.markdown(f'<div class="qty-big" style="color:{color}">{qty}</div>', unsafe_allow_html=True)
    st.markdown(f"**{label}** · scorta minima {minq}")

    def apply_delta(delta):
        d = get_df().copy()
        j = d.index[d["id"] == pid][0]
        d.at[j, "Giacenza"] = max(0, to_int(d.at[j, "Giacenza"], 0) + delta)
        set_df(d)

    row = st.columns(6)
    for col, delta, lab in zip(
        row,
        [-10, -5, -1, 1, 5, 10],
        ["-10", "-5", "-1", "+1", "+5", "+10"],
    ):
        with col:
            if st.button(lab, use_container_width=True, key=f"d_{delta}"):
                apply_delta(delta)
                st.rerun()

    st.divider()
    new_q = st.number_input("Imposta giacenza esatta", min_value=0, value=qty, step=1)
    if st.button("Imposta", type="primary"):
        d = get_df().copy()
        j = d.index[d["id"] == pid][0]
        d.at[j, "Giacenza"] = int(new_q)
        set_df(d)
        st.success("Giacenza aggiornata")
        st.rerun()


# ---------------------------------------------------------------------------
# ADD / EDIT
# ---------------------------------------------------------------------------
def page_add_edit(edit=False):
    st.title("Modifica prodotto" if edit else "Aggiungi prodotto")
    if st.button("HOME"):
        go("home")
        st.rerun()

    df = get_df()
    if df is None:
        # crea df vuoto
        set_df(
            pd.DataFrame(
                columns=[
                    "id", "Nome Prodotto", "Categoria", "Codice", "Codice a Barre",
                    "Prezzo", "Formato", "Giacenza", "Scorta minima", "Foto",
                ]
            )
        )
        df = get_df()

    initial = {
        "Nome Prodotto": "",
        "Categoria": "",
        "Codice": "",
        "Codice a Barre": "",
        "Prezzo": "",
        "Formato": "",
        "Giacenza": 0,
        "Scorta minima": 5,
        "Foto": "",
    }
    pid = None
    if edit:
        pid = st.session_state.selected_id
        match = df[df["id"] == pid]
        if match.empty:
            st.error("Prodotto non trovato")
            return
        r = match.iloc[0]
        for k in initial:
            initial[k] = r.get(k, initial[k])
            if pd.isna(initial[k]):
                initial[k] = "" if k not in ("Giacenza", "Scorta minima") else 0

    with st.form("form_prod"):
        nome = st.text_input("Nome prodotto", value=str(initial["Nome Prodotto"] or ""))
        cat = st.text_input("Categoria", value=str(initial["Categoria"] or ""))
        c1, c2 = st.columns(2)
        with c1:
            codice = st.text_input("Codice", value=str(initial["Codice"] or ""))
            prezzo = st.text_input("Prezzo", value=str(initial["Prezzo"] or ""))
            gia = st.number_input("Giacenza", min_value=0, value=to_int(initial["Giacenza"], 0))
        with c2:
            bar = st.text_input("Codice a barre", value=str(initial["Codice a Barre"] or ""))
            fmt = st.text_input("Formato / confezione", value=str(initial["Formato"] or ""))
            scorta = st.number_input("Scorta minima", min_value=0, value=to_int(initial["Scorta minima"], 5))
        foto = st.text_input("URL foto (opzionale)", value=str(initial["Foto"] or "") if initial["Foto"] else "")
        ok = st.form_submit_button("SALVA", type="primary", use_container_width=True)

    if ok:
        if not nome.strip():
            st.error("Il nome e' obbligatorio")
            return
        d = get_df().copy()
        foto_n = normalizza_url_immagine(foto) if foto else None
        if edit and pid is not None:
            j = d.index[d["id"] == pid][0]
            d.at[j, "Nome Prodotto"] = nome.strip()
            d.at[j, "Categoria"] = cat.strip() or "Senza categoria"
            d.at[j, "Codice"] = codice.strip()
            d.at[j, "Codice a Barre"] = bar.strip()
            d.at[j, "Prezzo"] = prezzo.strip()
            d.at[j, "Formato"] = fmt.strip()
            d.at[j, "Giacenza"] = int(gia)
            d.at[j, "Scorta minima"] = int(scorta)
            d.at[j, "Foto"] = foto_n
            set_df(d)
            st.success("Prodotto aggiornato")
            st.session_state.selected_id = pid
            go("detail")
            st.rerun()
        else:
            new_id = int(d["id"].max()) + 1 if len(d) and "id" in d.columns else 1
            nuova = {
                "id": new_id,
                "Nome Prodotto": nome.strip(),
                "Categoria": cat.strip() or "Senza categoria",
                "Codice": codice.strip(),
                "Codice a Barre": bar.strip(),
                "Prezzo": prezzo.strip(),
                "Formato": fmt.strip(),
                "Giacenza": int(gia),
                "Scorta minima": int(scorta),
                "Foto": foto_n,
            }
            d = pd.concat([d, pd.DataFrame([nuova])], ignore_index=True)
            set_df(d)
            st.success("Prodotto aggiunto")
            go("home")
            st.rerun()


# ---------------------------------------------------------------------------
# TABELLA editabile
# ---------------------------------------------------------------------------
def page_table():
    st.title("Tabella modificabile")
    if st.button("HOME"):
        go("home")
        st.rerun()
    df = get_df()
    if df is None:
        st.warning("Carica i dati prima.")
        return

    show_cols = [
        "Foto", "Nome Prodotto", "Categoria", "Giacenza", "Scorta minima",
        "Codice", "Codice a Barre", "Prezzo", "Formato",
    ]
    base = df.copy()
    for c in show_cols:
        if c not in base.columns:
            base[c] = None

    edited = st.data_editor(
        base[show_cols + ["id"]],
        use_container_width=True,
        hide_index=True,
        num_rows="dynamic",
        column_config={
            "Foto": st.column_config.ImageColumn("Foto", width="small"),
            "Giacenza": st.column_config.NumberColumn("Giacenza", min_value=0, step=1),
            "Scorta minima": st.column_config.NumberColumn("Scorta minima", min_value=0, step=1),
            "id": st.column_config.NumberColumn("id", disabled=True),
        },
        key="editor_main",
    )

    if st.button("Applica modifiche tabella", type="primary"):
        out = edited.copy()
        # riassegna id mancanti
        if out["id"].isna().any():
            max_id = int(df["id"].max()) if len(df) else 0
            for i in out.index:
                if pd.isna(out.at[i, "id"]):
                    max_id += 1
                    out.at[i, "id"] = max_id
        out["id"] = out["id"].astype(int)
        out["Giacenza"] = out["Giacenza"].apply(lambda x: to_int(x, 0))
        out["Scorta minima"] = out["Scorta minima"].apply(lambda x: to_int(x, 5))
        out["Foto"] = out["Foto"].apply(lambda x: normalizza_url_immagine(x) if pd.notna(x) else None)
        set_df(out)
        st.success("Modifiche applicate in memoria. Scarica l'Excel dalla sidebar per salvarle.")
        st.rerun()


# ---------------------------------------------------------------------------
# EXPORT
# ---------------------------------------------------------------------------
def page_export():
    st.title("Esporta / Salva")
    if st.button("HOME"):
        go("home")
        st.rerun()
    df = get_df()
    if df is None:
        st.warning("Nessun dato.")
        return

    st.write("Le modifiche sono in questa sessione del browser. Per conservarle:")
    xbytes = df_to_xlsx_bytes(df)
    st.download_button(
        "Scarica Excel aggiornato (.xlsx)",
        data=xbytes,
        file_name=f"magazzino_{datetime.now():%Y%m%d_%H%M%S}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        type="primary",
        use_container_width=True,
    )

    # HTML semplice
    total, ok, low, empty = counts(df)
    rows_html = []
    for _, r in df.sort_values("Nome Prodotto").iterrows():
        qty = to_int(r["Giacenza"], 0)
        minq = to_int(r["Scorta minima"], 5)
        _, lab, _ = status_of(qty, minq)
        cls = "ok" if lab == "OK" else ("low" if lab == "POCO" else "empty")
        foto = r.get("Foto") or ""
        img = f'<img src="{foto}" width="60">' if isinstance(foto, str) and foto.startswith("http") else ""
        nome = str(r["Nome Prodotto"]).replace("<", "&lt;")
        rows_html.append(
            f'<tr class="{cls}"><td>{img}</td><td>{nome}</td><td>{r["Categoria"]}</td>'
            f"<td>{qty}</td><td>{minq}</td><td>{lab}</td></tr>"
        )
    html = f"""<!DOCTYPE html><html lang="it"><head><meta charset="utf-8">
<title>Magazzino Cecap</title>
<style>
body{{font-family:Arial;margin:24px}} table{{border-collapse:collapse;width:100%}}
th,td{{border:1px solid #333;padding:8px;text-align:center}} th{{background:#42A5F5;color:#fff}}
tr.ok{{background:#C8E6C9}} tr.low{{background:#FFF9C4}} tr.empty{{background:#FFCDD2}}
</style></head><body>
<h1>Magazzino Cecap</h1>
<p>OK:{ok} Poco:{low} Finiti:{empty} Tot:{total} — {datetime.now():%d/%m/%Y %H:%M}</p>
<table><tr><th>Foto</th><th>Prodotto</th><th>Categoria</th><th>Giacenza</th><th>Min</th><th>Stato</th></tr>
{''.join(rows_html)}</table></body></html>"""
    st.download_button(
        "Scarica report HTML stampabile",
        data=html.encode("utf-8"),
        file_name=f"magazzino_{datetime.now():%Y%m%d_%H%M%S}.html",
        mime="text/html",
        use_container_width=True,
    )

    st.info(
        "Dopo il download puoi caricare l'Excel su Google Drive "
        "(sostituisci il file del foglio oppure importalo) per aggiornare la fonte."
    )


# ---------------------------------------------------------------------------
# Router
# ---------------------------------------------------------------------------
page = st.session_state.page
if page == "home":
    page_home()
elif page == "view":
    page_view()
elif page == "detail":
    page_detail()
elif page == "move":
    page_move_select()
elif page == "move_one":
    page_move_one()
elif page == "add":
    page_add_edit(edit=False)
elif page == "edit":
    page_add_edit(edit=True)
elif page == "table":
    page_table()
elif page == "export":
    page_export()
else:
    page_home()