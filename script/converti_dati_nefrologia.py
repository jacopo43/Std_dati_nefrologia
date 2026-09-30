"""Conversione e standardizzazione del file nefrologico.

Il programma legge un file Excel con un foglio di riepilogo e un foglio per
paziente. Il foglio di riepilogo non viene esportato: serve soltanto come
riferimento della struttura originale. A ciascun foglio-paziente viene assegnato
un identificativo progressivo anonimo.

Output: un file .xlsx con un unico foglio ``dati_standardizzati``.
"""

from __future__ import annotations

import argparse
import json
import re
import unicodedata
from dataclasses import dataclass
from pathlib import Path
from typing import Iterable

import numpy as np
import pandas as pd

# Configurazione -----
COLONNE_OUTPUT = [
    "id_paziente", "tipo_record", "indice_record",
    "eta_anni", "fumo", "alcool", "data_raccolta_terapia", "data_revisione_farmacologica",
    "punteggio_acb_pre", "numero_comorbidita_pre", "numero_farmaci_pre",
    "numero_interazioni_cd_pre", "numero_farmaci_inappropriati_beers_pre",
    "numero_farmaci_inappropriati_start_pre", "numero_farmaci_inappropriati_stopp_pre",
    "comorbidita_pre", "testo_comorbidita_originale",
    "data_ricognizione", "farmaco_pre", "codice_atc_pre", "dosaggio_pre",
    "unita_dosaggio_pre", "somministrazioni_giornaliere_pre", "via_somministrazione_pre",
    "testo_terapia_originale",
    "farmaco_interagente_1", "farmaco_interagente_2", "tipo_interazione",
    "motivo_interazione", "conseguenza_interazione_1", "conseguenza_interazione_2",
    "conseguenza_interazione_3", "consiglio_clinico",
]
COLONNE_CONTROLLO = [voce for colonna_output in COLONNE_OUTPUT for voce in (colonna_output, f"fonte_{colonna_output}")]
FRASI_NEGAZIONE = (
    "no", "non", "nessun", "nessuna", "nega", "assenza di", "senza",
    "senza evidenza di", "escluso", "negativo per",
)
FRASI_FAMILIARITA = ("familiarita per", "anamnesi familiare", "madre con", "padre con")
FRASI_INCERTEZZA = ("sospetto", "possibile", "probabile", "da escludere")
PRIORITA_INTERAZIONI = {
    "PROLUNGAMENTO QT": 10,
    "ALTERAZIONE ELETTROLITICA": 10,
    "NEFROTOSSICITÀ": 10,
    "EPATOTOSSICITÀ": 10,
    "IPERGLICEMIA": 20,
    "AUMENTO RISCHIO DI SANGUINAMENTO": 20,
    "EMORRAGIA E SANGUINAMENTO": 20,
    "DIMINUZIONE EFFICACIA": 40,
    "ALTERAZIONE ASSORBIMENTO": 45,
    "AUMENTO DELL'ESPOSIZIONE O DELL'EFFETTO DI UNO DEI FARMACI": 50,
    "AUMENTO DELLA TOSSICITÀ DI UNO DEI FARMACI": 90,
}


# Utilità generali -----
def pulisci_testo(valore):
    if valore is None or pd.isna(valore):
        return None
    testo = str(valore).replace("\x00", " ").strip()
    return testo or None

def normalizza_testo(valore) -> str:
    testo = pulisci_testo(valore)
    if not testo:
        return ""
    testo = unicodedata.normalize("NFKD", testo.lower())
    testo = "".join(c for c in testo if not unicodedata.combining(c))
    testo = testo.replace("’", "'")
    testo = re.sub(r"[^a-z0-9]+", " ", testo)
    return re.sub(r"\s+", " ", testo).strip()

def contiene_termine(testo: str, termine: str) -> bool:
    return bool(termine and re.search(rf"(?<!\w){re.escape(termine)}(?!\w)", testo))

def contesto_precedente(testo: str, inizio: int, massimo_parole: int = 7) -> str:
    return " ".join(testo[:inizio].strip().split()[-massimo_parole:])

def match_valido(testo: str, termine: str) -> bool:
    """Accetta un termine soltanto se non è negato, familiare o incerto."""
    trovato = False
    for occorrenza in re.finditer(rf"(?<!\w){re.escape(termine)}(?!\w)", testo):
        trovato = True
        contesto = contesto_precedente(testo, occorrenza.start())
        if any(contiene_termine(contesto, x) for x in FRASI_NEGAZIONE):
            continue
        if any(x in contesto for x in FRASI_FAMILIARITA):
            continue
        if any(contiene_termine(contesto, x) for x in FRASI_INCERTEZZA):
            continue
        return True
    return False if trovato else False

def dividi_termini(valore, separatori=r"[;,]") -> set[str]:
    testo = pulisci_testo(valore)
    if not testo:
        return set()
    return {normalizza_testo(x) for x in re.split(separatori, testo) if len(normalizza_testo(x)) >= 2}

def motore_excel(percorso: Path) -> str:
    estensione = percorso.suffix.lower()
    if estensione == ".xls":
        return "xlrd"
    if estensione == ".xlsx":
        return "openpyxl"
    raise ValueError("Sono supportati soltanto file .xls e .xlsx")

def prima_valida(serie: Iterable):
    for valore in serie:
        valore = pulisci_testo(valore)
        if valore is not None:
            return valore
    return None

def data_iso(valore):
    if valore is None or pd.isna(valore):
        return None
    data = pd.to_datetime(valore, dayfirst=True, errors="coerce")
    if pd.isna(data):
        return None
    return data.strftime("%Y-%m-%d")

def calcola_eta(data_nascita, data_riferimento):
    nascita = pd.to_datetime(data_nascita, dayfirst=True, errors="coerce")
    riferimento = pd.to_datetime(data_riferimento, dayfirst=True, errors="coerce")
    if pd.isna(nascita) or pd.isna(riferimento):
        return np.nan
    return int(riferimento.year - nascita.year - ((riferimento.month, riferimento.day) < (nascita.month, nascita.day)))

def colonna(df: pd.DataFrame, nome: str) -> pd.Series:
    obiettivo = normalizza_testo(nome)
    for c in df.columns:
        base = re.sub(r"__\d+$", "", str(c))
        if normalizza_testo(base) == obiettivo:
            return df[c]
    return pd.Series([np.nan] * len(df), index=df.index)

def colonne_univoche(colonne):
    conteggi = {}
    risultato = []
    for col in colonne:
        nome = pulisci_testo(col) or "senza_nome"
        conteggi[nome] = conteggi.get(nome, 0) + 1
        risultato.append(nome if conteggi[nome] == 1 else f"{nome}__{conteggi[nome]}")
    return risultato

def fonte_valori(nome_colonna: str, valori: Iterable, aggiunta: str | None = None) -> str | None:
    """Restituisce una descrizione leggibile dei valori sorgente non vuoti."""
    puliti = []
    for valore in valori:
        valore_pulito = pulisci_testo(valore)
        if valore_pulito is not None and valore_pulito not in puliti:
            puliti.append(valore_pulito)
    if not puliti and not aggiunta:
        return None
    parti = []
    if puliti:
        parti.append(f"{nome_colonna}: " + " || ".join(puliti))
    if aggiunta:
        parti.append(aggiunta)
    return " | ".join(parti)

def percentuale_mancanti_dataframe(df: pd.DataFrame) -> float:
    if df.empty or df.size == 0:
        return 0.0
    mancante = df.apply(lambda col: col.map(lambda x: pulisci_testo(x) is None))
    return float(mancante.to_numpy().mean() * 100)

def percentuale_mancanti_originale(pazienti: dict[str, pd.DataFrame]) -> float:
    mancanti = 0
    totale = 0
    for df in pazienti.values():
        # Rimuove soltanto righe/colonne completamente vuote del template,
        # mantenendo i mancanti reali nelle aree informative.
        utile = df.dropna(how="all").dropna(axis=1, how="all")
        if utile.empty:
            continue
        for valore in utile.to_numpy().ravel():
            totale += 1
            mancanti += int(pulisci_testo(valore) is None)
    return (mancanti / totale * 100) if totale else 0.0

def salva_excel_semplice(df: pd.DataFrame, percorso: Path, nome_foglio: str):
    """Salva un .xlsx senza tabelle/pivot/colori; solo intestazioni in grassetto."""
    from openpyxl.styles import Alignment, Border, Font, PatternFill

    percorso.parent.mkdir(parents=True, exist_ok=True)
    with pd.ExcelWriter(percorso, engine="openpyxl") as writer:
        df.to_excel(writer, sheet_name=nome_foglio, index=False)
        ws = writer.book[nome_foglio]
        for cella in ws[1]:
            # Rimuove anche lo stile predefinito che pandas applica alle intestazioni.
            cella.font = Font(name="Calibri", size=11, bold=True)
            cella.fill = PatternFill(fill_type=None)
            cella.border = Border()
            cella.alignment = Alignment()
            cella.number_format = "General"


# Dizionari -----
@dataclass(frozen=True)
class Farmaco:
    principio_attivo: str
    codice_atc: str | None
@dataclass(frozen=True)
class VoceDizionario:
    categoria: str
    termini: tuple[str, ...]
    radici: tuple[str, ...]
    ordine: int


class Dizionari:
    def __init__(self, cartella: Path):
        self.farmaci: dict[str, Farmaco] = {}
        self.indice_farmaci: dict[str, list[str]] = {}
        self.comorbidita: list[VoceDizionario] = []
        self.interazioni: list[VoceDizionario] = []
        self._carica_farmaci(cartella / "dizionario_farmaci.xlsx")
        self._carica_comorbidita(cartella / "dizionario_comorbidita.xlsx")
        self._carica_interazioni(cartella / "dizionario_interazioni.xlsx")

    def _carica_farmaci(self, percorso: Path):
        df = pd.read_excel(percorso, engine="openpyxl").fillna("")
        richieste = {"principio_attivo", "sinonimi"}
        if not richieste.issubset(df.columns):
            raise ValueError(f"Nel dizionario farmaci mancano: {sorted(richieste - set(df.columns))}")
        canonici = []
        alias = []
        for _, riga in df.iterrows():
            nome = pulisci_testo(riga.get("principio_attivo"))
            if not nome:
                continue
            info = Farmaco(nome, pulisci_testo(riga.get("codice_atc")))
            can = normalizza_testo(nome)
            if can:
                canonici.append((can, info))
            termini = dividi_termini(riga.get("sinonimi")) | dividi_termini(riga.get("radici_ricerca"))
            alias.extend((termine, info) for termine in termini)
        for termine, info in alias:
            self.farmaci.setdefault(termine, info)
        for termine, info in canonici:
            self.farmaci[termine] = info
        for alias_norm in self.farmaci:
            for token in set(alias_norm.split()):
                if len(token) >= 3:
                    self.indice_farmaci.setdefault(token, []).append(alias_norm)
        for token in self.indice_farmaci:
            self.indice_farmaci[token].sort(key=len, reverse=True)

    def _carica_comorbidita(self, percorso: Path):
        df = pd.read_excel(percorso, engine="openpyxl").fillna("")
        for ordine, (_, riga) in enumerate(df.iterrows()):
            categoria = pulisci_testo(riga.get("categoria_standard"))
            if not categoria:
                continue
            uso = normalizza_testo(riga.get("uso_automatico"))
            consentito = uso.startswith("si")
            bloccati = dividi_termini(riga.get("sinonimi_da_revisionare")) | dividi_termini(riga.get("sinonimi_non_automatici"))
            termini = (dividi_termini(riga.get("sinonimi")) - bloccati) if consentito else set()
            radici = (dividi_termini(riga.get("radici_ricerca")) - bloccati) if consentito else set()
            self.comorbidita.append(VoceDizionario(categoria, tuple(sorted(termini, key=len, reverse=True)), tuple(sorted(radici, key=len, reverse=True)), ordine))

    def _carica_interazioni(self, percorso: Path):
        df = pd.read_excel(percorso, engine="openpyxl").fillna("")
        for ordine, (_, riga) in enumerate(df.iterrows()):
            categoria = pulisci_testo(riga.get("conseguenza_standard"))
            if not categoria:
                continue
            uso = normalizza_testo(riga.get("uso_automatico"))
            bloccati = dividi_termini(riga.get("termini_da_revisionare"), r",") | dividi_termini(riga.get("termini_non_automatici"), r",")
            termini = dividi_termini(riga.get("termini_associati"), r",") - bloccati if uso.startswith("si") else set()
            self.interazioni.append(VoceDizionario(categoria, tuple(sorted(termini, key=len, reverse=True)), (), ordine))

    def trova_farmaco(self, testo) -> Farmaco | None:
        norm = normalizza_testo(testo)
        if not norm:
            return None
        if norm in self.farmaci:
            return self.farmaci[norm]
        candidati = set()
        for token in norm.split():
            candidati.update(self.indice_farmaci.get(token, ()))
        for alias in sorted(candidati, key=len, reverse=True):
            if contiene_termine(norm, alias):
                return self.farmaci[alias]
        return None

    @staticmethod
    def _corrisponde(norm: str, voce: VoceDizionario) -> bool:
        for termine in voce.termini:
            if contiene_termine(norm, termine) and match_valido(norm, termine):
                return True
        for radice in voce.radici:
            # Le radici possono essere prefissi (es. ``ipertens`` -> ``ipertensione``),
            # ma devono iniziare all'inizio di una parola. In questo modo ``asm``
            # riconosce ``asma``/``asmatico`` ma NON ``spasmo``.
            modello_radice = rf"(?<!\w){re.escape(radice)}"
            for occ in re.finditer(modello_radice, norm):
                contesto = contesto_precedente(norm, occ.start())
                if any(contiene_termine(contesto, x) for x in FRASI_NEGAZIONE + FRASI_INCERTEZZA):
                    continue
                if any(x in contesto for x in FRASI_FAMILIARITA):
                    continue
                return True
        return False

    def trova_comorbidita(self, testo) -> list[str]:
        norm = normalizza_testo(testo)
        if not norm:
            return []
        trovate = [v.categoria for v in self.comorbidita if self._corrisponde(norm, v)]
        # Le infezioni vengono riunite nella categoria generale prevista dal dizionario.
        infezioni = [x for x in trovate if x.startswith("INFEZIONE") or "MALATTIA INFETTIVA" in x]
        if infezioni:
            preferita = "INFEZIONE/MALATTIA INFETTIVA" if "INFEZIONE/MALATTIA INFETTIVA" in trovate else infezioni[0]
            trovate = [x for x in trovate if x not in infezioni] + [preferita]
        # Categorie generiche che non devono prevalere su categorie specifiche.
        if "FIBRILLAZIONE ATRIALE" in trovate:
            trovate = [x for x in trovate if x != "ARITMIA"]
        if "INSUFFICIENZA CARDIACA" in trovate:
            trovate = [x for x in trovate if x != "ALTRE CARDIOPATIE"]
        risultato = []
        for x in trovate:
            if x != "ALTRO" and x not in risultato:
                risultato.append(x)
        return risultato

    def trova_conseguenze(self, testo, massimo=3) -> list[str]:
        norm = normalizza_testo(testo)
        if not norm:
            return []
        voci = [v for v in self.interazioni if self._corrisponde(norm, v)]
        nomi = {v.categoria for v in voci}
        if re.search(r"\brischio\b.{0,30}\b(sanguin|emorrag)", norm):
            nomi.add("AUMENTO RISCHIO DI SANGUINAMENTO")
            nomi.discard("EMORRAGIA E SANGUINAMENTO")
        if "EPATOTOSSICITÀ" in nomi:
            nomi.discard("AUMENTO TRANSAMINASI")
        specifiche = {x for x in nomi if PRIORITA_INTERAZIONI.get(x, 50) <= 30}
        if specifiche:
            nomi.discard("AUMENTO DELL'ESPOSIZIONE O DELL'EFFETTO DI UNO DEI FARMACI")
            nomi.discard("AUMENTO DELLA TOSSICITÀ DI UNO DEI FARMACI")
        ordine_originale = {v.categoria: v.ordine for v in voci}
        return sorted(nomi, key=lambda x: (PRIORITA_INTERAZIONI.get(x, 50), ordine_originale.get(x, 9999)))[:massimo]


# Terapia -----
MODELLO_DOSAGGIO = re.compile(r"(?P<valore>\d+(?:[.,]\d+)?)\s*(?P<unita>kg|mg|mcg|ug|µg|g|ui|iu)(?=$|[^a-zA-Zµ])", re.I)
MODELLO_VIA = re.compile(r"\b(per\s+os|orale|os|ev|iv|im|sc|sottocute|sottocutanea|inalatoria|topica|transdermica)\b", re.I)

def somministrazioni_giornaliere(testo: str) -> float:
    grezzo = str(testo).lower().replace("’", "'")
    norm = normalizza_testo(testo)
    for modello in [r"(\d+(?:[.,]\d+)?)\s*volte\s*(?:al|a|il)?\s*(?:giorno|die)", r"(\d+(?:[.,]\d+)?)\s*(?:x|per)\s*(?:die|giorno)"]:
        m = re.search(modello, grezzo, flags=re.I)
        if m:
            return float(m.group(1).replace(",", "."))
    m = re.search(r"ogni\s*(\d+(?:[.,]\d+)?)\s*ore?", grezzo, flags=re.I)
    if m:
        ore = float(m.group(1).replace(",", "."))
        return round(24 / ore, 6) if ore > 0 else np.nan
    if re.search(r"(?:/|ogni\s*)24\s*ore?", grezzo, flags=re.I):
        return 1.0
    orari = re.findall(r"\bore?\s*(\d{1,2})(?::([0-5]\d))?", grezzo, flags=re.I)
    if orari:
        return float(len({(int(ora), minuto or "00") for ora, minuto in orari}))
    if any(x in norm for x in ["tre volte al giorno", "ter die", "tid", "mattina pomeriggio sera"]):
        return 3.0
    if any(x in norm for x in ["due volte al giorno", "bis die", "bid", "mattina e sera", "mattino e sera"]):
        return 2.0
    if any(x in norm for x in ["una volta al giorno", "1 die", "qd", "al mattino", "alla sera", "la sera", "la mattina"]):
        return 1.0
    if re.search(r"\b\d+(?:[.,]\d+)?\s*(?:cp|cpr|compress[ae]|capsul[ae]|cerott[oi])\b", grezzo, flags=re.I):
        return 1.0
    return np.nan

def analizza_terapia(testo, dizionari: Dizionari):
    testo = pulisci_testo(testo)
    if not testo:
        return None
    farmaco = dizionari.trova_farmaco(testo)
    dose = MODELLO_DOSAGGIO.search(testo)
    via = MODELLO_VIA.search(testo)
    return {
        "farmaco_pre": farmaco.principio_attivo if farmaco else np.nan,
        "codice_atc_pre": farmaco.codice_atc if farmaco else np.nan,
        "dosaggio_pre": float(dose.group("valore").replace(",", ".")) if dose else np.nan,
        "unita_dosaggio_pre": dose.group("unita").lower() if dose else np.nan,
        "somministrazioni_giornaliere_pre": somministrazioni_giornaliere(testo),
        "via_somministrazione_pre": via.group().strip() if via else np.nan,
        "testo_terapia_originale": testo,
    }


# Lettura e trasformazione -----
def leggi_fogli_paziente(percorso: Path) -> dict[str, pd.DataFrame]:
    motore = motore_excel(percorso)
    libro = pd.ExcelFile(percorso, engine=motore)
    risultato = {}
    for foglio in libro.sheet_names:
        if normalizza_testo(foglio) in {"riassunto", "summary", "riepilogo"}:
            continue
        df = pd.read_excel(percorso, sheet_name=foglio, header=1, engine=motore)
        df.columns = colonne_univoche(df.columns)
        risultato[foglio] = df
    if not risultato:
        raise ValueError("Non sono stati trovati fogli-paziente nel file sorgente.")
    return risultato

def conta_farmaci_in_testo(serie: pd.Series, dizionari: Dizionari) -> int:
    trovati = set()
    for valore in serie:
        info = dizionari.trova_farmaco(valore)
        if info:
            trovati.add(normalizza_testo(info.principio_attivo))
    return len(trovati)

def crea_base_riepilogo(df: pd.DataFrame, dizionari: Dizionari) -> tuple[dict, dict]:
    nascita = prima_valida(colonna(df, "DDN"))
    data_visita = prima_valida(colonna(df, "DATA VISITA AMB"))
    data_revisione = prima_valida(colonna(df, "DATA VALUTAZIONE"))
    riferimento_eta = data_revisione or data_visita

    testi_mh = [x for x in colonna(df, "MEDICAL HISTORY") if pulisci_testo(x)]
    comorbidita = []
    for testo in testi_mh:
        for categoria in dizionari.trova_comorbidita(testo):
            if categoria not in comorbidita and categoria not in {"FUMO", "ALCOL", "ALCOOL"}:
                comorbidita.append(categoria)
    tutte_categorie = [c for testo in testi_mh for c in dizionari.trova_comorbidita(testo)]

    gravita_originali = [x for x in colonna(df, "GRAVITA' INTERAZIONI") if pulisci_testo(x)]
    gravita = [normalizza_testo(x).upper() for x in gravita_originali]
    numero_cd = sum(x in {"C", "D"} for x in gravita)

    terapia = colonna(df, "TERAPIA IN CORSO DI RICOVERO")
    terapia_non_vuota = [x for x in terapia if pulisci_testo(x)]
    beers = [x for x in colonna(df, "CRITERI DI BEERS") if pulisci_testo(x)]
    start = [x for x in colonna(df, "CRITERI START") if pulisci_testo(x)]
    stopp = [x for x in colonna(df, "CRITERI STOPP") if pulisci_testo(x)]
    acb = prima_valida(colonna(df, "ACB Score"))

    riepilogo = {
        "eta_anni": calcola_eta(nascita, riferimento_eta),
        "fumo": int("FUMO" in tutte_categorie),
        "alcool": int(any(x in tutte_categorie for x in {"ALCOL", "ALCOOL"})),
        "data_raccolta_terapia": data_iso(data_visita),
        "data_revisione_farmacologica": data_iso(data_revisione),
        "punteggio_acb_pre": pd.to_numeric(acb, errors="coerce"),
        "numero_comorbidita_pre": len(comorbidita),
        "numero_farmaci_pre": len(terapia_non_vuota),
        "numero_interazioni_cd_pre": numero_cd,
        "numero_farmaci_inappropriati_beers_pre": conta_farmaci_in_testo(pd.Series(beers), dizionari),
        "numero_farmaci_inappropriati_start_pre": conta_farmaci_in_testo(pd.Series(start), dizionari),
        "numero_farmaci_inappropriati_stopp_pre": conta_farmaci_in_testo(pd.Series(stopp), dizionari),
    }

    riferimento_nome = "DATA VALUTAZIONE" if data_revisione else "DATA VISITA AMB"
    fonti = {
        "eta_anni": fonte_valori("DDN", [nascita], fonte_valori(riferimento_nome, [riferimento_eta])),
        "fumo": fonte_valori("MEDICAL HISTORY", testi_mh, "dizionario_comorbidita.xlsx"),
        "alcool": fonte_valori("MEDICAL HISTORY", testi_mh, "dizionario_comorbidita.xlsx"),
        "data_raccolta_terapia": fonte_valori("DATA VISITA AMB", [data_visita]),
        "data_revisione_farmacologica": fonte_valori("DATA VALUTAZIONE", [data_revisione]),
        "punteggio_acb_pre": fonte_valori("ACB Score", [acb]),
        "numero_comorbidita_pre": fonte_valori("MEDICAL HISTORY", testi_mh, "dizionario_comorbidita.xlsx"),
        "numero_farmaci_pre": fonte_valori("TERAPIA IN CORSO DI RICOVERO", terapia_non_vuota),
        "numero_interazioni_cd_pre": fonte_valori("GRAVITA' INTERAZIONI", gravita_originali),
        "numero_farmaci_inappropriati_beers_pre": fonte_valori("CRITERI DI BEERS", beers, "dizionario_farmaci.xlsx"),
        "numero_farmaci_inappropriati_start_pre": fonte_valori("CRITERI START", start, "dizionario_farmaci.xlsx"),
        "numero_farmaci_inappropriati_stopp_pre": fonte_valori("CRITERI STOPP", stopp, "dizionario_farmaci.xlsx"),
    }
    return riepilogo, fonti

def righe_comorbidita(df: pd.DataFrame, dizionari: Dizionari):
    viste = set()
    righe = []
    for testo in colonna(df, "MEDICAL HISTORY"):
        testo_pulito = pulisci_testo(testo)
        if not testo_pulito:
            continue
        for categoria in dizionari.trova_comorbidita(testo_pulito):
            if categoria in {"FUMO", "ALCOL", "ALCOOL"} or categoria in viste:
                continue
            viste.add(categoria)
            specifici = {"comorbidita_pre": categoria, "testo_comorbidita_originale": testo_pulito}
            fonte = f"MEDICAL HISTORY: {testo_pulito} | dizionario_comorbidita.xlsx"
            fonti = {"comorbidita_pre": fonte, "testo_comorbidita_originale": f"MEDICAL HISTORY: {testo_pulito}"}
            righe.append((specifici, fonti))
    return righe

def righe_terapia(df: pd.DataFrame, dizionari: Dizionari, data_ricognizione, fonte_data_ricognizione):
    righe = []
    for testo in colonna(df, "TERAPIA IN CORSO DI RICOVERO"):
        analisi = analizza_terapia(testo, dizionari)
        if analisi:
            analisi["data_ricognizione"] = data_ricognizione
            testo_originale = analisi["testo_terapia_originale"]
            base = f"TERAPIA IN CORSO DI RICOVERO: {testo_originale}"
            fonti = {
                "data_ricognizione": fonte_data_ricognizione,
                "farmaco_pre": base + " | dizionario_farmaci.xlsx",
                "codice_atc_pre": base + " | dizionario_farmaci.xlsx",
                "dosaggio_pre": base,
                "unita_dosaggio_pre": base,
                "somministrazioni_giornaliere_pre": base,
                "via_somministrazione_pre": base,
                "testo_terapia_originale": base,
            }
            righe.append((analisi, fonti))
    return righe

def righe_interazioni(df: pd.DataFrame, dizionari: Dizionari):
    a_originale = colonna(df, 'FARMACO interagente "A"')
    b_originale = colonna(df, 'FARMACO interagente "B"')
    a = a_originale.ffill()
    b = b_originale.ffill()
    gravita = colonna(df, "GRAVITA' INTERAZIONI")
    effetti = colonna(df, "POSSIBILI EFFETTI INTERAZIONE")
    comportamento = colonna(df, "COMPORTAMENTO CLINICO")

    righe = []
    for i in df.index:
        fa, fb = pulisci_testo(a.loc[i]), pulisci_testo(b.loc[i])
        tipo = pulisci_testo(gravita.loc[i])
        motivo = pulisci_testo(effetti.loc[i])
        consiglio = pulisci_testo(comportamento.loc[i])
        if not any([fa, fb, tipo, motivo, consiglio]):
            continue
        conseguenze = dizionari.trova_conseguenze(motivo, massimo=3)
        riga = {
            "farmaco_interagente_1": fa,
            "farmaco_interagente_2": fb,
            "tipo_interazione": tipo,
            "motivo_interazione": motivo,
            "conseguenza_interazione_1": conseguenze[0] if len(conseguenze) > 0 else np.nan,
            "conseguenza_interazione_2": conseguenze[1] if len(conseguenze) > 1 else np.nan,
            "conseguenza_interazione_3": conseguenze[2] if len(conseguenze) > 2 else np.nan,
            "consiglio_clinico": consiglio,
        }
        fonte_a = fonte_valori('FARMACO interagente "A"', [a_originale.loc[i]])
        if fonte_a is None and fa:
            fonte_a = f'FARMACO interagente "A": {fa} (ereditato dalla riga precedente per cella unita/vuota)'
        fonte_b = fonte_valori('FARMACO interagente "B"', [b_originale.loc[i]])
        if fonte_b is None and fb:
            fonte_b = f'FARMACO interagente "B": {fb} (ereditato dalla riga precedente per cella unita/vuota)'
        fonte_effetti = fonte_valori("POSSIBILI EFFETTI INTERAZIONE", [motivo])
        fonti = {
            "farmaco_interagente_1": fonte_a,
            "farmaco_interagente_2": fonte_b,
            "tipo_interazione": fonte_valori("GRAVITA' INTERAZIONI", [tipo]),
            "motivo_interazione": fonte_effetti,
            "conseguenza_interazione_1": (fonte_effetti + " | dizionario_interazioni.xlsx") if fonte_effetti and len(conseguenze) > 0 else None,
            "conseguenza_interazione_2": (fonte_effetti + " | dizionario_interazioni.xlsx") if fonte_effetti and len(conseguenze) > 1 else None,
            "conseguenza_interazione_3": (fonte_effetti + " | dizionario_interazioni.xlsx") if fonte_effetti and len(conseguenze) > 2 else None,
            "consiglio_clinico": fonte_valori("COMPORTAMENTO CLINICO", [consiglio]),
        }
        righe.append((riga, fonti))

    coppie_con_dettagli = {
        (normalizza_testo(r[0]["farmaco_interagente_1"]), normalizza_testo(r[0]["farmaco_interagente_2"]))
        for r in righe if any(pulisci_testo(r[0][k]) for k in ["tipo_interazione", "motivo_interazione", "consiglio_clinico"])
    }
    pulite = []
    visti = set()
    for riga, fonti in righe:
        coppia = (normalizza_testo(riga["farmaco_interagente_1"]), normalizza_testo(riga["farmaco_interagente_2"]))
        solo_coppia = not any(pulisci_testo(riga[k]) for k in ["tipo_interazione", "motivo_interazione", "consiglio_clinico"])
        if solo_coppia and coppia in coppie_con_dettagli:
            continue
        chiave = tuple("" if pd.isna(riga.get(c)) else str(riga.get(c)) for c in riga)
        if chiave not in visti:
            visti.add(chiave)
            pulite.append((riga, fonti))
    return pulite

def aggiungi_record(lista, lista_controllo, id_paziente, tipo_record, indice,
                     riepilogo, fonti_riepilogo, specifici=None, fonti_specifici=None):
    riga = {c: np.nan for c in COLONNE_OUTPUT}
    riga.update(riepilogo)
    riga.update({"id_paziente": id_paziente, "tipo_record": tipo_record, "indice_record": indice})
    if specifici:
        riga.update(specifici)
    lista.append(riga)

    fonti = {c: None for c in COLONNE_OUTPUT}
    fonti.update(fonti_riepilogo)
    fonti.update({
        "id_paziente": "ordine progressivo dei fogli-paziente; dati anagrafici non esportati",
        "tipo_record": "tipo di informazione derivato dalla struttura dell'output",
        "indice_record": "ordine progressivo del record nel dominio del paziente",
    })
    if fonti_specifici:
        fonti.update(fonti_specifici)

    riga_controllo = {}
    for c in COLONNE_OUTPUT:
        riga_controllo[c] = riga.get(c, np.nan)
        riga_controllo[f"fonte_{c}"] = fonti.get(c)
    lista_controllo.append(riga_controllo)

def converti(percorso_input: Path, cartella_risorse: Path, percorso_output: Path,
             percorso_output_controllo: Path | None = None):
    dizionari = Dizionari(cartella_risorse)
    pazienti = leggi_fogli_paziente(percorso_input)
    tutte_righe = []
    tutte_righe_controllo = []

    for id_paziente, (_, df) in enumerate(pazienti.items(), start=1):
        riepilogo, fonti_riepilogo = crea_base_riepilogo(df, dizionari)
        aggiungi_record(tutte_righe, tutte_righe_controllo, id_paziente, "riepilogo", 1,
                        riepilogo, fonti_riepilogo)

        for indice, (riga, fonti) in enumerate(righe_comorbidita(df, dizionari), start=1):
            aggiungi_record(tutte_righe, tutte_righe_controllo, id_paziente, "comorbidita", indice,
                            riepilogo, fonti_riepilogo, riga, fonti)

        fonte_data = fonti_riepilogo.get("data_revisione_farmacologica")
        for indice, (riga, fonti) in enumerate(
            righe_terapia(df, dizionari, riepilogo["data_revisione_farmacologica"], fonte_data), start=1
        ):
            aggiungi_record(tutte_righe, tutte_righe_controllo, id_paziente, "terapia", indice,
                            riepilogo, fonti_riepilogo, riga, fonti)

        for indice, (riga, fonti) in enumerate(righe_interazioni(df, dizionari), start=1):
            aggiungi_record(tutte_righe, tutte_righe_controllo, id_paziente, "interazione", indice,
                            riepilogo, fonti_riepilogo, riga, fonti)

    output = pd.DataFrame(tutte_righe, columns=COLONNE_OUTPUT)
    controllo = pd.DataFrame(tutte_righe_controllo, columns=COLONNE_CONTROLLO)

    if percorso_output_controllo is None:
        percorso_output_controllo = percorso_output.with_name(percorso_output.stem + "_controllo_fonti.xlsx")

    salva_excel_semplice(output, percorso_output, "dati_standardizzati")
    salva_excel_semplice(controllo, percorso_output_controllo, "dati_con_fonti")

    missing_originale = percentuale_mancanti_originale(pazienti)
    missing_convertito = percentuale_mancanti_dataframe(output)
    print(
        f"Missingness complessiva (celle vuote): originale = {missing_originale:.2f}% | "
        f"convertito = {missing_convertito:.2f}%"
    )
    return output, controllo

def principale():
    parser = argparse.ArgumentParser(description="Converte il file nefrologico in un unico dataset standardizzato.")
    parser.add_argument("input", type=Path, help="File sorgente .xls o .xlsx")
    parser.add_argument("--risorse", type=Path, default=Path(__file__).resolve().parent, help="Cartella contenente i tre dizionari")
    parser.add_argument("--output", type=Path, default=Path("dati_nefrologia_standardizzati.xlsx"), help="File .xlsx pulito da creare")
    parser.add_argument("--output-controllo", type=Path, default=None, help="File .xlsx di controllo con una colonna fonte per ogni variabile")
    argomenti = parser.parse_args()
    _, controllo = converti(argomenti.input, argomenti.risorse, argomenti.output, argomenti.output_controllo)
    percorso_controllo = argomenti.output_controllo or argomenti.output.with_name(argomenti.output.stem + "_controllo_fonti.xlsx")
    print(f"File pulito creato: {argomenti.output.resolve()}")
    print(f"File di controllo creato: {percorso_controllo.resolve()}")


if __name__ == "__main__":
    principale()
