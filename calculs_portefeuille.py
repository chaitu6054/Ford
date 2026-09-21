"""Rapport 3 — Portefeuille : comptes arrivant à échéance par ligne de modèle.

Source : fichier "Portefeuille Complet à fin <mois>" (structure supposée,
colonnes résolues via config.PORTEFEUILLE_COLONNES). Alimente la colonne
"Accounts Expiring" du rapport ERA. Sans ce fichier, on approxime avec
DateMaturite du fichier Remarketing (voir calculs_era).
"""

from __future__ import annotations

from datetime import date

import pandas as pd

from . import config as C
from . import workbook as W
from .calculs_era import ligne_era


def charger_portefeuille(chemin: str) -> pd.DataFrame:
    """Charger portefeuille (feuille « Portefeuille Complet » si présente)."""
    try:
        xls = pd.ExcelFile(chemin, engine="openpyxl")
        cible = next(
            (
                s
                for s in xls.sheet_names
                if W.nettoyer_nom_colonne(s).casefold() == "portefeuille complet"
            ),
            None,
        )
        feuille = cible if cible is not None else C.PORTEFEUILLE_FEUILLE
        df = pd.read_excel(xls, sheet_name=feuille, engine="openpyxl")
        df.columns = [W.nettoyer_nom_colonne(c) for c in df.columns]
        return df
    except Exception:
        raise


def preparer_portefeuille(df_brut: pd.DataFrame) -> tuple[pd.DataFrame, pd.DataFrame]:
    """Retourne (df préparé, exceptions). Filtre TCM + Ford quand les colonnes existent."""
    try:
        df = df_brut.copy()
        c = W.resoudre_colonnes(df, C.PORTEFEUILLE_COLONNES)
        exc = []
        if not c["DATE_MATURITE"]:
            raise ValueError(
                "Portefeuille : colonne de date de maturité introuvable "
                f"(attendu l'un de {C.PORTEFEUILLE_COLONNES['DATE_MATURITE']})"
            )
        brut = df[c["DATE_MATURITE"]]
        df["_maturite"] = W.serie_date(brut)
        for idx in df.index[brut.notna() & df["_maturite"].isna()]:
            exc.append(
                {
                    "ligne": int(idx) + 2,
                    "contrat": df.at[idx, c["CONTRAT"]] if c["CONTRAT"] else None,
                    "motif": "date de maturité invalide",
                }
            )
        vide = pd.Series("", index=df.index, dtype="string")
        plan = df[c["PLAN"]].astype("string").str.strip().str.casefold() if c["PLAN"] else vide
        marque = (
            df[c["MARQUE"]].astype("string").str.strip().str.casefold() if c["MARQUE"] else vide
        )
        df["_tcm"] = (
            plan.isin(C.PORTEFEUILLE_PLAN_TCM).fillna(True if not c["PLAN"] else False).astype(bool)
        )
        df["_ford"] = (
            marque.isin(C.PORTEFEUILLE_MARQUE_FORD)
            .fillna(True if not c["MARQUE"] else False)
            .astype(bool)
        )
        retenus = int((df["_tcm"] & df["_ford"]).sum())
        if retenus == 0 and len(df) > 0:
            exc.append(
                {
                    "ligne": None,
                    "contrat": None,
                    "motif": (
                        f"Portefeuille : 0/{len(df)} contrats retenus par le filtre "
                        f"TCM+Ford (plan={c['PLAN']!r}, marque={c['MARQUE']!r}) — "
                        "vérifier les valeurs de ces colonnes"
                    ),
                }
            )
        modele = df[c["MODEL"]].astype("string").str.strip() if c["MODEL"] else vide
        df["_ligne_era"] = [ligne_era(m, False) for m in modele.fillna("")]
        df["_annee"], df["_mois"] = df["_maturite"].dt.year, df["_maturite"].dt.month
        return df, pd.DataFrame(exc, columns=["ligne", "contrat", "motif"])
    except Exception:
        raise


def comptes_echeance(df: pd.DataFrame, date_arrete: date) -> pd.DataFrame:
    """Nombre de contrats TCM Ford arrivant à maturité dans le mois, par ligne ERA."""
    try:
        d = pd.Timestamp(date_arrete)
        per = df[df["_tcm"] & df["_ford"] & (df["_annee"] == d.year) & (df["_mois"] == d.month)]
        lignes = C.ERA_ICE + [m for m in C.ERA_BEV if m not in C.ERA_ICE] + [C.ERA_INCONNU]
        out = (
            per.groupby("_ligne_era")
            .size()
            .reindex(lignes)
            .fillna(0)
            .astype(int)
            .to_frame("Accounts Expiring")
        )
        out.loc["Total"] = int(out["Accounts Expiring"].sum())
        return out
    except Exception:
        raise


def echeances_par_mois(df: pd.DataFrame, annee: int) -> pd.DataFrame:
    """Vue annuelle : lignes ERA × mois (utile pour le contrôle et la projection)."""
    try:
        per = df[df["_tcm"] & df["_ford"] & (df["_annee"] == annee)]
        if per.empty:
            return pd.DataFrame()
        return (
            pd.crosstab(per["_ligne_era"], per["_mois"])
            .reindex(columns=range(1, 13))
            .fillna(0)
            .astype(int)
        )
    except Exception:
        raise


def tables_rapport_3(df: pd.DataFrame, date_arrete: date) -> dict[str, pd.DataFrame]:
    """Tables rapport 3."""
    try:
        return {
            "R3 Comptes a echeance": comptes_echeance(df, date_arrete),
            "R3 Echeances par mois": echeances_par_mois(df, pd.Timestamp(date_arrete).year),
        }
    except Exception:
        raise
