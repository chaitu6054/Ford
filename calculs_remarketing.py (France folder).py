"""Rapport 1 — France Remarketing Report (Remarketing = TCM, Repossession = O&R)."""

from __future__ import annotations

from dataclasses import dataclass, field
from datetime import date

import numpy as np
import pandas as pd

from . import config as C
from . import workbook as W


@dataclass
class Prepare:
    """Prepare."""

    df: pd.DataFrame
    exceptions: pd.DataFrame
    colonnes: dict = field(default_factory=dict)
    date_arrete: pd.Timestamp = None


# --------------------------------------------------------------------------- #
# Préparation
# --------------------------------------------------------------------------- #


def preparer_remarketing(
    df_brut: pd.DataFrame, date_arrete: date, fsa_vins: set[str] | None = None
) -> Prepare:
    """Preparer remarketing."""
    try:
        df = df_brut.copy()
        d_arr = pd.Timestamp(date_arrete).normalize()
        c = W.resoudre_colonnes(df)
        exc: list[dict] = []

        def note(
            masque: pd.Series, motif: str, detail: pd.Series | None = None
        ) -> None:
            """Note."""
            try:
                for idx in df.index[masque.fillna(False).astype(bool)]:
                    e = {
                        "ligne": int(idx) + 2,
                        "contrat": df.at[idx, c["CONTRAT"]] if c["CONTRAT"] else None,
                        "vin": df.at[idx, c["VIN"]] if c["VIN"] else None,
                        "motif": motif,
                    }
                    if detail is not None:
                        v = detail.at[idx]
                        e["valeur"] = None if pd.isna(v) else str(v)
                    exc.append(e)
            except Exception:
                raise

        for k in (
            "VALEUR_REVENTE",
            "ARGUS_PROF",
            "VFMG",
            "BALANCE_US",
            "KM",
            "COUVERT",
            "QUOTAS",
            "MONTANT_FACTURER",
            "FRAIS_VENTE",
            "FRAIS_REMISE",
            "CONTRAT_TERMS",
        ):
            if c[k]:
                brut = df[c[k]]
                df[c[k]] = W.serie_numerique(brut)
                note(
                    brut.notna() & df[c[k]].isna() & (brut.astype(str).str.strip() != ""),
                    f"valeur non numérique dans '{c[k]}'",
                    brut,
                )
        for k in ("DATE_RETOUR", "DATE_REVENTE", "DATE_MATURITE"):
            if c[k]:
                brut = df[c[k]]
                df[c[k]] = W.serie_date(brut)
                note(
                    brut.notna() & df[c[k]].isna() & (brut.astype(str).str.strip() != ""),
                    f"date invalide dans '{c[k]}'",
                )

        vide = pd.Series("", index=df.index, dtype="string")
        rest = (
            df[c["RESTITUABLE"]].astype("string").str.strip().str.upper()
            if c["RESTITUABLE"]
            else vide
        )
        df["_exclu_restituable"] = (
            rest.isin({r.upper() for r in C.RESTITUABLE_EXCLUS}).fillna(False).astype(bool)
        )

        statut = (
            df[c["STATUT"]].astype("string").str.strip().str.casefold() if c["STATUT"] else vide
        )
        df["_vendu"] = statut.str.startswith(C.STATUT_VENDU).fillna(False).astype(bool)

        vend = (
            df[c["VENDEUR"]].astype("string").str.strip().str.casefold() if c["VENDEUR"] else vide
        )
        df["_canal"] = vend.map(C.CANAUX)
        note(
            df["_vendu"] & df["_canal"].isna() & vend.notna() & (vend != ""),
            "Vendeur non mappé vers un canal",
        )
        note(df["_vendu"] & (vend.isna() | (vend == "")), "vendu sans Vendeur")

        modele = df[c["MODEL"]].astype("string").str.strip() if c["MODEL"] else vide
        df["_modele"] = modele.str.upper()
        fuel = df[c["FUEL"]].astype("string").str.strip().str.casefold() if c["FUEL"] else vide
        df["_bev"] = fuel.isin(C.FUEL_BEV).fillna(False).astype(bool)

        revente = (
            df[c["VALEUR_REVENTE"]] if c["VALEUR_REVENTE"] else pd.Series(np.nan, index=df.index)
        )
        argus = df[c["ARGUS_PROF"]] if c["ARGUS_PROF"] else pd.Series(np.nan, index=df.index)
        vfmg = df[c["VFMG"]] if c["VFMG"] else pd.Series(np.nan, index=df.index)
        df["_performance"] = np.where((argus > 0) & revente.notna(), revente / argus, np.nan)
        df["_frais_2pct"] = C.TAUX_RETURN_FEES * vfmg.fillna(revente)
        if c["KM"] and c["COUVERT"] and c["QUOTAS"]:
            df["_km_excedentaires"] = df[c["KM"]] - df[c["COUVERT"]] * df[c["QUOTAS"]]
        montant = (
            df[c["MONTANT_FACTURER"]] if c["MONTANT_FACTURER"] else pd.Series(0.0, index=df.index)
        )
        df["_perte_profit"] = df["_frais_2pct"] + montant.fillna(0)
        note(df["_vendu"] & df[c["DATE_REVENTE"]].isna(), "vendu (Statut=Solde) sans date revente")
        note(df["_vendu"] & argus.isna(), "vendu sans Cote ARGUS prof")

        d_ret, d_rev = df[c["DATE_RETOUR"]], df[c["DATE_REVENTE"]]
        df["_retourne_a_date"] = d_ret.notna() & (d_ret <= d_arr)
        df["_vendu_a_date"] = df["_vendu"] & d_rev.notna() & (d_rev <= d_arr)
        df["_en_stock_a_date"] = (
            df["_retourne_a_date"] & ~df["_vendu_a_date"] & ~df["_exclu_restituable"]
        )
        df["_stock_age"] = np.where(df["_en_stock_a_date"], (d_arr - d_ret).dt.days, np.nan)
        df["_bucket"] = W.bucket(df["_stock_age"])
        note(df["_en_stock_a_date"] & (df["_stock_age"] < 0), "ancienneté négative")
        note(df["_en_stock_a_date"] & (df["_stock_age"] > 360), "ancienneté > 360 jours")
        note(df["_vendu"] & d_ret.isna(), "vendu sans Date de retour")
        df["_annee_retour"], df["_mois_retour"] = d_ret.dt.year, d_ret.dt.month
        df["_annee_vente"], df["_mois_vente"] = d_rev.dt.year, d_rev.dt.month

        vins = df[c["VIN"]].astype("string").str.strip().str.upper() if c["VIN"] else vide
        df["_fsa"] = vins.isin(fsa_vins or set()).fillna(False).astype(bool)
        couleur = df["_couleur"].astype("string") if "_couleur" in df else vide
        df["_statut_l"] = couleur.str.contains(C.LIBELLE_STATUT_L, na=False).astype(bool)
        df["_produits_en_attente"] = False  # TODO: Sold Check 1/2/3 (onglet PIVOT de Soraya)

        if c["CONTRAT"]:
            dup = df[c["CONTRAT"]].notna() & df[c["CONTRAT"]].duplicated(keep=False)
            note(dup, "contrat en double")

        return Prepare(
            df=df,
            exceptions=pd.DataFrame(exc, columns=["ligne", "contrat", "vin", "motif"]),
            colonnes=c,
            date_arrete=d_arr,
        )
    except Exception:
        raise


def preparer_repossession(df_brut: pd.DataFrame, date_arrete: date) -> Prepare:
    """Preparer repossession."""
    try:
        df = df_brut.copy()
        d_arr = pd.Timestamp(date_arrete).normalize()
        c = W.resoudre_colonnes(df)
        for k in ("VALEUR_REVENTE", "BALANCE_US", "VFMG"):
            if c[k]:
                df[c[k]] = W.serie_numerique(df[c[k]])
        for k in ("DATE_RETOUR", "DATE_REVENTE"):
            if c[k]:
                df[c[k]] = W.serie_date(df[c[k]])
        balance = df[c["BALANCE_US"]] if c["BALANCE_US"] else pd.Series(np.nan, index=df.index)
        revente = (
            df[c["VALEUR_REVENTE"]] if c["VALEUR_REVENTE"] else pd.Series(np.nan, index=df.index)
        )
        df["_perte_profit"] = balance - revente.fillna(0)  # formule confirmée sur la feuille
        d_ret, d_rev = df[c["DATE_RETOUR"]], df[c["DATE_REVENTE"]]
        df["_exclu_restituable"] = False
        df["_vendu"] = d_rev.notna()
        df["_retourne_a_date"] = d_ret.notna() & (d_ret <= d_arr)
        df["_vendu_a_date"] = d_rev.notna() & (d_rev <= d_arr)
        df["_en_stock_a_date"] = df["_retourne_a_date"] & ~df["_vendu_a_date"]
        df["_stock_age"] = np.where(df["_en_stock_a_date"], (d_arr - d_ret).dt.days, np.nan)
        df["_bucket"] = W.bucket(df["_stock_age"])
        df["_fsa"], df["_statut_l"], df["_produits_en_attente"] = False, False, False
        df["_annee_vente"], df["_mois_vente"] = d_rev.dt.year, d_rev.dt.month
        df["_annee_retour"], df["_mois_retour"] = d_ret.dt.year, d_ret.dt.month
        return Prepare(
            df=df,
            exceptions=pd.DataFrame(columns=["ligne", "contrat", "vin", "motif"]),
            colonnes=c,
            date_arrete=d_arr,
        )
    except Exception:
        raise


# --------------------------------------------------------------------------- #
# Tableaux du rapport
# --------------------------------------------------------------------------- #


def _perimetre(p: Prepare) -> pd.DataFrame:
    try:
        return p.df[~p.df["_exclu_restituable"]]
    except Exception:
        raise


def canaux_disposition(p: Prepare) -> pd.DataFrame:
    """Canaux disposition."""
    try:
        df = _perimetre(p)
        annee, mois = p.date_arrete.year, p.date_arrete.month
        v = df[df["_vendu"] & (df["_annee_vente"] == annee)]
        m, ytd = v[v["_mois_vente"] == mois], v[v["_mois_vente"] <= mois]
        out = pd.DataFrame(index=C.LIGNES_CANAUX)
        out["Ford (mois)"] = (
            m.groupby("_canal").size().reindex(C.LIGNES_CANAUX).fillna(0).astype(int)
        )
        out["Non-Ford (mois)"] = 0
        out["Total (mois)"] = out["Ford (mois)"]
        out["YTD"] = ytd.groupby("_canal").size().reindex(C.LIGNES_CANAUX).fillna(0).astype(int)
        out["Performance vs Market Value (%)"] = (
            m.groupby("_canal")["_performance"].mean().reindex(C.LIGNES_CANAUX)
        )
        tm, ty = int(out["Total (mois)"].sum()), int(out["YTD"].sum())
        out["% mois"] = out["Total (mois)"] / tm if tm else np.nan
        out["% YTD"] = out["YTD"] / ty if ty else np.nan
        out.loc["Total"] = [
            tm,
            0,
            tm,
            ty,
            m["_performance"].mean(),
            1.0 if tm else np.nan,
            1.0 if ty else np.nan,
        ]
        return out
    except Exception:
        raise


def mouvements(p: Prepare, libelle: str = "TCM") -> pd.DataFrame:
    """Mouvements."""
    try:
        df = _perimetre(p)
        fin = p.date_arrete + pd.offsets.MonthEnd(0)
        debut = fin - pd.offsets.MonthBegin(1)
        d_ret, d_rev = df[p.colonnes["DATE_RETOUR"]], df[p.colonnes["DATE_REVENTE"]]
        retours = int(((d_ret >= debut) & (d_ret <= fin)).sum())
        ventes = int((df["_vendu"] & (d_rev >= debut) & (d_rev <= fin)).sum())
        avant = int(
            (
                d_ret.notna() & (d_ret < debut) & ~(df["_vendu"] & d_rev.notna() & (d_rev < debut))
            ).sum()
        )
        fin_stock = int(df["_en_stock_a_date"].sum())
        out = pd.DataFrame(
            {libelle: [avant, retours, ventes, fin_stock]},
            index=[
                "Inventory Prior Period",
                "Returns in Period",
                "Sold in Period",
                "Inventory End of Period",
            ],
        )
        out["contrôle"] = ["", "", "", "OK" if avant + retours - ventes == fin_stock else "ECART"]
        return out
    except Exception:
        raise


def ageing(
    p: Prepare, exclure_attente: bool = False, exclure_statut_l: bool = False, libelle: str = "TCM"
) -> pd.DataFrame:
    """Ageing."""
    try:
        st = _perimetre(p)
        st = st[st["_en_stock_a_date"]]
        if exclure_attente:
            st = st[~st["_produits_en_attente"]]
        if exclure_statut_l:
            st = st[~st["_statut_l"]]

        def ligne(sub: pd.DataFrame) -> list:
            """Ligne."""
            try:
                counts = (
                    sub["_bucket"]
                    .value_counts()
                    .reindex(C.LIBELLES_BUCKETS)
                    .fillna(0)
                    .astype(int)
                    .tolist()
                )
                return counts + [
                    len(sub),
                    round(float(sub["_stock_age"].mean()), 0) if len(sub) else np.nan,
                ]
            except Exception:
                raise

        return pd.DataFrame(
            [ligne(st), ligne(st[st["_fsa"]])],
            index=[libelle, "FSA"],
            columns=C.LIBELLES_BUCKETS + ["Total units", "Avg Age"],
        )
    except Exception:
        raise


def kpis(p: Prepare) -> pd.DataFrame:
    """Kpis."""
    try:
        df = _perimetre(p)
        st = df[df["_en_stock_a_date"]]
        vendus = int(df["_vendu_a_date"].sum())
        return pd.DataFrame(
            {
                "valeur": [
                    p.date_arrete.date(),
                    len(st),
                    vendus,
                    len(st) + vendus,
                    int(st["_fsa"].sum()),
                    round(float(st["_stock_age"].mean()), 1) if len(st) else np.nan,
                    (
                        round(float(st[st["_fsa"]]["_stock_age"].mean()), 1)
                        if st["_fsa"].any()
                        else np.nan
                    ),
                ]
            },
            index=[
                "Current Stock (date)",
                "Unsold Stock",
                "Sold Units",
                "Total",
                "FSA",
                "Average Age",
                "Average Age (FSA)",
            ],
        )
    except Exception:
        raise


def retours_par_mois(p: Prepare) -> pd.DataFrame:
    """Retours par mois."""
    try:
        df = _perimetre(p)
        return (
            df[df["_annee_retour"] == p.date_arrete.year]
            .groupby("_mois_retour")
            .size()
            .reindex(range(1, 13))
            .fillna(0)
            .astype(int)
            .to_frame("Nombre de VIN")
        )
    except Exception:
        raise


def ventes_par_mois(p: Prepare) -> pd.DataFrame:
    """Ventes par mois."""
    try:
        df = _perimetre(p)
        return (
            df[df["_vendu"] & (df["_annee_vente"] == p.date_arrete.year)]
            .groupby("_mois_vente")
            .size()
            .reindex(range(1, 13))
            .fillna(0)
            .astype(int)
            .to_frame("Nombre de Contrat")
        )
    except Exception:
        raise


def ventes_par_modele(p: Prepare) -> pd.DataFrame:
    """Ventes par modele."""
    try:
        df = _perimetre(p)
        v = df[df["_vendu"] & (df["_annee_vente"] == p.date_arrete.year)]
        if v.empty:
            return pd.DataFrame(index=range(1, 13))
        return (
            pd.crosstab(v["_mois_vente"], v["_modele"]).reindex(range(1, 13)).fillna(0).astype(int)
        )
    except Exception:
        raise


def stock_par_etape(p: Prepare) -> pd.DataFrame:
    """Stock par etape."""
    try:
        df = _perimetre(p)
        st = df[df["_en_stock_a_date"]]
        col = p.colonnes.get("ETAPE")
        if not col:
            return pd.DataFrame(columns=["Nombre de Contrat"])
        return st[col].fillna("(vide)").value_counts().to_frame("Nombre de Contrat")
    except Exception:
        raise


def tables_rapport_1(remk: Prepare, repo: Prepare | None) -> dict[str, pd.DataFrame]:
    """Tables rapport 1."""
    try:
        t = {
            "R1 Canaux": canaux_disposition(remk),
            "R1 Mouvements": mouvements(remk),
            "R1 Ageing": ageing(remk),
            "R1 Ageing sans attente": ageing(remk, exclure_attente=True),
            "R1 Ageing sans attente ni L": ageing(
                remk, exclure_attente=True, exclure_statut_l=True
            ),
            "R1 KPIs": kpis(remk),
            "R1 Retours par mois": retours_par_mois(remk),
            "R1 Ventes par mois": ventes_par_mois(remk),
            "R1 Ventes par modele": ventes_par_modele(remk),
            "R1 Stock par Etape": stock_par_etape(remk),
        }
        if repo is not None:
            t["R1 Mouvements O&R"] = mouvements(repo, libelle="O&R")
            t["R1 Ageing O&R"] = ageing(repo, libelle="OTHER/REPOS")
        return t
    except Exception:
        raise
