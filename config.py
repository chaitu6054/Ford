"""Configuration France Inventory Reporting. Rien n'est traduit : les noms restent en français."""

from __future__ import annotations

import os

# ---- GCS ------------------------------------------------------------------
GCP_CONN_ID = "google_cloud_default"
BUCKET_INPUT = "prj-ml-cap-rv-uk-d-tg-rv-fr-inv-input"
BUCKET_OUTPUT = "prj-ml-cap-rv-uk-d-tg-rv-fr-inv-output"
MODELE_FICHIER_REMARKETING = "remarketing_{month}.xlsx"
MODELE_FICHIER_PORTEFEUILLE = "portefeuille_{month}.xlsx"  # optionnel
MODELE_FICHIER_TAUX = "taux_{month}.xlsx"  # optionnel (Bookkeeping Rates)
PREFIXE_MARQUEUR = "_processed"

# ---- Notifications ---------------------------------------------------------
SMTP_CONN_ID = "smtp_default"
MAIL_FROM = "EURVSMTP@FORD.COM"
MAIL_TO = {"dev": ["EURV_FR_OPERATIONS_DEV@FORD.COM"], "prod": ["EURV_FR_OPERATIONS_PROD@FORD.COM"]}
MAIL_TO_ECHEC_SUPPLEMENT: list[str] = []  # contacts FBS à ajouter sur les échecs (renseigner)
SUJET = "France Remarketing workbook reconciliation {month} - {statut}"


def environnement() -> str:
    """Environnement."""
    try:
        env = os.environ.get("FR_INVENTORY_ENV")
        if env:
            return env
        return (
            "dev"
            if "astrodev"
            in os.environ.get(
                "AIRFLOW__API__BASE_URL", os.environ.get("AIRFLOW__WEBSERVER__BASE_URL", "astrodev")
            )
            else "prod"
        )
    except Exception:
        raise


# ---- Feuilles du classeur de base -----------------------------------------
FEUILLE_REMARKETING = "Remarketing"
FEUILLE_REPOSSESSION = "Repossession"
FEUILLE_CODE_COULEUR = "code couleur"
PREFIXE_FEUILLE_FSA = "FSA-UPDATE"

# ---- Règles métier confirmées par Soraya (15/09/2026) -----------------------
RESTITUABLE_EXCLUS = {"NON"}  # OUI + NC inclus
CANAUX = {
    "vp auto": "Physical Auctions",
    "vpauto": "Physical Auctions",
    "bca": "Physical Auctions",
    "autorola": "Physical Auctions",
    "concilian": "Physical Auctions",
    "dealer": "Franchised Dealers",
    "sica": "Franchised Dealers",
}
LIGNES_CANAUX = [
    "Franchised Dealers",
    "Physical Auctions",
    "Closed bidding",
    "PACE / FLEETPACE",
    "Traders",
    "Export",
    "Internet",
    "Write Offs / Insurance",
    "Retail / Employees",
    "Others",
]
STATUT_VENDU = "solde"
LIBELLE_STATUT_L = "statut l"  # couleur orange dans l'onglet code couleur
DATE_MIN_VALIDE = "1950-01-01"
TAUX_RETURN_FEES = 0.02

BUCKETS = [
    (0, 30, "0-30 Days"),
    (31, 60, "31-60 Days"),
    (61, 90, "61-90 Days"),
    (91, 120, "91-120 Days"),
    (121, 150, "121-150 Days"),
    (151, 180, "151-180 Days"),
    (181, 360, "181-360 Days"),
    (361, 10**9, "360+ Days"),
]
LIBELLES_BUCKETS = [b[2] for b in BUCKETS]

# ---- Colonnes (noms tels qu'ils apparaissent dans le classeur) --------------
COLONNES = {
    "CONTRAT": ("Contrat", "Numero contrat"),
    "VIN": ("VIN",),
    "STATUT": ("Statut",),
    "ETAPE": ("Etape",),
    "DATE_RETOUR": ("Date de retour", "Date d'inscription de dossier"),
    "RESTITUABLE": ("Restituable",),
    "VENDEUR": ("Vendeur",),
    "MODEL": ("Model", "Modele"),
    "FUEL": ("Fuel Type", "Engine (a remplir Risk)"),
    "DATE_REVENTE": ("date revente",),
    "VALEUR_REVENTE": ("Valeur revente TTC", "Valeur revente"),
    "ARGUS_PROF": ("Cote ARGUS prof",),
    "VFMG": ("VFMG",),
    "BALANCE_US": ("Balance US",),
    "KM": ("Kilométrage fait par le client", "km"),
    "COUVERT": ("COUVERT BLESS PAR AN",),
    "QUOTAS": ("quotos", "quotas"),
    "MONTANT_FACTURER": ("Montant a facturer le client", "Montant a facturer le client2"),
    "FRAIS_VENTE": ("Frais à la Vente",),
    "FRAIS_REMISE": ("Frais remise en etat",),
    "DATE_MATURITE": ("DateMaturite", "Date Maturite"),
    "CONTRAT_TERMS": ("Contrat Terms (mois)",),
}
COLONNES_REQUISES_REMARKETING = (
    "Contrat",
    "VIN",
    "Statut",
    "Date de retour",
    "Restituable",
    "Vendeur",
    "Model",
    "date revente",
    "Valeur revente TTC",
    "Cote ARGUS prof",
    "VFMG",
)

# ---- Rapport ERA (France TCM Ford) -------------------------------------------
ERA_ICE = [
    "EcoSport",
    "Fiesta",
    "Focus",
    "Puma",
    "Kuga",
    "Mondeo",
    "Mustang",
    "S-Max",
    "Galaxy",
    "Tourneo Courier",
    "Tourneo Connect",
    "Ranger",
    "Transit Courier",
    "Transit Connect",
    "Transit Custom",
    "Transit Van",
    "Tourneo Custom",
    "Explorer",
]
ERA_BEV = [
    "Mustang Mach-E",
    "Explorer",
    "Capri",
    "Puma Gen-E",
    "Tourneo Courier",
    "Transit Courier",
    "Transit Custom",
    "Transit Van",
    "Tourneo Custom",
]
ERA_USED = ["Used ICE Vehicles", "Used BEV Vehicles"]
ERA_INCONNU = "Unknown/ Not recognised"
# Model (fichier de base) -> ligne ERA. Les modèles en majuscules viennent de la colonne Model.
ERA_MODELES = {
    "puma": "Puma",
    "kuga": "Kuga",
    "focus": "Focus",
    "fiesta": "Fiesta",
    "mach-e": "Mustang Mach-E",
    "mustang mach-e": "Mustang Mach-E",
    "explorer": "Explorer",
    "capri": "Capri",
    "puma gen-e": "Puma Gen-E",
    "ecosport": "EcoSport",
    "mondeo": "Mondeo",
    "mustang": "Mustang",
    "s-max": "S-Max",
    "galaxy": "Galaxy",
    "ranger": "Ranger",
    "tourneo courier": "Tourneo Courier",
    "tourneo connect": "Tourneo Connect",
    "tourneo custom": "Tourneo Custom",
    "transit courier": "Transit Courier",
    "transit connect": "Transit Connect",
    "transit custom": "Transit Custom",
    "transit van": "Transit Van",
    "transit": "Transit Van",
}
FUEL_BEV = {"electric", "électrique", "bev"}
ERA_COLONNES = [
    "Month End Inventory Position",
    "No. of vehicles returned to FCE",
    "Accounts Expiring",
    "Return rate (%)",
    "No. of vehicles disposed of in month",
    "Sales Proceeds",
    "Disposal / Remarketing Costs",
    "Excess Mileage Charges",
    "Damage / Other Costs",
    "Insurance Other",
    "Net Sales Proceeds",
    "Net Placement LEV (OFP)",
    "Average disposal price",
    "Average placement LEV (OFP)",
    "Profit per Unit",
    "Total Profit",
]
ERA_COLONNES_MONTANT = ERA_COLONNES[5:]  # converties en $ dans la feuille OUTPUT
TAUX_CHANGE_DEFAUT = None  # EUR -> USD ; None = feuille OUTPUT non produite

# ---- Portefeuille (fichier optionnel) -----------------------------------------
PORTEFEUILLE_FEUILLE = 0
PORTEFEUILLE_COLONNES = {
    "CONTRAT": ("Contrat", "Numero contrat", "N° contrat", "Numéro de contrat"),
    "VIN": ("VIN", "N° de série", "Numero de serie"),
    "MODEL": ("Model", "Modele", "Modèle"),
    "DATE_MATURITE": (
        "DateMaturite",
        "Date Maturite",
        "Date de maturité",
        "Date fin contrat",
        "Maturité",
    ),
    "PLAN": ("Plan", "Produit", "Type produit", "Code Plan"),
    "MARQUE": ("Marque", "Brand"),
}
PORTEFEUILLE_PLAN_TCM = {"tcm"}
PORTEFEUILLE_MARQUE_FORD = {"ford"}
