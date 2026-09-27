import sys
import shutil
#!/usr/bin/env python3
BOOMBOOK_VERSION = "1.6.2"


import streamlit as st
import sqlite3
import pandas as pd
from datetime import datetime, date
import time
from pathlib import Path
from io import BytesIO
import json
import urllib.request
import re
import calendar
import zipfile as _xlsx_zipfile
from xml.sax.saxutils import escape as _xml_escape

try:
    from reportlab.lib.pagesizes import letter
    from reportlab.lib import colors
    from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
    from reportlab.lib.units import inch
    from reportlab.platypus import (
        SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle, HRFlowable
    )
    HAS_REPORTLAB = True
except ImportError:
    HAS_REPORTLAB = False

DATA_DIR = Path.home() / "Documents" / "The BoomBook"
DATA_DIR.mkdir(parents=True, exist_ok=True)
DB_PATH = DATA_DIR / "turf_fertilizer.db"

PRODUCT_CATALOG = [
    {"brand": "Simplot / BEST", "name": "Turf Supreme 16-6-8", "n_pct": 16.0, "p_pct": 6.0, "k_pct": 8.0, "type": "Granular", "notes": "Homogeneous pellet, Fe + S."},
    {"brand": "Simplot / BEST", "name": "Mini Turf 16-8-8", "n_pct": 16.0, "p_pct": 8.0, "k_pct": 8.0, "type": "Granular", "notes": "SGN 150 mini."},
    {"brand": "Simplot / BEST", "name": "Nitra King 21-2-4", "n_pct": 21.0, "p_pct": 2.0, "k_pct": 4.0, "type": "Granular", "notes": "Ammoniacal + nitrate N."},
    {"brand": "Simplot / BEST", "name": "Triple Pro 15-15-15", "n_pct": 15.0, "p_pct": 15.0, "k_pct": 15.0, "type": "Granular", "notes": "1:1:1 balanced."},
    {"brand": "Simplot / BEST", "name": "Pro-Prills 12-8-16", "n_pct": 12.0, "p_pct": 8.0, "k_pct": 16.0, "type": "Granular", "notes": "High K."},
    {"brand": "Simplot / BEST", "name": "NK Select 33-0-6 w/ GAL-XeONE", "n_pct": 33.0, "p_pct": 0.0, "k_pct": 6.0, "type": "Granular", "notes": "Controlled-release, zero P."},
    {"brand": "Simplot / BEST", "name": "Nitrex 20-1-5", "n_pct": 20.0, "p_pct": 1.0, "k_pct": 5.0, "type": "Granular", "notes": "High Fe."},
    {"brand": "Simplot / BEST", "name": "Short Kut 18-4-18 w/ X-Cote", "n_pct": 18.0, "p_pct": 4.0, "k_pct": 18.0, "type": "Granular", "notes": "PCSCU slow-release."},
    {"brand": "Simplot / BEST", "name": "BEST 23-3-6 w/ Iron + UFLEXX", "n_pct": 23.0, "p_pct": 3.0, "k_pct": 6.0, "type": "Granular", "notes": "Stabilized N + Fe."},
    {"brand": "Simplot / BEST", "name": "Ammonium Sulfate 21-0-0", "n_pct": 21.0, "p_pct": 0.0, "k_pct": 0.0, "type": "Granular", "notes": "Quick-release N + S."},
    {"brand": "Simplot / BEST", "name": "Agropell 12-12-12", "n_pct": 12.0, "p_pct": 12.0, "k_pct": 12.0, "type": "Granular", "notes": "Balanced pellet."},
    {"brand": "Simplot / BEST", "name": "Evergreen 18-5-0", "n_pct": 18.0, "p_pct": 5.0, "k_pct": 0.0, "type": "Granular", "notes": "High Fe."},
    {"brand": "Simplot", "name": "Extreme Green 20", "n_pct": 0.0, "p_pct": 0.0, "k_pct": 0.0, "type": "Soluble Powder", "notes": "Iron 20%, Sulfur 12%. Color only."},
    {"brand": "Floratine", "name": "X-Factor 18-3-6", "n_pct": 18.0, "p_pct": 3.0, "k_pct": 6.0, "type": "Liquid Foliar", "notes": "PowerPlay technology."},
    {"brand": "Floratine", "name": "X-Factor 12-0-12 (SRN)", "n_pct": 12.0, "p_pct": 0.0, "k_pct": 12.0, "type": "Liquid Foliar", "notes": "Slow-release N."},
    {"brand": "Floratine", "name": "X-Factor 7-7-7", "n_pct": 7.0, "p_pct": 7.0, "k_pct": 7.0, "type": "Liquid Foliar", "notes": "Balanced foliar."},
    {"brand": "Floratine", "name": "X-Factor 4-4-16", "n_pct": 4.0, "p_pct": 4.0, "k_pct": 16.0, "type": "Liquid Foliar", "notes": "High K foliar."},
    {"brand": "Floratine", "name": "X-Factor 25-0-0 (SRN)", "n_pct": 25.0, "p_pct": 0.0, "k_pct": 0.0, "type": "Liquid Foliar", "notes": "Methylene urea SRN."},
    {"brand": "Floratine", "name": "X-Factor 28-0-0", "n_pct": 28.0, "p_pct": 0.0, "k_pct": 0.0, "type": "Liquid Foliar", "notes": "UAN-based foliar N."},
    {"brand": "Floratine", "name": "X-Factor 12-6-0", "n_pct": 12.0, "p_pct": 6.0, "k_pct": 0.0, "type": "Liquid Foliar", "notes": "N + P foliar."},
    {"brand": "Floratine", "name": "P-48 (10-48-8)", "n_pct": 10.0, "p_pct": 48.0, "k_pct": 8.0, "type": "Soluble Powder", "notes": "High P for roots."},
    {"brand": "Floratine", "name": "PK Fight 0-22-28", "n_pct": 0.0, "p_pct": 22.0, "k_pct": 28.0, "type": "Liquid Foliar", "notes": "Phosphite + K."},
    {"brand": "Floratine", "name": "Moxie 6-0-4", "n_pct": 6.0, "p_pct": 0.0, "k_pct": 4.0, "type": "Liquid Foliar", "notes": "Foliar nitrates."},
    {"brand": "Floratine", "name": "Power 12-6-0", "n_pct": 12.0, "p_pct": 6.0, "k_pct": 0.0, "type": "Liquid Foliar", "notes": "PowerPlay organic acids."},
    {"brand": "Floratine", "name": "X-Factor 0-0-22", "n_pct": 0.0, "p_pct": 0.0, "k_pct": 22.0, "type": "Liquid Foliar", "notes": "Potassium thiosulfate."},
    {"brand": "Floratine", "name": "X-Factor 23-0-0 + Mo", "n_pct": 23.0, "p_pct": 0.0, "k_pct": 0.0, "type": "Liquid Foliar", "notes": "High N + Mo."},
    {"brand": "Floratine", "name": "X-Factor 24-0-0 + Mo", "n_pct": 24.0, "p_pct": 0.0, "k_pct": 0.0, "type": "Liquid Foliar", "notes": "High N + Mo."},
    {"brand": "The Andersons", "name": "Contec DG 12-3-12", "n_pct": 12.0, "p_pct": 3.0, "k_pct": 12.0, "type": "Granular", "notes": "Dispersing granule greens/tees."},
    {"brand": "The Andersons", "name": "Contec DG 18-9-18", "n_pct": 18.0, "p_pct": 9.0, "k_pct": 18.0, "type": "Granular", "notes": "Balanced DG MUtech."},
    {"brand": "The Andersons", "name": "Contec DG 13-0-26", "n_pct": 13.0, "p_pct": 0.0, "k_pct": 26.0, "type": "Granular", "notes": "High K 100% MUtech."},
    {"brand": "The Andersons", "name": "Contec DG 17-0-17", "n_pct": 17.0, "p_pct": 0.0, "k_pct": 17.0, "type": "Granular", "notes": "50% MUtech."},
    {"brand": "The Andersons", "name": "Contec DG 19-0-15", "n_pct": 19.0, "p_pct": 0.0, "k_pct": 15.0, "type": "Granular", "notes": "100% MUtech + micros."},
    {"brand": "The Andersons", "name": "Contec DG 12-24-8", "n_pct": 12.0, "p_pct": 24.0, "k_pct": 8.0, "type": "Granular", "notes": "High P starter."},
    {"brand": "The Andersons", "name": "16-0-6 Fairway", "n_pct": 16.0, "p_pct": 0.0, "k_pct": 6.0, "type": "Granular", "notes": "Fairway grade."},
    {"brand": "The Andersons", "name": "18-0-12 Fairway", "n_pct": 18.0, "p_pct": 0.0, "k_pct": 12.0, "type": "Granular", "notes": "Fairway grade."},
    {"brand": "The Andersons", "name": "21-0-5", "n_pct": 21.0, "p_pct": 0.0, "k_pct": 5.0, "type": "Granular", "notes": "Standard SGN."},
    {"brand": "The Andersons", "name": "Foltec SG 24-0-8", "n_pct": 24.0, "p_pct": 0.0, "k_pct": 8.0, "type": "Soluble Powder", "notes": "Sprayable soluble."},
    {"brand": "The Andersons", "name": "Foltec SG 8-0-24", "n_pct": 8.0, "p_pct": 0.0, "k_pct": 24.0, "type": "Soluble Powder", "notes": "High K soluble."},
    {"brand": "The Andersons", "name": "Foltec SG 8-24-8", "n_pct": 8.0, "p_pct": 24.0, "k_pct": 8.0, "type": "Soluble Powder", "notes": "High P soluble."},
    {"brand": "The Andersons", "name": "Foltec SG 16-0-16", "n_pct": 16.0, "p_pct": 0.0, "k_pct": 16.0, "type": "Soluble Powder", "notes": "Balanced N-K soluble."},
    {"brand": "The Andersons", "name": "HCU 44-0-0", "n_pct": 44.0, "p_pct": 0.0, "k_pct": 0.0, "type": "Granular", "notes": "Humic Coated Urea."},
    {"brand": "Lebanon", "name": "Country Club 18-0-18", "n_pct": 18.0, "p_pct": 0.0, "k_pct": 18.0, "type": "Granular", "notes": "Meth-Ex slow release."},
    {"brand": "Lebanon", "name": "Country Club 12-24-8", "n_pct": 12.0, "p_pct": 24.0, "k_pct": 8.0, "type": "Granular", "notes": "High P starter."},
    {"brand": "Lebanon", "name": "Country Club 25-0-12", "n_pct": 25.0, "p_pct": 0.0, "k_pct": 12.0, "type": "Granular", "notes": "High N."},
    {"brand": "Lebanon", "name": "ProScape 19-0-19", "n_pct": 19.0, "p_pct": 0.0, "k_pct": 19.0, "type": "Granular", "notes": "Balanced N-K."},
    {"brand": "Scotts Pro", "name": "24-0-10", "n_pct": 24.0, "p_pct": 0.0, "k_pct": 10.0, "type": "Granular", "notes": "Professional turf fertilizer."},
    {"brand": "Scotts Pro", "name": "26-0-3", "n_pct": 26.0, "p_pct": 0.0, "k_pct": 3.0, "type": "Granular", "notes": "High N."},
    {"brand": "Scotts Pro", "name": "Starter 24-25-4", "n_pct": 24.0, "p_pct": 25.0, "k_pct": 4.0, "type": "Granular", "notes": "Starter fertilizer."},
    {"brand": "Thrive", "name": "Thrive 46-0-0", "n_pct": 46.0, "p_pct": 0.0, "k_pct": 0.0, "type": "Granular", "notes": "Urea. High analysis nitrogen."},
    {"brand": "Generic", "name": "Urea 46-0-0", "n_pct": 46.0, "p_pct": 0.0, "k_pct": 0.0, "type": "Granular", "notes": "Straight urea."},
    {"brand": "Generic", "name": "Ammonium Sulfate 21-0-0", "n_pct": 21.0, "p_pct": 0.0, "k_pct": 0.0, "type": "Granular", "notes": "Quick release N + sulfur."},
    {"brand": "Generic", "name": "Potassium Sulfate 0-0-50", "n_pct": 0.0, "p_pct": 0.0, "k_pct": 50.0, "type": "Granular", "notes": "Sulfate of potash."},
    {"brand": "Generic", "name": "MAP 11-52-0", "n_pct": 11.0, "p_pct": 52.0, "k_pct": 0.0, "type": "Granular", "notes": "Monoammonium phosphate."},
    {"brand": "Generic", "name": "20-20-20 Water Soluble", "n_pct": 20.0, "p_pct": 20.0, "k_pct": 20.0, "type": "Soluble Powder", "notes": "General purpose water soluble."},
    {"brand": "Generic", "name": "15-30-15 Starter", "n_pct": 15.0, "p_pct": 30.0, "k_pct": 15.0, "type": "Granular", "notes": "High P starter blend."},
    {"brand": "Generic", "name": "10-10-10", "n_pct": 10.0, "p_pct": 10.0, "k_pct": 10.0, "type": "Granular", "notes": "Balanced."},
    {"brand": "Generic", "name": "16-4-8", "n_pct": 16.0, "p_pct": 4.0, "k_pct": 8.0, "type": "Granular", "notes": "Common turf ratio."},
    {"brand": "Generic", "name": "28-3-10", "n_pct": 28.0, "p_pct": 3.0, "k_pct": 10.0, "type": "Granular", "notes": "High N blend."},
    {"brand": "Generic", "name": "32-0-4", "n_pct": 32.0, "p_pct": 0.0, "k_pct": 4.0, "type": "Granular", "notes": "High N."},
    {"brand": 'Syngenta', "name": 'Daconil Action', "n_pct": 0.0, "p_pct": 0.0, "k_pct": 0.0, "type": 'Fungicide', "notes": 'AI: chlorothalonil 53.94% + acibenzolar-S-methyl 0.11%; FRAC M05 + P1; EPA Reg. 100-1364; golf-course turf. Verify current label before use.'},
    {"brand": 'Syngenta', "name": 'Heritage Action', "n_pct": 0.0, "p_pct": 0.0, "k_pct": 0.0, "type": 'Fungicide', "notes": 'AI: azoxystrobin 50% + acibenzolar-S-methyl 1.18%; FRAC 11 + P1; EPA Reg. 100-1550; golf-course turf. Verify current label before use.'},
    {"brand": 'Syngenta', "name": 'Posterity XT', "n_pct": 0.0, "p_pct": 0.0, "k_pct": 0.0, "type": 'Fungicide', "notes": 'AI: azoxystrobin 6% + pydiflumetofen 1% + propiconazole 10.1%; EPA Reg. 100-1654; golf greens/tees/fairways. Verify current label before use.'},
    {"brand": 'Syngenta', "name": 'Secure Action', "n_pct": 0.0, "p_pct": 0.0, "k_pct": 0.0, "type": 'Fungicide', "notes": 'AI: fluazinam 40%; EPA Reg. 100-1633; golf-course turf. Verify current label before use.'},
    {"brand": 'Syngenta', "name": 'Renown Fungicide', "n_pct": 0.0, "p_pct": 0.0, "k_pct": 0.0, "type": 'Fungicide', "notes": 'AI: azoxystrobin 3% + chlorothalonil 45%; EPA Reg. 100-1315; golf-course turf. Verify current label before use.'},
    {"brand": 'BASF', "name": 'Maxtima Fungicide', "n_pct": 0.0, "p_pct": 0.0, "k_pct": 0.0, "type": 'Fungicide', "notes": 'AI: mefentrifluconazole 34.93%; EPA Reg. 7969-404; turf disease control. Verify current label and state registration before use.'},
    {"brand": 'BASF', "name": 'Encartis Fungicide', "n_pct": 0.0, "p_pct": 0.0, "k_pct": 0.0, "type": 'Fungicide', "notes": 'AI: chlorothalonil 54.14% + boscalid 2.2%; EPA Reg. 7969-348; golf greens/tees/fairways. Verify current label before use.'},
    {"brand": 'Nufarm', "name": 'Traction / Sector Fungicide', "n_pct": 0.0, "p_pct": 0.0, "k_pct": 0.0, "type": 'Fungicide', "notes": 'AI: fluazinam 17% + tebuconazole 17.6%; EPA Reg. 228-739; golf greens/tees/fairways/rough. Verify current label before use.'},
    {"brand": 'FMC', "name": 'QuickSilver T&O Herbicide', "n_pct": 0.0, "p_pct": 0.0, "k_pct": 0.0, "type": 'Herbicide', "notes": 'AI: carfentrazone-ethyl 21.3%; HRAC 14; EPA Reg. 279-3265; turf/golf use including broadleaf weeds and silvery thread moss. Verify current label before use.'},
    {"brand": 'Envu', "name": 'Specticle FLO', "n_pct": 0.0, "p_pct": 0.0, "k_pct": 0.0, "type": 'Herbicide', "notes": 'AI: indaziflam 7.4%; EPA Reg. 432-1608; turf/ornamental uses. Verify turf species, golf-course site and current label before use.'},
    {"brand": 'Valent', "name": 'Certainty Turf Herbicide', "n_pct": 0.0, "p_pct": 0.0, "k_pct": 0.0, "type": 'Herbicide', "notes": 'AI: sulfosulfuron 75%; EPA Reg. 59639-226; golf-course turf sites including creeping bentgrass and bermudagrass. Verify current label before use.'},
    {"brand": 'Envu', "name": 'Tribute Total', "n_pct": 0.0, "p_pct": 0.0, "k_pct": 0.0, "type": 'Herbicide', "notes": 'AI: thiencarbazone-methyl 9.9% + foramsulfuron 19.8% + halosulfuron-methyl 30.8%; EPA Reg. 432-1519; turf herbicide. Verify current label/turf restrictions before use.'},
    {"brand": 'BASF', "name": 'Finale Herbicide', "n_pct": 0.0, "p_pct": 0.0, "k_pct": 0.0, "type": 'Herbicide', "notes": 'AI: glufosinate 11.33%; EPA record includes golf-course sand traps and dormant bermudagrass uses. Verify current EPA registration/label before use.'},
    {"brand": 'Envu', "name": 'Sencor 75% Turf Herbicide', "n_pct": 0.0, "p_pct": 0.0, "k_pct": 0.0, "type": 'Herbicide', "notes": 'AI: metribuzin 75%; EPA Reg. 432-1469; bermudagrass golf-course use. Verify current label before use.'},
    {"brand": 'The Andersons', "name": '0.067% Acelepryn Insecticide Plus Fertilizer', "n_pct": 0.0, "p_pct": 0.0, "k_pct": 0.0, "type": 'Insecticide', "notes": 'AI: chlorantraniliprole 0.067%; EPA Reg. 9198-247; golf-course turf including rough/fairways. Product may carry fertilizer analysis by formulation; enter NPK separately if needed. Verify current label.'},
    {"brand": 'Corteva', "name": 'Conserve SC Turf and Ornamental', "n_pct": 0.0, "p_pct": 0.0, "k_pct": 0.0, "type": 'Insecticide', "notes": 'AI: spinosad 11.6%; EPA Reg. 62719-291; turf pests include annual bluegrass weevil, black cutworm and black turfgrass ataenius. Verify current label and golf-course site before use.'},
    {"brand": 'Syngenta', "name": 'Tuque exoGEM', "n_pct": 0.0, "p_pct": 0.0, "k_pct": 0.0, "type": 'Fungicide', "notes": 'AI: fludioxonil 7.1% + benzovindiflupyr 0.7%; FRAC 12 + 7; EPA Reg. 100-1733; broad-spectrum fungicide for prevention/control of turf diseases on golf courses; professional applicators. Verify current label.'},
    {"brand": 'Quali-Pro', "name": 'Fluazinam 40 / Quali-Pro', "n_pct": 0.0, "p_pct": 0.0, "k_pct": 0.0, "type": 'Fungicide', "notes": 'AI: fluazinam 40%; FRAC 29; EPA Reg. 53883-454; EPA lists golf-course rough, courses, fairways, greens and tees. Verify current label.'},
    {"brand": 'Nufarm', "name": 'Sector Fungicide', "n_pct": 0.0, "p_pct": 0.0, "k_pct": 0.0, "type": 'Fungicide', "notes": 'AI: fluazinam 17% + tebuconazole 17.6%; FRAC 29 + 3; EPA Reg. 228-739; golf-course turf including rough, fairways, greens and tees. Verify current label.'},
    {"brand": 'Nufarm', "name": 'Sektor Fungicide', "n_pct": 0.0, "p_pct": 0.0, "k_pct": 0.0, "type": 'Fungicide', "notes": 'AI: fluazinam 17% + tebuconazole 17.6%; FRAC 29 + 3; alternate brand under EPA Reg. 228-739; golf-course turf. Verify current label.'},
    {"brand": 'Generic', "name": 'Thiophanate Methyl 4.5F AG', "n_pct": 0.0, "p_pct": 0.0, "k_pct": 0.0, "type": 'Fungicide', "notes": 'AI: thiophanate-methyl; FRAC 1; EPA turf label includes golf greens, tees and fairways; professional/certified applicator restrictions may apply. Verify exact product EPA registration/current label.'},
    {"brand": 'Tee-Off', "name": 'Tee-Off 4.5F', "n_pct": 0.0, "p_pct": 0.0, "k_pct": 0.0, "type": 'Fungicide', "notes": 'AI: thiophanate-methyl 46.2%; FRAC 1; EPA Reg. 83070-1; EPA lists bentgrass golf fairways, greens and tees. Verify current label.'},
    {"brand": 'UPL', "name": 'Topsin 4.5FL Fungicide', "n_pct": 0.0, "p_pct": 0.0, "k_pct": 0.0, "type": 'Fungicide', "notes": 'AI: thiophanate-methyl 45%; FRAC 1; EPA Reg. 8033-122; turf/ornamental fungicide. Verify current golf-turf directions and current label.'},
    {"brand": 'Generic', "name": 'Azoxystrobin 2.08 lb SC', "n_pct": 0.0, "p_pct": 0.0, "k_pct": 0.0, "type": 'Fungicide', "notes": 'AI: azoxystrobin 22.9%; FRAC 11; EPA Reg. 71532-35. Broad-spectrum azoxystrobin formulation; verify exact golf-course turf site and current label before use.'},
    {"brand": 'Generic', "name": 'Azoxystrobin + Tebuconazole Turf Fungicide', "n_pct": 0.0, "p_pct": 0.0, "k_pct": 0.0, "type": 'Fungicide', "notes": 'AI: azoxystrobin 18.2% + tebuconazole 27.3%; FRAC 11 + 3; EPA Reg. 60063-62; turf and ornamental fungicide. Verify current golf-course directions.'},
    {"brand": 'Generic', "name": 'Prothioconazole + Azoxystrobin Turf', "n_pct": 0.0, "p_pct": 0.0, "k_pct": 0.0, "type": 'Fungicide', "notes": 'AI: prothioconazole + azoxystrobin; FRAC 3 + 11; EPA Reg. 60063-100; EPA label states golf-course turf only and professional applicators only. Verify current label.'},
    {"brand": 'Syngenta', "name": 'Posterity XT / Headway XT', "n_pct": 0.0, "p_pct": 0.0, "k_pct": 0.0, "type": 'Fungicide', "notes": 'AI: azoxystrobin 6% + pydiflumetofen 1% + propiconazole 10.1%; FRAC 11 + 7 + 3; EPA Reg. 100-1654; golf courses, fairways, greens and tees. Verify current label.'},
    {"brand": 'Syngenta', "name": 'Posterity Peak', "n_pct": 0.0, "p_pct": 0.0, "k_pct": 0.0, "type": 'Fungicide', "notes": 'Alternate brand under EPA Reg. 100-1654; azoxystrobin + pydiflumetofen + propiconazole; golf courses/fairways/greens/tees. Verify current label.'},
    {"brand": 'Generic', "name": 'Chlorothalonil Turf Fungicide', "n_pct": 0.0, "p_pct": 0.0, "k_pct": 0.0, "type": 'Fungicide', "notes": 'AI: chlorothalonil; FRAC M05; EPA recognizes non-residential turf/golf-course uses. Product-specific registrations and 2025 mitigation changes vary; verify exact product/current label.'},
    {"brand": 'Generic', "name": 'Propiconazole Turf Fungicide', "n_pct": 0.0, "p_pct": 0.0, "k_pct": 0.0, "type": 'Fungicide', "notes": 'AI: propiconazole; FRAC 3; multiple turf formulations exist. Add/use only against the exact product label in hand; verify golf-course sites and current EPA/state registration.'},
    {"brand": 'Generic', "name": 'Tebuconazole Turf Fungicide', "n_pct": 0.0, "p_pct": 0.0, "k_pct": 0.0, "type": 'Fungicide', "notes": 'AI: tebuconazole; FRAC 3; multiple turf formulations exist. Verify exact product EPA registration, golf-course use site and current label.'},
    {"brand": 'Generic', "name": 'Fluazinam Turf Fungicide', "n_pct": 0.0, "p_pct": 0.0, "k_pct": 0.0, "type": 'Fungicide', "notes": 'AI: fluazinam; FRAC 29; EPA registrations include golf-course greens, tees, fairways and rough for certain formulations. Verify exact product/current label.'},
    {"brand": 'Generic', "name": 'Imidacloprid 0.5G Insecticide', "n_pct": 0.0, "p_pct": 0.0, "k_pct": 0.0, "type": 'Insecticide', "notes": 'AI: imidacloprid 0.5%; IRAC 4A; EPA Reg. 42750-151; EPA lists ornamental turf golf courses; pests include ABW larvae, billbugs, black turfgrass ataenius, chinch bugs and white grubs. Verify current label.'},
    {"brand": 'Generic', "name": 'Bifenthrin 7.9% SC', "n_pct": 0.0, "p_pct": 0.0, "k_pct": 0.0, "type": 'Insecticide', "notes": 'AI: bifenthrin 7.9%; IRAC 3A; EPA Reg. 42750-369; turf pests include annual bluegrass weevil, billbugs and black turfgrass ataenius. Verify exact golf-course site/current label.'},
    {"brand": 'Generic', "name": 'Chlorantraniliprole T&O Insecticide', "n_pct": 0.0, "p_pct": 0.0, "k_pct": 0.0, "type": 'Insecticide', "notes": 'AI: chlorantraniliprole; IRAC 28; EPA golf-course/turf labels exist and include golf greens, tees/fairways restrictions in some states. Verify exact product/current label.'},
    {"brand": 'Generic', "name": 'NUL-3447 T&O Insecticide', "n_pct": 0.0, "p_pct": 0.0, "k_pct": 0.0, "type": 'Insecticide', "notes": 'EPA Reg. 228-768; EPA label includes golf courses, lawns and landscapes; turf pests include annual bluegrass weevil, billbugs and crane fly larvae. Verify active ingredient/current label before use.'},
    {"brand": 'Generic', "name": 'Spinosad Turf & Ornamental', "n_pct": 0.0, "p_pct": 0.0, "k_pct": 0.0, "type": 'Insecticide', "notes": 'AI: spinosad; IRAC 5; turf formulations include golf-course pest uses. Verify exact product EPA registration and current label.'},
    {"brand": 'Generic', "name": 'Indaziflam Turf Herbicide', "n_pct": 0.0, "p_pct": 0.0, "k_pct": 0.0, "type": 'Herbicide', "notes": 'AI: indaziflam; HRAC 29; EPA recognizes indaziflam turf uses including golf courses for certain formulations. Verify exact formulation/turf species/current label.'},
    {"brand": 'Generic', "name": 'Sulfosulfuron Turf Herbicide', "n_pct": 0.0, "p_pct": 0.0, "k_pct": 0.0, "type": 'Herbicide', "notes": 'AI: sulfosulfuron; HRAC 2; golf-course turf registrations exist for certain products such as Certainty. Verify exact product, turf species and current label.'},
    {"brand": 'Generic', "name": 'Carfentrazone-ethyl Turf Herbicide', "n_pct": 0.0, "p_pct": 0.0, "k_pct": 0.0, "type": 'Herbicide', "notes": 'AI: carfentrazone-ethyl; HRAC 14; turf/golf products include QuickSilver T&O. Verify exact product/current label.'},
    {"brand": 'Generic', "name": 'Glufosinate Turf/Noncrop Herbicide', "n_pct": 0.0, "p_pct": 0.0, "k_pct": 0.0, "type": 'Herbicide', "notes": 'AI: glufosinate; HRAC 10; some labels include limited golf-course/dormant turf uses. Do not assume general turf use; verify exact product, turf species and current label.'},
    {"brand": 'Generic', "name": 'Metribuzin Turf Herbicide', "n_pct": 0.0, "p_pct": 0.0, "k_pct": 0.0, "type": 'Herbicide', "notes": 'AI: metribuzin; HRAC 5; certain registrations include bermudagrass golf-course use. Verify exact product/current label.'},
    {"brand": 'GRIGG', "name": "Gary's Green 18-3-4", "n_pct": 18.0, "p_pct": 3.0, "k_pct": 4.0, "type": 'Liquid Foliar', "minors": 'Mg 0.50%; Cu 0.12%; Fe 1.00%; Mn 0.10%; Zn 0.10%', "notes": "Professional golf/turf fertilizer; verify current product label."},
    {"brand": 'GRIGG', "name": "Gary's Green Ultra 14-2-3", "n_pct": 14.0, "p_pct": 2.0, "k_pct": 3.0, "type": 'Liquid Foliar', "minors": 'Mg 0.50%', "notes": "Professional golf/turf fertilizer; verify current product label."},
    {"brand": 'GRIGG', "name": 'Nutra Green 5-10-5', "n_pct": 5.0, "p_pct": 10.0, "k_pct": 5.0, "type": 'Liquid Foliar', "minors": 'Mg 1.00%; B 0.12%; Cu 0.10%; Fe 1.00%; Mn 0.50%; Zn 0.10%', "notes": "Professional golf/turf fertilizer; verify current product label."},
    {"brand": 'GRIGG', "name": 'Advantage 5-0-1', "n_pct": 5.0, "p_pct": 0.0, "k_pct": 1.0, "type": 'Liquid Foliar', "minors": 'B 0.05%; Mn 0.75%; Mo 0.001%; Zn 0.75%', "notes": "Professional golf/turf fertilizer; verify current product label."},
    {"brand": 'GRIGG', "name": 'P-K Plus 3-5-17', "n_pct": 3.0, "p_pct": 5.0, "k_pct": 17.0, "type": 'Liquid Foliar', "minors": 'B 0.02%; Co 0.01%; Mo 0.001%', "notes": "Professional golf/turf fertilizer; verify current product label."},
    {"brand": 'GRIGG', "name": 'Tuff Turf 1-0-14', "n_pct": 1.0, "p_pct": 0.0, "k_pct": 14.0, "type": 'Liquid Foliar', "minors": 'Mg 0.50%; Fe 0.50%', "notes": "Professional golf/turf fertilizer; verify current product label."},
    {"brand": 'GRIGG', "name": 'Ultraplex 4-0-3', "n_pct": 4.0, "p_pct": 0.0, "k_pct": 3.0, "type": 'Liquid Foliar', "minors": 'Mg 0.50%; B 0.05%; Cu 0.05%; Fe 1.95%; Mn 0.40%; Zn 0.40%', "notes": "Professional golf/turf fertilizer; verify current product label."},
    {"brand": 'GRIGG', "name": 'Carboplex 6-4-4', "n_pct": 6.0, "p_pct": 4.0, "k_pct": 4.0, "type": 'Liquid Foliar', "minors": 'Fe 0.20%; Mn 0.05%; Zn 0.05%', "notes": "Professional golf/turf fertilizer; verify current product label."},
    {"brand": 'GRIGG', "name": 'Kelplex 1-2-2', "n_pct": 1.0, "p_pct": 2.0, "k_pct": 2.0, "type": 'Liquid Foliar', "minors": 'Fe 0.10%', "notes": "Professional golf/turf fertilizer; verify current product label."},
    {"brand": 'GRIGG', "name": 'Kal-B 8-0-4', "n_pct": 8.0, "p_pct": 0.0, "k_pct": 4.0, "type": 'Liquid Foliar', "minors": 'Ca 10.00%; B 0.05%', "notes": "Professional golf/turf fertilizer; verify current product label."},
    {"brand": 'GRIGG', "name": 'MegAleX 3-0-0', "n_pct": 3.0, "p_pct": 0.0, "k_pct": 0.0, "type": 'Liquid Foliar', "minors": '', "notes": "Professional golf/turf fertilizer; verify current product label."},
    {"brand": 'GRIGG', "name": 'Seven Iron 7-7-7', "n_pct": 7.0, "p_pct": 7.0, "k_pct": 7.0, "type": 'Granular', "minors": 'Ca 3.00%; S 7.00%; Fe 7.00%; Mn 1.50%; Zn 0.20%', "notes": "Professional golf/turf fertilizer; verify current product label."},
    {"brand": 'GRIGG', "name": 'Turf Rally 16-4-8', "n_pct": 16.0, "p_pct": 4.0, "k_pct": 8.0, "type": 'Granular', "minors": 'Ca 2.00%; S 6.00%; Fe 3.00%; Mn 0.20%; Zn 0.10%', "notes": "Professional golf/turf fertilizer; verify current product label."},
    {"brand": 'GRIGG', "name": 'Endurance 8-4-16 Greens Grade', "n_pct": 8.0, "p_pct": 4.0, "k_pct": 16.0, "type": 'Granular', "minors": 'Ca 8.00%; S 6.00%; Fe 3.00%; Mn 0.20%; Zn 0.10%', "notes": "Professional golf/turf fertilizer; verify current product label."},
    {"brand": 'GRIGG', "name": 'Suprema 12-0-12 + Micros', "n_pct": 12.0, "p_pct": 0.0, "k_pct": 12.0, "type": 'Liquid Foliar', "minors": 'Fe 1.00%; Mn 0.05%; Zn 0.05%', "notes": "Professional golf/turf fertilizer; verify current product label."},
    {"brand": 'LebanonTurf', "name": 'Country Club 34-0-0', "n_pct": 34.0, "p_pct": 0.0, "k_pct": 0.0, "type": 'Granular', "minors": '', "notes": "Professional golf/turf fertilizer; verify current product label."},
    {"brand": 'LebanonTurf', "name": 'Country Club 30-0-5', "n_pct": 30.0, "p_pct": 0.0, "k_pct": 5.0, "type": 'Granular', "minors": '', "notes": "Professional golf/turf fertilizer; verify current product label."},
    {"brand": 'LebanonTurf', "name": 'Country Club 27-0-5', "n_pct": 27.0, "p_pct": 0.0, "k_pct": 5.0, "type": 'Granular', "minors": '', "notes": "Professional golf/turf fertilizer; verify current product label."},
    {"brand": 'LebanonTurf', "name": 'Country Club 26-0-5', "n_pct": 26.0, "p_pct": 0.0, "k_pct": 5.0, "type": 'Granular', "minors": '', "notes": "Professional golf/turf fertilizer; verify current product label."},
    {"brand": 'LebanonTurf', "name": 'Country Club 25-0-15 + 1% Fe', "n_pct": 25.0, "p_pct": 0.0, "k_pct": 15.0, "type": 'Granular', "minors": 'Fe 1.00%', "notes": "Professional golf/turf fertilizer; verify current product label."},
    {"brand": 'LebanonTurf', "name": 'Country Club 24-0-12', "n_pct": 24.0, "p_pct": 0.0, "k_pct": 12.0, "type": 'Granular', "minors": '', "notes": "Professional golf/turf fertilizer; verify current product label."},
    {"brand": 'LebanonTurf', "name": 'Country Club 19-0-19 + 3% Fe', "n_pct": 19.0, "p_pct": 0.0, "k_pct": 19.0, "type": 'Granular', "minors": 'Fe 3.00%', "notes": "Professional golf/turf fertilizer; verify current product label."},
    {"brand": 'LebanonTurf', "name": 'Country Club 17-0-17', "n_pct": 17.0, "p_pct": 0.0, "k_pct": 17.0, "type": 'Granular', "minors": '', "notes": "Professional golf/turf fertilizer; verify current product label."},
    {"brand": 'LebanonTurf', "name": 'Country Club 16-0-8', "n_pct": 16.0, "p_pct": 0.0, "k_pct": 8.0, "type": 'Granular', "minors": '', "notes": "Professional golf/turf fertilizer; verify current product label."},
    {"brand": 'LebanonTurf', "name": 'Country Club 14-0-14', "n_pct": 14.0, "p_pct": 0.0, "k_pct": 14.0, "type": 'Granular', "minors": '', "notes": "Professional golf/turf fertilizer; verify current product label."},
    {"brand": 'LebanonTurf', "name": 'Country Club MD 12-24-8', "n_pct": 12.0, "p_pct": 24.0, "k_pct": 8.0, "type": 'Granular', "minors": '', "notes": "Professional golf/turf fertilizer; verify current product label."},
    {"brand": "Harrell's", "name": 'POLYON 27-4-10', "n_pct": 27.0, "p_pct": 4.0, "k_pct": 10.0, "type": 'Granular', "minors": 'S 6.588%; Fe 0.563%', "notes": "Professional golf/turf fertilizer; verify current product label."},

    {"brand": 'The Andersons', "name": 'Contec DG Mag-tec 0-0-12', "n_pct": 0.0, "p_pct": 0.0, "k_pct": 12.0, "type": 'Granular', "minors": 'Mg 24.00%', "notes": "Professional golf/turf nutrition product; verify current product label/guaranteed analysis."},
    {"brand": 'The Andersons', "name": 'Contec DG Kal-tec 0-0-13', "n_pct": 0.0, "p_pct": 0.0, "k_pct": 13.0, "type": 'Granular', "minors": 'Ca 9.20%; Mg 2.00%; Mn 1.50%', "notes": "Professional golf/turf nutrition product; verify current product label/guaranteed analysis."},
    {"brand": 'The Andersons', "name": 'Contec DG 0-0-25', "n_pct": 0.0, "p_pct": 0.0, "k_pct": 25.0, "type": 'Granular', "minors": 'Mg 4.00%; Mn 3.00%', "notes": "Professional golf/turf nutrition product; verify current product label/guaranteed analysis."},
    {"brand": 'The Andersons', "name": 'Contec DG 5-5-20', "n_pct": 5.0, "p_pct": 5.0, "k_pct": 20.0, "type": 'Granular', "minors": '', "notes": "Professional golf/turf nutrition product; verify current product label/guaranteed analysis."},
    {"brand": 'The Andersons', "name": 'Contec DG 6-0-12', "n_pct": 6.0, "p_pct": 0.0, "k_pct": 12.0, "type": 'Granular', "minors": '', "notes": "Professional golf/turf nutrition product; verify current product label/guaranteed analysis."},
    {"brand": 'The Andersons', "name": 'Contec DG 7-14-14', "n_pct": 7.0, "p_pct": 14.0, "k_pct": 14.0, "type": 'Granular', "minors": 'Mg 2.00%', "notes": "Professional golf/turf nutrition product; verify current product label/guaranteed analysis."},
    {"brand": 'The Andersons', "name": 'Contec DG 9-0-18', "n_pct": 9.0, "p_pct": 0.0, "k_pct": 18.0, "type": 'Granular', "minors": 'Mn 6.00%; Fe 0.30%', "notes": "Professional golf/turf nutrition product; verify current product label/guaranteed analysis."},
    {"brand": 'The Andersons', "name": 'Contec DG 9-4-9', "n_pct": 9.0, "p_pct": 4.0, "k_pct": 9.0, "type": 'Granular', "minors": '', "notes": "Professional golf/turf nutrition product; verify current product label/guaranteed analysis."},
    {"brand": 'The Andersons', "name": 'Contec DG 10-5-20', "n_pct": 10.0, "p_pct": 5.0, "k_pct": 20.0, "type": 'Granular', "minors": 'Fe 0.30%', "notes": "Professional golf/turf nutrition product; verify current product label/guaranteed analysis."},
    {"brand": 'The Andersons', "name": 'Contec DG 11-8-9', "n_pct": 11.0, "p_pct": 8.0, "k_pct": 9.0, "type": 'Granular', "minors": 'Mg 2.80%', "notes": "Professional golf/turf nutrition product; verify current product label/guaranteed analysis."},
    {"brand": 'The Andersons', "name": 'Contec DG 14-7-14', "n_pct": 14.0, "p_pct": 7.0, "k_pct": 14.0, "type": 'Granular', "minors": 'Fe 0.30%; Mn 0.14%', "notes": "Professional golf/turf nutrition product; verify current product label/guaranteed analysis."},
    {"brand": 'The Andersons', "name": 'Contec DG 15-0-15', "n_pct": 15.0, "p_pct": 0.0, "k_pct": 15.0, "type": 'Granular', "minors": 'Fe 0.30%; Mn 0.30%', "notes": "Professional golf/turf nutrition product; verify current product label/guaranteed analysis."},
    {"brand": 'The Andersons', "name": 'Contec DG 18-0-5', "n_pct": 18.0, "p_pct": 0.0, "k_pct": 5.0, "type": 'Granular', "minors": 'Mn 7.00%', "notes": "Professional golf/turf nutrition product; verify current product label/guaranteed analysis."},
    {"brand": 'The Andersons', "name": 'Contec DG 18-0-18', "n_pct": 18.0, "p_pct": 0.0, "k_pct": 18.0, "type": 'Granular', "minors": 'Fe 0.28%', "notes": "Professional golf/turf nutrition product; verify current product label/guaranteed analysis."},
    {"brand": 'The Andersons', "name": 'Contec DG 18-3-18', "n_pct": 18.0, "p_pct": 3.0, "k_pct": 18.0, "type": 'Granular', "minors": '', "notes": "Professional golf/turf nutrition product; verify current product label/guaranteed analysis."},
    {"brand": 'The Andersons', "name": 'Contec DG 18-24-5 Fairways', "n_pct": 18.0, "p_pct": 24.0, "k_pct": 5.0, "type": 'Granular', "minors": '', "notes": "Professional golf/turf nutrition product; verify current product label/guaranteed analysis."},
    {"brand": 'The Andersons', "name": 'Contec DG 20-0-16 Fairways', "n_pct": 20.0, "p_pct": 0.0, "k_pct": 16.0, "type": 'Granular', "minors": '', "notes": "Professional golf/turf nutrition product; verify current product label/guaranteed analysis."},
    {"brand": 'The Andersons', "name": 'Contec DG 24-0-10 Fairways', "n_pct": 24.0, "p_pct": 0.0, "k_pct": 10.0, "type": 'Granular', "minors": '', "notes": "Professional golf/turf nutrition product; verify current product label/guaranteed analysis."},
    {"brand": 'The Andersons', "name": 'Contec DG 28-0-6 Fairways', "n_pct": 28.0, "p_pct": 0.0, "k_pct": 6.0, "type": 'Granular', "minors": '', "notes": "Professional golf/turf nutrition product; verify current product label/guaranteed analysis."},
    {"brand": 'The Andersons', "name": 'HCU Humic Coated Urea 44-0-0', "n_pct": 44.0, "p_pct": 0.0, "k_pct": 0.0, "type": 'Granular', "minors": '', "notes": "Professional golf/turf nutrition product; verify current product label/guaranteed analysis."},
    {"brand": 'The Andersons', "name": 'SmartPhos DG 4-22-0', "n_pct": 4.0, "p_pct": 22.0, "k_pct": 0.0, "type": 'Granular', "minors": 'Mg 8.00%', "notes": "Professional golf/turf nutrition product; verify current product label/guaranteed analysis."},
    {"brand": 'Floratine', "name": 'P-48 10-48-8', "n_pct": 10.0, "p_pct": 48.0, "k_pct": 8.0, "type": 'Liquid Foliar', "minors": 'B 0.01%; Cu 0.04%; Fe 0.10%; Mn 0.04%; Zn 0.05%; Mo 0.0006%', "notes": "Professional golf/turf nutrition product; verify current product label/guaranteed analysis."},
    {"brand": 'Simplot', "name": 'Extreme Green 4.1L', "n_pct": 0.0, "p_pct": 0.0, "k_pct": 0.0, "type": 'Liquid Foliar', "minors": 'Fe 4.00%; Mn 1.00%; S 3.50%', "notes": "Professional golf/turf nutrition product; verify current product label/guaranteed analysis."},
    {"brand": 'Simplot BEST', "name": 'Turf Supreme Mini 16-6-8', "n_pct": 16.0, "p_pct": 6.0, "k_pct": 8.0, "type": 'Granular', "minors": 'S 16.00%; Fe 1.50%; Mn 0.20%', "notes": "Professional golf/turf nutrition product; verify current product label/guaranteed analysis."},
    {"brand": 'Simplot BEST', "name": 'NK Select 33-0-6', "n_pct": 33.0, "p_pct": 0.0, "k_pct": 6.0, "type": 'Granular', "minors": 'S 2.90%; Fe 3.25%', "notes": "Professional golf/turf nutrition product; verify current product label/guaranteed analysis."},
    {"brand": 'Simplot BEST', "name": '23-3-6 with Iron + UFLEXX', "n_pct": 23.0, "p_pct": 3.0, "k_pct": 6.0, "type": 'Granular', "minors": 'Fe 4.20%', "notes": "Professional golf/turf nutrition product; verify current product label/guaranteed analysis."},
    {"brand": 'LebanonTurf', "name": 'Country Club IV 0-0-25', "n_pct": 0.0, "p_pct": 0.0, "k_pct": 25.0, "type": 'Granular', "minors": 'Fe 2.00%; Mg 2.00%; Mn 2.00%', "notes": "Professional golf/turf nutrition product; verify current product label/guaranteed analysis."},
    {"brand": 'LebanonTurf', "name": 'Country Club IV 17-0-17', "n_pct": 17.0, "p_pct": 0.0, "k_pct": 17.0, "type": 'Granular', "minors": 'Fe 1.00%; Mn 0.50%', "notes": "Professional golf/turf nutrition product; verify current product label/guaranteed analysis."},
    {"brand": 'LebanonTurf', "name": 'Country Club IV 18-3-18', "n_pct": 18.0, "p_pct": 3.0, "k_pct": 18.0, "type": 'Granular', "minors": '', "notes": "Professional golf/turf nutrition product; verify current product label/guaranteed analysis."},
    {"brand": 'LebanonTurf', "name": 'Country Club IV 18-9-18', "n_pct": 18.0, "p_pct": 9.0, "k_pct": 18.0, "type": 'Granular', "minors": 'Fe 0.50%; Mg 0.50%; Mn 0.28%', "notes": "Professional golf/turf nutrition product; verify current product label/guaranteed analysis."},
    {"brand": 'LebanonTurf', "name": 'Country Club IV 24-3-12', "n_pct": 24.0, "p_pct": 3.0, "k_pct": 12.0, "type": 'Granular', "minors": '', "notes": "Professional golf/turf nutrition product; verify current product label/guaranteed analysis."},
    {"brand": 'LebanonTurf', "name": 'Country Club MD 18-0-18', "n_pct": 18.0, "p_pct": 0.0, "k_pct": 18.0, "type": 'Granular', "minors": 'Fe 1.60%; Mg 0.50%; Mn 0.50%', "notes": "Professional golf/turf nutrition product; verify current product label/guaranteed analysis."},
    {"brand": 'LebanonTurf', "name": 'Country Club Root Reviver 3-3-4', "n_pct": 3.0, "p_pct": 3.0, "k_pct": 4.0, "type": 'Granular', "minors": 'Fe 1.00%; Mg 1.50%', "notes": "Professional golf/turf nutrition product; verify current product label/guaranteed analysis."},
    {"brand": "Harrell's", "name": 'POLYON 16-0-8', "n_pct": 16.0, "p_pct": 0.0, "k_pct": 8.0, "type": 'Granular', "minors": '', "notes": "Professional golf/turf nutrition product; verify current product label/guaranteed analysis."},
    {"brand": "Harrell's", "name": 'POLYON 24-0-12', "n_pct": 24.0, "p_pct": 0.0, "k_pct": 12.0, "type": 'Granular', "minors": '', "notes": "Professional golf/turf nutrition product; verify current product label/guaranteed analysis."},
    {"brand": "Harrell's", "name": 'POLYON 32-0-14', "n_pct": 32.0, "p_pct": 0.0, "k_pct": 14.0, "type": 'Granular', "minors": '', "notes": "Professional golf/turf nutrition product; verify current product label/guaranteed analysis."},
]


def get_conn():
    conn = sqlite3.connect(DB_PATH)
    conn.row_factory = sqlite3.Row
    return conn


def init_db():
    conn = get_conn()
    cur = conn.cursor()

    # Header table for one application (date, area, weather, equipment)
    cur.execute("""
        CREATE TABLE IF NOT EXISTS applications (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            app_date TEXT NOT NULL,
            area_name TEXT,
            hole_number INTEGER,
            area_sqft REAL NOT NULL,
            weather TEXT,
            temp_f REAL,
            soil_temp_f REAL,
            wind_mph REAL,
            humidity_pct REAL,
            equipment TEXT,
            gpa REAL,
            nozzle TEXT,
            gear TEXT,
            speed TEXT,
            pressure TEXT,
            tanks REAL,
            gallons_per_tank REAL,
            bags REAL,
            rinsed TEXT,
            watered_in TEXT,
            watered_notes TEXT,
            notes TEXT,
            applicators TEXT,
            created_at TEXT DEFAULT CURRENT_TIMESTAMP
        )
    """)


    try:
        cur.execute("ALTER TABLE applications ADD COLUMN hole_number INTEGER")
    except Exception:
        pass

    # Line items - up to 12 products per application
    cur.execute("""
        CREATE TABLE IF NOT EXISTS application_products (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            application_id INTEGER NOT NULL,
            line_num INTEGER NOT NULL,
            brand TEXT,
            fertilizer_name TEXT,
            n_pct REAL NOT NULL DEFAULT 0,
            p_pct REAL NOT NULL DEFAULT 0,
            k_pct REAL NOT NULL DEFAULT 0,
            target_n_rate REAL,
            product_per_1000 REAL,
            product_lbs REAL NOT NULL DEFAULT 0,
            product_unit TEXT DEFAULT 'lbs',
            n_applied_lbs REAL NOT NULL DEFAULT 0,
            p_applied_lbs REAL NOT NULL DEFAULT 0,
            k_applied_lbs REAL NOT NULL DEFAULT 0,
            lot_number TEXT,
            minors TEXT,
            FOREIGN KEY (application_id) REFERENCES applications(id) ON DELETE CASCADE
        )
    """)
    # Migrate older DBs
    try:
        cur.execute("ALTER TABLE application_products ADD COLUMN product_unit TEXT DEFAULT 'lbs'")
    except Exception:
        pass

    cur.execute("""
        CREATE TABLE IF NOT EXISTS areas (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            name TEXT UNIQUE NOT NULL,
            sqft REAL NOT NULL,
            notes TEXT
        )
    """)

    cur.execute("""
        CREATE TABLE IF NOT EXISTS products (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            brand TEXT NOT NULL,
            name TEXT NOT NULL,
            n_pct REAL NOT NULL,
            p_pct REAL NOT NULL,
            k_pct REAL NOT NULL,
            type TEXT,
            notes TEXT,
            epa_reg_no TEXT DEFAULT '',
            active_ingredients TEXT DEFAULT '',
            resistance_group TEXT DEFAULT '',
            targets TEXT DEFAULT '',
            minors TEXT DEFAULT '',
            UNIQUE(brand, name)
        )
    """)

    # Add structured pesticide fields to older BoomBook databases without deleting data.
    cur.execute("PRAGMA table_info(products)")
    _product_cols = {row[1] for row in cur.fetchall()}
    for _col in ["epa_reg_no", "active_ingredients", "resistance_group", "targets", "minors"]:
        if _col not in _product_cols:
            cur.execute(f"ALTER TABLE products ADD COLUMN {_col} TEXT DEFAULT ''")

    # Keep the built-in catalog synchronized. INSERT OR IGNORE preserves user edits/custom products
    # while allowing newly bundled products to appear in an existing BoomBook database.
    for p in PRODUCT_CATALOG:
        cur.execute(
            "INSERT OR IGNORE INTO products (brand, name, n_pct, p_pct, k_pct, type, notes, minors) VALUES (?, ?, ?, ?, ?, ?, ?, ?)",
            (p["brand"], p["name"], p["n_pct"], p["p_pct"], p["k_pct"], p["type"], p["notes"], p.get("minors", ""))
        )
        if p.get("minors"):
            cur.execute(
                "UPDATE products SET minors=? WHERE brand=? AND name=? AND (minors IS NULL OR minors='')",
                (p["minors"], p["brand"], p["name"])
            )

    # Saved spray programs / favorite mixes. Stored separately from application history.
    cur.execute("""
        CREATE TABLE IF NOT EXISTS spray_programs (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            name TEXT UNIQUE NOT NULL,
            products_json TEXT NOT NULL,
            notes TEXT DEFAULT '',
            created_at TEXT DEFAULT CURRENT_TIMESTAMP,
            updated_at TEXT DEFAULT CURRENT_TIMESTAMP
        )
    """)

    # Searchable pesticide metadata. Rates are intentionally not stored here.
    _pesticide_meta = {
        "Posterity XT": ("100-1654", "Azoxystrobin 6%; Pydiflumetofen 1%; Propiconazole 10.1%", "FRAC 11 + 7 + 3", "Anthracnose; Brown patch; Dollar spot; Fairy ring; Gray leaf spot; Gray snow mold; Large patch; Leaf rust; Leaf spot; Melting out; Necrotic ring spot; Pink patch; Powdery mildew"),
        "Posterity XT / Headway XT": ("100-1654", "Azoxystrobin 6%; Pydiflumetofen 1%; Propiconazole 10.1%", "FRAC 11 + 7 + 3", "Anthracnose; Brown patch; Dollar spot; Fairy ring; Gray leaf spot; Gray snow mold; Large patch; Leaf rust; Leaf spot; Melting out; Necrotic ring spot; Pink patch; Powdery mildew"),
        "Posterity Peak": ("100-1654", "Azoxystrobin 6%; Pydiflumetofen 1%; Propiconazole 10.1%", "FRAC 11 + 7 + 3", "Anthracnose; Brown patch; Dollar spot; Fairy ring; Gray leaf spot; Gray snow mold; Large patch; Leaf rust; Leaf spot; Melting out; Necrotic ring spot; Pink patch; Powdery mildew"),
        "Encartis Fungicide": ("7969-348", "Chlorothalonil 54.14%; Boscalid 2.2%", "FRAC M05 + 7", "Algae; Anthracnose; Brown blight; Brown patch; Copper spot; Dollar spot; Gray leaf spot; Leaf spot; Melting out; Red thread; Stem rust"),
        "Azoxystrobin 2.08 lb SC": ("71532-35", "Azoxystrobin 22.9%", "FRAC 11", "Anthracnose; Brown patch; Dollar spot; Fairy ring; Gray leaf spot; Leaf spot; Rust"),
        "Azoxystrobin + Tebuconazole Turf Fungicide": ("60063-62", "Azoxystrobin 18.2%; Tebuconazole 27.3%", "FRAC 11 + 3", "Anthracnose; Brown patch; Dollar spot; Gray leaf spot; Leaf spot; Rust; Snow mold"),
        "Imidacloprid 0.5G Insecticide": ("42750-151", "Imidacloprid 0.5%", "IRAC 4A", "Annual bluegrass weevil larvae; ABW; Billbugs; Black turfgrass ataenius; Chinch bug; Cutworms; European crane fly larvae; Japanese beetle larvae; White grubs"),
        "QuickSilver T&O Herbicide": ("279-3265", "Carfentrazone-ethyl 21.3%", "HRAC 14", "Broadleaf weeds; Silvery thread moss"),
        "Certainty Turf Herbicide": ("59639-226", "Sulfosulfuron 75%", "HRAC 2", "Nutsedge; Kyllinga; Dallisgrass; Annual bluegrass; Selected broadleaf and grassy weeds"),
        "Specticle FLO": ("432-1608", "Indaziflam 7.4%", "HRAC 29", "Annual grasses; Annual broadleaf weeds; Crabgrass; Goosegrass; Annual bluegrass"),
        "Tribute Total": ("432-1519", "Thiencarbazone-methyl 9.9%; Foramsulfuron 19.8%; Halosulfuron-methyl 30.8%", "HRAC 2", "Annual bluegrass; Crabgrass; Goosegrass; Dallisgrass; Nutsedge; Kyllinga; Selected broadleaf weeds"),
        "Conserve SC Turf and Ornamental": ("62719-291", "Spinosad 11.6%", "IRAC 5", "Annual bluegrass weevil; ABW; Black cutworm; Black turfgrass ataenius; Armyworms; Sod webworms"),
        "0.067% Acelepryn Insecticide Plus Fertilizer": ("9198-247", "Chlorantraniliprole 0.067%", "IRAC 28", "Annual bluegrass weevil; ABW; Billbugs; Black turfgrass ataenius; White grubs; Caterpillars"),
        "Bifenthrin 7.9% SC": ("42750-369", "Bifenthrin 7.9%", "IRAC 3A", "Annual bluegrass weevil; ABW; Billbugs; Black turfgrass ataenius; Chinch bugs; Cutworms; Sod webworms"),
    }
    for _name, (_epa, _ai, _group, _targets) in _pesticide_meta.items():
        cur.execute(
            "UPDATE products SET epa_reg_no=?, active_ingredients=?, resistance_group=?, targets=? WHERE name=?",
            (_epa, _ai, _group, _targets, _name)
        )

    cur.execute("SELECT COUNT(*) FROM areas")
    if cur.fetchone()[0] == 0:
        for name, sqft, notes in [
            ("Greens", 25000, "All greens combined"),
            ("Tees", 15000, "All tees"),
            ("Fairways", 800000, "Main fairways"),
            ("Approaches", 40000, "Approach areas"),
            ("Rough", 500000, "Primary rough"),
            ("Driving Range", 100000, "Range turf"),
            ("Other", 10000, "Misc areas"),
        ]:
            cur.execute(
                "INSERT OR IGNORE INTO areas (name, sqft, notes) VALUES (?, ?, ?)",
                (name, sqft, notes)
            )

    conn.commit()
    conn.close()



SETTINGS_PATH = Path(__file__).parent / "boombook_settings.json"
DEFAULT_LAT = 34.9457
DEFAULT_LON = -111.6357

# Common TeeJet nozzle / tip families (from TeeJet spray tip catalog)
TEEJET_NOZZLES = [
    "",
    "XR TeeJet (Extended Range flat fan)",
    "XRC TeeJet (Extended Range compact)",
    "Turbo TeeJet (TT)",
    "Turbo TeeJet TwinJet (TTJ60)",
    "AIXR TeeJet (Air Induction XR)",
    "AI TeeJet (Air Induction)",
    "AIC TeeJet (Air Induction compact)",
    "AI3070 TeeJet (dual pattern AI)",
    "Air Induction Turbo TwinJet (AITTJ60)",
    "TTI TeeJet Induction (Turbo TeeJet Induction)",
    "TTI TwinJet (TTI60)",
    "AccuPulse TwinJet (APTJ) — PWM",
    "DG TeeJet (Drift Guard)",
    "Turbo FloodJet (TF)",
    "Flat Fan 80°",
    "Flat Fan 110°",
    "Dual fan / twin fan",
    "Boomless",
    "Other / custom",
]

EQUIPMENT_OPTIONS = [
    "",
    "Sprayer",
    "Lely",
    "Other",
]



def load_settings():
    if SETTINGS_PATH.exists():
        try:
            return json.loads(SETTINGS_PATH.read_text(encoding="utf-8"))
        except Exception:
            pass
    return {"lat": DEFAULT_LAT, "lon": DEFAULT_LON, "location_name": "Munds Park, AZ"}


def save_settings(data):
    SETTINGS_PATH.write_text(json.dumps(data, indent=2), encoding="utf-8")


def fetch_nws_weather(lat, lon):
    """Fetch current conditions from NWS (api.weather.gov, MapClick fallback)."""
    ua = {
        "User-Agent": "(TheBoomBook, pinewood-turf-log@local)",
        "Accept": "application/geo+json",
    }
    errors = []

    # --- Primary: api.weather.gov ---
    try:
        import ssl
        ctx = ssl.create_default_context()
        req = urllib.request.Request(
            f"https://api.weather.gov/points/{lat},{lon}", headers=ua
        )
        with urllib.request.urlopen(req, timeout=15, context=ctx) as r:
            points = json.loads(r.read().decode())
        stations_url = points["properties"]["observationStations"]
        req = urllib.request.Request(stations_url, headers=ua)
        with urllib.request.urlopen(req, timeout=15, context=ctx) as r:
            stations = json.loads(r.read().decode())
        if not stations.get("features"):
            raise RuntimeError("No observation stations near this point")
        sid = stations["features"][0]["properties"]["stationIdentifier"]
        sname = stations["features"][0]["properties"].get("name") or sid
        req = urllib.request.Request(
            f"https://api.weather.gov/stations/{sid}/observations/latest",
            headers=ua,
        )
        with urllib.request.urlopen(req, timeout=15, context=ctx) as r:
            obs = json.loads(r.read().decode())["properties"]

        def val(key):
            node = obs.get(key) or {}
            return node.get("value")

        temp_c = val("temperature")
        temp_f = round(temp_c * 9 / 5 + 32, 1) if temp_c is not None else None
        humidity = val("relativeHumidity")
        humidity = round(float(humidity), 0) if humidity is not None else None

        # windSpeed unit may be km/h or m/s depending on station feed
        wind_raw = val("windSpeed")
        wind_mph = None
        if wind_raw is not None:
            unit = ((obs.get("windSpeed") or {}).get("unitCode") or "").lower()
            if "km" in unit:
                wind_mph = round(float(wind_raw) * 0.621371, 1)
            elif "m_s" in unit or "m/s" in unit:
                wind_mph = round(float(wind_raw) * 2.23694, 1)
            else:
                # default treat as km/h (common on api.weather.gov)
                wind_mph = round(float(wind_raw) * 0.621371, 1)

        text = (obs.get("textDescription") or "").strip()
        if not text:
            text = _scrape_nws_conditions(lat, lon) or "Current observation"

        if temp_f is None and humidity is None:
            raise RuntimeError("Observation had no temperature or humidity")

        return {
            "temp_f": temp_f,
            "humidity": humidity,
            "wind_mph": wind_mph if wind_mph is not None else 0.0,
            "conditions": text,
            "station": sname,
            "updated": (obs.get("timestamp") or "")[:19].replace("T", " "),
            "error": None,
        }
    except Exception as e:
        errors.append(f"API: {e}")

    # --- Fallback: scrape MapClick page ---
    try:
        scraped = _scrape_nws_mapclick(lat, lon)
        if scraped and (scraped.get("temp_f") is not None or scraped.get("humidity") is not None):
            scraped["error"] = None
            scraped["note"] = "; ".join(errors) if errors else None
            return scraped
        errors.append("Scrape: no data parsed")
    except Exception as e:
        errors.append(f"Scrape: {e}")

    return {"error": " | ".join(errors), "temp_f": None}


def _scrape_nws_conditions(lat, lon):
    try:
        data = _scrape_nws_mapclick(lat, lon)
        return (data or {}).get("conditions")
    except Exception:
        return None


def _scrape_nws_mapclick(lat, lon):
    """Parse current conditions from forecast.weather.gov MapClick HTML."""
    url = f"https://forecast.weather.gov/MapClick.php?lat={lat}&lon={lon}"
    req = urllib.request.Request(
        url,
        headers={
            "User-Agent": "Mozilla/5.0 (compatible; TheBoomBook/1.0)",
            "Accept": "text/html",
        },
    )
    import ssl
    ctx = ssl.create_default_context()
    with urllib.request.urlopen(req, timeout=15, context=ctx) as r:
        html = r.read().decode("utf-8", errors="ignore")

    result = {"station": "NWS MapClick", "wind_mph": 0.0}

    m = re.search(r'myforecast-current-lrg[^>]*>\s*([0-9.]+)', html, re.I)
    if not m:
        m = re.search(r'([0-9]{1,3})\s*&deg;F', html)
    if m:
        result["temp_f"] = float(m.group(1))

    m = re.search(r'Humidity.*?>\s*([0-9.]+)\s*%', html, re.I | re.S)
    if m:
        result["humidity"] = float(m.group(1))

    m = re.search(r'Wind Speed.*?>\s*(?:[A-Za-z ]*)?([0-9.]+)\s*MPH', html, re.I | re.S)
    if m:
        result["wind_mph"] = float(m.group(1))

    m = re.search(r'class="myforecast-current"[^>]*>\s*([^<]{2,40})\s*<', html, re.I)
    if m:
        cond = re.sub(r'\s+', ' ', m.group(1)).strip()
        if cond.upper() != "NA":
            result["conditions"] = cond
    if not result.get("conditions"):
        m = re.search(r'<p class="short-desc">\s*(.*?)\s*</p>', html, re.I | re.S)
        if m:
            result["conditions"] = re.sub(r'<[^>]+>', '', m.group(1))
            result["conditions"] = re.sub(r'\s+', ' ', result["conditions"]).strip()
    if not result.get("conditions"):
        result["conditions"] = "Current observation"

    m = re.search(r'Last update</[^>]*>\s*([^<]+)', html, re.I)
    if m:
        result["updated"] = re.sub(r'\s+', ' ', m.group(1)).strip()[:40]

    if "temp_f" not in result and "humidity" not in result:
        return None
    return result


def calculate_rates(area_sqft, target_n, n_pct, p_pct, k_pct):
    if n_pct <= 0:
        return None
    if area_sqft <= 0:
        raise ValueError("Area must be greater than 0")
    product_per_1000 = target_n / (n_pct / 100.0)
    total_product = product_per_1000 * (area_sqft / 1000.0)
    return {
        "product_per_1000": round(product_per_1000, 3),
        "total_product": round(total_product, 2),
        "n_applied": round(total_product * (n_pct / 100.0), 3),
        "p_applied": round(total_product * (p_pct / 100.0), 3),
        "k_applied": round(total_product * (k_pct / 100.0), 3),
        "n_per_1000": round(target_n, 3),
        "p_per_1000": round(product_per_1000 * (p_pct / 100.0), 3),
        "k_per_1000": round(product_per_1000 * (k_pct / 100.0), 3),
    }


def automatic_database_backup():
    """Create at most one automatic DB backup per day and retain the newest 30."""
    backup_dir = DATA_DIR / "Backups"
    backup_dir.mkdir(parents=True, exist_ok=True)
    today_prefix = f"turf_fertilizer_{date.today().isoformat()}_"
    existing_today = list(backup_dir.glob(today_prefix + "*.db"))
    if not existing_today and DB_PATH.exists():
        stamp = datetime.now().strftime("%Y-%m-%d_%H%M%S")
        dest = backup_dir / f"turf_fertilizer_{stamp}.db"
        shutil.copy2(DB_PATH, dest)
    backups = sorted(backup_dir.glob("turf_fertilizer_*.db"), key=lambda x: x.stat().st_mtime, reverse=True)
    for old in backups[30:]:
        try: old.unlink()
        except Exception: pass
    backups = sorted(backup_dir.glob("turf_fertilizer_*.db"), key=lambda x: x.stat().st_mtime, reverse=True)
    return backups[0] if backups else None


def get_dashboard_data():
    conn = get_conn()
    apps = pd.read_sql_query("""
        SELECT a.id, a.app_date, a.area_name, a.hole_number, a.area_sqft,
               COUNT(p.id) AS num_products,
               COALESCE(SUM(p.n_applied_lbs),0) AS total_n,
               COALESCE(SUM(p.p_applied_lbs),0) AS total_p,
               COALESCE(SUM(p.k_applied_lbs),0) AS total_k
        FROM applications a LEFT JOIN application_products p ON p.application_id=a.id
        GROUP BY a.id ORDER BY a.app_date DESC, a.id DESC
    """, conn)
    conn.close()
    if apps.empty:
        return apps, apps, apps
    apps["date"] = pd.to_datetime(apps["app_date"], errors="coerce")
    now = pd.Timestamp.today()
    month = apps[(apps["date"].dt.year==now.year) & (apps["date"].dt.month==now.month)]
    year = apps[apps["date"].dt.year==now.year]
    return apps, month, year


def get_spray_programs():
    conn=get_conn()
    df=pd.read_sql_query("SELECT * FROM spray_programs ORDER BY name", conn)
    conn.close(); return df


def save_spray_program(name, products, notes=""):
    conn=get_conn()
    conn.execute("""INSERT INTO spray_programs(name,products_json,notes,updated_at)
                    VALUES(?,?,?,CURRENT_TIMESTAMP)
                    ON CONFLICT(name) DO UPDATE SET products_json=excluded.products_json,
                    notes=excluded.notes, updated_at=CURRENT_TIMESTAMP""",
                 (name.strip(), json.dumps(products), notes))
    conn.commit(); conn.close()


def delete_spray_program(program_id):
    conn=get_conn(); conn.execute("DELETE FROM spray_programs WHERE id=?",(program_id,)); conn.commit(); conn.close()


def application_to_cart(data):
    cart=[]
    for p in data.get("products",[]):
        cart.append({
            "brand": p.get("brand") or "", "fertilizer_name": p.get("fertilizer_name") or "",
            "n_pct": p.get("n_pct") or 0, "p_pct": p.get("p_pct") or 0, "k_pct": p.get("k_pct") or 0,
            "target_n_rate": p.get("target_n_rate") or 0, "product_per_1000": p.get("product_per_1000") or 0,
            "product_lbs": p.get("product_lbs") or 0, "product_unit": p.get("product_unit") or "lbs",
            "n_applied_lbs": p.get("n_applied_lbs") or 0, "p_applied_lbs": p.get("p_applied_lbs") or 0,
            "k_applied_lbs": p.get("k_applied_lbs") or 0, "lot_number": p.get("lot_number") or "",
            "minors": p.get("minors") or "",
        })
    return cart


def get_products(brand_filter=None):
    conn = get_conn()
    if brand_filter and brand_filter != "All":
        df = pd.read_sql_query(
            "SELECT * FROM products WHERE brand = ? ORDER BY brand, name",
            conn, params=(brand_filter,)
        )
    else:
        df = pd.read_sql_query("SELECT * FROM products ORDER BY brand, name", conn)
    conn.close()
    return df


def get_areas():
    conn = get_conn()
    df = pd.read_sql_query("SELECT * FROM areas ORDER BY name", conn)
    conn.close()
    return df


def save_area(name, sqft, notes=""):
    conn = get_conn()
    conn.execute("INSERT OR REPLACE INTO areas (name, sqft, notes) VALUES (?, ?, ?)", (name, sqft, notes))
    conn.commit()
    conn.close()


def delete_area(area_id):
    conn = get_conn()
    conn.execute("DELETE FROM areas WHERE id = ?", (area_id,))
    conn.commit()
    conn.close()


def add_custom_product(brand, name, n, p, k, ptype, notes):
    conn = get_conn()
    conn.execute(
        "INSERT OR REPLACE INTO products (brand, name, n_pct, p_pct, k_pct, type, notes) VALUES (?, ?, ?, ?, ?, ?, ?)",
        (brand, name, n, p, k, ptype, notes)
    )
    conn.commit()
    conn.close()


def save_application(header, products_list):
    """Save one application with multiple product lines (up to 12)."""
    conn = get_conn()
    cur = conn.cursor()
    cur.execute("""
        INSERT INTO applications (
            app_date, area_name, hole_number, area_sqft, weather, temp_f, soil_temp_f,
            wind_mph, humidity_pct, equipment, gpa, nozzle, gear, speed, pressure,
            tanks, gallons_per_tank, bags, rinsed, watered_in, watered_notes,
            notes, applicators
        ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
    """, (
        header["app_date"], header["area_name"], header.get("hole_number"),
        header["area_sqft"],
        header.get("weather"), header.get("temp_f"), header.get("soil_temp_f"),
        header.get("wind_mph"), header.get("humidity_pct"),
        header.get("equipment"), header.get("gpa"), header.get("nozzle"),
        header.get("gear"), header.get("speed"), header.get("pressure"),
        header.get("tanks"), header.get("gallons_per_tank"), header.get("bags"),
        header.get("rinsed"), header.get("watered_in"), header.get("watered_notes"),
        header.get("notes"), header.get("applicators"),
    ))
    app_id = cur.lastrowid
    for i, prod in enumerate(products_list[:12], start=1):
        cur.execute("""
            INSERT INTO application_products (
                application_id, line_num, brand, fertilizer_name,
                n_pct, p_pct, k_pct, target_n_rate, product_per_1000,
                product_lbs, product_unit, n_applied_lbs, p_applied_lbs, k_applied_lbs,
                lot_number, minors
            ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
        """, (
            app_id, i, prod.get("brand", ""), prod.get("fertilizer_name", ""),
            prod.get("n_pct", 0), prod.get("p_pct", 0), prod.get("k_pct", 0),
            prod.get("target_n_rate"), prod.get("product_per_1000"),
            prod.get("product_lbs", 0), prod.get("product_unit", "lbs"),
            prod.get("n_applied_lbs", 0),
            prod.get("p_applied_lbs", 0), prod.get("k_applied_lbs", 0),
            prod.get("lot_number", ""), prod.get("minors", ""),
        ))
    conn.commit()
    conn.close()
    return app_id


def get_applications_summary():
    conn = get_conn()
    df = pd.read_sql_query("""
        SELECT a.id, a.app_date, a.area_name, a.hole_number, a.area_sqft, a.weather, a.temp_f,
               COUNT(p.id) as num_products,
               COALESCE(SUM(p.product_lbs), 0) as total_product_lbs,
               COALESCE(SUM(p.n_applied_lbs), 0) as total_n,
               COALESCE(SUM(p.p_applied_lbs), 0) as total_p,
               COALESCE(SUM(p.k_applied_lbs), 0) as total_k
        FROM applications a
        LEFT JOIN application_products p ON p.application_id = a.id
        GROUP BY a.id
        ORDER BY a.app_date DESC, a.id DESC
    """, conn)
    conn.close()
    return df


def get_application_full(app_id):
    conn = get_conn()
    cur = conn.cursor()
    cur.execute("SELECT * FROM applications WHERE id = ?", (app_id,))
    row = cur.fetchone()
    if not row:
        conn.close()
        return None
    header = dict(row)
    cur.execute(
        "SELECT * FROM application_products WHERE application_id = ? ORDER BY line_num",
        (app_id,)
    )
    products = [dict(r) for r in cur.fetchall()]
    conn.close()
    return {"header": header, "products": products}


def delete_application(app_id):
    conn = get_conn()
    conn.execute("DELETE FROM application_products WHERE application_id = ?", (app_id,))
    conn.execute("DELETE FROM applications WHERE id = ?", (app_id,))
    conn.commit()
    conn.close()


def update_application(app_id, header, products_list):
    """Replace header + product lines for an existing application."""
    conn = get_conn()
    cur = conn.cursor()
    cur.execute("""
        UPDATE applications SET
            app_date=?, area_name=?, hole_number=?, area_sqft=?, weather=?,
            temp_f=?, soil_temp_f=?, wind_mph=?, humidity_pct=?,
            equipment=?, gpa=?, nozzle=?, gear=?, speed=?, pressure=?,
            tanks=?, gallons_per_tank=?, bags=?, rinsed=?, watered_in=?,
            watered_notes=?, notes=?, applicators=?
        WHERE id=?
    """, (
        header["app_date"], header["area_name"], header.get("hole_number"),
        header["area_sqft"], header.get("weather"),
        header.get("temp_f"), header.get("soil_temp_f"),
        header.get("wind_mph"), header.get("humidity_pct"),
        header.get("equipment"), header.get("gpa"), header.get("nozzle"),
        header.get("gear"), header.get("speed"), header.get("pressure"),
        header.get("tanks"), header.get("gallons_per_tank"), header.get("bags"),
        header.get("rinsed"), header.get("watered_in"), header.get("watered_notes"),
        header.get("notes"), header.get("applicators"),
        app_id,
    ))
    cur.execute("DELETE FROM application_products WHERE application_id = ?", (app_id,))
    for i, prod in enumerate(products_list[:12], start=1):
        cur.execute("""
            INSERT INTO application_products (
                application_id, line_num, brand, fertilizer_name,
                n_pct, p_pct, k_pct, target_n_rate, product_per_1000,
                product_lbs, product_unit, n_applied_lbs, p_applied_lbs, k_applied_lbs,
                lot_number, minors
            ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
        """, (
            app_id, i,
            prod.get("brand"), prod.get("fertilizer_name"),
            prod.get("n_pct"), prod.get("p_pct"), prod.get("k_pct"),
            prod.get("target_n_rate"), prod.get("product_per_1000"),
            prod.get("product_lbs"), prod.get("product_unit") or "lbs",
            prod.get("n_applied_lbs"), prod.get("p_applied_lbs"), prod.get("k_applied_lbs"),
            prod.get("lot_number"), prod.get("minors"),
        ))
    conn.commit()
    conn.close()
    return app_id



def get_period_npk_totals(app_date_str):
    """Return monthly and yearly N/P/K totals (lbs) through app_date inclusive."""
    conn = get_conn()
    cur = conn.cursor()
    # Normalize date
    d = (app_date_str or "")[:10]
    if not d:
        conn.close()
        return {"month": {"n": 0, "p": 0, "k": 0}, "year": {"n": 0, "p": 0, "k": 0}}

    year = d[:4]
    month = d[:7]  # YYYY-MM

    cur.execute("""
        SELECT
            COALESCE(SUM(p.n_applied_lbs), 0),
            COALESCE(SUM(p.p_applied_lbs), 0),
            COALESCE(SUM(p.k_applied_lbs), 0)
        FROM application_products p
        JOIN applications a ON a.id = p.application_id
        WHERE strftime('%Y-%m', a.app_date) = ?
          AND a.app_date <= ?
    """, (month, d))
    row = cur.fetchone()
    month_tots = {"n": float(row[0] or 0), "p": float(row[1] or 0), "k": float(row[2] or 0)}

    cur.execute("""
        SELECT
            COALESCE(SUM(p.n_applied_lbs), 0),
            COALESCE(SUM(p.p_applied_lbs), 0),
            COALESCE(SUM(p.k_applied_lbs), 0)
        FROM application_products p
        JOIN applications a ON a.id = p.application_id
        WHERE strftime('%Y', a.app_date) = ?
          AND a.app_date <= ?
    """, (year, d))
    row = cur.fetchone()
    year_tots = {"n": float(row[0] or 0), "p": float(row[1] or 0), "k": float(row[2] or 0)}
    conn.close()
    return {"month": month_tots, "year": year_tots, "month_label": month, "year_label": year}


def generate_application_sheet_pdf(app_data):
    header = app_data["header"]
    products = app_data["products"]
    buffer = BytesIO()
    doc = SimpleDocTemplate(
        buffer, pagesize=letter,
        leftMargin=0.5 * inch, rightMargin=0.5 * inch,
        topMargin=0.4 * inch, bottomMargin=0.4 * inch
    )
    styles = getSampleStyleSheet()
    title_style = ParagraphStyle(
        "TitleCustom", parent=styles["Heading1"],
        fontSize=14, spaceAfter=2, alignment=1, textColor=colors.HexColor("#1a3c2a")
    )
    sub_style = ParagraphStyle("Sub", parent=styles["Normal"], fontSize=8, alignment=1, spaceAfter=4)
    section = ParagraphStyle(
        "Section", parent=styles["Normal"], fontSize=9, spaceBefore=6, spaceAfter=3,
        textColor=colors.HexColor("#1a3c2a"), fontName="Helvetica-Bold"
    )
    normal = ParagraphStyle("Normal8", parent=styles["Normal"], fontSize=8)
    small = ParagraphStyle("Small7", parent=styles["Normal"], fontSize=7)

    story = []
    story.append(Paragraph("THE BOOMBOOK APPLICATION RECORD", title_style))
    story.append(Paragraph("Pinewood Country Club", sub_style))
    story.append(Paragraph(
        "395 E. Pinewood Blvd., Munds Park, AZ 86017 &nbsp;&nbsp; (928) 286-1370",
        sub_style
    ))
    story.append(HRFlowable(width="100%", thickness=1, color=colors.HexColor("#1a3c2a"), spaceAfter=6))

    area = (header.get("area_name") or "").lower()
    def chk(name):
        return "X" if name in area else "____"

    story.append(Paragraph(
        f"<b>Date of Application:</b> {header.get('app_date', '__________')} &nbsp;&nbsp;&nbsp; "
        f"<b>Application ID:</b> {header.get('id', '')}",
        normal
    ))
    story.append(Spacer(1, 3))
    hole = header.get("hole_number")
    hole_s = f"Hole {int(hole)}" if hole else "Hole ____"
    story.append(Paragraph(
        f"<b>Area Treated:</b>  Greens {chk('green')} &nbsp; Tees {chk('tee')} &nbsp; "
        f"FWYs {chk('fairway')} &nbsp; Rough {chk('rough')} &nbsp; "
        f"DR {chk('driving') or chk('range')} &nbsp; Other {chk('other') or '____'}"
        f" &nbsp;&nbsp; <b>{hole_s}</b>",
        normal
    ))
    story.append(Spacer(1, 4))

    story.append(Paragraph("APPLICATION INFORMATION:", section))
    eq = header.get("equipment") or "____________________"
    gpa = header.get("gpa")
    gpa_s = f"{gpa:g}" if gpa is not None else "______"
    nozzle = header.get("nozzle") or "______________"
    gear = header.get("gear") or "______"
    speed = header.get("speed") or "______"
    pressure = header.get("pressure") or "__________"
    tanks = header.get("tanks")
    tanks_s = f"{tanks:g}" if tanks is not None else "______"
    gpt = header.get("gallons_per_tank")
    gpt_s = f"{gpt:g}" if gpt is not None else "______"
    bags = header.get("bags")
    bags_s = f"{bags:g}" if bags is not None else "______"
    area_sqft = header.get("area_sqft") or 0

    story.append(Paragraph(
        f"<b>Equipment Used:</b> {eq} &nbsp;&nbsp; <b>GPA:</b> {gpa_s} &nbsp;&nbsp; <b>Nozzle Type:</b> {nozzle}",
        normal
    ))
    story.append(Paragraph(
        f"<b>Calibration:</b> GEAR: {gear} &nbsp; SPEED: {speed} &nbsp; PRESSURE / SETTING: {pressure}",
        normal
    ))
    story.append(Paragraph(
        f"<b>Total Area Treated:</b> {area_sqft:,.0f} sq ft &nbsp;&nbsp; "
        f"<b>Total # of Tanks:</b> {tanks_s} &nbsp;&nbsp; "
        f"<b>Gallons/Tank:</b> {gpt_s} &nbsp;&nbsp; <b>Bags:</b> {bags_s}",
        normal
    ))
    story.append(Spacer(1, 6))

    # Product table - only rows for products actually used (max 12)
    table_header = [
        Paragraph("<b>#</b>", small),
        Paragraph("<b>Fertilizer Name / Analysis / Bags-Jugs</b>", small),
        Paragraph("<b>Total Product</b>", small),
        Paragraph("<b>Rate / 1000</b>", small),
        Paragraph("<b>Unit</b>", small),
        Paragraph("<b>N</b>", small),
        Paragraph("<b>P</b>", small),
        Paragraph("<b>K</b>", small),
        Paragraph("<b>Minors</b>", small),
    ]
    rows = [table_header]
    total_n = total_p = total_k = 0.0
    total_by_unit = {"lbs": 0.0, "gal": 0.0, "fl oz": 0.0}
    # Prefer line_num order; fall back to list order
    ordered = sorted(
        [p for p in products if p],
        key=lambda p: (p.get("line_num") is None, p.get("line_num") or 0),
    )
    for i, prod in enumerate(ordered[:12], start=1):
        name = f"{prod.get('brand') or ''} {prod.get('fertilizer_name') or ''}".strip()
        analysis = f"{prod.get('n_pct', 0):g}-{prod.get('p_pct', 0):g}-{prod.get('k_pct', 0):g}"
        lot = prod.get("lot_number") or ""
        label = f"{name}<br/>{analysis}" + (f"<br/>{lot}" if lot else "")
        amount = prod.get("product_lbs") or 0
        rate = prod.get("product_per_1000") or 0
        unit = (prod.get("product_unit") or "lbs").strip().lower()
        if unit in ("gallon", "gallons", "gals"):
            unit = "gal"
        elif unit in ("oz", "floz", "fl. oz", "fl.oz", "fluid oz", "fluid ounces"):
            unit = "fl oz"
        elif unit not in total_by_unit:
            unit = "lbs"
        n = prod.get("n_applied_lbs") or 0
        p = prod.get("p_applied_lbs") or 0
        k = prod.get("k_applied_lbs") or 0
        total_n += n
        total_p += p
        total_k += k
        total_by_unit[unit] = total_by_unit.get(unit, 0.0) + float(amount)
        if unit == "gal":
            amt_str = f"{amount:.2f} gal"
        elif unit == "fl oz":
            amt_str = f"{amount:.1f} fl oz"
        else:
            amt_str = f"{amount:.1f} lbs"
        rows.append([
            Paragraph(str(i), small),
            Paragraph(label, small),
            Paragraph(amt_str, small),
            Paragraph(f"{rate:.2f}", small),
            Paragraph(unit, small),
            Paragraph(f"{n:.2f}", small),
            Paragraph(f"{p:.2f}", small),
            Paragraph(f"{k:.2f}", small),
            Paragraph(prod.get("minors") or "", small),
        ])

    # Build separated product totals (do not mix lbs + gal + fl oz)
    total_parts = []
    if total_by_unit.get("lbs", 0):
        total_parts.append(f"{total_by_unit['lbs']:.1f} lbs")
    if total_by_unit.get("gal", 0):
        total_parts.append(f"{total_by_unit['gal']:.2f} gal")
    if total_by_unit.get("fl oz", 0):
        total_parts.append(f"{total_by_unit['fl oz']:.1f} fl oz")
    total_prod_str = "<br/>".join(total_parts) if total_parts else "0"

    rows.append([
        Paragraph("", small),
        Paragraph("<b>Total</b>", small),
        Paragraph(f"<b>{total_prod_str}</b>", small),
        Paragraph("", small),
        Paragraph("", small),
        Paragraph(f"<b>{total_n:.2f}</b>", small),
        Paragraph(f"<b>{total_p:.2f}</b>", small),
        Paragraph(f"<b>{total_k:.2f}</b>", small),
        Paragraph("", small),
    ])

    col_w = [0.3*inch, 2.3*inch, 0.95*inch, 0.6*inch, 0.5*inch, 0.55*inch, 0.55*inch, 0.55*inch, 0.7*inch]
    t3 = Table(rows, colWidths=col_w)
    t3.setStyle(TableStyle([
        ("FONTSIZE", (0, 0), (-1, -1), 7),
        ("GRID", (0, 0), (-1, -1), 0.4, colors.grey),
        ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#e8f0e8")),
        ("BACKGROUND", (0, -1), (-1, -1), colors.HexColor("#f4f9f4")),
        ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
        ("ALIGN", (2, 0), (-1, -1), "CENTER"),
        ("TOPPADDING", (0, 0), (-1, -1), 2),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 2),
    ]))
    story.append(t3)
    story.append(Paragraph(
        "Note: Product totals are separated by unit (lbs / gal / fl oz). N-P-K totals are always in lbs.",
        small
    ))
    story.append(Spacer(1, 4))

    rinsed = header.get("rinsed") or "____"
    watered = header.get("watered_in") or "____"
    watered_notes = header.get("watered_notes") or "____________________"
    story.append(Paragraph(
        f"<b>Equipment & Containers Triple Rinsed & Disposed of Properly?</b>  {rinsed}",
        normal
    ))
    story.append(Paragraph(
        f"<b>Watered In?</b>  {watered} &nbsp;&nbsp; <b>How Long & When?</b> {watered_notes}",
        normal
    ))
    story.append(Spacer(1, 4))

    story.append(Paragraph("WEATHER CONDITIONS AT TIME OF APPLICATION:", section))
    temp = header.get("temp_f")
    soil = header.get("soil_temp_f")
    wind = header.get("wind_mph")
    humid = header.get("humidity_pct")
    weather = header.get("weather") or "__________"
    temp_s = f"{temp:.0f} F" if temp is not None else "______"
    soil_s = f"{soil:.0f} F" if soil is not None else "______"
    wind_s = f"{wind:.0f} mph" if wind is not None else "______"
    humid_s = f"{humid:.0f} %" if humid is not None else "______"
    story.append(Paragraph(
        f"<b>Air Temp:</b> {temp_s} &nbsp; <b>Soil Temp:</b> {soil_s} &nbsp; "
        f"<b>Wind:</b> {wind_s} &nbsp; <b>Humidity:</b> {humid_s} &nbsp; <b>Conditions:</b> {weather}",
        normal
    ))
    story.append(Spacer(1, 4))

    story.append(Paragraph("OBSERVATIONS & COMMENTS:", section))
    notes = header.get("notes") or ""
    story.append(Paragraph(notes.replace("\\n", "<br/>") if notes else "_______________________________________________", normal))
    story.append(Paragraph("_______________________________________________", normal))
    story.append(Spacer(1, 6))

    apps = header.get("applicators") or "________________________________"
    story.append(Paragraph(f"<b>Applicators:</b> {apps}", normal))
    story.append(Spacer(1, 4))
    story.append(Paragraph("<b>Calibration & Product Double Check Signature:</b> ________________________________", normal))
    story.append(Spacer(1, 6))
    story.append(Paragraph(
        "Monthly and Yearly totals below are auto-calculated from all applications logged through this application date.",
        normal
    ))
    story.append(Spacer(1, 8))

    # Auto-populate monthly and yearly N-P-K totals from the log
    period = get_period_npk_totals(str(header.get("app_date") or ""))
    mn, mp, mk = period["month"]["n"], period["month"]["p"], period["month"]["k"]
    yn, yp, yk = period["year"]["n"], period["year"]["p"], period["year"]["k"]
    month_label = period.get("month_label") or "Month"
    year_label = period.get("year_label") or "Year"

    totals_data = [
        [
            Paragraph(f"<b>Monthly Totals (lbs)</b><br/><font size='6'>{month_label} (through this app)</font>", small),
            Paragraph(f"<b>N</b><br/>{mn:.2f}", small),
            Paragraph(f"<b>P</b><br/>{mp:.2f}", small),
            Paragraph(f"<b>K</b><br/>{mk:.2f}", small),
        ],
        [
            Paragraph(f"<b>Yearly Totals (lbs)</b><br/><font size='6'>{year_label} YTD (through this app)</font>", small),
            Paragraph(f"<b>N</b><br/>{yn:.2f}", small),
            Paragraph(f"<b>P</b><br/>{yp:.2f}", small),
            Paragraph(f"<b>K</b><br/>{yk:.2f}", small),
        ],
    ]
    t4 = Table(totals_data, colWidths=[2.2*inch, 1.1*inch, 1.1*inch, 1.1*inch])
    t4.setStyle(TableStyle([
        ("FONTSIZE", (0, 0), (-1, -1), 8),
        ("GRID", (0, 0), (-1, -1), 0.5, colors.grey),
        ("BACKGROUND", (0, 0), (0, -1), colors.HexColor("#e8f0e8")),
        ("BACKGROUND", (1, 0), (-1, -1), colors.HexColor("#f4f9f4")),
        ("ALIGN", (1, 0), (-1, -1), "CENTER"),
        ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
        ("TOPPADDING", (0, 0), (-1, -1), 4),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 4),
    ]))
    story.append(t4)
    story.append(Spacer(1, 10))
    story.append(Paragraph(
        "Generated " + datetime.now().strftime("%Y-%m-%d %H:%M") + "  |  The BoomBook",
        ParagraphStyle("Footer", parent=small, textColor=colors.grey, alignment=1)
    ))

    doc.build(story)
    buffer.seek(0)
    return buffer.getvalue()




def get_yearly_totals_by_area(year=None):
    """N-P-K and product totals by area for a given year (or all years).
    Product amounts are separated by unit: lbs, gal, fl oz.
    """
    conn = get_conn()
    # Normalize unit with CASE so mixed labels still group correctly
    unit_lbs = """COALESCE(SUM(CASE
        WHEN LOWER(COALESCE(p.product_unit,'lbs')) IN ('lbs','lb','pounds') THEN p.product_lbs
        WHEN LOWER(COALESCE(p.product_unit,'lbs')) IN ('gal','gallon','gallons','gals','fl oz','oz','floz','fl. oz','fluid oz') THEN 0
        ELSE p.product_lbs END), 0)"""
    unit_gal = """COALESCE(SUM(CASE
        WHEN LOWER(COALESCE(p.product_unit,'')) IN ('gal','gallon','gallons','gals') THEN p.product_lbs
        ELSE 0 END), 0)"""
    unit_oz = """COALESCE(SUM(CASE
        WHEN LOWER(COALESCE(p.product_unit,'')) IN ('fl oz','oz','floz','fl. oz','fl.oz','fluid oz','fluid ounces') THEN p.product_lbs
        ELSE 0 END), 0)"""

    if year:
        df = pd.read_sql_query(f"""
            SELECT a.area_name,
                   COUNT(DISTINCT a.id) as applications,
                   {unit_lbs} as product_lbs,
                   {unit_gal} as product_gal,
                   {unit_oz} as product_fl_oz,
                   COALESCE(SUM(p.n_applied_lbs), 0) as n_lbs,
                   COALESCE(SUM(p.p_applied_lbs), 0) as p_lbs,
                   COALESCE(SUM(p.k_applied_lbs), 0) as k_lbs
            FROM applications a
            LEFT JOIN application_products p ON p.application_id = a.id
            WHERE strftime('%Y', a.app_date) = ?
            GROUP BY a.area_name
            ORDER BY a.area_name
        """, conn, params=(str(year),))
    else:
        df = pd.read_sql_query(f"""
            SELECT strftime('%Y', a.app_date) as year,
                   a.area_name,
                   COUNT(DISTINCT a.id) as applications,
                   {unit_lbs} as product_lbs,
                   {unit_gal} as product_gal,
                   {unit_oz} as product_fl_oz,
                   COALESCE(SUM(p.n_applied_lbs), 0) as n_lbs,
                   COALESCE(SUM(p.p_applied_lbs), 0) as p_lbs,
                   COALESCE(SUM(p.k_applied_lbs), 0) as k_lbs
            FROM applications a
            LEFT JOIN application_products p ON p.application_id = a.id
            GROUP BY year, a.area_name
            ORDER BY year DESC, a.area_name
        """, conn)
    conn.close()
    return df


def get_yearly_totals_by_product(year=None):
    """Product usage totals for a given year, amounts separated by unit."""
    conn = get_conn()
    unit_lbs = """COALESCE(SUM(CASE
        WHEN LOWER(COALESCE(p.product_unit,'lbs')) IN ('lbs','lb','pounds') THEN p.product_lbs
        WHEN LOWER(COALESCE(p.product_unit,'lbs')) IN ('gal','gallon','gallons','gals','fl oz','oz','floz','fl. oz','fluid oz') THEN 0
        ELSE p.product_lbs END), 0)"""
    unit_gal = """COALESCE(SUM(CASE
        WHEN LOWER(COALESCE(p.product_unit,'')) IN ('gal','gallon','gallons','gals') THEN p.product_lbs
        ELSE 0 END), 0)"""
    unit_oz = """COALESCE(SUM(CASE
        WHEN LOWER(COALESCE(p.product_unit,'')) IN ('fl oz','oz','floz','fl. oz','fl.oz','fluid oz','fluid ounces') THEN p.product_lbs
        ELSE 0 END), 0)"""
    # Also keep unit label for display when a product is only one unit
    unit_label = """MAX(COALESCE(p.product_unit, 'lbs'))"""

    if year:
        df = pd.read_sql_query(f"""
            SELECT COALESCE(p.brand, '') as brand,
                   p.fertilizer_name,
                   COUNT(*) as times_used,
                   {unit_label} as primary_unit,
                   {unit_lbs} as product_lbs,
                   {unit_gal} as product_gal,
                   {unit_oz} as product_fl_oz,
                   COALESCE(SUM(p.n_applied_lbs), 0) as n_lbs,
                   COALESCE(SUM(p.p_applied_lbs), 0) as p_lbs,
                   COALESCE(SUM(p.k_applied_lbs), 0) as k_lbs
            FROM application_products p
            JOIN applications a ON a.id = p.application_id
            WHERE strftime('%Y', a.app_date) = ?
            GROUP BY p.brand, p.fertilizer_name
            ORDER BY product_lbs DESC, product_gal DESC, product_fl_oz DESC
        """, conn, params=(str(year),))
    else:
        df = pd.read_sql_query(f"""
            SELECT strftime('%Y', a.app_date) as year,
                   COALESCE(p.brand, '') as brand,
                   p.fertilizer_name,
                   COUNT(*) as times_used,
                   {unit_label} as primary_unit,
                   {unit_lbs} as product_lbs,
                   {unit_gal} as product_gal,
                   {unit_oz} as product_fl_oz,
                   COALESCE(SUM(p.n_applied_lbs), 0) as n_lbs,
                   COALESCE(SUM(p.p_applied_lbs), 0) as p_lbs,
                   COALESCE(SUM(p.k_applied_lbs), 0) as k_lbs
            FROM application_products p
            JOIN applications a ON a.id = p.application_id
            GROUP BY year, p.brand, p.fertilizer_name
            ORDER BY year DESC, product_lbs DESC
        """, conn)
    conn.close()
    return df


def get_available_years():
    conn = get_conn()
    df = pd.read_sql_query(
        "SELECT DISTINCT strftime('%Y', app_date) as year FROM applications ORDER BY year DESC",
        conn
    )
    conn.close()
    years = df["year"].dropna().tolist()
    return years if years else [str(date.today().year)]


def generate_yearly_report_pdf(year, by_area_df, by_product_df):
    """Printable yearly totals report."""
    buffer = BytesIO()
    doc = SimpleDocTemplate(
        buffer, pagesize=letter,
        leftMargin=0.6*inch, rightMargin=0.6*inch,
        topMargin=0.5*inch, bottomMargin=0.5*inch
    )
    styles = getSampleStyleSheet()
    title_style = ParagraphStyle(
        "TitleY", parent=styles["Heading1"],
        fontSize=14, alignment=1, textColor=colors.HexColor("#1a3c2a"), spaceAfter=4
    )
    sub = ParagraphStyle("SubY", parent=styles["Normal"], fontSize=9, alignment=1, spaceAfter=8)
    section = ParagraphStyle(
        "SecY", parent=styles["Normal"], fontSize=10, spaceBefore=10, spaceAfter=4,
        textColor=colors.HexColor("#1a3c2a"), fontName="Helvetica-Bold"
    )
    normal = ParagraphStyle("NormY", parent=styles["Normal"], fontSize=8)
    small = ParagraphStyle("SmallY", parent=styles["Normal"], fontSize=7)

    story = []
    story.append(Paragraph("THE BOOMBOOK YEARLY TOTALS REPORT", title_style))
    story.append(Paragraph("Pinewood Country Club", sub))
    story.append(Paragraph(f"Year: {year}", sub))
    story.append(HRFlowable(width="100%", thickness=1, color=colors.HexColor("#1a3c2a"), spaceAfter=8))

    # Grand totals (product separated by unit)
    if not by_area_df.empty:
        tot_n = by_area_df["n_lbs"].sum()
        tot_p = by_area_df["p_lbs"].sum()
        tot_k = by_area_df["k_lbs"].sum()
        tot_lbs = by_area_df["product_lbs"].sum() if "product_lbs" in by_area_df.columns else 0
        tot_gal = by_area_df["product_gal"].sum() if "product_gal" in by_area_df.columns else 0
        tot_oz = by_area_df["product_fl_oz"].sum() if "product_fl_oz" in by_area_df.columns else 0
        tot_apps = by_area_df["applications"].sum()
        story.append(Paragraph("GRAND TOTALS", section))
        gt = [
            ["Apps", "Product lbs", "Product gal", "Product fl oz", "N (lbs)", "P2O5 (lbs)", "K2O (lbs)"],
            [f"{tot_apps:.0f}", f"{tot_lbs:,.1f}", f"{tot_gal:,.2f}", f"{tot_oz:,.1f}",
             f"{tot_n:,.1f}", f"{tot_p:,.1f}", f"{tot_k:,.1f}"],
        ]
        t = Table(gt, colWidths=[0.9*inch]*7)
        t.setStyle(TableStyle([
            ("FONTNAME", (0, 0), (-1, 0), "Helvetica-Bold"),
            ("FONTSIZE", (0, 0), (-1, -1), 8),
            ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#1a3c2a")),
            ("TEXTCOLOR", (0, 0), (-1, 0), colors.white),
            ("ALIGN", (0, 0), (-1, -1), "CENTER"),
            ("GRID", (0, 0), (-1, -1), 0.5, colors.grey),
            ("BACKGROUND", (0, 1), (-1, 1), colors.HexColor("#f4f9f4")),
            ("TOPPADDING", (0, 0), (-1, -1), 4),
            ("BOTTOMPADDING", (0, 0), (-1, -1), 4),
        ]))
        story.append(t)

    # By Area
    story.append(Paragraph("TOTALS BY AREA", section))
    if by_area_df.empty:
        story.append(Paragraph("No data for this year.", normal))
    else:
        area_rows = [["Area", "Apps", "lbs", "gal", "fl oz", "N lbs", "P2O5", "K2O"]]
        for _, r in by_area_df.iterrows():
            area_rows.append([
                str(r.get("area_name") or "-"),
                f"{r['applications']:.0f}",
                f"{r.get('product_lbs', 0):,.1f}",
                f"{r.get('product_gal', 0):,.2f}",
                f"{r.get('product_fl_oz', 0):,.1f}",
                f"{r['n_lbs']:,.1f}",
                f"{r['p_lbs']:,.1f}",
                f"{r['k_lbs']:,.1f}",
            ])
        t2 = Table(area_rows, colWidths=[1.3*inch, 0.55*inch, 0.7*inch, 0.65*inch, 0.7*inch, 0.75*inch, 0.75*inch, 0.75*inch])
        t2.setStyle(TableStyle([
            ("FONTNAME", (0, 0), (-1, 0), "Helvetica-Bold"),
            ("FONTSIZE", (0, 0), (-1, -1), 8),
            ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#e8f0e8")),
            ("GRID", (0, 0), (-1, -1), 0.4, colors.grey),
            ("ALIGN", (1, 0), (-1, -1), "CENTER"),
            ("TOPPADDING", (0, 0), (-1, -1), 3),
            ("BOTTOMPADDING", (0, 0), (-1, -1), 3),
        ]))
        story.append(t2)

    # By Product
    story.append(Paragraph("TOTALS BY PRODUCT", section))
    if by_product_df.empty:
        story.append(Paragraph("No product data for this year.", normal))
    else:
        prod_rows = [["Brand / Product", "Times", "lbs", "gal", "fl oz", "N lbs", "P2O5", "K2O"]]
        for _, r in by_product_df.iterrows():
            name = f"{r.get('brand') or ''} {r.get('fertilizer_name') or ''}".strip()
            prod_rows.append([
                name[:36],
                f"{r['times_used']:.0f}",
                f"{r.get('product_lbs', 0):,.1f}",
                f"{r.get('product_gal', 0):,.2f}",
                f"{r.get('product_fl_oz', 0):,.1f}",
                f"{r['n_lbs']:,.1f}",
                f"{r['p_lbs']:,.1f}",
                f"{r['k_lbs']:,.1f}",
            ])
        t3 = Table(prod_rows, colWidths=[1.5*inch, 0.5*inch, 0.65*inch, 0.6*inch, 0.65*inch, 0.7*inch, 0.65*inch, 0.65*inch])
        t3.setStyle(TableStyle([
            ("FONTNAME", (0, 0), (-1, 0), "Helvetica-Bold"),
            ("FONTSIZE", (0, 0), (-1, -1), 7),
            ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#e8f0e8")),
            ("GRID", (0, 0), (-1, -1), 0.4, colors.grey),
            ("ALIGN", (1, 0), (-1, -1), "CENTER"),
            ("TOPPADDING", (0, 0), (-1, -1), 2),
            ("BOTTOMPADDING", (0, 0), (-1, -1), 2),
        ]))
        story.append(t3)

    story.append(Spacer(1, 16))
    story.append(Paragraph(
        "Generated " + datetime.now().strftime("%Y-%m-%d %H:%M") + "  |  The BoomBook",
        ParagraphStyle("FootY", parent=small, textColor=colors.grey, alignment=1)
    ))
    story.append(Paragraph("Signature: ________________________  Date: ________", normal))

    doc.build(story)
    buffer.seek(0)
    return buffer.getvalue()


# ---------------------------------------------------------------------------
# UI
# ---------------------------------------------------------------------------
# App logo (place icon_boom_spray.png next to this script)
_LOGO_CANDIDATES = [
    Path(__file__).parent / "icon_boom_spray_new.ico",
    Path(__file__).parent / "icon_boom_spray.ico",
    Path(__file__).parent / "icon_boom_spray.png",
    Path(__file__).parent / "icon_grass_drop.png",
    Path(__file__).parent / "icon_combined.png",
    Path(__file__).parent / "icon_spray_nozzle.png",
    Path(__file__).parent / "icon.png",
]
_APP_LOGO = next((p for p in _LOGO_CANDIDATES if p.exists()), None)

st.set_page_config(
    page_title="The BoomBook — Turf Spray Log",
    page_icon=str(_APP_LOGO) if _APP_LOGO else "seedling",
    layout="wide",
)
init_db()
if "backup_checked" not in st.session_state:
    try:
        st.session_state.last_backup = str(automatic_database_backup() or "")
    except Exception as _backup_error:
        st.session_state.last_backup = ""
        st.session_state.backup_error = str(_backup_error)
    st.session_state.backup_checked = True

# Header styling
# Caddyshack-style title: script + green fill + light outline
st.markdown("""

<style>
@import url('https://fonts.googleapis.com/css2?family=Lobster&family=Source+Sans+3:wght@400;600;700&display=swap');

h1.app-title {
    font-family: 'Lobster', 'Brush Script MT', cursive !important;
    font-size: 5.0rem !important;
    font-weight: 400 !important;
    font-style: normal !important;
    color: #1f7a3a !important;
    margin-bottom: 0.1rem !important;
    letter-spacing: 0.02em;
    line-height: 1.05 !important;
    /* Movie-logo outline: white edge + gray depth (Caddyshack style, green fill) */
    -webkit-text-stroke: 2px #f8f8f8;
    paint-order: stroke fill;
    text-shadow:
        1px 1px 0 #c8c8c8,
        2px 2px 0 #a0a0a0,
        3px 3px 0 #787878,
        4px 4px 0 #505050;
    text-transform: none;
}
p.app-subtitle {
    font-family: 'Source Sans 3', 'Segoe UI', sans-serif !important;
    font-size: 1.25rem !important;
    font-weight: 600 !important;
    color: #2d5a3d !important;
    margin-top: 0.2rem !important;
    margin-bottom: 0.25rem !important;
}
p.app-tagline {
    font-family: 'Source Sans 3', 'Segoe UI', sans-serif !important;
    font-size: 0.95rem !important;
    color: #5a6b5e !important;
    margin-top: 0 !important;
}
</style>
""", unsafe_allow_html=True)

# Header weather strip (cached in session; refresh on demand)
_settings = load_settings()
if "header_wx" not in st.session_state:
    st.session_state.header_wx = None

def _refresh_header_weather():
    wx = fetch_nws_weather(
        float(_settings.get("lat", DEFAULT_LAT)),
        float(_settings.get("lon", DEFAULT_LON)),
    )
    if not wx.get("error") or wx.get("temp_f") is not None:
        st.session_state.header_wx = wx
        # Keep calculator fields in sync when header loads
        if "wx" not in st.session_state:
            st.session_state.wx = {
                "temp_f": 70.0, "humidity": 50.0, "wind_mph": 5.0,
                "conditions": "Clear", "soil_temp": 65.0,
            }
        if wx.get("temp_f") is not None:
            st.session_state.wx["temp_f"] = float(wx["temp_f"])
            st.session_state.wx["soil_temp"] = float(wx["temp_f"]) - 5.0
        if wx.get("humidity") is not None:
            st.session_state.wx["humidity"] = float(wx["humidity"])
        if wx.get("wind_mph") is not None:
            st.session_state.wx["wind_mph"] = float(wx["wind_mph"])
        if wx.get("conditions"):
            st.session_state.wx["conditions"] = str(wx["conditions"])
    return wx

# Auto-load once per browser session (retry allowed via button)
if st.session_state.header_wx is None:
    try:
        _refresh_header_weather()
    except Exception as _e:
        st.session_state.header_wx = {"error": str(_e), "temp_f": None}

_hwx = st.session_state.header_wx or {}

# Header: title | weather | logo
_h1, _h2, _h3 = st.columns([4.2, 2.5, 1.8], gap="medium")
with _h1:
    st.markdown("""
<h1 class="app-title">The BoomBook</h1>
<div style="
    font-family: 'Brush Script MT', 'Segoe Script', 'Lucida Handwriting', cursive;
    font-size: 1.15rem;
    font-style: italic;
    color: #1f6f3d;
    margin-top: -10px;
    margin-left: 12px;
    margin-bottom: 12px;
">by Ryan Duffy</div>
""", unsafe_allow_html=True)
    st.markdown('<p class="app-subtitle">Pinewood Country Club</p>', unsafe_allow_html=True)
    st.markdown('<p class="app-tagline">Spray rates · N-P-K logs · printable boom sheets</p>', unsafe_allow_html=True)
with _h2:
    loc = _settings.get("location_name") or "Munds Park"
    if _hwx.get("temp_f") is not None:
        cond = _hwx.get("conditions") or "—"
        temp = _hwx.get("temp_f")
        hum = _hwx.get("humidity")
        wind = _hwx.get("wind_mph")
        updated = _hwx.get("updated") or ""
        station = _hwx.get("station") or "NWS"
        hum_s = f"{hum:.0f}%" if hum is not None else "—"
        wind_s = f"{wind:.0f} mph" if wind is not None else "—"
        st.markdown(
            f"""
            <div style="
                background: linear-gradient(135deg, #e8f5ec 0%, #f4faf6 100%);
                border: 1px solid #b7d4c0;
                border-radius: 12px;
                padding: 7px 11px;
                margin-top: 14px;
            ">
              <div style="font-size:0.75rem;color:#5a6b5e;font-weight:600;letter-spacing:0.04em;">
                LIVE WEATHER · {loc}
              </div>
              <div style="font-size:1.65rem;font-weight:700;color:#1a3c2a;line-height:1.2;">
                {temp:.0f}°F
                <span style="font-size:1rem;font-weight:600;color:#2d5a3d;">&nbsp;{cond}</span>
              </div>
              <div style="font-size:0.9rem;color:#2d5a3d;margin-top:2px;">
                Wind {wind_s} &nbsp;·&nbsp; Humidity {hum_s}
              </div>
              <div style="font-size:0.7rem;color:#7a8b7e;margin-top:4px;">
                {station}{(' · ' + updated) if updated else ''}
              </div>
            </div>
            """,
            unsafe_allow_html=True,
        )
        if st.button("Refresh weather", key="hdr_wx_refresh"):
            st.session_state.header_wx = None
            _refresh_header_weather()
            st.rerun()
    else:
        err = _hwx.get("error") or "Could not reach NWS"
        st.warning(f"Weather unavailable — {err}")
        if st.button("Try load weather", key="hdr_wx_try"):
            st.session_state.header_wx = None  # force retry
            try:
                _refresh_header_weather()
            except Exception as _e:
                st.session_state.header_wx = {"error": str(_e), "temp_f": None}
            st.rerun()
with _h3:
    if _APP_LOGO:
        st.markdown("<div style='height:6px'></div>", unsafe_allow_html=True)
        st.image(str(_APP_LOGO), width=320)
    else:
        st.caption("Place icon_boom_spray.ico next to turf_fertilizer_app.py")


def get_application_export_detail():
    conn = get_conn()
    df = pd.read_sql_query('''
        SELECT a.id AS application_id, a.app_date, a.area_name, a.hole_number, a.area_sqft,
               a.weather, a.temp_f, a.soil_temp_f, a.wind_mph, a.humidity_pct,
               a.equipment, a.gpa, a.nozzle, a.gear, a.speed, a.pressure,
               a.tanks, a.gallons_per_tank, a.bags, a.rinsed, a.watered_in,
               a.watered_notes, a.notes, a.applicators,
               p.line_num, p.brand, p.fertilizer_name, p.n_pct, p.p_pct, p.k_pct,
               p.target_n_rate, p.product_per_1000, p.product_lbs, p.product_unit,
               p.n_applied_lbs, p.p_applied_lbs, p.k_applied_lbs, p.lot_number, p.minors
        FROM applications a
        LEFT JOIN application_products p ON p.application_id=a.id
        ORDER BY a.app_date DESC, a.id DESC, p.line_num
    ''', conn)
    conn.close()
    return df

def _xlsx_col_name(n):
    s=''
    while n:
        n,r=divmod(n-1,26); s=chr(65+r)+s
    return s

def dataframe_to_xlsx_bytes(sheets):
    out=BytesIO()
    def cell(row,col,value):
        ref=f'{_xlsx_col_name(col)}{row}'
        if value is None or (isinstance(value,float) and pd.isna(value)):
            return f'<c r="{ref}" t="inlineStr"><is><t></t></is></c>'
        if isinstance(value,(int,float)) and not isinstance(value,bool):
            return f'<c r="{ref}"><v>{value}</v></c>'
        return f'<c r="{ref}" t="inlineStr"><is><t>{_xml_escape(str(value))}</t></is></c>'
    clean=[(re.sub(r'[:\\\\/?*\\[\\]]','_',str(n))[:31] or 'Sheet',d.copy()) for n,d in sheets.items()]
    with _xlsx_zipfile.ZipFile(out,'w',_xlsx_zipfile.ZIP_DEFLATED) as z:
        overrides=''.join(f'<Override PartName="/xl/worksheets/sheet{i}.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>' for i in range(1,len(clean)+1))
        z.writestr('[Content_Types].xml','<?xml version="1.0" encoding="UTF-8"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>'+overrides+'<Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/></Types>')
        z.writestr('_rels/.rels','<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/></Relationships>')
        sx=''.join(f'<sheet name="{_xml_escape(n)}" sheetId="{i}" r:id="rId{i}"/>' for i,(n,d) in enumerate(clean,1))
        z.writestr('xl/workbook.xml','<?xml version="1.0" encoding="UTF-8"?><workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets>'+sx+'</sheets></workbook>')
        rels=''.join(f'<Relationship Id="rId{i}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet{i}.xml"/>' for i in range(1,len(clean)+1))
        z.writestr('xl/_rels/workbook.xml.rels','<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'+rels+f'<Relationship Id="rId{len(clean)+1}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/></Relationships>')
        z.writestr('xl/styles.xml','<?xml version="1.0" encoding="UTF-8"?><styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><fonts count="1"><font><sz val="11"/><name val="Calibri"/></font></fonts><fills count="1"><fill><patternFill patternType="none"/></fill></fills><borders count="1"><border/></borders><cellStyleXfs count="1"><xf/></cellStyleXfs><cellXfs count="1"><xf xfId="0"/></cellXfs></styleSheet>')
        for si,(name,df) in enumerate(clean,1):
            rows=['<row r="1">'+''.join(cell(1,j,c) for j,c in enumerate(df.columns,1))+'</row>']
            for ri,vals in enumerate(df.itertuples(index=False,name=None),2):
                rows.append(f'<row r="{ri}">'+''.join(cell(ri,j,v) for j,v in enumerate(vals,1))+'</row>')
            z.writestr(f'xl/worksheets/sheet{si}.xml','<?xml version="1.0" encoding="UTF-8"?><worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData>'+''.join(rows)+'</sheetData></worksheet>')
    return out.getvalue()

def render_month_calendar(apps_df, year, month):
    weeks=calendar.Calendar(firstweekday=6).monthdayscalendar(year,month)
    by_day={}
    if not apps_df.empty:
        tmp=apps_df.copy(); tmp['_dt']=pd.to_datetime(tmp['app_date'],errors='coerce')
        tmp=tmp[(tmp['_dt'].dt.year==year)&(tmp['_dt'].dt.month==month)]
        for day,grp in tmp.groupby(tmp['_dt'].dt.day): by_day[int(day)]=grp
    h=['<style>.bbcal{width:100%;border-collapse:separate;border-spacing:5px;table-layout:fixed}.bbcal th{padding:7px;color:#355744}.bbcal td{vertical-align:top;height:105px;border:1px solid #d9e4dc;border-radius:9px;padding:7px;background:#fff}.bbcal td.empty{background:#f7f8f7}.bbday{font-weight:700;color:#173f2a}.bbevent{font-size:.76rem;background:#edf6f0;border-radius:5px;padding:3px 5px;margin:3px 0}</style><table class="bbcal"><tr>']
    h += [f'<th>{x}</th>' for x in ['Sun','Mon','Tue','Wed','Thu','Fri','Sat']]; h.append('</tr>')
    for week in weeks:
        h.append('<tr>')
        for day in week:
            if not day: h.append('<td class="empty"></td>'); continue
            h.append(f'<td><div class="bbday">{day}</div>')
            grp=by_day.get(day)
            if grp is not None:
                for _,r in grp.head(3).iterrows():
                    hole=f" · H{r['hole_number']}" if pd.notna(r.get('hole_number')) and str(r.get('hole_number')).strip() else ''
                    h.append(f'<div class="bbevent">#{int(r["id"])} · {_xml_escape(str(r["area_name"]))}{_xml_escape(hole)} · {int(r["num_products"])} product(s)</div>')
                if len(grp)>3: h.append(f'<small>+{len(grp)-3} more</small>')
            h.append('</td>')
        h.append('</tr>')
    h.append('</table>')
    return ''.join(h)


# ---------- BoomBook Update Center ----------
UPDATE_MANIFEST_URL = "https://raw.githubusercontent.com/sukaphree7/The-BoomBook/main/update.json"

def _version_tuple(v):
    try:
        return tuple(int(x) for x in str(v).strip().lstrip("vV").split("."))
    except Exception:
        return (0,)

def check_for_boombook_update():
    result = {
        "current_version": BOOMBOOK_VERSION, "latest_version": BOOMBOOK_VERSION,
        "update_available": False, "download_url": "", "source_url": "",
        "sha256": "", "notes": "", "checked_at": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
        "configured": bool(UPDATE_MANIFEST_URL.strip()), "error": "",
    }
    if not result["configured"]:
        return result
    try:
        import urllib.request, json
        req = urllib.request.Request(UPDATE_MANIFEST_URL, headers={"User-Agent": f"The-BoomBook/{BOOMBOOK_VERSION}"})
        with urllib.request.urlopen(req, timeout=8) as response:
            manifest = json.loads(response.read().decode("utf-8"))
        latest = str(manifest.get("version", BOOMBOOK_VERSION))
        result.update(
            latest_version=latest,
            download_url=str(manifest.get("download_url", "") or ""),
            source_url=str(manifest.get("source_url", "") or ""),
            sha256=str(manifest.get("sha256", "") or "").lower(),
            notes=str(manifest.get("notes", "") or ""),
            update_available=_version_tuple(latest) > _version_tuple(BOOMBOOK_VERSION),
        )
    except Exception as exc:
        result["error"] = str(exc)
    return result

def install_boombook_update(status):
    import hashlib, os, subprocess, tempfile, urllib.request
    source_url = status.get("source_url", "")
    expected = status.get("sha256", "").lower()
    if not source_url or not expected:
        return False, "This release does not include an automatic-install payload."
    try:
        work = Path(tempfile.gettempdir()) / "BoomBookUpdate"
        work.mkdir(parents=True, exist_ok=True)
        payload = work / "turf_fertilizer_app.py"
        req = urllib.request.Request(source_url, headers={"User-Agent": f"The-BoomBook/{BOOMBOOK_VERSION}"})
        with urllib.request.urlopen(req, timeout=20) as response:
            data = response.read()
        actual = hashlib.sha256(data).hexdigest().lower()
        if actual != expected:
            return False, "Update verification failed. Nothing was installed."
        payload.write_bytes(data)
        helper = work / "install_boombook_update.ps1"
        helper.write_text(r'''$ErrorActionPreference="Stop"
Start-Sleep -Seconds 3
$install=Join-Path $env:LOCALAPPDATA "Programs\The BoomBook"
$exe=Join-Path $install "The BoomBook.exe"
$target=Join-Path $install "_internal\turf_fertilizer_app.py"
$payload=Join-Path $env:TEMP "BoomBookUpdate\turf_fertilizer_app.py"
$db=Join-Path ([Environment]::GetFolderPath("MyDocuments")) "The BoomBook\turf_fertilizer.db"
$backup=Join-Path ([Environment]::GetFolderPath("MyDocuments")) "The BoomBook\Backups"
Get-CimInstance Win32_Process -ErrorAction SilentlyContinue | Where-Object {$_.ExecutablePath -and $_.ExecutablePath.StartsWith($install,[StringComparison]::OrdinalIgnoreCase)} | ForEach-Object {Stop-Process -Id $_.ProcessId -Force -ErrorAction SilentlyContinue}
Start-Sleep -Seconds 2
if(Test-Path $db){New-Item -ItemType Directory -Path $backup -Force|Out-Null;$stamp=Get-Date -Format "yyyy-MM-dd_HHmmss";Copy-Item $db (Join-Path $backup "turf_fertilizer_BEFORE_AUTO_UPDATE_$stamp.db") -Force}
Copy-Item $payload $target -Force
Start-Process $exe -WorkingDirectory $install
''', encoding="utf-8-sig")
        flags = 0
        if os.name == "nt":
            flags = getattr(subprocess, "CREATE_NEW_PROCESS_GROUP", 0) | getattr(subprocess, "DETACHED_PROCESS", 0)
        startupinfo = None
        if os.name == "nt":
            startupinfo = subprocess.STARTUPINFO()
            startupinfo.dwFlags |= subprocess.STARTF_USESHOWWINDOW
            startupinfo.wShowWindow = 0
        subprocess.Popen(
            ["powershell.exe", "-WindowStyle", "Hidden", "-NoLogo", "-NoProfile",
             "-ExecutionPolicy", "Bypass", "-File", str(helper)],
            creationflags=flags, close_fds=True, startupinfo=startupinfo
        )
        return True, "Update downloaded and verified. BoomBook will restart automatically."
    except Exception as exc:
        return False, str(exc)


# Session state for multi-product cart / edit mode
if "cart" not in st.session_state:
    st.session_state.cart = []
if "editing_id" not in st.session_state:
    st.session_state.editing_id = None
if "prefill" not in st.session_state:
    st.session_state.prefill = None

tab_dash, tab_calc, tab_log, tab_calendar, tab_exports, tab_programs, tab_yearly, tab_products, tab_areas, tab_update, tab_help = st.tabs(
    ["Dashboard", "Calculator", "Application Log & Sheets", "Calendar", "Exports", "Saved Spray Programs", "Yearly Totals", "Products", "Saved Areas", "Update", "Help"]
)

# ========== DASHBOARD ==========
with tab_dash:
    st.subheader("Season Dashboard")
    all_apps, month_apps, year_apps = get_dashboard_data()
    d1,d2,d3,d4 = st.columns(4)
    d1.metric("Applications this month", len(month_apps))
    d2.metric("Applications YTD", len(year_apps))
    d3.metric("N applied YTD", f"{year_apps['total_n'].sum():.1f} lbs" if not year_apps.empty else "0.0 lbs")
    d4.metric("K₂O applied YTD", f"{year_apps['total_k'].sum():.1f} lbs" if not year_apps.empty else "0.0 lbs")
    m1,m2,m3 = st.columns(3)
    m1.metric("N this month", f"{month_apps['total_n'].sum():.1f} lbs" if not month_apps.empty else "0.0 lbs")
    m2.metric("P₂O₅ this month", f"{month_apps['total_p'].sum():.1f} lbs" if not month_apps.empty else "0.0 lbs")
    m3.metric("K₂O this month", f"{month_apps['total_k'].sum():.1f} lbs" if not month_apps.empty else "0.0 lbs")
    st.markdown("#### Recent applications")
    if all_apps.empty:
        st.info("No applications recorded yet.")
    else:
        recent=all_apps.head(8)[["id","app_date","area_name","hole_number","num_products","total_n","total_p","total_k"]].copy()
        st.dataframe(recent.rename(columns={"id":"ID","app_date":"Date","area_name":"Area","hole_number":"Hole","num_products":"Products","total_n":"N lbs","total_p":"P₂O₅ lbs","total_k":"K₂O lbs"}),use_container_width=True,hide_index=True)
    st.markdown("#### Data protection")
    if st.session_state.get("last_backup"):
        st.success(f"Automatic database backup active · latest: {Path(st.session_state.last_backup).name} · retaining newest 30 backups")
    elif st.session_state.get("backup_error"):
        st.warning(f"Backup warning: {st.session_state.backup_error}")
    st.caption(f"Database: {DB_PATH}")

# ========== CALCULATOR ==========
with tab_calc:
    st.subheader("Build an Application (up to 12 products)")

    products_df = get_products()
    brand_list = ["-- Manual entry --"] + sorted(products_df["brand"].unique().tolist())
    areas_df = get_areas()

    # --- Header section ---
    pf = st.session_state.prefill
    if st.session_state.editing_id and pf:
        st.success(
            f"Loaded **#{st.session_state.editing_id}** from history "
            f"({pf.get('app_date', '')} · {pf.get('area_name', '')}"
            f"{(' · Hole ' + str(pf['hole_number'])) if pf.get('hole_number') else ''}). "
            f"{len(st.session_state.cart)} product(s) in the cart — adjust below, then Update or Save as new."
        )

    st.markdown("#### 1. Application header")
    c1, c2 = st.columns(2)
    with c1:
        area_options = ["-- Custom --"] + (areas_df["name"].tolist() if not areas_df.empty else [])
        selected_area = st.selectbox("Area", area_options)
        if selected_area != "-- Custom --" and not areas_df.empty:
            row = areas_df[areas_df["name"] == selected_area].iloc[0]
            default_sqft = float(row["sqft"])
            area_name = selected_area
        else:
            default_sqft = 5000.0
            area_name = st.text_input("Area name", value="Greens")

        # Hole number 1-18 for in-play areas
        hole_areas = ("greens", "tees", "fairways", "fairway", "rough", "approaches", "approach")
        area_for_hole = (area_name or selected_area or "").lower()
        needs_hole = any(a in area_for_hole for a in hole_areas)
        hole_number = None
        if needs_hole:
            hole_number = st.selectbox(
                "Hole number",
                options=["All / Combined"] + list(range(1, 19)),
                help="Select which hole (1–18), or All / Combined for the whole area.",
                key="hole_select",
            )
            if hole_number == "All / Combined":
                hole_number = None
            else:
                hole_number = int(hole_number)

        unit = st.radio("Unit", ["Square feet", "Acres"], horizontal=True)
        if unit == "Acres":
            area_input = st.number_input("Area (acres)", min_value=0.01, value=0.115, step=0.01, format="%.3f")
            area_sqft = area_input * 43560
            st.caption(f"= {area_sqft:,.0f} sq ft")
        else:
            area_sqft = st.number_input("Area (sq ft)", min_value=1.0, value=default_sqft, step=100.0, format="%.0f")
        log_date = st.date_input("Application date", value=date.today())

    with c2:
        settings = load_settings()
        if "wx" not in st.session_state:
            st.session_state.wx = {
                "temp_f": 70.0,
                "humidity": 50.0,
                "wind_mph": 5.0,
                "conditions": "Clear",
                "soil_temp": 65.0,
            }

        wx_box = st.container()
        loc_col, btn_col = st.columns([3, 2])
        with loc_col:
            st.caption(
                f"NWS location: {settings.get('location_name', 'Munds Park')} "
                f"({settings.get('lat', DEFAULT_LAT):.4f}, {settings.get('lon', DEFAULT_LON):.4f})"
            )
        with btn_col:
            if st.button("Load weather from NWS", type="secondary", use_container_width=True):
                with st.spinner("Fetching NWS conditions..."):
                    wx = fetch_nws_weather(
                        float(settings.get("lat", DEFAULT_LAT)),
                        float(settings.get("lon", DEFAULT_LON)),
                    )
                if wx.get("error") and not wx.get("temp_f"):
                    st.error(f"Weather fetch failed: {wx['error']}")
                else:
                    if wx.get("temp_f") is not None:
                        st.session_state.wx["temp_f"] = float(wx["temp_f"])
                    if wx.get("humidity") is not None:
                        st.session_state.wx["humidity"] = float(wx["humidity"])
                    if wx.get("wind_mph") is not None:
                        st.session_state.wx["wind_mph"] = float(wx["wind_mph"])
                    if wx.get("conditions"):
                        st.session_state.wx["conditions"] = str(wx["conditions"])
                    # rough soil estimate = air - 5F if not measured
                    if wx.get("temp_f") is not None:
                        st.session_state.wx["soil_temp"] = float(wx["temp_f"]) - 5.0
                    st.success(
                        f"Loaded from {wx.get('station', 'NWS')}"
                        + (f" · {wx['updated']}" if wx.get("updated") else "")
                    )
                    st.rerun()

        cond_options = ["Clear", "Partly cloudy", "Overcast", "Light rain", "Windy", "Other"]
        cur_cond = st.session_state.wx.get("conditions") or "Clear"
        # If NWS text not in list, put it as custom first option
        if cur_cond not in cond_options:
            cond_options = [cur_cond] + cond_options
        weather = st.selectbox(
            "Weather / conditions",
            cond_options,
            index=cond_options.index(cur_cond) if cur_cond in cond_options else 0,
        )
        w1, w2 = st.columns(2)
        with w1:
            temp_f = st.number_input(
                "Air Temp (F)", value=float(st.session_state.wx.get("temp_f", 70.0)), step=1.0
            )
            wind_mph = st.number_input(
                "Wind (mph)", value=float(st.session_state.wx.get("wind_mph", 5.0)), step=1.0
            )
        with w2:
            soil_temp = st.number_input(
                "Soil Temp (F)", value=float(st.session_state.wx.get("soil_temp", 65.0)), step=1.0
            )
            humidity = st.number_input(
                "Humidity (%)", value=float(st.session_state.wx.get("humidity", 50.0)), step=5.0
            )

        with st.expander("Change weather location (saved once)"):
            la = st.number_input("Latitude", value=float(settings.get("lat", DEFAULT_LAT)), format="%.4f")
            lo = st.number_input("Longitude", value=float(settings.get("lon", DEFAULT_LON)), format="%.4f")
            lname = st.text_input("Location name", value=settings.get("location_name", "Munds Park, AZ"))
            if st.button("Save location"):
                save_settings({"lat": la, "lon": lo, "location_name": lname})
                st.success("Location saved. Click Load weather from NWS.")
                st.rerun()
            st.caption(
                "Default is Munds Park / Pinewood (34.9457, -111.6357) — "
                "same as forecast.weather.gov MapClick for your course."
            )

    with st.expander("Equipment / Calibration (optional)"):
        e1, e2, e3 = st.columns(3)
        with e1:
            equipment = st.selectbox("Equipment Used", EQUIPMENT_OPTIONS, key="equip_select")
            if equipment == "Other":
                equipment = st.text_input("Equipment (specify)", value="", key="equip_other")
            gear = st.selectbox(
                "Gear",
                ["", "N/A", "Auto", "1", "2", "3", "4", "5", "6", "Other"],
                key="gear_select",
            )
            if gear == "Other":
                gear = st.text_input("Gear (specify)", value="", key="gear_other")
            tanks = st.number_input("Total # of Tanks", min_value=0.0, value=0.0, step=1.0)
        with e2:
            gpa = st.number_input("GPA", min_value=0.0, value=0.0, step=0.5)
            speed = st.selectbox(
                "Speed",
                ["", "N/A", "Auto"] + [f"{i} mph" for i in range(2, 16)] + ["Other"],
                key="speed_select",
            )
            if speed == "Other":
                speed = st.text_input("Speed (specify)", value="", key="speed_other")
            gallons_tank = st.number_input("Gallons/Tank", min_value=0.0, value=0.0, step=1.0)
        with e3:
            nozzle = st.selectbox(
                "Nozzle Type (TeeJet)",
                TEEJET_NOZZLES,
                key="nozzle_select",
                help="Common TeeJet tip families from the TeeJet spray tip catalog.",
            )
            if nozzle == "Other / custom":
                nozzle = st.text_input("Nozzle (specify)", value="", key="nozzle_other")
            pressure = st.selectbox(
                "Pressure / Setting",
                ["", "N/A", "Auto"] + [f"{i} psi" for i in range(20, 105, 5)] + ["Other"],
                key="pressure_select",
            )
            if pressure == "Other":
                pressure = st.text_input("Pressure / Setting (specify)", value="", key="pressure_other")
            bags = st.number_input("Bags", min_value=0.0, value=0.0, step=1.0)
        rinsed = st.selectbox("Triple rinsed?", ["", "Yes", "No"])
        watered_in = st.selectbox("Watered in?", ["", "Yes", "No"])
        watered_notes = st.text_input("How long & when watered?", value="")

    _pf_notes = (pf.get("notes") or "") if pf else ""
    _pf_app = (pf.get("applicators") or "") if pf else ""
    notes = st.text_area("Observations / Comments", value=_pf_notes, height=60)
    applicators = st.text_input("Applicators", value=_pf_app)

    st.divider()
    st.markdown("#### 2. Add products to this application")

    # Show current cart with per-line delete
    if st.session_state.cart:
        st.markdown(f"**Products in this application: {len(st.session_state.cart)} / 12**")
        for idx, item in enumerate(st.session_state.cart):
            name = item.get("fertilizer_name") or "Product"
            brand = item.get("brand") or ""
            unit = item.get("product_unit") or "lbs"
            amt = item.get("product_lbs")
            try:
                amt_s = f"{float(amt):g}" if amt is not None else "—"
            except (TypeError, ValueError):
                amt_s = str(amt) if amt is not None else "—"
            n = item.get("n_applied_lbs")
            p = item.get("p_applied_lbs")
            k = item.get("k_applied_lbs")
            try:
                npk = f"N {float(n):g} · P {float(p):g} · K {float(k):g}" if n is not None else ""
            except (TypeError, ValueError):
                npk = ""
            label = f"**{idx + 1}.** {brand} — {name} · {amt_s} {unit}"
            if npk:
                label += f" · {npk}"
            c_left, c_right = st.columns([6, 1])
            with c_left:
                st.markdown(label)
            with c_right:
                if st.button("Delete", key=f"del_cart_{idx}", use_container_width=True):
                    st.session_state.cart.pop(idx)
                    st.rerun()
        if st.button("Clear all products from this application"):
            st.session_state.cart = []
            st.rerun()
    else:
        st.info("No products added yet. Add up to 12 below.")

    if len(st.session_state.cart) < 12:
        st.markdown("**Add one product:**")
        # Build searchable product labels: "Brand | Product (N-P-K)"
        product_options = ["-- Manual entry --"]
        product_lookup = {}  # label -> row dict
        for _, row in products_df.iterrows():
            label = f"{row['brand']}  |  {row['name']}  ({row['n_pct']:g}-{row['p_pct']:g}-{row['k_pct']:g})"
            product_options.append(label)
            product_lookup[label] = row

        prod_choice = st.selectbox(
            "Search / select product",
            product_options,
            key="calc_product_search",
            help="Type to search by brand or product name. Brand fills in automatically."
        )

        if prod_choice != "-- Manual entry --":
            prow = product_lookup[prod_choice]
            n_pct = float(prow["n_pct"])
            p_pct = float(prow["p_pct"])
            k_pct = float(prow["k_pct"])
            fert_name = prow["name"]
            brand = prow["brand"]
            db_minors = str(prow.get("minors", "") or "") if "minors" in prow.index or hasattr(prow, "get") else ""
            try:
                db_minors = str(prow["minors"]) if pd.notna(prow.get("minors")) else ""
            except Exception:
                db_minors = ""
            st.text_input("Brand (auto-filled)", value=brand, disabled=True, key="auto_brand_display")
            st.caption(
                f"{n_pct:g}-{p_pct:g}-{k_pct:g}  |  {prow.get('type', '') or ''}  |  {prow.get('notes', '') or ''}"
            )
            if db_minors:
                st.caption(f"Minors: {db_minors}")
            is_catalog = True
        else:
            brand = st.text_input("Brand", value="", key="man_brand")
            fert_name = st.text_input("Product name", value="", key="man_name")
            c1, c2, c3 = st.columns(3)
            with c1:
                n_pct = st.number_input("N%", value=0.0, min_value=0.0, max_value=100.0, step=0.1, key="man_n")
            with c2:
                p_pct = st.number_input("P%", value=0.0, min_value=0.0, max_value=100.0, step=0.1, key="man_p")
            with c3:
                k_pct = st.number_input("K%", value=0.0, min_value=0.0, max_value=100.0, step=0.1, key="man_k")
            db_minors = ""
            is_catalog = False
            prow = None

        # Unit selection: lbs (granular), gal or fl oz (liquids)
        ptype_str = ""
        try:
            if is_catalog and prow is not None:
                ptype_str = str(prow.get("type", "") or "")
        except Exception:
            ptype_str = ""
        is_liquid_default = any(x in ptype_str.lower() for x in ["liquid", "foliar", "pgr", "fungicide"])
        default_unit = "fl oz" if is_liquid_default else "lbs"
        unit_choice = st.selectbox(
            "Product unit",
            ["lbs", "gal", "fl oz"],
            index=["lbs", "gal", "fl oz"].index(default_unit),
            key="product_unit_choice",
            help="Granular/soluble = lbs. Liquids = gal or fl oz."
        )

        if unit_choice == "lbs":
            rate_hint = "lb product / 1,000 sq ft"
            total_hint = "Total product (lbs)"
        elif unit_choice == "gal":
            rate_hint = "fl oz product / 1,000 sq ft"
            total_hint = "Total product (gallons)"
        else:  # fl oz
            rate_hint = "fl oz product / 1,000 sq ft"
            total_hint = "Total product (fl oz)"

        rate_mode = st.radio(
            "How to set rate",
            [
                "Target N (lb N / 1,000 sq ft)",
                f"Direct product rate ({rate_hint})",
                total_hint,
            ],
            horizontal=True,
            key="rate_mode"
        )

        if rate_mode.startswith("Target N"):
            target_n = st.number_input("Target N (lb / 1,000 sq ft)", min_value=0.0, value=0.0, step=0.05, key="tn")
            if n_pct > 0:
                results = calculate_rates(area_sqft, target_n, n_pct, p_pct, k_pct)
                results["product_unit"] = "lbs"
            else:
                results = None
                st.warning("This product has 0% N. Use Direct product rate or Total amount instead.")
        elif rate_mode.startswith("Direct product rate"):
            if unit_choice == "lbs":
                rate_1000 = st.number_input("lb product / 1,000 sq ft", min_value=0.0, value=3.0, step=0.1, key="r1000")
                total_prod = rate_1000 * (area_sqft / 1000.0)
                results = {
                    "product_per_1000": rate_1000,
                    "total_product": round(total_prod, 2),
                    "product_unit": "lbs",
                    "n_applied": round(total_prod * n_pct / 100.0, 3),
                    "p_applied": round(total_prod * p_pct / 100.0, 3),
                    "k_applied": round(total_prod * k_pct / 100.0, 3),
                    "n_per_1000": round(rate_1000 * n_pct / 100.0, 3),
                }
            else:
                # rate always in fl oz / 1k for liquids
                rate_1000 = st.number_input("fl oz product / 1,000 sq ft", min_value=0.0, value=3.0, step=0.1, key="r1000")
                total_fl_oz = rate_1000 * (area_sqft / 1000.0)
                if unit_choice == "gal":
                    total_prod = round(total_fl_oz / 128.0, 3)
                    unit_out = "gal"
                else:
                    total_prod = round(total_fl_oz, 1)
                    unit_out = "fl oz"
                # approx NPK lbs using density ~8.34 lb/gal
                gal_equiv = total_fl_oz / 128.0
                results = {
                    "product_per_1000": rate_1000,
                    "total_product": total_prod,
                    "product_unit": unit_out,
                    "n_applied": round(gal_equiv * 8.34 * (n_pct / 100.0), 3) if n_pct else 0,
                    "p_applied": round(gal_equiv * 8.34 * (p_pct / 100.0), 3) if p_pct else 0,
                    "k_applied": round(gal_equiv * 8.34 * (k_pct / 100.0), 3) if k_pct else 0,
                    "n_per_1000": round((rate_1000 / 128.0) * 8.34 * (n_pct / 100.0), 3) if n_pct else 0,
                }
            target_n = results.get("n_per_1000", 0)
        else:
            # Total product mode
            if unit_choice == "lbs":
                total_prod = st.number_input("Total product (lbs)", min_value=0.0, value=50.0, step=1.0, key="tprod")
                rate_1000 = total_prod / (area_sqft / 1000.0) if area_sqft else 0
                results = {
                    "product_per_1000": round(rate_1000, 3),
                    "total_product": round(total_prod, 2),
                    "product_unit": "lbs",
                    "n_applied": round(total_prod * n_pct / 100.0, 3),
                    "p_applied": round(total_prod * p_pct / 100.0, 3),
                    "k_applied": round(total_prod * k_pct / 100.0, 3),
                    "n_per_1000": round(rate_1000 * n_pct / 100.0, 3),
                }
            elif unit_choice == "gal":
                total_prod = st.number_input("Total product (gallons)", min_value=0.0, value=5.0, step=0.1, key="tprod")
                rate_1000 = (total_prod * 128.0) / (area_sqft / 1000.0) if area_sqft else 0
                results = {
                    "product_per_1000": round(rate_1000, 3),
                    "total_product": round(total_prod, 3),
                    "product_unit": "gal",
                    "n_applied": round(total_prod * 8.34 * (n_pct / 100.0), 3) if n_pct else 0,
                    "p_applied": round(total_prod * 8.34 * (p_pct / 100.0), 3) if p_pct else 0,
                    "k_applied": round(total_prod * 8.34 * (k_pct / 100.0), 3) if k_pct else 0,
                    "n_per_1000": round((rate_1000 / 128.0) * 8.34 * (n_pct / 100.0), 3) if n_pct else 0,
                }
            else:  # fl oz
                total_prod = st.number_input("Total product (fl oz)", min_value=0.0, value=128.0, step=1.0, key="tprod")
                rate_1000 = total_prod / (area_sqft / 1000.0) if area_sqft else 0
                gal_equiv = total_prod / 128.0
                results = {
                    "product_per_1000": round(rate_1000, 3),
                    "total_product": round(total_prod, 1),
                    "product_unit": "fl oz",
                    "n_applied": round(gal_equiv * 8.34 * (n_pct / 100.0), 3) if n_pct else 0,
                    "p_applied": round(gal_equiv * 8.34 * (p_pct / 100.0), 3) if p_pct else 0,
                    "k_applied": round(gal_equiv * 8.34 * (k_pct / 100.0), 3) if k_pct else 0,
                    "n_per_1000": round((rate_1000 / 128.0) * 8.34 * (n_pct / 100.0), 3) if n_pct else 0,
                }
            target_n = results.get("n_per_1000", 0)

        # Packaging: Bags (50 lb) or Jugs (2.5 gal) — replaces Lot #
        BAG_LB = 50.0
        JUG_GAL = 2.5
        qty_choices = [
            ("", None),
            ("1/4", 0.25),
            ("1/2", 0.5),
            ("3/4", 0.75),
            ("1", 1.0),
            ("1.5", 1.5),
            ("2", 2.0),
            ("2.5", 2.5),
            ("3", 3.0),
            ("4", 4.0),
            ("5", 5.0),
            ("6", 6.0),
            ("8", 8.0),
            ("10", 10.0),
            ("12", 12.0),
            ("15", 15.0),
            ("20", 20.0),
        ]
        pk1, pk2 = st.columns(2)
        with pk1:
            pack_type = st.selectbox(
                "Packaging",
                ["— not using bags/jugs —", "Bags (50 lb each)", "Jugs (2.5 gal each)"],
                key="pack_type",
                help="Bags auto-fill total lbs (50 lb/bag). Jugs auto-fill total gallons (2.5 gal/jug).",
            )
        with pk2:
            qty_label = st.selectbox(
                "Quantity",
                [q[0] for q in qty_choices],
                key="pack_qty",
                disabled=(pack_type.startswith("—")),
            )
        qty_val = dict(qty_choices).get(qty_label)
        pack_note = ""
        if pack_type.startswith("Bags") and qty_val is not None:
            pack_total = round(qty_val * BAG_LB, 2)
            pack_unit = "lbs"
            pack_note = f"{qty_label} bag(s) × 50 lb = {pack_total:g} lbs"
            # Override results totals from packaging
            if results is None:
                results = {}
            results["total_product"] = pack_total
            results["product_unit"] = "lbs"
            results["product_per_1000"] = round(pack_total / (area_sqft / 1000.0), 3) if area_sqft else 0
            results["n_applied"] = round(pack_total * n_pct / 100.0, 3)
            results["p_applied"] = round(pack_total * p_pct / 100.0, 3)
            results["k_applied"] = round(pack_total * k_pct / 100.0, 3)
            results["n_per_1000"] = round(results["product_per_1000"] * n_pct / 100.0, 3)
            st.caption(pack_note)
        elif pack_type.startswith("Jugs") and qty_val is not None:
            pack_total = round(qty_val * JUG_GAL, 3)
            pack_unit = "gal"
            pack_note = f"{qty_label} jug(s) × 2.5 gal = {pack_total:g} gal"
            if results is None:
                results = {}
            results["total_product"] = pack_total
            results["product_unit"] = "gal"
            # rate as fl oz / 1000 for consistency with liquid rates
            total_fl_oz = pack_total * 128.0
            results["product_per_1000"] = round(total_fl_oz / (area_sqft / 1000.0), 3) if area_sqft else 0
            results["n_applied"] = round(pack_total * 8.34 * (n_pct / 100.0), 3) if n_pct else 0
            results["p_applied"] = round(pack_total * 8.34 * (p_pct / 100.0), 3) if p_pct else 0
            results["k_applied"] = round(pack_total * 8.34 * (k_pct / 100.0), 3) if k_pct else 0
            results["n_per_1000"] = round((results["product_per_1000"] / 128.0) * 8.34 * (n_pct / 100.0), 3) if n_pct else 0
            st.caption(pack_note)
        else:
            pack_note = ""

        # Store packaging description in lot_number field for PDF / log
        lot_number = pack_note if pack_note else ""

        if is_catalog and db_minors:
            minors = st.text_input("Minors", value=db_minors, key="minors_auto",
                                   help="Pre-filled from product database; you can edit for this application.")
        else:
            minors = st.text_input("Minors (optional)", value="", key="minors")

        if results and results.get("total_product") is not None:
            unit = results.get("product_unit", "lbs")
            m1, m2, m3, m4 = st.columns(4)
            m1.metric(f"Total product ({unit})", f"{results['total_product']:.2f}")
            if unit == "lbs":
                rate_lbl = "lb / 1k"
            else:
                rate_lbl = "fl oz / 1k"
            m2.metric(f"Rate ({rate_lbl})", f"{results.get('product_per_1000', 0):.2f}")
            m3.metric("N applied (lbs)", f"{results.get('n_applied', 0):.2f}")
            m4.metric("P / K (lbs)", f"{results.get('p_applied', 0):.2f} / {results.get('k_applied', 0):.2f}")

        if st.button("Add this product to application", type="secondary"):
            if results is None or results.get("total_product") is None:
                st.error("Set a rate or choose Bags/Jugs quantity first.")
            else:
                st.session_state.cart.append({
                    "brand": brand,
                    "fertilizer_name": fert_name or f"{n_pct:g}-{p_pct:g}-{k_pct:g}",
                    "n_pct": n_pct, "p_pct": p_pct, "k_pct": k_pct,
                    "target_n_rate": target_n if rate_mode.startswith("Target") else results.get("n_per_1000"),
                    "product_per_1000": results["product_per_1000"],
                    "product_lbs": results["total_product"],
                    "product_unit": results.get("product_unit", "lbs"),
                    "n_applied_lbs": results["n_applied"],
                    "p_applied_lbs": results["p_applied"],
                    "k_applied_lbs": results["k_applied"],
                    "lot_number": lot_number,
                    "minors": minors,
                })
                st.success(f"Added ({len(st.session_state.cart)}/12)")
                st.rerun()
    else:
        st.warning("Maximum 12 products reached. Remove some or save the application.")

    st.divider()
    st.markdown("#### 3. Save the full application")
    if st.session_state.editing_id:
        st.info(
            f"Editing **application #{st.session_state.editing_id}**. "
            "Products and fields are loaded from history. "
            "Update that record, or save as a new application."
        )
        if st.button("Cancel edit (start blank)", key="cancel_edit"):
            st.session_state.editing_id = None
            st.session_state.prefill = None
            st.session_state.cart = []
            st.rerun()

    save_cols = st.columns(2 if st.session_state.editing_id else 1)
    do_save_new = save_cols[0].button(
        "Save as new application", type="primary", use_container_width=True
    )
    do_update = False
    if st.session_state.editing_id:
        do_update = save_cols[1].button(
            f"Update application #{st.session_state.editing_id}",
            type="secondary",
            use_container_width=True,
        )

    if do_save_new or do_update:
        if not st.session_state.cart:
            st.error("Add at least one product first.")
        else:
            header = {
                "app_date": log_date.isoformat(),
                "area_name": area_name if selected_area == "-- Custom --" else selected_area,
                "hole_number": hole_number,
                "area_sqft": area_sqft,
                "weather": weather,
                "temp_f": temp_f,
                "soil_temp_f": soil_temp,
                "wind_mph": wind_mph,
                "humidity_pct": humidity,
                "equipment": equipment,
                "gpa": gpa if gpa else None,
                "nozzle": nozzle,
                "gear": gear,
                "speed": speed,
                "pressure": pressure,
                "tanks": tanks if tanks else None,
                "gallons_per_tank": gallons_tank if gallons_tank else None,
                "bags": bags if bags else None,
                "rinsed": rinsed,
                "watered_in": watered_in,
                "watered_notes": watered_notes,
                "notes": notes,
                "applicators": applicators,
            }
            if do_update and st.session_state.editing_id:
                app_id = update_application(
                    st.session_state.editing_id, header, st.session_state.cart
                )
                st.session_state.cart = []
                st.session_state.editing_id = None
                st.session_state.prefill = None
                st.success(f"Application #{app_id} updated. Go to Log tab to print.")
            else:
                app_id = save_application(header, st.session_state.cart)
                st.session_state.cart = []
                st.session_state.editing_id = None
                st.session_state.prefill = None
                st.success(f"Application #{app_id} saved with products. Go to Log tab to print.")
            st.rerun()

# ========== LOG ==========
with tab_log:
    st.subheader("Application History & Printable Sheets")
    df = get_applications_summary()

    if df.empty:
        st.info("No applications yet. Build one in the Calculator tab.")
    else:
        s1, s2, s3, s4 = st.columns(4)
        s1.metric("Applications", len(df))
        s2.metric("Total N", f"{df['total_n'].sum():.1f} lbs")
        s3.metric("Total P2O5", f"{df['total_p'].sum():.1f} lbs")
        s4.metric("Total K2O", f"{df['total_k'].sum():.1f} lbs")

        st.dataframe(
            df.rename(columns={
                "app_date": "Date", "area_name": "Area", "hole_number": "Hole", "area_sqft": "Sq Ft",
                "num_products": "# Products", "total_product_lbs": "Product lbs",
                "total_n": "N lbs", "total_p": "P2O5 lbs", "total_k": "K2O lbs",
            }),
            use_container_width=True, hide_index=True,
        )

        st.markdown("#### Load for adjustment")
        load_id = st.number_input(
            "Application ID to load into Calculator",
            min_value=1, step=1,
            value=int(df["id"].iloc[0]),
            key="load_app_id",
        )
        if st.button("Load into Calculator for adjustment", type="primary"):
            data = get_application_full(int(load_id))
            if not data:
                st.error(f"Application #{load_id} not found.")
            else:
                h = data["header"]
                # Map product rows into cart shape
                cart = []
                for p in data["products"]:
                    cart.append({
                        "brand": p.get("brand") or "",
                        "fertilizer_name": p.get("fertilizer_name") or "",
                        "n_pct": p.get("n_pct") or 0,
                        "p_pct": p.get("p_pct") or 0,
                        "k_pct": p.get("k_pct") or 0,
                        "target_n_rate": p.get("target_n_rate") or 0,
                        "product_per_1000": p.get("product_per_1000") or 0,
                        "product_lbs": p.get("product_lbs") or 0,
                        "product_unit": p.get("product_unit") or "lbs",
                        "n_applied_lbs": p.get("n_applied_lbs") or 0,
                        "p_applied_lbs": p.get("p_applied_lbs") or 0,
                        "k_applied_lbs": p.get("k_applied_lbs") or 0,
                        "lot_number": p.get("lot_number") or "",
                        "minors": p.get("minors") or "",
                    })
                st.session_state.cart = cart
                st.session_state.editing_id = int(load_id)
                st.session_state.prefill = h
                # Weather fields for calculator
                st.session_state.wx = {
                    "temp_f": float(h["temp_f"]) if h.get("temp_f") is not None else 70.0,
                    "humidity": float(h["humidity_pct"]) if h.get("humidity_pct") is not None else 50.0,
                    "wind_mph": float(h["wind_mph"]) if h.get("wind_mph") is not None else 5.0,
                    "conditions": h.get("weather") or "Clear",
                    "soil_temp": float(h["soil_temp_f"]) if h.get("soil_temp_f") is not None else 65.0,
                }
                st.success(
                    f"Loaded application #{load_id} — open the **Calculator** tab to adjust, "
                    "then Save as new or Update existing."
                )
                st.rerun()

        st.markdown("#### Duplicate an application")
        dup_id = st.number_input("Application ID to duplicate", min_value=1, step=1, value=int(df["id"].iloc[0]), key="dup_app_id")
        if st.button("Duplicate into Calculator", use_container_width=True):
            data = get_application_full(int(dup_id))
            if not data:
                st.error(f"Application #{dup_id} not found.")
            else:
                st.session_state.cart = application_to_cart(data)
                st.session_state.editing_id = None
                h = dict(data["header"])
                h["app_date"] = date.today().isoformat()
                st.session_state.prefill = h
                st.success(f"Application #{dup_id} duplicated. It is now a NEW unsaved application in Calculator; the original will not be changed.")
                st.rerun()

        st.markdown("#### Printable Application Sheet")
        sheet_id = st.number_input("Application ID to print", min_value=1, step=1,
                                   value=int(df["id"].iloc[0]), key="print_app_id")
        app_data = get_application_full(int(sheet_id))
        if app_data:
            h = app_data["header"]
            prods = app_data["products"]
            st.markdown(
                f"**Preview:** {h.get('app_date')} | {h.get('area_name')}" + (f" · Hole {h['hole_number']}" if h.get("hole_number") else "") + " | "
                f"{len(prods)} product(s) | "
                f"N {sum(p.get('n_applied_lbs') or 0 for p in prods):.2f} / "
                f"P {sum(p.get('p_applied_lbs') or 0 for p in prods):.2f} / "
                f"K {sum(p.get('k_applied_lbs') or 0 for p in prods):.2f} lbs"
            )
            if HAS_REPORTLAB:
                pdf_bytes = generate_application_sheet_pdf(app_data)
                st.download_button(
                    "Download Printable PDF Sheet",
                    data=pdf_bytes,
                    file_name=f"application_sheet_{h['id']}_{h.get('app_date', '')}.pdf",
                    mime="application/pdf",
                    type="primary",
                )
            else:
                st.warning("Install reportlab for PDF: pip install reportlab")
        else:
            st.warning("No application with that ID.")

        st.divider()
        del_id = st.number_input("Delete application ID", min_value=1, step=1, value=1, key="del_app")
        if st.button("Delete selected ID"):
            delete_application(int(del_id))
            st.success(f"Deleted #{del_id}")
            st.rerun()

        csv = df.to_csv(index=False)
        st.download_button("Download summary CSV", data=csv,
                           file_name=f"turf_log_{date.today().isoformat()}.csv", mime="text/csv")



# ========== CALENDAR ==========
with tab_calendar:
    st.subheader("Application Calendar")
    cal_df=get_applications_summary(); today=date.today()
    c1,c2=st.columns(2)
    with c1: cal_year=st.selectbox("Year",list(range(today.year-5,today.year+2)),index=5,key="cal_year")
    with c2:
        month_names=list(calendar.month_name)[1:]
        cal_month_name=st.selectbox("Month",month_names,index=today.month-1,key="cal_month")
        cal_month=month_names.index(cal_month_name)+1
    st.markdown(f"### {calendar.month_name[cal_month]} {cal_year}")
    st.markdown(render_month_calendar(cal_df,int(cal_year),int(cal_month)),unsafe_allow_html=True)
    md=cal_df.copy()
    if not md.empty:
        md['_dt']=pd.to_datetime(md['app_date'],errors='coerce')
        md=md[(md['_dt'].dt.year==int(cal_year))&(md['_dt'].dt.month==int(cal_month))].drop(columns=['_dt'])
    st.markdown("#### Applications this month")
    if md.empty: st.info("No applications recorded for this month.")
    else: st.dataframe(md,use_container_width=True,hide_index=True)

# ========== EXPORTS ==========
with tab_exports:
    st.subheader("Excel / CSV Exports")
    st.caption("Choose a date range and export complete BoomBook application records.")
    summary_export=get_applications_summary(); detail_export=get_application_export_detail()
    e1,e2=st.columns(2)
    with e1: export_start=st.date_input("Start date",value=date(date.today().year,1,1),key="export_start")
    with e2: export_end=st.date_input("End date",value=date.today(),key="export_end")
    if export_start>export_end: st.error("Start date must be on or before end date.")
    else:
        def _ef(df):
            if df.empty:return df.copy()
            d=pd.to_datetime(df['app_date'],errors='coerce').dt.date
            return df[(d>=export_start)&(d<=export_end)].copy()
        s_exp=_ef(summary_export); d_exp=_ef(detail_export)
        a,b=st.columns(2); a.metric("Applications",len(s_exp)); b.metric("Product lines",len(d_exp))
        stamp=f"{export_start.isoformat()}_to_{export_end.isoformat()}"
        wb=dataframe_to_xlsx_bytes({"Application Summary":s_exp,"Product Detail":d_exp})
        st.download_button("Download Excel Workbook (.xlsx)",wb,f"BoomBook_Export_{stamp}.xlsx","application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",type="primary",use_container_width=True)
        c1,c2=st.columns(2)
        with c1: st.download_button("Application Summary CSV",s_exp.to_csv(index=False).encode("utf-8-sig"),f"BoomBook_Application_Summary_{stamp}.csv","text/csv",use_container_width=True)
        with c2: st.download_button("Product Detail CSV",d_exp.to_csv(index=False).encode("utf-8-sig"),f"BoomBook_Product_Detail_{stamp}.csv","text/csv",use_container_width=True)
        st.markdown("#### Export preview")
        if s_exp.empty: st.info("No applications in the selected date range.")
        else: st.dataframe(s_exp.head(50),use_container_width=True,hide_index=True)

# ========== SAVED SPRAY PROGRAMS ==========
with tab_programs:
    st.subheader("Saved Spray Programs")
    st.caption("Save frequently used product mixes, then load them into the Calculator with one click. Rates can still be adjusted before saving an application.")
    programs_df = get_spray_programs()
    left,right = st.columns([1.2,1])
    with left:
        st.markdown("#### Save current Calculator mix")
        program_name = st.text_input("Program name", placeholder="Example: Weekly Greens Foliar")
        program_notes = st.text_input("Program notes", placeholder="Optional")
        st.caption(f"Current Calculator cart: {len(st.session_state.cart)} product(s)")
        if st.button("Save / Update Spray Program", type="primary", use_container_width=True):
            if not program_name.strip(): st.error("Enter a program name.")
            elif not st.session_state.cart: st.error("Add products in Calculator first, then save the mix here.")
            else:
                save_spray_program(program_name, st.session_state.cart, program_notes)
                st.success(f"Saved spray program: {program_name}"); st.rerun()
    with right:
        st.markdown("#### Load a saved program")
        if programs_df.empty:
            st.info("No saved spray programs yet.")
        else:
            names=programs_df["name"].tolist()
            chosen=st.selectbox("Saved program", names)
            row=programs_df[programs_df["name"]==chosen].iloc[0]
            try: mix=json.loads(row["products_json"])
            except Exception: mix=[]
            st.write(f"**{len(mix)} product(s)**")
            if row.get("notes"): st.caption(str(row["notes"]))
            for x in mix:
                st.write(f"• {x.get('brand','')} — {x.get('fertilizer_name','')} · {x.get('product_per_1000',0):g} {x.get('product_unit','lbs')}/1,000")
            if st.button("Load Program into Calculator", type="primary", use_container_width=True):
                st.session_state.cart = mix
                st.session_state.editing_id = None
                st.session_state.prefill = None
                st.success(f"Loaded {chosen}. Open Calculator to adjust rates or save the application."); st.rerun()
            if st.button("Delete Saved Program", use_container_width=True):
                delete_spray_program(int(row["id"])); st.rerun()

# ========== YEARLY TOTALS ==========
with tab_yearly:
    st.subheader("Yearly Product & N-P-K Totals by Area")
    st.caption("Track fertilizer use and nutrient totals for each area over the year. Printable report available.")

    years = get_available_years()
    year_sel = st.selectbox("Year", years, index=0)

    by_area = get_yearly_totals_by_area(year_sel)
    by_product = get_yearly_totals_by_product(year_sel)

    if by_area.empty and by_product.empty:
        st.info(f"No applications recorded for {year_sel}. Log applications in the Calculator tab.")
    else:
        # Summary metrics
        if not by_area.empty:
            m1, m2, m3, m4, m5, m6, m7 = st.columns(7)
            m1.metric("Applications", f"{by_area['applications'].sum():.0f}")
            m2.metric("Product (lbs)", f"{by_area['product_lbs'].sum():,.0f}")
            m3.metric("Product (gal)", f"{by_area.get('product_gal', pd.Series([0])).sum():,.1f}")
            m4.metric("Product (fl oz)", f"{by_area.get('product_fl_oz', pd.Series([0])).sum():,.0f}")
            m5.metric("Total N", f"{by_area['n_lbs'].sum():,.1f} lbs")
            m6.metric("Total P2O5", f"{by_area['p_lbs'].sum():,.1f} lbs")
            m7.metric("Total K2O", f"{by_area['k_lbs'].sum():,.1f} lbs")

        st.markdown("#### Totals by Area")
        if not by_area.empty:
            st.dataframe(
                by_area.rename(columns={
                    "area_name": "Area",
                    "applications": "Apps",
                    "product_lbs": "lbs",
                    "product_gal": "gal",
                    "product_fl_oz": "fl oz",
                    "n_lbs": "N lbs",
                    "p_lbs": "P2O5 lbs",
                    "k_lbs": "K2O lbs",
                }),
                use_container_width=True, hide_index=True,
            )
        else:
            st.write("No area data.")

        st.markdown("#### Totals by Product")
        if not by_product.empty:
            st.dataframe(
                by_product.rename(columns={
                    "brand": "Brand",
                    "fertilizer_name": "Product",
                    "times_used": "Times used",
                    "product_lbs": "lbs",
                    "product_gal": "gal",
                    "product_fl_oz": "fl oz",
                    "n_lbs": "N lbs",
                    "p_lbs": "P2O5 lbs",
                    "k_lbs": "K2O lbs",
                }),
                use_container_width=True, hide_index=True,
            )
        else:
            st.write("No product data.")

        st.divider()
        st.markdown("#### Printable / Export")
        c1, c2 = st.columns(2)
        with c1:
            if HAS_REPORTLAB:
                pdf_bytes = generate_yearly_report_pdf(year_sel, by_area, by_product)
                st.download_button(
                    f"Download {year_sel} Yearly Totals PDF",
                    data=pdf_bytes,
                    file_name=f"yearly_totals_{year_sel}.pdf",
                    mime="application/pdf",
                    type="primary",
                )
            else:
                st.warning("Install reportlab for PDF export.")
        with c2:
            # Combined CSV
            area_csv = by_area.to_csv(index=False) if not by_area.empty else ""
            prod_csv = by_product.to_csv(index=False) if not by_product.empty else ""
            combined = f"=== TOTALS BY AREA ({year_sel}) ===\n{area_csv}\n=== TOTALS BY PRODUCT ({year_sel}) ===\n{prod_csv}"
            st.download_button(
                f"Download {year_sel} Yearly Totals CSV",
                data=combined,
                file_name=f"yearly_totals_{year_sel}.csv",
                mime="text/csv",
            )


# ========== PRODUCTS ==========
with tab_products:
    st.subheader("Product Database")
    st.caption("Search fertilizers and golf/turf pesticide reference records. Always verify the current pesticide label before use.")
    _all_products = get_products()
    _known_pesticide_types = {"Fungicide", "Herbicide", "Insecticide"}
    _all_products["Category"] = _all_products["type"].apply(lambda x: x if x in _known_pesticide_types else "Fertilizer")

    _search = st.text_input(
        "Search products / targets",
        placeholder="Try: dollar spot, ABW, nutsedge, azoxystrobin, FRAC 11, or an EPA number",
        key="product_database_search",
    ).strip()

    f1, f2 = st.columns(2)
    with f1:
        category_filter = st.selectbox("Filter by category", ["All", "Fertilizer", "Fungicide", "Herbicide", "Insecticide"], key="product_category_filter")
    _category_view = _all_products if category_filter == "All" else _all_products[_all_products["Category"] == category_filter]
    with f2:
        _brands = sorted(_category_view["brand"].dropna().unique().tolist())
        brand_filter = st.selectbox("Filter by brand", ["All"] + _brands, key="product_brand_filter")
    pdf = _category_view if brand_filter == "All" else _category_view[_category_view["brand"] == brand_filter]

    if _search:
        _q = _search.lower()
        _cols = ["brand","name","type","notes","epa_reg_no","active_ingredients","resistance_group","targets","minors"]
        _mask = pdf[_cols].fillna("").astype(str).apply(lambda row: row.str.lower().str.contains(_q, regex=False).any(), axis=1)
        pdf = pdf[_mask]

    st.caption(f"{len(pdf)} product(s) shown")
    _display_cols = ["Category","brand","name","active_ingredients","resistance_group","epa_reg_no","targets","n_pct","p_pct","k_pct","minors","notes"]
    st.dataframe(
        pdf[_display_cols].rename(columns={
            "brand":"Brand","name":"Product","active_ingredients":"Active Ingredient(s)",
            "resistance_group":"FRAC / HRAC / IRAC","epa_reg_no":"EPA Reg. #","targets":"Targets",
            "n_pct":"N %","p_pct":"P2O5 %","k_pct":"K2O %","minors":"Minors / Secondary Nutrients","notes":"Label / Notes"
        }),
        use_container_width=True, hide_index=True
    )

    with st.expander("Add a custom product"):
        with st.form("add_custom_product_form"):
            cb = st.text_input("Brand", value="Custom")
            cn = st.text_input("Product name")
            c1, c2, c3 = st.columns(3)
            with c1:
                cn_pct = st.number_input("N%", value=0.0, min_value=0.0, max_value=100.0, step=0.1)
            with c2:
                cp_pct = st.number_input("P%", value=0.0, min_value=0.0, max_value=100.0, step=0.1)
            with c3:
                ck_pct = st.number_input("K%", value=0.0, min_value=0.0, max_value=100.0, step=0.1)
            ctype = st.selectbox("Type", ["Granular", "Liquid Foliar", "Soluble Powder", "Fungicide", "Herbicide", "Insecticide", "Other"])
            cnotes = st.text_input("Notes")
            if st.form_submit_button("Add product"):
                if cn.strip():
                    add_custom_product(cb, cn.strip(), cn_pct, cp_pct, ck_pct, ctype, cnotes)
                    st.success(f"Added {cn}")
                    st.rerun()

# ========== AREAS ==========
with tab_areas:
    st.subheader("Saved Turf Areas")
    with st.form("add_area"):
        a_name = st.text_input("Area name")
        a_sqft = st.number_input("Size (sq ft)", min_value=1.0, value=5000.0, step=100.0)
        a_notes = st.text_input("Notes")
        if st.form_submit_button("Save Area"):
            if a_name.strip():
                save_area(a_name.strip(), a_sqft, a_notes)
                st.success(f"Saved {a_name}")
                st.rerun()
    areas = get_areas()
    if not areas.empty:
        st.dataframe(areas, use_container_width=True, hide_index=True)
        del_area_id = st.number_input("Area ID to delete", min_value=1, step=1, key="del_area")
        if st.button("Delete area"):
            delete_area(int(del_area_id))
            st.success("Deleted")
            st.rerun()


# ========== UPDATE CENTER ==========
with tab_update:
    st.subheader("BoomBook Update Center")
    st.caption("Keep The BoomBook current without replacing your application database.")

    u1, u2 = st.columns(2)
    u1.metric("Installed version", f"v{BOOMBOOK_VERSION}")
    if "bb_update_status" in st.session_state:
        last = st.session_state["bb_update_status"].get("checked_at", "Not checked")
    else:
        last = "Not checked"
    u2.metric("Last checked", last)

    if st.button("Check for Updates", type="primary", use_container_width=True):
        with st.spinner("Checking for BoomBook updates..."):
            st.session_state["bb_update_status"] = check_for_boombook_update()

    status = st.session_state.get("bb_update_status")
    if status:
        if status.get("error"):
            st.error("BoomBook could not check the update server.")
            st.caption(status["error"])
        elif not status.get("configured"):
            st.info(
                "The in-app Update Center is installed and ready. "
                "The update feed has not been connected yet, so this version will continue "
                "to use the downloaded BoomBook update packages."
            )
        elif status.get("update_available"):
            st.success(f"BoomBook v{status['latest_version']} is available.")
            if status.get("notes"):
                st.write(status["notes"])
            if status.get("source_url") and status.get("sha256"):
                if st.button(
                    f"Install BoomBook v{status['latest_version']}",
                    type="primary", use_container_width=True, key="install_boombook_update",
                ):
                    with st.spinner("Downloading and verifying update..."):
                        ok, message = install_boombook_update(status)
                    if ok:
                        st.success(message)
                        st.info("BoomBook will close and reopen in a few seconds. Your database is backed up first.")
                    else:
                        st.error("Automatic update could not start.")
                        st.caption(message)
            elif status.get("download_url"):
                st.link_button(
                    f"Download BoomBook v{status['latest_version']}",
                    status["download_url"],
                    use_container_width=True,
                )
        else:
            st.success(f"You're up to date — BoomBook v{BOOMBOOK_VERSION}.")

    st.markdown("#### Update safety")
    st.write(
        "BoomBook updates are designed to preserve the database in Documents → "
        "The BoomBook. Update packages also make a backup before replacing application files."
    )
    st.caption(
        "Once an update feed is connected, this page can announce new releases automatically "
        "and provide the install package directly inside BoomBook."
    )


# ========== HELP ==========
with tab_help:
    st.subheader("How to use multi-product applications")
    st.markdown("""
### Building an application with multiple fertilizers

1. Set the **header** (date, area, weather, equipment).
2. **Add products** one at a time (up to 10). Each can use:
   - Target N rate (lb N / 1,000 sq ft), or
   - Direct product rate (lb product / 1,000 sq ft), or
   - Total product lbs
3. Click **Save Application** when finished. All products are stored under one Application ID.
4. In the **Log** tab, download a PDF that matches the Pinewood form layout (rows 1-10).

### Products included
Simplot/BEST, Floratine, The Andersons, Lebanon, Scotts Pro, Thrive 46-0-0, Extreme Green 20, and common generics.

### Safety
Do not apply more than about 1.0 lb quick-release N per 1,000 sq ft in a single application.
""")
